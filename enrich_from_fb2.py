import gspread
import os
import json
import sys
import re
import time
import requests
import xml.etree.ElementTree as ET

# ========== НАСТРОЙКИ ==========
SPREADSHEET_ID = "1kcG0TG4GZtSM2mypjgvNDUpIbLfIvcmW80_hBKA11nw"
YANDEX_PUBLIC_FOLDER_URL = "https://disk.yandex.ru/d/zMxF4nXHPkIVCQ"
SHEET_TITLE = "Книги"

# Обновлять все строки или только те, где есть пустые поля
ONLY_EMPTY = False  # True — обновлять только пустые поля, False — перезаписывать все

# Максимальная длина описания
MAX_DESC_LEN = 2000
# ==============================

def get_gspread_client():
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if creds_json:
        try:
            creds_dict = json.loads(creds_json)
            return gspread.service_account_from_dict(creds_dict)
        except Exception as e:
            print(f"Ошибка парсинга GOOGLE_CREDENTIALS_JSON: {e}")
            sys.exit(1)
    else:
        try:
            return gspread.service_account(filename="credentials.json")
        except FileNotFoundError:
            print("Файл credentials.json не найден.")
            sys.exit(1)

def extract_public_key(url):
    m = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
    return m.group(1) if m else None

def get_yandex_files_with_download(public_url):
    """Получает список FB2-файлов с прямой ссылкой на скачивание каждого."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    }
    # Список файлов — передаём полную ссылку (это работает!)
    api_url = "https://cloud-api.yandex.net/v1/disk/public/resources"
    params = {"public_key": public_url, "limit": 1000, "sort": "name"}
    resp = requests.get(api_url, params=params, headers=headers, timeout=30)
    if resp.status_code != 200:
        print(f"Ошибка при получении списка файлов: {resp.status_code} — {resp.text}")
        return []
    items = resp.json().get('_embedded', {}).get('items', [])
    print(f"Найдено {len(items)} элементов на Яндекс.Диске.")

    result = []
    for idx, item in enumerate(items, start=1):
        if item.get('type') != 'file':
            continue
        fname = item.get('name', '')
        if not fname.lower().endswith('.fb2'):
            continue

        # Получаем прямую ссылку на скачивание — передаём полную ссылку и path
        dl_api = "https://cloud-api.yandex.net/v1/disk/public/resources/download"
        dl_params = {"public_key": public_url, "path": item['path']}
        try:
            dl_resp = requests.get(dl_api, params=dl_params, headers=headers, timeout=30)
            if dl_resp.status_code == 200:
                download_url = dl_resp.json().get('href')
                result.append({
                    'name': fname,
                    'path': item.get('path', ''),
                    'download_url': download_url
                })
            else:
                print(f"  Не удалось получить ссылку для {fname}: {dl_resp.status_code}")
        except Exception as e:
            print(f"  Ошибка при получении ссылки для {fname}: {e}")
        if idx % 20 == 0:
            time.sleep(1)
    print(f"Получено прямых ссылок на FB2: {len(result)}")
    return result

def parse_fb2(content_bytes):
    """Парсит FB2 и возвращает {'title', 'author', 'description'}."""
    result = {'title': '', 'author': '', 'description': ''}

    text = None
    for enc in ('utf-8', 'windows-1251', 'koi8-r'):
        try:
            text = content_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        return result

    text = text.lstrip('\ufeff')

    try:
        root = ET.fromstring(text)
    except ET.ParseError:
        for marker in ('<FictionBook', '<fictionbook'):
            idx = text.find(marker)
            if idx >= 0:
                try:
                    root = ET.fromstring(text[idx:])
                    break
                except ET.ParseError:
                    continue
        else:
            return result

    ns_candidates = [
        'http://www.gribuser.ru/xml/fictionbook/2.0',
        'http://www.fictionbook.org/FictionBook2/Encodings',
        '',
    ]

    def find_el(parent, tag):
        for ns in ns_candidates:
            if ns:
                el = parent.find(f'.//{{{ns}}}{tag}')
                if el is not None:
                    return el
            else:
                el = parent.find(f'.//{tag}')
                if el is not None:
                    return el
        return None

    # === Название книги (book-title) ===
    title_el = find_el(root, 'book-title')
    if title_el is not None and title_el.text:
        result['title'] = title_el.text.strip()

    # === Автор ===
    title_info = find_el(root, 'title-info')
    author_el = None
    if title_info is not None:
        for ns in ns_candidates:
            if ns:
                author_el = title_info.find(f'{{{ns}}}author')
            else:
                author_el = title_info.find('author')
            if author_el is not None:
                break
    if author_el is None:
        author_el = find_el(root, 'author')

    if author_el is not None:
        first = last = middle = nickname = ''
        for ns in ns_candidates:
            prefix = f'{{{ns}}}' if ns else ''
            f = author_el.find(f'{prefix}first-name')
            l = author_el.find(f'{prefix}last-name')
            m = author_el.find(f'{prefix}middle-name')
            n = author_el.find(f'{prefix}nickname')
            if f is not None and f.text:
                first = f.text.strip()
            if l is not None and l.text:
                last = l.text.strip()
            if m is not None and m.text:
                middle = m.text.strip()
            if n is not None and n.text:
                nickname = n.text.strip()
            if first or last or middle or nickname:
                break
        parts = []
        if last:
            parts.append(last)
        if first:
            parts.append(first)
        if middle:
            parts.append(middle)
        if not parts and nickname:
            parts.append(nickname)
        result['author'] = ' '.join(parts)

    # === Описание (annotation) ===
    ann_el = find_el(root, 'annotation')
    if ann_el is not None:
        parts = []
        for sub in ann_el.iter():
            tag = sub.tag.split('}')[-1]
            if tag in ('subtitle', 'title', 'section'):
                continue
            if sub.text and sub.text.strip():
                parts.append(sub.text.strip())
        description = ' '.join(parts)
        if len(description) > MAX_DESC_LEN:
            description = description[:MAX_DESC_LEN] + '...'
        result['description'] = description

    return result

def normalize(name):
    if not name:
        return ''
    name = os.path.splitext(name)[0]
    name = re.sub(r'\s*\([^)]*\)\s*$', '', name)
    name = name.strip().lower()
    name = re.sub(r'\s+', ' ', name)
    return name

def main():
    if not YANDEX_PUBLIC_FOLDER_URL:
        print("YANDEX_PUBLIC_FOLDER_URL не задан.")
        return

    print("Подключение к Google Sheets...")
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet(SHEET_TITLE)
    print(f"Лист '{SHEET_TITLE}' открыт.")

    all_values = worksheet.get_all_values()
    if len(all_values) < 2:
        print("Таблица пустая.")
        return
    headers = all_values[0]
    print(f"Заголовки: {headers}")

    try:
        col_title = headers.index("Название")
        col_author = headers.index("Автор")
        col_desc = headers.index("Описание")
    except ValueError as e:
        print(f"Не найдены нужные колонки: {e}")
        return

    # Карта нормализованных названий из таблицы -> номер строки
    name_to_row = {}
    for i, row in enumerate(all_values[1:], start=2):
        if col_title < len(row):
            n = normalize(row[col_title])
            if n:
                name_to_row[n] = i

    print(f"Строк с названиями в таблице: {len(name_to_row)}")

    # Получаем FB2-файлы (передаём полную ссылку)
    files = get_yandex_files_with_download(YANDEX_PUBLIC_FOLDER_URL)
    if not files:
        print("FB2-файлы не получены.")
        return

    updates = []
    processed = 0
    skipped = 0
    failed = 0

    for f in files:
        norm = normalize(f['name'])
        if norm not in name_to_row:
            skipped += 1
            continue

        row_num = name_to_row[norm]
        row = all_values[row_num - 1]

        current_title = row[col_title] if col_title < len(row) else ''
        current_author = row[col_author] if col_author < len(row) else ''
        current_desc = row[col_desc] if col_desc < len(row) else ''

        # Если ONLY_EMPTY и все поля заполнены — пропускаем
        if ONLY_EMPTY and current_author.strip() and current_desc.strip() and current_title.strip():
            skipped += 1
            continue

        # Скачиваем FB2
        try:
            file_resp = requests.get(f['download_url'], timeout=60)
            if file_resp.status_code != 200:
                print(f"  {f['name']}: ошибка скачивания {file_resp.status_code}")
                failed += 1
                continue
            content = file_resp.content
        except Exception as e:
            print(f"  {f['name']}: {e}")
            failed += 1
            continue

        # Парсим
        try:
            parsed = parse_fb2(content)
        except Exception as e:
            print(f"  {f['name']}: ошибка парсинга {e}")
            failed += 1
            continue

        fb2_title = parsed.get('title', '').strip()
        fb2_author = parsed.get('author', '').strip()
        fb2_desc = parsed.get('description', '').strip()

        if not fb2_title and not fb2_author and not fb2_desc:
            print(f"  {f['name']}: FB2 пуст")
            skipped += 1
            continue

        # Применяем: только пустые поля или все
        if ONLY_EMPTY:
            final_title = current_title.strip() or fb2_title
            final_author = current_author.strip() or fb2_author
            final_desc = current_desc.strip() or fb2_desc
        else:
            final_title = fb2_title or current_title.strip()
            final_author = fb2_author or current_author.strip()
            final_desc = fb2_desc or current_desc.strip()

        updates.append({
            'row': row_num,
            'title': final_title,
            'author': final_author,
            'description': final_desc
        })
        processed += 1
        print(f"  ✓ {fb2_title[:60]} | {fb2_author[:60]}")

        time.sleep(0.3)

    print(f"\nИтого:")
    print(f"  Обработано: {processed}")
    print(f"  Пропущено: {skipped}")
    print(f"  Ошибок: {failed}")

    if not updates:
        print("Нечего записывать.")
        return

    print(f"\nОбновляем таблицу ({len(updates)} строк)...")

    title_data = []
    author_data = []
    desc_data = []

    col_letter_title = chr(65 + col_title)
    col_letter_author = chr(65 + col_author)
    col_letter_desc = chr(65 + col_desc)

    for u in updates:
        title_data.append({'range': f'{col_letter_title}{u["row"]}', 'values': [[u['title']]]})
        author_data.append({'range': f'{col_letter_author}{u["row"]}', 'values': [[u['author']]]})
        desc_data.append({'range': f'{col_letter_desc}{u["row"]}', 'values': [[u['description']]]})

    print("Обновляем колонку 'Название'...")
    for i in range(0, len(title_data), 50):
        worksheet.batch_update(title_data[i:i+50])
        time.sleep(1)

    print("Обновляем колонку 'Автор'...")
    for i in range(0, len(author_data), 50):
        worksheet.batch_update(author_data[i:i+50])
        time.sleep(1)

    print("Обновляем колонку 'Описание'...")
    for i in range(0, len(desc_data), 50):
        worksheet.batch_update(desc_data[i:i+50])
        time.sleep(1)

    print("Готово! Таблица обновлена.")

if __name__ == "__main__":
    main()
