import gspread
import os
import json
import sys
import re
import time
import requests
import xml.etree.ElementTree as ET
from urllib.parse import quote

# ========== НАСТРОЙКИ ==========
SPREADSHEET_ID = "1kcG0TG4GZtSM2mypjgvNDUpIbLfIvcmW80_hBKA11nw"
YANDEX_PUBLIC_FOLDER_URL = "https://disk.yandex.ru/d/zMxF4nXHPkIVCQ"
SHEET_TITLE = "Книги"
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

def normalize(name):
    """Нормализует название для сравнения (убирает расширение, скобки, лишние пробелы, нижний регистр)."""
    if not name:
        return ''
    name = os.path.splitext(name)[0]
    name = re.sub(r'\s*\([^)]*\)\s*$', '', name)
    name = name.strip().lower()
    name = re.sub(r'\s+', ' ', name)
    # Убираем все виды тире/дефисов в один символ
    name = re.sub(r'[—–]', '-', name)
    # Убираем возможные невидимые символы
    name = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', name)
    return name

def get_yandex_files_with_download(public_url):
    """Получает список FB2-файлов с прямой ссылкой на скачивание."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    }
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

        dl_api = "https://cloud-api.yandex.net/v1/disk/public/resources/download"
        dl_params = {"public_key": public_url, "path": item['path']}
        try:
            dl_resp = requests.get(dl_api, params=dl_params, headers=headers, timeout=30)
            if dl_resp.status_code == 200:
                download_url = dl_resp.json().get('href')
                result.append({
                    'name': fname,
                    'download_url': download_url
                })
            else:
                print(f"  {fname}: не удалось получить ссылку ({dl_resp.status_code})")
        except Exception as e:
            print(f"  {fname}: ошибка {e}")
        if idx % 20 == 0:
            time.sleep(0.5)
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

    # Название книги
    title_el = find_el(root, 'book-title')
    if title_el is not None and title_el.text:
        result['title'] = title_el.text.strip()

    # Автор
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

    # Описание
    ann_el = find_el(root, 'annotation')
    if ann_el is not None:
        parts = []
        for sub in ann_el.iter():
            tag = sub.tag.split('}')[-1]
            if tag in ('subtitle', 'title', 'section', 'image'):
                continue
            if sub.text and sub.text.strip():
                parts.append(sub.text.strip())
        description = ' '.join(parts)
        if len(description) > MAX_DESC_LEN:
            description = description[:MAX_DESC_LEN] + '...'
        result['description'] = description

    return result

def build_download_link(public_url, file_name):
    """Формирует ссылку на файл в публичной папке."""
    public_key_match = re.search(r'/d/([a-zA-Z0-9_-]+)', public_url)
    if not public_key_match:
        return public_url
    public_key = public_key_match.group(1)
    encoded = quote(file_name, safe='')
    return f"https://disk.yandex.ru/d/{public_key}?path=/{encoded}"

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
    headers = all_values[0] if all_values else []
    print(f"Заголовки: {headers}")

    try:
        col_title = headers.index("Название")
        col_author = headers.index("Автор")
        col_desc = headers.index("Описание")
        col_format = headers.index("Формат")
        col_link = headers.index("Ссылка для скачивания")
    except ValueError as e:
        print(f"Не найдены нужные колонки: {e}")
        return

    # Карта: нормализованное название -> номер строки (1-based)
    name_to_row = {}
    # Запоминаем строки, у которых пустой автор или описание
    rows_need_update = {}  # row_num -> {'title': ..., 'author_empty': bool, 'desc_empty': bool}
    for i, row in enumerate(all_values[1:], start=2):
        if col_title < len(row):
            title = row[col_title].strip()
            if title:
                n = normalize(title)
                name_to_row[n] = i
                author_empty = not (col_author < len(row) and row[col_author].strip())
                desc_empty = not (col_desc < len(row) and row[col_desc].strip())
                if author_empty or desc_empty:
                    rows_need_update[i] = {
                        'title': title,
                        'author_empty': author_empty,
                        'desc_empty': desc_empty
                    }

    print(f"Строк с названиями в таблице: {len(name_to_row)}")
    print(f"Строк, требующих заполнения автора/описания: {len(rows_need_update)}")

    # Получаем файлы с Яндекс.Диска
    files = get_yandex_files_with_download(YANDEX_PUBLIC_FOLDER_URL)
    if not files:
        print("Файлы не получены.")
        return

    new_rows = []          # для добавления
    updates_by_row = {}    # row_num -> {'title': ..., 'author': ..., 'description': ...}

    processed_new = 0
    processed_update = 0
    failed = 0

    for idx, f in enumerate(files, start=1):
        fname = f['name']
        norm = normalize(fname)

        # Скачиваем FB2
        try:
            file_resp = requests.get(f['download_url'], timeout=60)
            if file_resp.status_code != 200:
                print(f"  {fname}: ошибка скачивания {file_resp.status_code}")
                failed += 1
                continue
            content = file_resp.content
        except Exception as e:
            print(f"  {fname}: ошибка {e}")
            failed += 1
            continue

        # Парсим FB2
        try:
            parsed = parse_fb2(content)
        except Exception as e:
            print(f"  {fname}: ошибка парсинга {e}")
            failed += 1
            continue

        fb2_title = parsed.get('title', '').strip()
        fb2_author = parsed.get('author', '').strip()
        fb2_desc = parsed.get('description', '').strip()

        # Если title из FB2 пуст — используем имя файла без расширения
        if not fb2_title:
            fb2_title = os.path.splitext(fname)[0]
            fb2_title = re.sub(r'\s*\([^)]*\)\s*$', '', fb2_title).strip()

        ext = os.path.splitext(fname)[1].lstrip('.').lower()
        download_link = build_download_link(YANDEX_PUBLIC_FOLDER_URL, fname)

        # Ищем в таблице
        if norm in name_to_row:
            # Файл уже есть в таблице — возможно, обновим поля
            row_num = name_to_row[norm]
            if row_num in rows_need_update:
                info = rows_need_update[row_num]
                new_author = info['author_empty'] and fb2_author
                new_desc = info['desc_empty'] and fb2_desc
                if new_author or new_desc:
                    updates_by_row[row_num] = {
                        'title': info['title'],   # оставляем как есть
                        'author': fb2_author if new_author else '',
                        'description': fb2_desc if new_desc else '',
                        'set_author': new_author,
                        'set_desc': new_desc
                    }
                    processed_update += 1
                    print(f"  [ОБНОВЛЕНИЕ] {info['title'][:50]} | Автор: {fb2_author[:40]} | Описание: {fb2_desc[:40]}...")
        else:
            # Новый файл — добавляем
            new_rows.append({
                'title': fb2_title,
                'author': fb2_author,
                'description': fb2_desc,
                'format': ext,
                'link': download_link
            })
            processed_new += 1
            print(f"  [НОВОЕ] {fb2_title[:60]} | {fb2_author[:40]}")

        if idx % 10 == 0:
            time.sleep(0.3)

    print(f"\nИтого:")
    print(f"  Новых книг: {processed_new}")
    print(f"  Обновлений существующих: {processed_update}")
    print(f"  Ошибок: {failed}")

    # Записываем новые строки
    if new_rows:
        print(f"\nДобавляем {len(new_rows)} новых строк...")
        rows_to_append = []
        for r in new_rows:
            row = []
            for header in headers:
                if header == "Название":
                    row.append(r['title'])
                elif header == "Автор":
                    row.append(r['author'])
                elif header == "Описание":
                    row.append(r['description'])
                elif header == "Формат":
                    row.append(r['format'])
                elif header == "Ссылка для скачивания":
                    row.append(r['link'])
                else:
                    row.append("")
            rows_to_append.append(row)
        try:
            worksheet.append_rows(rows_to_append, value_input_option='RAW')
            print(f"  ✓ Добавлено {len(rows_to_append)} строк.")
        except gspread.exceptions.APIError as e:
            if '429' in str(e):
                print("  Превышена квота, ждём 60 секунд...")
                time.sleep(60)
                worksheet.append_rows(rows_to_append, value_input_option='RAW')
                print(f"  ✓ Добавлено {len(rows_to_append)} строк (со 2-й попытки).")
            else:
                raise

    # Обновляем существующие строки
    if updates_by_row:
        print(f"\nОбновляем {len(updates_by_row)} существующих строк...")
        col_letter_author = chr(65 + col_author)
        col_letter_desc = chr(65 + col_desc)

        author_batch = []
        desc_batch = []
        for row_num, u in updates_by_row.items():
            if u['set_author']:
                author_batch.append({'range': f'{col_letter_author}{row_num}', 'values': [[u['author']]]})
            if u['set_desc']:
                desc_batch.append({'range': f'{col_letter_desc}{row_num}', 'values': [[u['description']]]})

        # Отправляем пачками по 50
        if author_batch:
            print(f"  Обновляем авторов ({len(author_batch)} строк)...")
            for i in range(0, len(author_batch), 50):
                worksheet.batch_update(author_batch[i:i+50])
                time.sleep(1)
        if desc_batch:
            print(f"  Обновляем описания ({len(desc_batch)} строк)...")
            for i in range(0, len(desc_batch), 50):
                worksheet.batch_update(desc_batch[i:i+50])
                time.sleep(1)
        print(f"  ✓ Обновлено {len(updates_by_row)} строк.")

    print("\nГотово!")

if __name__ == "__main__":
    main()
