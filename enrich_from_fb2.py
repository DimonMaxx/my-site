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

# Скачивать только те файлы, у которых в таблице пуст "Автор" ИЛИ "Описание"
ONLY_EMPTY = True

# Максимальная длина описания в ячейке (Google Sheets имеет лимит 50000, но для удобства ограничим)
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

def get_yandex_files_with_download(public_key):
    """Получает список файлов с прямой ссылкой на скачивание каждого."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    }
    # Список файлов
    api_url = "https://cloud-api.yandex.net/v1/disk/public/resources"
    params = {"public_key": public_key, "limit": 1000, "sort": "name"}
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
        # Получаем прямую ссылку на скачивание
        dl_api = "https://cloud-api.yandex.net/v1/disk/public/resources/download"
        dl_params = {"public_key": public_key, "path": item['path']}
        try:
            dl_resp = requests.get(dl_api, params=dl_params, headers=headers, timeout=30)
            if dl_resp.status_code == 200:
                download_url = dl_resp.json().get('href')
                result.append({
                    'name': item.get('name', ''),
                    'path': item.get('path', ''),
                    'download_url': download_url
                })
            else:
                print(f"  Не удалось получить ссылку для {item.get('name')}: {dl_resp.status_code}")
        except Exception as e:
            print(f"  Ошибка при получении ссылки для {item.get('name')}: {e}")
        # Небольшая пауза, чтобы не спамить API
        if idx % 20 == 0:
            time.sleep(1)
    print(f"Получено прямых ссылок: {len(result)}")
    return result

def parse_fb2(content_bytes):
    """Парсит FB2 и возвращает (author, description)."""
    author = ''
    description = ''

    # FB2 может быть в UTF-8 или Windows-1251
    text = None
    for enc in ('utf-8', 'windows-1251', 'koi8-r'):
        try:
            text = content_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        return '', ''

    # Убираем BOM
    text = text.lstrip('\ufeff')

    try:
        root = ET.fromstring(text)
    except ET.ParseError as e:
        # Попробуем ещё раз, убрав всё до первого <FictionBook
        idx = text.find('<FictionBook')
        if idx == -1:
            idx = text.find('<fictionbook')
        if idx > 0:
            try:
                root = ET.fromstring(text[idx:])
            except ET.ParseError:
                return '', ''
        else:
            return '', ''

    # Пространство имён может быть, может не быть
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

    def find_text(parent, tag):
        el = find_el(parent, tag)
        return el.text.strip() if (el is not None and el.text) else ''

    # === Автор ===
    author_el = find_el(root, 'author')
    if author_el is not None:
        first = find_text(author_el, 'first-name')
        last = find_text(author_el, 'last-name')
        middle = find_text(author_el, 'middle-name')
        nickname = find_text(author_el, 'nickname')
        parts = [p for p in [last, first, middle] if p]
        if not parts and nickname:
            parts = [nickname]
        author = ' '.join(parts)

    # === Описание (annotation) ===
    ann_el = find_el(root, 'annotation')
    if ann_el is not None:
        # Собираем все текстовые узлы
        texts = []
        for el in ann_el.iter():
            # Пропускаем служебные теги
            if el.tag.split('}')[-1] in ('section', 'title', 'subtitle'):
                continue
            if el.text and el.text.strip():
                texts.append(el.text.strip())
        description = ' '.join(texts)

    # Обрезаем описание
    if len(description) > MAX_DESC_LEN:
        description = description[:MAX_DESC_LEN] + '...'

    return author, description

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

    public_key = extract_public_key(YANDEX_PUBLIC_FOLDER_URL)
    if not public_key:
        print("Не удалось извлечь public_key.")
        return

    print("Подключение к Google Sheets...")
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet(SHEET_TITLE)
    print(f"Лист '{SHEET_TITLE}' открыт.")

    # Читаем все данные
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

    # Индекс норм. названия -> номер строки (в all_values, где 0 — заголовок)
    name_to_row = {}
    for i, row in enumerate(all_values[1:], start=2):  # start=2 — реальный номер строки в листе
        if col_title < len(row):
            n = normalize(row[col_title])
            if n:
                name_to_row[n] = i

    print(f"Строк с названиями в таблице: {len(name_to_row)}")

    # Получаем список файлов с Яндекс.Диска
    files = get_yandex_files_with_download(public_key)
    if not files:
        print("Файлы не получены.")
        return

    # Определяем, какие файлы нужно обогатить
    updates = []  # список (row_number, author, description)
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

        # Если ONLY_EMPTY — пропускаем строки, где уже заполнены автор И описание
        current_author = row[col_author] if col_author < len(row) else ''
        current_desc = row[col_desc] if col_desc < len(row) else ''
        if ONLY_EMPTY and current_author.strip() and current_desc.strip():
            skipped += 1
            continue

        # Скачиваем файл
        try:
            print(f"Скачиваем: {f['name']}")
            file_resp = requests.get(f['download_url'], timeout=60)
            if file_resp.status_code != 200:
                print(f"  Ошибка скачивания: {file_resp.status_code}")
                failed += 1
                continue
            content = file_resp.content
        except Exception as e:
            print(f"  Ошибка: {e}")
            failed += 1
            continue

        # Парсим FB2
        try:
            author, description = parse_fb2(content)
        except Exception as e:
            print(f"  Ошибка парсинга: {e}")
            failed += 1
            continue

        # Если автор или описание не найдены — пропускаем
        if not author and not description:
            print(f"  Пусто в FB2")
            skipped += 1
            continue

        # Обновляем только пустые поля
        final_author = current_author.strip() or author
        final_desc = current_desc.strip() or description

        updates.append({
            'row': row_num,
            'author': final_author,
            'description': final_desc
        })
        processed += 1
        print(f"  ✓ Автор: {author[:60]}... | Описание: {description[:60]}...")

        # Пауза, чтобы не перегревать Яндекс.Диск
        time.sleep(0.3)

    print(f"\nИтого:")
    print(f"  Обработано и подготовлено к записи: {processed}")
    print(f"  Пропущено: {skipped}")
    print(f"  Ошибок: {failed}")

    if not updates:
        print("Нечего записывать.")
        return

    # Формируем batch_update
    print(f"\nОбновляем таблицу ({len(updates)} строк)...")
    batch_data = []
    for u in updates:
        # Автор — колонка B (2), Описание — колонка C (3)
        # Учитываем, что индексы заголовков могут отличаться — используем реальные колонки
        col_letter_author = chr(65 + col_author)  # A=65
        col_letter_desc = chr(65 + col_desc)
        batch_data.append({
            'range': f'{col_letter_author}{u["row"]}:{col_letter_desc}{u["row"]}',
            'values': [[u['author'], u['description']]]
        })

    # gspread позволяет обновлять пачкой через batch_update
    # Но если колонки Автор и Описание не соседние, нужно два запроса
    if col_author == col_desc - 1:
        # Соседние — одним диапазоном
        worksheet.batch_update(batch_data)
    else:
        # Не соседние — обновляем по одной колонке
        author_data = []
        desc_data = []
        col_letter_author = chr(65 + col_author)
        col_letter_desc = chr(65 + col_desc)
        for u in updates:
            author_data.append({'range': f'{col_letter_author}{u["row"]}', 'values': [[u['author']]]})
            desc_data.append({'range': f'{col_letter_desc}{u["row"]}', 'values': [[u['description']]]})
        worksheet.batch_update(author_data)
        worksheet.batch_update(desc_data)

    print("Готово! Таблица обновлена.")

if __name__ == "__main__":
    main()
