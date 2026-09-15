# yandex_disk_sync.py
# Синхронизация Google Sheets с Яндекс.Диском.
# - Книги: folder = имя папки автора
# - Программы: folder = имя подпапки (Скрипты/Макросы/...)
# - FB2: парсим автора/описание/обложку из метаданных
# - TXT/DOC/DOCX/RTF: автора берём из имени папки, контент не парсим

import os
import sys
import time
import base64
import hashlib
import requests
import xml.etree.ElementTree as ET
from urllib.parse import quote

from common import (
    SPREADSHEET_ID,
    SECTION_TO_SHEET,
    get_gspread_client,
    normalize,
)

# ========== НАСТРОЙКИ ==========
# Публичные ссылки на папки Яндекс.Диска по разделам.
YANDEX_SOURCES = {
    'books':    'https://disk.yandex.ru/d/zMxF4nXHPkIVCQ',
    'programs': 'https://disk.yandex.ru/d/EjUHvm6mUcgVMw',
}

SUPABASE_URL = "https://rmoonebbvpmvthvpcmpt.supabase.co"
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
COVERS_BUCKET = "covers"

MAX_DESC_LEN = 2000

# Форматы книг (для заполнения колонки «Формат»)
BOOK_EXTS = {'.fb2', '.epub', '.pdf', '.djvu', '.mobi', '.txt', '.doc', '.docx', '.rtf'}
# Форматы, которые ПАРСИМ (только FB2 — XML)
PARSEABLE_EXTS = {'.fb2'}

FIELDS_FOR_SECTION = {
    "programs": ["title", "folder", "description", "version", "size", "download_link"],
    "books":    ["title", "folder", "author", "description", "format", "download_link", "cover"],
    "news":     ["title", "date", "body"],
    "articles": ["title", "date", "body"],
    "movies":   ["title", "folder", "year", "description", "download_link"],
    "music":    ["title", "folder", "artist", "year", "description", "download_link"],
    "games":    ["title", "folder", "platform", "year", "description", "download_link"],
    "misc":     ["title", "description", "download_link"],
}

RU_LABELS = {
    "title": "Название",
    "folder": "Папка",
    "description": "Описание",
    "version": "Версия",
    "size": "Размер (МБ)",
    "author": "Автор",
    "format": "Формат",
    "year": "Год",
    "artist": "Исполнитель",
    "platform": "Платформа",
    "date": "Дата",
    "body": "Текст",
    "download_link": "Ссылка для скачивания",
    "cover": "Обложка",
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
}

# ==============================


def col_num_to_letter(n):
    """1 → A, 26 → Z, 27 → AA и т.д."""
    result = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def list_yandex_recursive(public_url, sub_path=None):
    """
    Рекурсивный обход публичной папки Яндекс.Диска.
    Возвращает список: { name, path, parts, size, ext, download_url }
    """
    api_url = "https://cloud-api.yandex.net/v1/disk/public/resources"
    params = {"public_key": public_url, "limit": 1000, "sort": "name"}
    if sub_path:
        params["path"] = sub_path

    try:
        resp = requests.get(api_url, params=params, headers=HEADERS, timeout=30)
    except Exception as e:
        print(f"  Ошибка запроса {sub_path or '/'}: {e}")
        return []

    if resp.status_code != 200:
        print(f"  Ошибка списка {sub_path or '/'}: {resp.status_code}")
        return []

    items = resp.json().get('_embedded', {}).get('items', [])
    result = []

    for item in items:
        item_type = item.get('type')
        item_path = item.get('path', '')
        item_name = item.get('name', '')

        if item_type == 'dir':
            result.extend(list_yandex_recursive(public_url, item_path))
        elif item_type == 'file':
            ext = ''
            if '.' in item_name:
                ext = '.' + item_name.rsplit('.', 1)[1].lower()

            parts = [p for p in item_path.split('/') if p]

            result.append({
                'name': item_name,
                'path': item_path,
                'parts': parts,
                'size': item.get('size', 0),
                'ext': ext,
                'download_url': None,
            })

    return result


def compute_folder(parts):
    """
    parts = части пути внутри публичной папки:
      ['АРКАДИЙ СТРУГАЦКИЙ', 'книга.fb2']            → 'АРКАДИЙ СТРУГАЦКИЙ'
      ['Скрипты', 'macro.exe']                       → 'Скрипты'
      ['Скрипты', 'Подпапка', 'macro.exe']           → 'Скрипты / Подпапка'
      ['file.fb2']                                    → ''
    """
    if len(parts) <= 1:
        return ''
    return ' / '.join(parts[:-1])


def get_download_url(public_url, path):
    """Временная ссылка на скачивание."""
    dl_api = "https://cloud-api.yandex.net/v1/disk/public/resources/download"
    params = {"public_key": public_url, "path": path}
    try:
        resp = requests.get(dl_api, params=params, headers=HEADERS, timeout=30)
        if resp.status_code == 200:
            return resp.json().get('href')
    except Exception as e:
        print(f"  Ошибка получения download URL для {path}: {e}")
    return None


def make_permanent_link(public_url, path):
    """Постоянная ссылка через публичную папку."""
    return f"{public_url}?path={quote(path)}"


def parse_fb2(content_bytes):
    """Парсит FB2. Возвращает title, author, description, cover_data, cover_ext."""
    result = {'title': '', 'author': '', 'description': '', 'cover_data': None, 'cover_ext': ''}
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

    t_el = find_el(root, 'book-title')
    if t_el is not None and t_el.text:
        result['title'] = t_el.text.strip()

    ti = find_el(root, 'title-info')
    a_el = None
    if ti is not None:
        for ns in ns_candidates:
            a_el = ti.find(f'{{{ns}}}author') if ns else ti.find('author')
            if a_el is not None:
                break
    if a_el is None:
        a_el = find_el(root, 'author')
    if a_el is not None:
        parts = []
        for tag in ('last-name', 'first-name', 'middle-name'):
            for ns in ns_candidates:
                prefix = f'{{{ns}}}' if ns else ''
                el = a_el.find(f'{prefix}{tag}')
                if el is not None and el.text:
                    parts.append(el.text.strip())
                    break
        result['author'] = ' '.join(parts)

    ann = find_el(root, 'annotation')
    if ann is not None:
        parts = []
        for sub in ann.iter():
            tag = sub.tag.split('}')[-1]
            if tag in ('subtitle', 'title', 'section', 'image'):
                continue
            if sub.text and sub.text.strip():
                parts.append(sub.text.strip())
        desc = ' '.join(parts)
        if len(desc) > MAX_DESC_LEN:
            desc = desc[:MAX_DESC_LEN] + '...'
        result['description'] = desc

    cover = find_el(root, 'coverpage')
    if cover is not None:
        img_el = None
        for ns in ns_candidates:
            prefix = f'{{{ns}}}' if ns else ''
            img_el = cover.find(f'.//{prefix}image')
            if img_el is not None:
                break
        if img_el is not None:
            href = None
            for attr in img_el.attrib:
                if attr.endswith('href'):
                    href = img_el.attrib[attr]
                    break
            if href and href.startswith('#'):
                binary_id = href[1:]
                for ns in ns_candidates:
                    prefix = f'{{{ns}}}' if ns else ''
                    for b_el in root.iter(f'{prefix}binary'):
                        if b_el.attrib.get('id') == binary_id:
                            content_type = b_el.attrib.get('content-type', 'image/jpeg')
                            ext_img = 'jpg'
                            if 'png' in content_type:
                                ext_img = 'png'
                            elif 'gif' in content_type:
                                ext_img = 'gif'
                            elif 'webp' in content_type:
                                ext_img = 'webp'
                            try:
                                data = base64.b64decode(b_el.text.strip())
                                result['cover_data'] = data
                                result['cover_ext'] = ext_img
                            except Exception as e:
                                print(f"  Ошибка декодирования обложки: {e}")
                            break
                    if result['cover_data']:
                        break
    return result


def upload_cover_to_supabase(cover_data, ext, book_title):
    if not SUPABASE_SERVICE_KEY:
        print("  SUPABASE_SERVICE_ROLE_KEY не задан — пропускаем загрузку обложки.")
        return None
    filename = f"{hashlib.md5(book_title.encode('utf-8')).hexdigest()}.{ext}"
    url = f"{SUPABASE_URL}/storage/v1/object/{COVERS_BUCKET}/{filename}"
    headers = {
        "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
        "Content-Type": f"image/{ext}",
        "x-upsert": "true",
        "apikey": SUPABASE_SERVICE_KEY
    }
    try:
        resp = requests.put(url, headers=headers, data=cover_data, timeout=30)
        if resp.status_code in (200, 201):
            return f"{SUPABASE_URL}/storage/v1/object/public/{COVERS_BUCKET}/{filename}"
        else:
            print(f"  Ошибка загрузки обложки: {resp.status_code} — {resp.text[:200]}")
    except Exception as e:
        print(f"  Ошибка загрузки обложки: {e}")
    return None


def ensure_columns(worksheet, headers, required_fields):
    headers = list(headers)
    for field in required_fields:
        ru = RU_LABELS.get(field)
        if ru and ru not in headers:
            col_idx = len(headers) + 1
            col_letter = col_num_to_letter(col_idx)
            worksheet.update(f'{col_letter}1', [[ru]], value_input_option='RAW')
            headers.append(ru)
            print(f"    + создана колонка '{ru}' (позиция {col_letter})")
    return headers


def process_sheet(worksheet, section_key, files, public_url):
    sheet_name = worksheet.title
    print(f"\n  Лист '{sheet_name}': {len(files)} файлов")

    all_values = worksheet.get_all_values()
    if not all_values:
        print(f"    Лист пустой.")
        return

    headers = list(all_values[0])
    required_fields = FIELDS_FOR_SECTION.get(section_key, [])
    headers = ensure_columns(worksheet, headers, required_fields)

    all_values = worksheet.get_all_values()
    headers = list(all_values[0])

    col_idx = {}
    for field, ru_label in RU_LABELS.items():
        if ru_label in headers:
            col_idx[field] = headers.index(ru_label)

    if 'title' not in col_idx:
        print(f"    На листе нет колонки 'Название', пропускаем.")
        return

    title_to_row = {}
    for i, row in enumerate(all_values[1:], start=2):
        if col_idx['title'] < len(row):
            n = normalize(row[col_idx['title']])
            if n:
                title_to_row[n] = i

    updated_cells = 0
    for f in files:
        norm_name = normalize(f['name'])
        row_num = title_to_row.get(norm_name)
        if row_num is None:
            continue

        row = all_values[row_num - 1] if row_num - 1 < len(all_values) else []

        # Парсим FB2 только если это .fb2 и есть пустые поля
        need_parse = False
        if section_key == "books" and f['ext'] in PARSEABLE_EXTS:
            for field in ['author', 'description', 'cover']:
                if field in col_idx:
                    v = row[col_idx[field]] if col_idx[field] < len(row) else ''
                    if not v or not str(v).strip():
                        need_parse = True
                        break

        fb2_data = None
        if need_parse:
            if not f['download_url']:
                f['download_url'] = get_download_url(public_url, f['path'])
            if f['download_url']:
                try:
                    r = requests.get(f['download_url'], timeout=60)
                    if r.status_code == 200:
                        fb2_data = parse_fb2(r.content)
                except Exception as e:
                    print(f"      Ошибка парсинга {f['name']}: {e}")

        updates = {}

        # Папка
        if 'folder' in col_idx and f['folder']:
            v = row[col_idx['folder']] if col_idx['folder'] < len(row) else ''
            if not str(v).strip():
                updates['folder'] = f['folder']

        # Ссылка
        if 'download_link' in col_idx:
            v = row[col_idx['download_link']] if col_idx['download_link'] < len(row) else ''
            if not str(v).strip():
                updates['download_link'] = make_permanent_link(public_url, f['path'])

        # Формат (для книг)
        if 'format' in col_idx and f['ext']:
            v = row[col_idx['format']] if col_idx['format'] < len(row) else ''
            if not str(v).strip():
                updates['format'] = f['ext'].lstrip('.')

        # Размер (для программ)
        if 'size' in col_idx and f['size'] > 0:
            v = row[col_idx['size']] if col_idx['size'] < len(row) else ''
            if not str(v).strip():
                updates['size'] = str(round(f['size'] / 1024 / 1024, 1))

        # Автор
        if 'author' in col_idx:
            v = row[col_idx['author']] if col_idx['author'] < len(row) else ''
            if not str(v).strip():
                # 1) Из метаданных FB2
                if fb2_data and fb2_data.get('author'):
                    updates['author'] = fb2_data['author']
                # 2) Из имени папки (для txt/doc/docx/rtf и как fallback)
                elif f['folder']:
                    updates['author'] = f['folder'].split(' / ')[0]

        # Описание (только FB2)
        if fb2_data and 'description' in col_idx and fb2_data.get('description'):
            v = row[col_idx['description']] if col_idx['description'] < len(row) else ''
            if not str(v).strip():
                updates['description'] = fb2_data['description']

        # Обложка (только FB2)
        if fb2_data and 'cover' in col_idx and fb2_data.get('cover_data'):
            v = row[col_idx['cover']] if col_idx['cover'] < len(row) else ''
            if not str(v).strip():
                url = upload_cover_to_supabase(
                    fb2_data['cover_data'],
                    fb2_data.get('cover_ext', 'jpg'),
                    fb2_data.get('title') or f['name']
                )
                if url:
                    updates['cover'] = url

        for field, value in updates.items():
            cell = f"{col_num_to_letter(col_idx[field] + 1)}{row_num}"
            try:
                worksheet.update(cell, [[value]], value_input_option='RAW')
                updated_cells += 1
            except Exception as e:
                print(f"      Ошибка обновления {cell}: {e}")

    print(f"    Обновлено ячеек у существующих строк: {updated_cells}")

    # Новые записи
    new_rows = []
    new_titles = set()

    for f in files:
        norm_name = normalize(f['name'])
        if norm_name in title_to_row:
            continue
        if norm_name in new_titles:
            continue
        new_titles.add(norm_name)

        title = os.path.splitext(f['name'])[0]
        fb2_data = {}

        need_parse = (section_key == "books" and f['ext'] in PARSEABLE_EXTS)
        if need_parse:
            if not f['download_url']:
                f['download_url'] = get_download_url(public_url, f['path'])
            if f['download_url']:
                try:
                    r = requests.get(f['download_url'], timeout=60)
                    if r.status_code == 200:
                        fb2_data = parse_fb2(r.content)
                        if fb2_data.get('title'):
                            title = fb2_data['title']
                except Exception as e:
                    print(f"      Ошибка парсинга {f['name']}: {e}")

        row_values = {
            'title': title,
            'folder': f['folder'],
            'download_link': make_permanent_link(public_url, f['path']),
        }
        if section_key == "books":
            row_values['format'] = f['ext'].lstrip('.') if f['ext'] else ''
            # Автор: из FB2, иначе из папки
            if fb2_data.get('author'):
                row_values['author'] = fb2_data['author']
            elif f['folder']:
                row_values['author'] = f['folder'].split(' / ')[0]
            if fb2_data.get('description'):
                row_values['description'] = fb2_data['description']
            if fb2_data.get('cover_data'):
                url = upload_cover_to_supabase(
                    fb2_data['cover_data'],
                    fb2_data.get('cover_ext', 'jpg'),
                    title
                )
                if url:
                    row_values['cover'] = url
        elif section_key == "programs":
            if f['size'] > 0:
                row_values['size'] = str(round(f['size'] / 1024 / 1024, 1))
            row_values['version'] = ''

        max_col = max(col_idx.values()) if col_idx else 0
        row_array = [''] * (max_col + 1)
        for field, value in row_values.items():
            if field in col_idx:
                row_array[col_idx[field]] = value
        new_rows.append(row_array)

    if new_rows:
        try:
            max_len = max(len(r) for r in new_rows)
            for r in new_rows:
                while len(r) < max_len:
                    r.append('')
            worksheet.append_rows(new_rows, value_input_option='RAW')
            print(f"    Добавлено новых строк: {len(new_rows)}")
        except Exception as e:
            print(f"    Ошибка добавления строк: {e}")


def main():
    print("Подключение к Google Sheets...")
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    for section_key, public_url in YANDEX_SOURCES.items():
        section_ru = SECTION_TO_SHEET.get(section_key)
        if not section_ru:
            print(f"\nРаздел '{section_key}' не найден в SECTION_TO_SHEET, пропускаем.")
            continue

        print(f"\n=== Раздел: {section_ru} ({section_key}) ===")
        print(f"Источник: {public_url}")

        files = list_yandex_recursive(public_url)
        print(f"Найдено файлов: {len(files)}")

        if not files:
            continue

        for f in files:
            f['folder'] = compute_folder(f['parts'])

        from collections import Counter
        folders_count = Counter(f['folder'] or '(без папки)' for f in files)
        print(f"Папок: {len(folders_count)}")
        for folder_name, cnt in sorted(folders_count.items()):
            print(f"  • {folder_name}: {cnt} файлов")

        try:
            worksheet = sh.worksheet(section_ru)
        except Exception:
            print(f"Лист '{section_ru}' не найден, пропускаем.")
            continue

        try:
            process_sheet(worksheet, section_key, files, public_url)
        except Exception as e:
            print(f"Ошибка обработки листа '{section_ru}': {e}")

    print("\nГотово!")


if __name__ == "__main__":
    main()
