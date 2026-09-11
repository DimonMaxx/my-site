import gspread
import os
import json
import sys
import re
import time
import base64
import requests
import xml.etree.ElementTree as ET
from urllib.parse import quote

# ========== НАСТРОЙКИ ==========
SPREADSHEET_ID = "1kcG0TG4GZtSM2mypjgvNDUpIbLfIvcmW80_hBKA11nw"
YANDEX_PUBLIC_FOLDER_URL = "https://disk.yandex.ru/d/zMxF4nXHPkIVCQ"
SHEET_TITLE = "Книги"

SUPABASE_URL = "https://rmoonebbvpmvthvpcmpt.supabase.co"
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
COVERS_BUCKET = "covers"

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

def get_yandex_files_with_download(public_url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    }
    api_url = "https://cloud-api.yandex.net/v1/disk/public/resources"
    params = {"public_key": public_url, "limit": 1000, "sort": "name"}
    resp = requests.get(api_url, params=params, headers=headers, timeout=30)
    if resp.status_code != 200:
        print(f"Ошибка списка файлов: {resp.status_code}")
        return []
    items = resp.json().get('_embedded', {}).get('items', [])
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
                result.append({
                    'name': fname,
                    'download_url': dl_resp.json().get('href')
                })
        except Exception as e:
            print(f"  Ошибка {fname}: {e}")
        if idx % 20 == 0:
            time.sleep(0.5)
    print(f"Получено файлов: {len(result)}")
    return result

def parse_fb2(content_bytes):
    """Возвращает словарь {title, author, description, cover_data, cover_ext}."""
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
                if el is not None: return el
            else:
                el = parent.find(f'.//{tag}')
                if el is not None: return el
        return None

    # Название
    t_el = find_el(root, 'book-title')
    if t_el is not None and t_el.text:
        result['title'] = t_el.text.strip()

    # Автор
    ti = find_el(root, 'title-info')
    a_el = None
    if ti is not None:
        for ns in ns_candidates:
            a_el = ti.find(f'{{{ns}}}author') if ns else ti.find('author')
            if a_el is not None: break
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

    # Описание
    ann = find_el(root, 'annotation')
    if ann is not None:
        parts = []
        for sub in ann.iter():
            tag = sub.tag.split('}')[-1]
            if tag in ('subtitle', 'title', 'section', 'image'): continue
            if sub.text and sub.text.strip():
                parts.append(sub.text.strip())
        desc = ' '.join(parts)
        if len(desc) > MAX_DESC_LEN:
            desc = desc[:MAX_DESC_LEN] + '...'
        result['description'] = desc

    # Обложка
    cover = find_el(root, 'coverpage')
    if cover is not None:
        # Ищем тег <image> внутри coverpage
        img_el = None
        for ns in ns_candidates:
            prefix = f'{{{ns}}}' if ns else ''
            img_el = cover.find(f'.//{prefix}image')
            if img_el is not None: break
        if img_el is not None:
            # Атрибут l:href или href
            href = None
            for attr in img_el.attrib:
                if attr.endswith('href'):
                    href = img_el.attrib[attr]
                    break
            if href and href.startswith('#'):
                binary_id = href[1:]
                # Ищем <binary id="...">
                for ns in ns_candidates:
                    prefix = f'{{{ns}}}' if ns else ''
                    for b_el in root.iter(f'{prefix}binary'):
                        if b_el.attrib.get('id') == binary_id:
                            content_type = b_el.attrib.get('content-type', 'image/jpeg')
                            ext = 'jpg'
                            if 'png' in content_type: ext = 'png'
                            elif 'gif' in content_type: ext = 'gif'
                            elif 'webp' in content_type: ext = 'webp'
                            try:
                                data = base64.b64decode(b_el.text.strip())
                                result['cover_data'] = data
                                result['cover_ext'] = ext
                            except Exception as e:
                                print(f"  Ошибка декодирования обложки: {e}")
                            break
                    if result['cover_data']:
                        break
    return result

def normalize(name):
    if not name: return ''
    name = os.path.splitext(name)[0]
    name = re.sub(r'\s*\([^)]*\)\s*$', '', name)
    name = name.strip().lower()
    name = re.sub(r'\s+', ' ', name)
    name = re.sub(r'[—–]', '-', name)
    name = re.sub(r'[^\w\s\-]', ' ', name)
    return re.sub(r'\s+', ' ', name).strip()

def upload_cover_to_supabase(cover_data, ext, book_slug):
    """Загружает обложку в Supabase Storage, возвращает публичный URL."""
    if not SUPABASE_SERVICE_KEY:
        print("  SUPABASE_SERVICE_ROLE_KEY не задан — пропускаем загрузку обложки.")
        return None
    filename = f"{book_slug}.{ext}"
    url = f"{SUPABASE_URL}/storage/v1/object/{COVERS_BUCKET}/{filename}"
    headers = {
        "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
        "Content-Type": f"image/{ext}",
        "x-upsert": "true"
    }
    try:
        resp = requests.put(url, headers=headers, data=cover_data, timeout=30)
        if resp.status_code in (200, 201):
            public_url = f"{SUPABASE_URL}/storage/v1/object/public/{COVERS_BUCKET}/{filename}"
            return public_url
        else:
            print(f"  Ошибка загрузки обложки: {resp.status_code} — {resp.text[:100]}")
    except Exception as e:
        print(f"  Ошибка загрузки обложки: {e}")
    return None

def main():
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
        col_cover = headers.index("Обложка") if "Обложка" in headers else -1
    except ValueError as e:
        print(f"Не найдены колонки: {e}")
        return

    if col_cover == -1:
        print("Колонка 'Обложка' не найдена. Добавьте её в таблицу.")
        return

    # Собираем существующие книги
    name_to_row = {}
    for i, row in enumerate(all_values[1:], start=2):
        if col_title < len(row):
            n = normalize(row[col_title])
            if n: name_to_row[n] = i

    files = get_yandex_files_with_download(YANDEX_PUBLIC_FOLDER_URL)
    if not files:
        print("Файлы не получены.")
        return

    updated = 0
    added = 0
    skipped = 0

    for f in files:
        norm = normalize(f['name'])
        fb2_title = ''
        # Скачиваем FB2
        try:
            resp = requests.get(f['download_url'], timeout=60)
            if resp.status_code != 200:
                continue
            parsed = parse_fb2(resp.content)
            fb2_title = parsed.get('title', '').strip()
        except Exception as e:
            print(f"  Ошибка парсинга {f['name']}: {e}")
            continue

        # Ищем строку в таблице
        row_num = None
        for db_title, rnum in name_to_row.items():
            if db_title == norm: 
                row_num = rnum; break
        if row_num is None and fb2_title:
            norm2 = normalize(fb2_title)
            if norm2 in name_to_row:
                row_num = name_to_row[norm2]

        if row_num:
            # Обновляем только обложку, если пуста
            row = all_values[row_num - 1]
            current_cover = row[col_cover] if col_cover < len(row) else ''
            if current_cover.strip():
                skipped += 1
                continue
            cover_data = parsed.get('cover_data')
            cover_ext = parsed.get('cover_ext', 'jpg')
            if not cover_data:
                skipped += 1
                continue
            slug = re.sub(r'[^\w\-]+', '-', normalize(fb2_title or f['name']))[:80]
            url = upload_cover_to_supabase(cover_data, cover_ext, slug)
            if url:
                col_letter = chr(65 + col_cover)
                worksheet.update(f'{col_letter}{row_num}', [[url]], value_input_option='RAW')
                print(f"  ✓ Обложка для: {fb2_title or f['name'][:40]}")
                updated += 1
                time.sleep(0.3)
        else:
            skipped += 1

    print(f"\nИтого: обновлено обложек: {updated}, пропущено: {skipped}")

if __name__ == "__main__":
    main()
