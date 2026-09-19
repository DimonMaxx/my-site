# yandex_disk_sync.py
# Синхронизация Google Sheets с Яндекс.Диском + обогащение описаний.
# - Книги: folder = имя папки автора
# - Программы: folder = имя подпапки (Скрипты/Макросы/...)
# - FB2: парсим автора/описание/обложку из метаданных
# - TXT: парсим аннотацию, если она есть в первых 30 КБ
# - DOC/DOCX/RTF: автора берём из имени папки, контент не парсим
# - Для книг без описания — запрос во внешние источники:
#     FantLab → Wikipedia → OpenLibrary → Google Books

import os
import re
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


# ============================================================
# НАСТРОЙКИ
# ============================================================

# Публичные ссылки на папки Яндекс.Диска.
# ВАЖНО: передаём ИМЕННО ПОЛНУЮ ССЫЛКУ — публичный API так и ждёт.
YANDEX_SOURCES = {
    'books':    'https://disk.yandex.ru/d/zMxF4nXHPkIVCQ',
    'programs': 'https://disk.yandex.ru/d/EjUHvm6mUcgVMw',
}

SUPABASE_URL = "https://rmoonebbvpmvthvpcmpt.supabase.co"
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
COVERS_BUCKET = "covers"

MAX_DESC_LEN = 2000

BOOK_EXTS = {'.fb2', '.epub', '.pdf', '.djvu', '.mobi', '.txt',
             '.doc', '.docx', '.rtf'}
PARSEABLE_EXTS = {'.fb2', '.txt'}

FIELDS_FOR_SECTION = {
    "programs": ["title", "folder", "description", "version", "size", "download_link"],
    "books":    ["title", "folder", "author", "description", "format",
                 "download_link", "cover"],
    "news":     ["title", "date", "body"],
    "articles": ["title", "date", "body"],
    "movies":   ["title", "folder", "year", "description", "download_link"],
    "music":    ["title", "folder", "artist", "year", "description", "download_link"],
    "games":    ["title", "folder", "platform", "year", "description", "download_link"],
    "misc":     ["title", "description", "download_link"],
}

RU_LABELS = {
    "title":         "Название",
    "folder":        "Папка",
    "description":   "Описание",
    "version":       "Версия",
    "size":          "Размер (МБ)",
    "author":        "Автор",
    "format":        "Формат",
    "year":          "Год",
    "artist":        "Исполнитель",
    "platform":      "Платформа",
    "date":          "Дата",
    "body":          "Текст",
    "download_link": "Ссылка для скачивания",
    "cover":         "Обложка",
}

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/120.0 Safari/537.36")
}

HTTP_TIMEOUT = 30
SLEEP_BETWEEN_REQUESTS = 0.4


# ============================================================
# УТИЛИТЫ
# ============================================================

def col_num_to_letter(n):
    """1 → A, 26 → Z, 27 → AA."""
    result = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _normalize_for_match(text):
    if not text:
        return ""
    t = str(text).lower().replace("ё", "е")
    t = re.sub(r"[^\w\s]", " ", t, flags=re.UNICODE)
    t = re.sub(r"\s+", " ", t).strip()
    return t


_STOP_WORDS = {"или", "как", "для", "при", "над", "под", "без", "про",
               "the", "and", "for", "with", "from", "that", "this"}

BOOK_MARKERS = (
    "книга", "роман", "повесть", "рассказ", "произведение",
    "сборник", "трилогия", "эпопея", "цикл", "литератур",
    "novel", "book", "story",
)


def _title_phrase_matches(title, text):
    """Совпадает ли название книги с текстом (для валидации внешних источников)."""
    title_norm = _normalize_for_match(title).rstrip(" .")
    text_norm  = _normalize_for_match(text)
    if not title_norm or not text_norm:
        return False

    title_words_all = title_norm.split()
    if len(title_words_all) >= 2 and title_norm in text_norm:
        return True

    title_words = [w for w in title_words_all if len(w) >= 4]
    if not title_words:
        return title_norm in text_norm

    text_words = set(text_norm.split())
    matched = [w for w in title_words if w in text_words]
    if not matched:
        return False
    ratio = len(matched) / len(title_words)
    if len(matched) >= 2 and ratio >= 0.6:
        text_l = text.lower()
        return any(m in text_l for m in BOOK_MARKERS)
    return False


# ============================================================
# ЯНДЕКС.ДИСК — обход и скачивание
# ============================================================

def list_yandex_recursive(public_url, sub_path=None):
    """
    Рекурсивный обход публичной папки Яндекс.Диска.
    Возвращает список: {name, path, parts, size, ext, download_url}
    """
    api_url = "https://cloud-api.yandex.net/v1/disk/public/resources"
    params = {"public_key": public_url, "limit": 1000, "sort": "name"}
    if sub_path:
        params["path"] = sub_path

    try:
        resp = requests.get(api_url, params=params, headers=HEADERS, timeout=HTTP_TIMEOUT)
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
    if len(parts) <= 1:
        return ''
    return ' / '.join(parts[:-1])


def get_download_url(public_url, path):
    """Временная ссылка на скачивание файла через публичный API."""
    dl_api = "https://cloud-api.yandex.net/v1/disk/public/resources/download"
    params = {"public_key": public_url, "path": path}
    try:
        resp = requests.get(dl_api, params=params, headers=HEADERS, timeout=HTTP_TIMEOUT)
        if resp.status_code == 200:
            return resp.json().get('href')
    except Exception as e:
        print(f"  Ошибка получения download URL для {path}: {e}")
    return None


def make_permanent_link(public_url, path):
    return f"{public_url}?path={quote(path)}"


# ============================================================
# РАЗБОР FB2
# ============================================================

def parse_fb2(content_bytes):
    """Возвращает {title, author, description, cover_data, cover_ext}."""
    result = {'title': '', 'author': '', 'description': '',
              'cover_data': None, 'cover_ext': ''}

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
            prefix = f'{{{ns}}}' if ns else ''
            a_el = ti.find(f'{prefix}author')
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


# ============================================================
# РАЗБОР TXT (аннотация из первых 30 КБ)
# ============================================================

_TXT_FIELD_AUTHOR = re.compile(
    r"^\s*(?:Автор|Author|АВТОР)\s*[:\-]\s*(.+?)\s*$",
    re.IGNORECASE | re.MULTILINE,
)
_TXT_FIELD_TITLE = re.compile(
    r"^\s*(?:Название|Title|НАЗВАНИЕ|Книга|Book)\s*[:\-]\s*(.+?)\s*$",
    re.IGNORECASE | re.MULTILINE,
)
_TXT_FIELD_ANNOT = re.compile(
    r"^\s*(?:Аннотация|Annotation|Описание|Description|"
    r"Аннотация книги|Краткое описание)\s*[:\-]\s*(.+?)(?=\n\s*\n|\Z)",
    re.IGNORECASE | re.MULTILINE | re.DOTALL,
)


def _decode_bytes(data):
    for enc in ("utf-8-sig", "utf-8", "cp1251", "koi8-r", "cp866"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="ignore")


def parse_txt(content_bytes):
    """Возвращает {title, author, description} — что удалось найти."""
    result = {'title': '', 'author': '', 'description': ''}

    head = content_bytes[:30_000]
    text = _decode_bytes(head)

    m = _TXT_FIELD_AUTHOR.search(text)
    if m:
        result['author'] = m.group(1).strip()[:300]

    m = _TXT_FIELD_TITLE.search(text)
    if m:
        result['title'] = m.group(1).strip()[:300]

    m = _TXT_FIELD_ANNOT.search(text)
    if m:
        desc = re.sub(r"\s+", " ", m.group(1)).strip()
        if len(desc) >= 30:
            if len(desc) > MAX_DESC_LEN:
                desc = desc[:MAX_DESC_LEN] + '...'
            result['description'] = desc

    return result


# ============================================================
# ОБОГАЩЕНИЕ ЧЕРЕЗ ВНЕШНИЕ ИСТОЧНИКИ
# ============================================================

def _fantlab_lookup(title, author):
    """Ищет описание через API FantLab."""
    if not title:
        return {}

    def _search(q, limit=5):
        try:
            r = requests.get(
                "https://api.fantlab.ru/search-works",
                params={"q": q, "onlymatches": 1},
                headers={"User-Agent": HEADERS["User-Agent"]},
                timeout=15,
            )
            if r.status_code != 200:
                return []
            data = r.json()
            if isinstance(data, dict):
                return (data.get("matches") or [])[:limit]
            if isinstance(data, list):
                return data[:limit]
        except Exception:
            pass
        return []

    query = title + (f" {author}" if author else "")
    matches = _search(query)
    time.sleep(SLEEP_BETWEEN_REQUESTS)
    if not matches:
        matches = _search(title)
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    if not matches:
        return {}

    # Выбираем наиболее релевантный
    title_l = (title or "").lower()
    surname = ""
    if author:
        parts = [p for p in re.split(r"[,\s]+", author.strip()) if p]
        if parts:
            surname = max(parts, key=len).lower()

    best = None
    best_score = -1
    for m in matches:
        score = 0
        names = " ".join(filter(None, [
            m.get("rusname", ""),
            m.get("name", ""),
            m.get("fullname", ""),
        ])).lower()
        if title_l and title_l in names:
            score += 3
        elif _title_phrase_matches(title, names):
            score += 2
        elif title_l and any(w in names for w in title_l.split() if len(w) > 3):
            score += 1
        if surname:
            authors_str = " ".join(filter(None, [
                m.get("all_autor_rusname", ""),
                m.get("autor1_rusname", ""),
                m.get("autor2_rusname", ""),
                m.get("autor3_rusname", ""),
            ])).lower()
            if surname in authors_str:
                score += 3
        if score > best_score:
            best_score = score
            best = m

    if not best or best_score < 2:
        return {}

    work_id = best.get("work_id")
    if not work_id:
        return {}

    try:
        r = requests.get(
            f"https://api.fantlab.ru/work/{work_id}/extended",
            headers={"User-Agent": HEADERS["User-Agent"]},
            timeout=15,
        )
        if r.status_code != 200:
            return {}
        work = r.json() or {}
    except Exception:
        return {}
    time.sleep(SLEEP_BETWEEN_REQUESTS)

    desc = work.get("work_description") or work.get("work_description_author") or ""
    if not desc:
        return {}
    desc = re.sub(r"<[^>]+>", "", desc)
    desc = re.sub(r"\[/?[a-zA-Z_]+\]", "", desc)
    desc = re.sub(r"\s+", " ", desc).strip()
    if len(desc) < 50:
        return {}
    if len(desc) > MAX_DESC_LEN:
        desc = desc[:MAX_DESC_LEN] + "..."

    cover = ""
    img = work.get("image") or {}
    if isinstance(img, dict):
        cover = img.get("url") or ""
        if cover and not cover.startswith("http"):
            cover = "https://fantlab.ru" + cover

    return {"description": desc, "cover": cover, "source": "fantlab"}


def _wikipedia_lookup(title, author):
    if not title:
        return {}

    def _request(params):
        try:
            r = requests.get(
                "https://ru.wikipedia.org/w/api.php",
                params=params,
                headers={"User-Agent": HEADERS["User-Agent"]},
                timeout=15,
            )
            if r.status_code != 200:
                return {}
            return r.json()
        except Exception:
            return {}

    def _search_page(q):
        data = _request({
            "action": "query", "format": "json", "list": "search",
            "srsearch": q, "srlimit": 1, "srnamespace": 0,
        })
        hits = (data.get("query") or {}).get("search") or []
        if not hits:
            return None, None
        return hits[0].get("pageid"), hits[0].get("title")

    def _extract(pageid):
        data = _request({
            "action": "query", "format": "json", "prop": "extracts",
            "pageids": pageid, "explaintext": 1, "exintro": 1, "redirects": 1,
        })
        pages = (data.get("query") or {}).get("pages") or {}
        for _, p in pages.items():
            t = (p.get("extract") or "").strip()
            if t:
                return t
        return ""

    queries = [
        f'"{title}" роман',
        f'"{title}" книга',
        f'"{title}" повесть',
        f'{title} (роман)',
        f'{title} (книга)',
        f'{title} (повесть)',
    ]
    if author:
        queries.append(f'{title} {author} роман')

    seen = set()
    for q in queries:
        pageid, page_title = _search_page(q)
        if not pageid or pageid in seen:
            time.sleep(SLEEP_BETWEEN_REQUESTS)
            continue
        seen.add(pageid)

        # отсеиваем страницы-биографии
        pt = (page_title or "").strip()
        if re.match(r"^[А-ЯЁ][а-яё]+\s*,\s*[А-ЯЁ][а-яё]+", pt):
            time.sleep(SLEEP_BETWEEN_REQUESTS)
            continue

        text = _extract(pageid)
        if not text or len(text) < 150:
            time.sleep(SLEEP_BETWEEN_REQUESTS)
            continue

        if not _title_phrase_matches(title, text):
            time.sleep(SLEEP_BETWEEN_REQUESTS)
            continue

        # берём первые 2–3 абзаца
        paras = [p.strip() for p in text.split("\n") if p.strip()][:3]
        desc = "\n\n".join(paras)[:MAX_DESC_LEN]
        if desc:
            return {"description": desc, "cover": "", "source": "wikipedia"}
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    return {}


def _google_books_lookup(title, author):
    if not title:
        return {}
    api_key = os.environ.get("GOOGLE_BOOKS_API_KEY", "")
    q = f'intitle:"{title}"'
    if author:
        q += f' inauthor:"{author}"'
    params = {"q": q, "maxResults": 3, "printType": "books"}
    if api_key:
        params["key"] = api_key
    try:
        r = requests.get("https://www.googleapis.com/books/v1/volumes",
                         params=params,
                         headers={"User-Agent": HEADERS["User-Agent"]},
                         timeout=15)
        if r.status_code != 200:
            return {}
        items = (r.json() or {}).get("items") or []
    except Exception:
        return {}
    for it in items:
        info = it.get("volumeInfo", {})
        desc = (info.get("description") or "").strip()
        if not desc:
            continue
        if len(desc) > MAX_DESC_LEN:
            desc = desc[:MAX_DESC_LEN] + "..."
        cover = (info.get("imageLinks") or {}).get("thumbnail", "")
        return {"description": desc, "cover": cover, "source": "google_books"}
    return {}


def enrich_book(title, author, cache):
    """
    Пытается получить описание у внешних источников.
    Возвращает {description, cover, source} или {}. Кэширует результат.
    """
    ck = hashlib.sha1(
        f"{(title or '').lower()}|{(author or '').lower()}".encode()
    ).hexdigest()

    if ck in cache:
        return cache[ck]

    meta = {}
    for fn in (_fantlab_lookup, _wikipedia_lookup, _google_books_lookup):
        try:
            res = fn(title, author)
            if res and res.get("description"):
                meta = res
                break
        except Exception:
            pass
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    cache[ck] = meta
    return meta


# ============================================================
# ЗАГРУЗКА ОБЛОЖЕК В SUPABASE
# ============================================================

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
        "apikey": SUPABASE_SERVICE_KEY,
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


# ============================================================
# GOOGLE SHEETS
# ============================================================

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


def _fetch_file_content(f, public_url, max_bytes=None):
    """
    Скачивает содержимое файла через публичный API Яндекс.Диска.
    Для fb2 нужен весь файл (в конце binary-обложки),
    для txt достаточно первых 30–50 КБ.
    """
    if not f['download_url']:
        f['download_url'] = get_download_url(public_url, f['path'])
    if not f['download_url']:
        return b""

    try:
        if max_bytes:
            r = requests.get(f['download_url'], headers=HEADERS,
                             timeout=60, stream=True)
            r.raise_for_status()
            buf = bytearray()
            try:
                for chunk in r.iter_content(chunk_size=16384):
                    if not chunk:
                        break
                    buf.extend(chunk)
                    if len(buf) >= max_bytes:
                        break
            finally:
                r.close()
            return bytes(buf)
        else:
            r = requests.get(f['download_url'], headers=HEADERS, timeout=60)
            if r.status_code == 200:
                return r.content
    except Exception as e:
        print(f"      Ошибка скачивания {f['name']}: {e}")
    return b""


# ============================================================
# ОБРАБОТКА ЛИСТА
# ============================================================

def process_sheet(worksheet, section_key, files, public_url, cache):
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
    enriched_count = 0

    for f in files:
        norm_name = normalize(f['name'])
        row_num = title_to_row.get(norm_name)
        if row_num is None:
            continue

        row = all_values[row_num - 1] if row_num - 1 < len(all_values) else []

        # ---- Нужно ли парсить файл ----
        need_parse = False
        if section_key == "books" and f['ext'] in PARSEABLE_EXTS:
            for field in ['author', 'description', 'cover']:
                if field in col_idx:
                    v = row[col_idx[field]] if col_idx[field] < len(row) else ''
                    if not v or not str(v).strip():
                        need_parse = True
                        break

        file_data = {}
        if need_parse:
            max_b = 30_000 if f['ext'] == '.txt' else None
            content = _fetch_file_content(f, public_url, max_bytes=max_b)
            if content:
                if f['ext'] == '.fb2':
                    file_data = parse_fb2(content)
                elif f['ext'] == '.txt':
                    file_data = parse_txt(content)

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

        # Формат
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
                if file_data.get('author'):
                    updates['author'] = file_data['author']
                elif f['folder']:
                    updates['author'] = f['folder'].split(' / ')[0]

        # Описание — из файла
        if 'description' in col_idx and file_data.get('description'):
            v = row[col_idx['description']] if col_idx['description'] < len(row) else ''
            if not str(v).strip():
                updates['description'] = file_data['description']

        # Обложка — из файла
        if 'cover' in col_idx and file_data.get('cover_data'):
            v = row[col_idx['cover']] if col_idx['cover'] < len(row) else ''
            if not str(v).strip():
                url = upload_cover_to_supabase(
                    file_data['cover_data'],
                    file_data.get('cover_ext', 'jpg'),
                    file_data.get('title') or f['name']
                )
                if url:
                    updates['cover'] = url

        # ---- ОБОГАЩЕНИЕ ЧЕРЕЗ ВНЕШНИЕ ИСТОЧНИКИ ----
        # Только для книг, и только если описания всё ещё нет.
        if section_key == "books" and 'description' in col_idx:
            has_desc_now = bool(updates.get('description')) or bool(
                (row[col_idx['description']] if col_idx['description'] < len(row)
                 else '').strip()
            )
            if not has_desc_now:
                book_title = os.path.splitext(f['name'])[0]
                book_author = (
                    updates.get('author')
                    or (row[col_idx['author']] if 'author' in col_idx and col_idx['author'] < len(row) else '')
                    or (f['folder'].split(' / ')[0] if f['folder'] else '')
                )
                ext_meta = enrich_book(book_title, book_author, cache)
                if ext_meta.get('description'):
                    updates['description'] = ext_meta['description']
                    enriched_count += 1
                if ('cover' in col_idx
                        and ext_meta.get('cover')
                        and not updates.get('cover')):
                    current_cover = row[col_idx['cover']] if col_idx['cover'] < len(row) else ''
                    if not str(current_cover).strip():
                        updates['cover'] = ext_meta['cover']

        for field, value in updates.items():
            cell = f"{col_num_to_letter(col_idx[field] + 1)}{row_num}"
            try:
                worksheet.update(cell, [[value]], value_input_option='RAW')
                updated_cells += 1
            except Exception as e:
                print(f"      Ошибка обновления {cell}: {e}")

    print(f"    Обновлено ячеек у существующих строк: {updated_cells}")
    print(f"    Обогащено из внешних источников: {enriched_count}")

    # ---- Новые записи ----
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
        file_data = {}

        if section_key == "books" and f['ext'] in PARSEABLE_EXTS:
            max_b = 30_000 if f['ext'] == '.txt' else None
            content = _fetch_file_content(f, public_url, max_bytes=max_b)
            if content:
                if f['ext'] == '.fb2':
                    file_data = parse_fb2(content)
                elif f['ext'] == '.txt':
                    file_data = parse_txt(content)
                if file_data.get('title'):
                    title = file_data['title']

        row_values = {
            'title': title,
            'folder': f['folder'],
            'download_link': make_permanent_link(public_url, f['path']),
        }

        if section_key == "books":
            row_values['format'] = f['ext'].lstrip('.') if f['ext'] else ''
            if file_data.get('author'):
                row_values['author'] = file_data['author']
            elif f['folder']:
                row_values['author'] = f['folder'].split(' / ')[0]
            if file_data.get('description'):
                row_values['description'] = file_data['description']
            if file_data.get('cover_data'):
                url = upload_cover_to_supabase(
                    file_data['cover_data'],
                    file_data.get('cover_ext', 'jpg'),
                    title
                )
                if url:
                    row_values['cover'] = url

            # Обогащение, если описания нет
            if not row_values.get('description'):
                ext_meta = enrich_book(
                    title, row_values.get('author', ''), cache
                )
                if ext_meta.get('description'):
                    row_values['description'] = ext_meta['description']
                    enriched_count += 1
                if ext_meta.get('cover') and not row_values.get('cover'):
                    row_values['cover'] = ext_meta['cover']

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


# ============================================================
# MAIN
# ============================================================

def main():
    print("Подключение к Google Sheets...")
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    cache = {}

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
            process_sheet(worksheet, section_key, files, public_url, cache)
        except Exception as e:
            print(f"Ошибка обработки листа '{section_ru}': {e}")

    print("\nГотово!")


if __name__ == "__main__":
    main()
