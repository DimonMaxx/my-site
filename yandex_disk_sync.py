# yandex_disk_sync.py
# Синхронизация Google Sheets с Яндекс.Диском.
#
# Логика:
#   Книги:
#     - Создаёт новые строки для новых файлов.
#     - Обогащает из FB2/TXT/DOCX + Google Books / OpenLibrary / Wikipedia / DuckDuckGo.
#     - Защита от дубликатов через existing_keys.
#   Программы:
#     - Создаёт новые строки для новых файлов (Название = имя файла).
#     - НЕ ищет в интернете, НЕ парсит содержимое.
#     - Заполняет только Папку, Размер, Формат, Ссылку.
#
# Кэш: _content/_enrichment_cache.json

import os
import re
import sys
import time
import json
import base64
import hashlib
import requests
import xml.etree.ElementTree as ET
from io import BytesIO
from urllib.parse import quote

from common import (
    SPREADSHEET_ID,
    SECTION_TO_SHEET,
    get_gspread_client,
    normalize,
)

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
    print("ПРЕДУПРЕЖДЕНИЕ: python-docx не установлен, .docx не будет парситься.")

# ========== НАСТРОЙКИ ==========
YANDEX_SOURCES = {
    'books':    'https://disk.yandex.ru/d/zMxF4nXHPkIVCQ',
    'programs': 'https://disk.yandex.ru/d/EjUHvm6mUcgVMw',
}

SUPABASE_URL = "https://rmoonebbvpmvthvpcmpt.supabase.co"
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
COVERS_BUCKET = "covers"

MAX_DESC_LEN = 2000

BOOK_EXTS = {'.fb2', '.epub', '.pdf', '.djvu', '.mobi', '.txt', '.doc', '.docx', '.rtf'}
PARSEABLE_EXTS = {'.fb2', '.txt', '.docx'}

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

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:120.0) Gecko/20100101 Firefox/120.0",
]

_ua_index = 0


def get_headers():
    global _ua_index
    _ua_index = (_ua_index + 1) % len(USER_AGENTS)
    return {"User-Agent": USER_AGENTS[_ua_index]}


# ========== Кэш ==========
CACHE_FILE = "_content/_enrichment_cache.json"
MAX_WEB_REQUESTS = 1500
WEB_SEARCH_DELAY = 0.25
WEB_REQUEST_TIMEOUT = 15

_web_requests_made = 0


def load_cache():
    if not os.path.exists(CACHE_FILE):
        return {}
    try:
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"  Ошибка чтения кэша: {e}")
        return {}


def save_cache(cache):
    try:
        os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False, indent=2, sort_keys=True)
        print(f"  Кэш сохранён: {len(cache)} записей")
    except Exception as e:
        print(f"  Ошибка сохранения кэша: {e}")


def cache_key(title, author=None):
    t = normalize(title or '')
    a = normalize(author or '')
    return f"{t}|{a}"


# ========== Утилиты ==========
def col_num_to_letter(n):
    result = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def decode_bytes(content_bytes):
    for enc in ('utf-8', 'windows-1251', 'koi8-r', 'iso-8859-5'):
        try:
            return content_bytes.decode(enc)
        except UnicodeDecodeError:
            continue
    return None


def is_title_clean(title):
    """Название «красивое» — без подчёркиваний, не имя файла, без маркеров серии."""
    if not title or len(title) < 2:
        return False
    if '_' in title:
        return False
    if title.count('.') > 3:
        return False
    if re.search(r'\b(том|книга|часть|выпуск|серия)\s*\d+', title, re.IGNORECASE):
        return False
    if re.search(r'\.(fb2|txt|docx?|rtf|pdf|epub|djvu|mobi)$', title, re.IGNORECASE):
        return False
    return True


# ========== Извлечение из имени файла ==========
def extract_from_filename(filename):
    base = os.path.splitext(filename)[0].strip()
    base = re.sub(r'\s*_OCR\s*$', '', base, flags=re.IGNORECASE)
    base = re.sub(r'\s*\(Si\)\s*$', '', base)
    base = re.sub(r'\s*\[СИ\]\s*$', '', base)
    base = re.sub(r'\s*\[СИ\]\.?\s*$', '', base)
    base = base.replace('_', ' ')
    base = re.sub(r'\s+', ' ', base).strip()

    author = ''
    title = base

    m = re.match(r'^([^\-—–]{2,60})\s*[-—–]\s*(.+)$', base)
    if m:
        author = m.group(1).strip()
        title = m.group(2).strip()
        title = re.sub(r'^Серия\s*[-—–]\s*', '', title).strip()
        title = re.sub(r'\s*\(серия\s*\d+\)\s*$', '', title, flags=re.IGNORECASE).strip()
        return {'author': author, 'title': title}

    m = re.match(r'^(.+?)\s*\(([^)]+)\)\s*$', base)
    if m:
        possible_author = m.group(2).strip()
        if len(possible_author.split()) <= 4 and possible_author[:1].isupper():
            if not re.match(r'^(том|книга|серия|выпуск|часть|\d)', possible_author, re.IGNORECASE):
                return {'author': possible_author, 'title': m.group(1).strip()}

    m = re.match(r'^([А-ЯЁA-Z][^.]{3,40})\.\s+(.+)$', base)
    if m:
        return {'author': m.group(1).strip(), 'title': m.group(2).strip()}

    m = re.match(r'^(.+?)\s*\.\s*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)\s*$', base)
    if m:
        return {'author': m.group(2).strip(), 'title': m.group(1).strip()}

    title = re.sub(r'\s+(том|книга|часть|выпуск)\s*\d+\s*$', '', title, flags=re.IGNORECASE).strip()
    return {'author': author, 'title': title}


# ========== TXT ==========
def extract_from_txt(content_bytes):
    result = {'author': '', 'title': '', 'description': ''}
    head = content_bytes[:8192]
    text = decode_bytes(head)
    if not text:
        return result

    m = re.search(r'^\s*Автор\s*[:\-]\s*(.+)$', text, re.IGNORECASE | re.MULTILINE)
    if m:
        result['author'] = m.group(1).strip()[:80]

    m = re.search(r'^\s*(Название|Заголовок)\s*[:\-]\s*(.+)$', text, re.IGNORECASE | re.MULTILINE)
    if m:
        result['title'] = m.group(2).strip()[:200]

    m = re.search(r'^\s*Аннотация\s*[:\-]\s*(.+?)(?:\n\n|\Z)', text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
    if m:
        result['description'] = re.sub(r'\s+', ' ', m.group(1)).strip()[:MAX_DESC_LEN]

    if not result['author'] or not result['title']:
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        if len(lines) >= 2:
            first = lines[0]
            second = lines[1]
            if not result['author'] and len(first) < 60 and not first.endswith('.'):
                result['author'] = first[:80]
            if not result['title'] and len(second) < 200:
                result['title'] = second[:200]

    return result


# ========== DOCX ==========
def extract_from_docx(content_bytes):
    result = {'author': '', 'title': '', 'description': ''}
    if not HAS_DOCX:
        return result

    try:
        doc = Document(BytesIO(content_bytes))
    except Exception as e:
        print(f"      Ошибка чтения DOCX: {e}")
        return result

    paragraphs = []
    for p in doc.paragraphs[:60]:
        t = p.text.strip()
        if t:
            paragraphs.append(t)
        if len(paragraphs) >= 30:
            break

    if not paragraphs:
        return result

    for line in paragraphs[:10]:
        m = re.match(r'^Автор\s*[:\-]\s*(.+)$', line, re.IGNORECASE)
        if m and not result['author']:
            result['author'] = m.group(1).strip()[:80]

        m = re.match(r'^(Название|Заголовок)\s*[:\-]\s*(.+)$', line, re.IGNORECASE)
        if m and not result['title']:
            result['title'] = m.group(2).strip()[:200]

        m = re.match(r'^Аннотация\s*[:\-]\s*(.+)$', line, re.IGNORECASE)
        if m and not result['description']:
            result['description'] = m.group(1).strip()[:MAX_DESC_LEN]

    if not result['title'] and paragraphs:
        result['title'] = paragraphs[0][:200]
    if not result['author'] and len(paragraphs) >= 2:
        second = paragraphs[1]
        if len(second) < 60 and not second.endswith('.'):
            result['author'] = second[:80]

    if not result['description']:
        for line in paragraphs[1:]:
            if len(line) > 100:
                result['description'] = line[:MAX_DESC_LEN]
                break

    return result


# ========== FB2 ==========
def parse_fb2(content_bytes):
    result = {'title': '', 'author': '', 'description': '', 'cover_data': None, 'cover_ext': ''}
    text = decode_bytes(content_bytes)
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


# ========== Поиск ==========
def _clean_search_title(title):
    t = title
    t = re.sub(r'\s*\([^)]*\)\s*$', '', t)
    t = re.sub(r'\s*\[[^\]]*\]\s*$', '', t)
    t = re.sub(r'\s*[.,;:]\s*$', '', t)
    t = re.sub(r'\s+(том|книга|часть|выпуск)\s*\d+\s*$', '', t, flags=re.IGNORECASE)
    return t.strip()


def search_google_books(title, author=None):
    global _web_requests_made
    if _web_requests_made >= MAX_WEB_REQUESTS:
        return None

    q = f'intitle:"{_clean_search_title(title)}"'
    if author:
        q += f'+inauthor:"{author}"'

    url = "https://www.googleapis.com/books/v1/volumes"
    params = {"q": q, "maxResults": 3, "printType": "books"}

    try:
        r = requests.get(url, params=params, headers=get_headers(), timeout=WEB_REQUEST_TIMEOUT)
        _web_requests_made += 1
    except Exception as e:
        print(f"      Google Books ошибка: {e}")
        return None

    if r.status_code != 200:
        return None

    data = r.json()
    items = data.get('items', [])
    if not items:
        if author:
            return search_google_books(title, author=None)
        return None

    vi = items[0].get('volumeInfo', {})

    cover_url = None
    img_links = vi.get('imageLinks', {})
    if img_links:
        cover_url = img_links.get('thumbnail') or img_links.get('smallThumbnail')
        if cover_url:
            cover_url = cover_url.replace('http://', 'https://')
            cover_url = re.sub(r'&zoom=\d+', '', cover_url)
            cover_url = re.sub(r'&edge=curl', '', cover_url)

    return {
        'title': vi.get('title', ''),
        'subtitle': vi.get('subtitle', ''),
        'author': ', '.join(vi.get('authors', [])[:3]),
        'description': vi.get('description', '')[:MAX_DESC_LEN],
        'cover_url': cover_url,
        'source': 'google_books',
    }


def search_openlibrary(title, author=None):
    global _web_requests_made
    if _web_requests_made >= MAX_WEB_REQUESTS:
        return None

    url = "https://openlibrary.org/search.json"
    params = {"title": _clean_search_title(title), "limit": 3}
    if author:
        params["author"] = author

    try:
        r = requests.get(url, params=params, headers=get_headers(), timeout=WEB_REQUEST_TIMEOUT)
        _web_requests_made += 1
    except Exception as e:
        print(f"      OpenLibrary ошибка: {e}")
        return None

    if r.status_code != 200:
        return None

    data = r.json()
    docs = data.get('docs', [])
    if not docs:
        if author:
            return search_openlibrary(title, author=None)
        return None

    doc = docs[0]
    cover_url = None
    if doc.get('cover_i'):
        cover_url = f"https://covers.openlibrary.org/b/id/{doc['cover_i']}-L.jpg"

    description = ''
    if doc.get('first_sentence'):
        fs = doc['first_sentence']
        if isinstance(fs, list):
            description = ' '.join(fs)[:MAX_DESC_LEN]
        else:
            description = str(fs)[:MAX_DESC_LEN]

    return {
        'title': doc.get('title', ''),
        'subtitle': '',
        'author': ', '.join(doc.get('author_name', [])[:3]),
        'description': description,
        'cover_url': cover_url,
        'source': 'openlibrary',
    }


def search_wikipedia(title, author=None):
    global _web_requests_made
    if _web_requests_made >= MAX_WEB_REQUESTS:
        return None

    clean_title = _clean_search_title(title)
    results = []
    for lang in ('ru', 'en'):
        if _web_requests_made >= MAX_WEB_REQUESTS:
            break
        search_q = clean_title
        if author:
            search_q += f' {author}'
        try:
            sr = requests.get(
                f"https://{lang}.wikipedia.org/w/api.php",
                params={
                    "action": "opensearch",
                    "search": search_q,
                    "limit": 3,
                    "format": "json",
                },
                headers=get_headers(),
                timeout=WEB_REQUEST_TIMEOUT,
            )
            _web_requests_made += 1
            if sr.status_code == 200:
                sr_data = sr.json()
                if len(sr_data) >= 2:
                    for page_title in sr_data[1][:2]:
                        results.append((lang, page_title))
        except Exception as e:
            print(f"      Wikipedia ({lang}) ошибка поиска: {e}")

    for lang, page_title in results:
        if _web_requests_made >= MAX_WEB_REQUESTS:
            break
        try:
            summ_url = f"https://{lang}.wikipedia.org/api/rest_v1/page/summary/{quote(page_title)}"
            r = requests.get(summ_url, headers=get_headers(), timeout=WEB_REQUEST_TIMEOUT)
            _web_requests_made += 1
            if r.status_code != 200:
                continue
            d = r.json()
            extract = d.get('extract', '')
            if not extract or len(extract) < 50:
                continue
            page_title_clean = re.sub(r'\s*\([^)]*\)\s*$', '', d.get('title', page_title))
            desc = d.get('description', '')
            author_found = ''
            if desc:
                m = re.search(r'(?:роман|повесть|книга|произведение)\s+([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+)+)', desc)
                if m:
                    author_found = m.group(1).strip()
            return {
                'title': page_title_clean,
                'subtitle': '',
                'author': author_found,
                'description': extract[:MAX_DESC_LEN],
                'cover_url': None,
                'source': f'wikipedia_{lang}',
            }
        except Exception as e:
            print(f"      Wikipedia ({lang}) ошибка summary: {e}")

    return None


def search_duckduckgo(title, author=None):
    global _web_requests_made
    if _web_requests_made >= MAX_WEB_REQUESTS:
        return None

    clean_title = _clean_search_title(title)
    query = clean_title
    if author:
        query += f' {author}'
    query += ' книга'

    try:
        r = requests.get(
            "https://api.duckduckgo.com/",
            params={
                "q": query,
                "format": "json",
                "no_html": 1,
                "no_redirect": 1,
            },
            headers=get_headers(),
            timeout=WEB_REQUEST_TIMEOUT,
        )
        _web_requests_made += 1
    except Exception as e:
        print(f"      DuckDuckGo ошибка: {e}")
        return None

    if r.status_code != 200:
        return None

    d = r.json()
    abstract = d.get('AbstractText', '')
    if not abstract or len(abstract) < 50:
        return None

    heading = d.get('Heading', '')
    return {
        'title': heading or clean_title,
        'subtitle': '',
        'author': '',
        'description': abstract[:MAX_DESC_LEN],
        'cover_url': None,
        'source': 'duckduckgo',
    }


def enrich_from_web(title, author, cache):
    key = cache_key(title, author)
    if key in cache:
        return cache[key]

    key_no_author = cache_key(title, '')
    if key_no_author in cache:
        return cache[key_no_author]

    result = {}

    result = search_google_books(title, author) or {}
    if not result or not result.get('title'):
        r2 = search_openlibrary(title, author) or {}
        if r2.get('title'):
            result = r2

    need_more = (
        not result
        or not result.get('author')
        or not result.get('description')
        or len(result.get('description', '')) < 100
    )
    if need_more:
        wiki = search_wikipedia(title, author) or {}
        if wiki:
            if not result:
                result = wiki
            else:
                if not result.get('author') and wiki.get('author'):
                    result['author'] = wiki['author']
                if not result.get('description') and wiki.get('description'):
                    result['description'] = wiki['description']
                elif wiki.get('description') and len(wiki['description']) > len(result.get('description', '')):
                    result['description'] = wiki['description']
                if not result.get('title') and wiki.get('title'):
                    result['title'] = wiki['title']

    if not result or not result.get('description'):
        ddg = search_duckduckgo(title, author) or {}
        if ddg:
            if not result:
                result = ddg
            else:
                if not result.get('description') and ddg.get('description'):
                    result['description'] = ddg['description']
                if not result.get('title') and ddg.get('title'):
                    result['title'] = ddg['title']

    if not result:
        result = {}

    cache[key] = result
    cache[key_no_author] = result

    time.sleep(WEB_SEARCH_DELAY)
    return result


# ========== Обложки ==========
def upload_cover_bytes_to_supabase(cover_data, ext, key):
    if not SUPABASE_SERVICE_KEY:
        return None
    filename = f"{hashlib.md5(key.encode('utf-8')).hexdigest()}.{ext}"
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
        print(f"  Ошибка загрузки обложки: {resp.status_code} — {resp.text[:200]}")
    except Exception as e:
        print(f"  Ошибка загрузки обложки: {e}")
    return None


def download_and_upload_cover(cover_url, key):
    if not cover_url:
        return None
    try:
        r = requests.get(cover_url, headers=get_headers(), timeout=20)
        if r.status_code != 200:
            return None
        data = r.content
        if len(data) < 500:
            return None
        ext = 'jpg'
        ct = r.headers.get('Content-Type', '').lower()
        if 'png' in ct:
            ext = 'png'
        elif 'webp' in ct:
            ext = 'webp'
        elif 'gif' in ct:
            ext = 'gif'
        return upload_cover_bytes_to_supabase(data, ext, key)
    except Exception as e:
        print(f"      Ошибка скачивания обложки: {e}")
        return None


# ========== Яндекс.Диск ==========
def list_yandex_recursive(public_url, sub_path=None):
    api_url = "https://cloud-api.yandex.net/v1/disk/public/resources"
    params = {"public_key": public_url, "limit": 1000, "sort": "name"}
    if sub_path:
        params["path"] = sub_path

    try:
        resp = requests.get(api_url, params=params, headers=get_headers(), timeout=30)
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
    dl_api = "https://cloud-api.yandex.net/v1/disk/public/resources/download"
    params = {"public_key": public_url, "path": path}
    try:
        resp = requests.get(dl_api, params=params, headers=get_headers(), timeout=30)
        if resp.status_code == 200:
            return resp.json().get('href')
    except Exception as e:
        print(f"  Ошибка download URL для {path}: {e}")
    return None


def make_permanent_link(public_url, path):
    return f"{public_url}?path={quote(path)}"


# ========== Ключи файла ==========
def get_all_file_keys(f):
    """
    Возвращает множество нормализованных ключей для файла:
    - имя файла без расширения
    - название, извлечённое из имени
    - вариант с автором из имени
    """
    keys = set()

    base = os.path.splitext(f['name'])[0]
    keys.add(normalize(base))

    parsed = extract_from_filename(f['name'])
    if parsed.get('title'):
        keys.add(normalize(parsed['title']))

    # Вариант "Автор Название" (слитно) — для случаев, когда в таблице стоит полное имя файла
    if parsed.get('author') and parsed.get('title'):
        keys.add(normalize(f"{parsed['author']} {parsed['title']}"))

    return keys


# ========== Обогащение книги ==========
def enrich_book_file(f, public_url, cache):
    """Полное обогащение для книги: имя файла + содержимое + интернет."""
    result = {
        'title': '',
        'author': '',
        'description': '',
        'cover_url': '',
        'format': f['ext'].lstrip('.') if f['ext'] else '',
    }

    # 1. Из имени файла
    from_name = extract_from_filename(f['name'])
    if from_name.get('title'):
        result['title'] = from_name['title']
    if from_name.get('author'):
        result['author'] = from_name['author']

    # 2. Из содержимого
    file_data = None
    if f['ext'] in PARSEABLE_EXTS:
        if not f['download_url']:
            f['download_url'] = get_download_url(public_url, f['path'])
        if f['download_url']:
            try:
                r = requests.get(f['download_url'], timeout=60)
                if r.status_code == 200:
                    file_data = r.content
            except Exception as e:
                print(f"      Ошибка скачивания {f['name']}: {e}")

    if file_data:
        if f['ext'] == '.fb2':
            parsed = parse_fb2(file_data)
            if parsed.get('title'):
                result['title'] = parsed['title']
            if parsed.get('author'):
                result['author'] = parsed['author']
            if parsed.get('description'):
                result['description'] = parsed['description']
            if parsed.get('cover_data'):
                cover_ext = parsed.get('cover_ext', 'jpg')
                url = upload_cover_bytes_to_supabase(
                    parsed['cover_data'], cover_ext, result['title'] or f['name']
                )
                if url:
                    result['cover_url'] = url
        elif f['ext'] == '.txt':
            parsed = extract_from_txt(file_data)
            if parsed.get('title'):
                result['title'] = parsed['title'] or result['title']
            if parsed.get('author'):
                result['author'] = parsed['author'] or result['author']
            if parsed.get('description'):
                result['description'] = parsed['description']
        elif f['ext'] == '.docx':
            parsed = extract_from_docx(file_data)
            if parsed.get('title'):
                result['title'] = parsed['title'] or result['title']
            if parsed.get('author'):
                result['author'] = parsed['author'] or result['author']
            if parsed.get('description'):
                result['description'] = parsed['description']

    # 3. Из интернета
    author_is_folder = False
    if f['folder'] and result['author']:
        if normalize(result['author']) == normalize(f['folder'].split(' / ')[0]):
            author_is_folder = True

    need_web = (
        not is_title_clean(result['title'])
        or not result['author']
        or author_is_folder
        or not result['description']
    )

    if need_web and result['title']:
        search_author = result['author']
        if author_is_folder:
            search_author = None

        web_data = enrich_from_web(result['title'], search_author, cache)
        if web_data:
            web_title = web_data.get('title', '')
            if web_title:
                if not result['title'] or not is_title_clean(result['title']):
                    result['title'] = web_title
                elif len(web_title) > len(result['title']) and not is_title_clean(result['title']):
                    result['title'] = web_title

            web_author = web_data.get('author', '')
            if web_author and (not result['author'] or author_is_folder):
                result['author'] = web_author
            elif web_author and not author_is_folder and len(web_author) > len(result['author']):
                result['author'] = web_author

            if web_data.get('description') and not result['description']:
                result['description'] = web_data['description']

            if not result['cover_url'] and web_data.get('cover_url'):
                url = download_and_upload_cover(web_data['cover_url'], result['title'])
                if url:
                    result['cover_url'] = url

    # Финальная очистка
    if result['title']:
        result['title'] = re.sub(r'\s+', ' ', result['title']).strip()
        result['title'] = result['title'].rstrip('.')

    return result


def enrich_program_file(f):
    """
    Для программ — только имя файла + формат + размер.
    Никакого парсинга содержимого и поиска в интернете.
    """
    title = os.path.splitext(f['name'])[0]
    title = title.replace('_', ' ').strip()
    title = re.sub(r'\s+', ' ', title).strip()

    return {
        'title': title,
        'author': '',
        'description': '',
        'cover_url': '',
        'format': f['ext'].lstrip('.') if f['ext'] else '',
    }


# ========== Работа с листом ==========
def ensure_columns(worksheet, headers, required_fields):
    headers = list(headers)
    for field in required_fields:
        ru = RU_LABELS.get(field)
        if ru and ru not in headers:
            col_idx = len(headers) + 1
            col_letter = col_num_to_letter(col_idx)
            worksheet.update(values=[[ru]], range_name=f'{col_letter}1', value_input_option='RAW')
            headers.append(ru)
            print(f"    + создана колонка '{ru}' (позиция {col_letter})")
    return headers


def build_indexes(all_values, col_idx):
    """
    Строит:
    - existing_keys: множество всех нормализованных ключей (названия из таблицы)
    - title_to_row: ключ → номер строки
    """
    existing_keys = set()
    title_to_row = {}

    for i, row in enumerate(all_values[1:], start=2):
        if col_idx['title'] < len(row):
            n = normalize(row[col_idx['title']])
            if n:
                existing_keys.add(n)
                if n not in title_to_row:
                    title_to_row[n] = i
    return existing_keys, title_to_row


def update_row(worksheet, row_num, col_idx, updates):
    """Применяет обновления к строке. Возвращает число обновлённых ячеек."""
    count = 0
    for field, value in updates.items():
        if field not in col_idx:
            continue
        cell = f"{col_num_to_letter(col_idx[field] + 1)}{row_num}"
        try:
            worksheet.update(values=[[value]], range_name=cell, value_input_option='RAW')
            count += 1
        except Exception as e:
            print(f"      Ошибка обновления {cell}: {e}")
    return count


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

    # Индексы
    existing_keys, title_to_row = build_indexes(all_values, col_idx)

    # Дополнительно добавляем в existing_keys все "варианты" из имён файлов,
    # которые уже есть в таблице — по названию из строки
    # (на случай, если название в таблице = «красивое», а имя файла — «грязное»)

    updated_cells = 0

    # ---------- Проход 1: обновляем существующие строки ----------
    for f in files:
        # Ищем строку по любому из ключей файла
        row_num = None
        for k in get_all_file_keys(f):
            if k in title_to_row:
                row_num = title_to_row[k]
                break

        if row_num is None:
            continue

        row = all_values[row_num - 1] if row_num - 1 < len(all_values) else []

        current_author = row[col_idx['author']] if 'author' in col_idx and col_idx['author'] < len(row) else ''
        current_desc = row[col_idx['description']] if 'description' in col_idx and col_idx['description'] < len(row) else ''
        current_cover = row[col_idx['cover']] if 'cover' in col_idx and col_idx['cover'] < len(row) else ''
        current_title = row[col_idx['title']] if col_idx['title'] < len(row) else ''

        author_is_folder = False
        if f['folder'] and current_author and normalize(current_author) == normalize(f['folder'].split(' / ')[0]):
            author_is_folder = True

        title_is_clean = is_title_clean(current_title)

        # Для программ — обогащаем только ссылку/папку/размер/формат
        if section_key == 'programs':
            data = enrich_program_file(f)
            need_enrich = (
                not str(row[col_idx.get('folder', 0)] if 'folder' in col_idx and col_idx['folder'] < len(row) else '').strip()
                or not str(row[col_idx.get('size', 0)] if 'size' in col_idx and col_idx['size'] < len(row) else '').strip()
                or not str(row[col_idx.get('download_link', 0)] if 'download_link' in col_idx and col_idx['download_link'] < len(row) else '').strip()
            )
        else:
            # Книги и другие разделы
            data = None  # посчитаем позже, если понадобится
            need_enrich = (
                not str(current_author).strip()
                or not str(current_desc).strip()
                or not str(current_cover).strip()
                or author_is_folder
                or not title_is_clean
            )

        if not need_enrich:
            continue

        # Для книг — полное обогащение
        if section_key != 'programs':
            data = enrich_book_file(f, public_url, cache)

        updates = {}

        if 'folder' in col_idx and f['folder']:
            v = row[col_idx['folder']] if col_idx['folder'] < len(row) else ''
            if not str(v).strip():
                updates['folder'] = f['folder']

        if 'download_link' in col_idx:
            v = row[col_idx['download_link']] if col_idx['download_link'] < len(row) else ''
            if not str(v).strip():
                updates['download_link'] = make_permanent_link(public_url, f['path'])

        if 'format' in col_idx and data['format']:
            v = row[col_idx['format']] if col_idx['format'] < len(row) else ''
            if not str(v).strip():
                updates['format'] = data['format']

        if 'size' in col_idx and f['size'] > 0:
            v = row[col_idx['size']] if col_idx['size'] < len(row) else ''
            if not str(v).strip():
                updates['size'] = str(round(f['size'] / 1024 / 1024, 1))

        if section_key != 'programs':
            # Автор
            if 'author' in col_idx and data['author']:
                if not str(current_author).strip() or author_is_folder:
                    updates['author'] = data['author']

            # Название
            if 'title' in col_idx and data['title']:
                if not title_is_clean:
                    if normalize(data['title']) != normalize(current_title):
                        updates['title'] = data['title']
                elif len(data['title']) > len(current_title) + 5:
                    if normalize(data['title']) != normalize(current_title):
                        updates['title'] = data['title']

            # Описание
            if 'description' in col_idx and data['description'] and not str(current_desc).strip():
                updates['description'] = data['description']

            # Обложка
            if 'cover' in col_idx and data['cover_url'] and not str(current_cover).strip():
                updates['cover'] = data['cover_url']

        # Применяем
        updated_cells += update_row(worksheet, row_num, col_idx, updates)

        # Если название обновилось — добавим новый ключ в existing_keys и title_to_row
        if 'title' in updates:
            new_norm = normalize(updates['title'])
            existing_keys.add(new_norm)
            if new_norm not in title_to_row:
                title_to_row[new_norm] = row_num
            # Также добавим все варианты старого названия
            existing_keys.add(normalize(current_title))

    print(f"    Обновлено ячеек у существующих строк: {updated_cells}")

    # ---------- Проход 2: новые строки ----------
    new_rows = []
    new_keys_local = set()

    for f in files:
        file_keys = get_all_file_keys(f)
        # Если хотя бы один ключ файла уже известен — пропускаем
        if any(k in existing_keys for k in file_keys):
            continue
        if any(k in new_keys_local for k in file_keys):
            continue
        new_keys_local.update(file_keys)

        if section_key == 'programs':
            data = enrich_program_file(f)
        else:
            data = enrich_book_file(f, public_url, cache)

        row_values = {
            'title': data['title'] or os.path.splitext(f['name'])[0],
            'folder': f['folder'],
            'download_link': make_permanent_link(public_url, f['path']),
        }
        if section_key == 'books':
            row_values['format'] = data['format']
            if data['author']:
                row_values['author'] = data['author']
            elif f['folder']:
                row_values['author'] = f['folder'].split(' / ')[0]
            if data['description']:
                row_values['description'] = data['description']
            if data['cover_url']:
                row_values['cover'] = data['cover_url']
        elif section_key == 'programs':
            if f['size'] > 0:
                row_values['size'] = str(round(f['size'] / 1024 / 1024, 1))

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


# ========== Main ==========
def main():
    global _web_requests_made

    print("Подключение к Google Sheets...")
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    print("Загрузка кэша обогащения...")
    cache = load_cache()
    print(f"  Записей в кэше: {len(cache)}")

    try:
        for section_key, public_url in YANDEX_SOURCES.items():
            section_ru = SECTION_TO_SHEET.get(section_key)
            if not section_ru:
                print(f"\nРаздел '{section_key}' не найден, пропускаем.")
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

            try:
                worksheet = sh.worksheet(section_ru)
            except Exception:
                print(f"Лист '{section_ru}' не найден, пропускаем.")
                continue

            try:
                process_sheet(worksheet, section_key, files, public_url, cache)
            except Exception as e:
                print(f"Ошибка обработки листа '{section_ru}': {e}")

    finally:
        print("\nСохранение кэша обогащения...")
        save_cache(cache)
        print(f"Веб-запросов сделано за прогон: {_web_requests_made}")

    print("\nГотово!")


if __name__ == "__main__":
    main()
