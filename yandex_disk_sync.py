#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
yandex_disk_sync.py
Синхронизация Яндекс.Диска → Google Sheets.

Разделы и их настройки читаются из таблицы site_sections (Supabase),
поэтому добавлять/менять разделы можно через админ-панель без правки Python.

Режимы ускорения:
  • SYNC_SECTIONS=programs,music   — синхронизировать только перечисленные
    разделы (через запятую). Пусто → все активные разделы.
  • SKIP_ENRICHMENT=1              — не ходить во внешние API (Fantlab,
    Wikipedia, OpenLibrary, Google Books, LLM). Используются только данные
    из fb2/txt. Сильно ускоряет книги при отладке.

Другие фичи:
  • PRESERVE_USER_EDITS — сохраняет ручные правки пользователя
    (название, описание, автор и т.п.) при обновлении существующих строк.
  • manual_override (per-section) — существующие строки раздела вообще
    не перезаписываются, добавляются только новые файлы.
  • strip_prefix_mode (per-section: auto | always | never) — управляет
    отрезанием yandex_path от путей файлов.
  • BACKUP_BEFORE_SYNC — перед записью изменений создаётся резервная
    копия листа (_backup_<section>_<timestamp>), хранятся последние N.
  • Детект «осиротевших» строк (файлы исчезли с Диска, строки остались
    в Sheets) с сохранением в Supabase-таблицу sync_orphans,
    включая полные значения строк и заголовки (для diff-view).
  • ensure_headers полностью перезаписывает первую строку до максимальной
    ширины листа — устраняет «хвосты» старых колонок и дубликаты
    заголовков.
"""

import os
import io
import re
import json
import time
import base64
import hashlib
import datetime
import traceback
import requests
import xml.etree.ElementTree as ET
from urllib.parse import quote

try:
    import yadisk
except ImportError:
    yadisk = None

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    gspread = None
    Credentials = None

try:
    from supabase import create_client as supa_create_client
except ImportError:
    supa_create_client = None


# ============================================================
# КОНФИГУРАЦИЯ
# ============================================================

SPREADSHEET_ID   = os.environ.get("SPREADSHEET_ID", "")
SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME", "НаполнениеСайта")

SUPABASE_URL = "https://rmoonebbvpmvthvpcmpt.supabase.co"
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
COVERS_BUCKET = "covers"

# Общие списки расширений
ALLOWED_EXTS  = {".fb2", ".epub", ".pdf", ".djvu", ".mobi", ".txt",
                 ".doc", ".docx", ".rtf"}
PROGRAM_EXTS  = {".rar", ".zip", ".7z", ".xlsm", ".xlsx", ".xls",
                 ".ods", ".odt", ".exe", ".msi", ".bat", ".ps1",
                 ".py", ".sh"}
MUSIC_EXTS    = {".mp3", ".flac", ".wav", ".ogg", ".m4a", ".aac",
                 ".wma", ".opus", ".ape", ".aiff", ".alac"}
PARSEABLE_EXTS = {".fb2", ".txt"}

# Внешние источники обогащения
GOOGLE_BOOKS_API = "https://www.googleapis.com/books/v1/volumes"
FANTLAB_SEARCH   = "https://api.fantlab.ru/search-works"
FANTLAB_WORK     = "https://api.fantlab.ru/work/{id}/extended"
WIKI_API         = f"https://{os.environ.get('WIKI_LANG', 'ru')}.wikipedia.org/w/api.php"
OPENLIBRARY_API  = "https://openlibrary.org/search.json"
YANDEX_GPT_URL   = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion"
OPENAI_URL       = "https://api.openai.com/v1/chat/completions"

HTTP_TIMEOUT    = 30
SLEEP_BETWEEN   = 0.4
SLEEP_BEFORE_LISTDIR = 1.5

MAX_DESC_LEN    = 2000
MAX_TXT_HEAD    = 30_000

USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/120.0 Safari/537.36")
HEADERS = {"User-Agent": USER_AGENT}

BOOK_MARKERS = (
    "книга", "роман", "повесть", "рассказ", "произведение",
    "сборник", "трилогия", "эпопея", "цикл", "литератур",
    "novel", "book", "story",
)

_CP866_JUNK = set("╞░─┘╦╪╟┌┐└┴┬├┤│╫╬═║╔╗╚╝")

RU_TO_EN = {
    "Название":              "title",
    "Автор":                 "author",
    "Исполнитель":           "artist",
    "Год":                   "year",
    "Описание":              "description",
    "Формат":                "format",
    "Размер (МБ)":           "size",
    "Ссылка для скачивания": "download_link",
    "Обложка":               "cover",
    "Папка":                 "folder",
    "Версия":                "version",
    "Платформа":             "platform",
}

# Соответствие ключа колонки (как в Supabase) → русский заголовок
EN_TO_RU = {v: k for k, v in RU_TO_EN.items()}


# ============================================================
# РЕЖИМЫ СИНХРОНИЗАЦИИ
# ============================================================

# PRESERVE_USER_EDITS=1 (по умолчанию):
#   при обновлении существующей строки непустые ячейки сохраняются
#   (ручные правки), пустые — заполняются из Диска.
# PRESERVE_USER_EDITS=0:
#   строка перезаписывается целиком (старое поведение).
PRESERVE_USER_EDITS = os.environ.get("PRESERVE_USER_EDITS", "1") == "1"

# BACKUP_BEFORE_SYNC=1 (по умолчанию): перед записью изменений
#   создаётся резервная копия листа.
BACKUP_BEFORE_SYNC  = os.environ.get("BACKUP_BEFORE_SYNC", "1") == "1"

# Сколько последних резервных копий хранить для каждого раздела.
BACKUP_KEEP_COUNT   = int(os.environ.get("BACKUP_KEEP_COUNT", "3"))

# STRIP_PREFIX_MODE — глобальный fallback, если у раздела нет
# собственного strip_prefix_mode:
#   'always' — всегда отрезать yandex_path от путей файлов.
#   'never'  — не отрезать.
#   'auto'   — старая логика через is_public_root_link.
STRIP_PREFIX_MODE_DEFAULT = os.environ.get("STRIP_PREFIX_MODE", "always").lower()

# SYNC_SECTIONS — список ключей разделов через запятую.
# Пусто → синхронизировать все активные разделы.
_raw_sync_sections = os.environ.get("SYNC_SECTIONS", "").strip()
SYNC_SECTIONS_FILTER = [
    s.strip() for s in _raw_sync_sections.split(",") if s.strip()
] if _raw_sync_sections else []

# SKIP_ENRICHMENT=1 — не ходить во внешние API (Fantlab, Wikipedia,
# OpenLibrary, Google Books, LLM). Использовать только данные fb2/txt.
SKIP_ENRICHMENT = os.environ.get("SKIP_ENRICHMENT", "0") == "1"

# Таблица Supabase, в которую пишутся «осиротевшие» строки.
SYNC_ORPHANS_TABLE  = "sync_orphans"

# Колонки, которые ВСЕГДА перезаписываются данными из Яндекс.Диска:
# их значения вычисляются из файла, ручные правки бессмысленны.
ALWAYS_UPDATE_HEADERS = {
    "Ссылка для скачивания", "Ссылка",
    "Размер (МБ)", "Размер",
    "Формат",
    "download_link", "link", "size", "format",
}

# Кандидаты имён колонки «Название» — для извлечения title
# при формировании списка осиротевших строк.
TITLE_HEADER_CANDIDATES = ("Название", "название", "Title", "title")


# ============================================================
# ДИАГНОСТИКА
# ============================================================

class Diag:
    def __init__(self, section_label=""):
        self.section_label = section_label
        self.found     = 0
        self.kept      = 0
        self.updated   = 0
        self.added     = 0
        self.from_file = 0
        self.sources   = {}
        self.skipped   = {"bad_ext": 0, "garbage": 0, "wrong_ext": 0,
                          "dup_link": 0}
        self.skip_samples = []
        # Дополнительные счётчики
        self.preserved_edits = 0
        self.unchanged_rows  = 0
        self.orphans_found   = 0
        self.backup_name     = ""

    def report(self):
        print("\n  ── ДИАГНОСТИКА ──")
        print(f"  Найдено файлов:        {self.found}")
        print(f"  Оставлено к записи:    {self.kept}")
        print(f"  Обновлено строк:       {self.updated}")
        print(f"  Добавлено строк:       {self.added}")
        print(f"  (описаний из файлов):  {self.from_file}")
        if self.preserved_edits:
            print(f"  Сохранено ручных правок: {self.preserved_edits}")
        if self.unchanged_rows:
            print(f"  Без изменений:         {self.unchanged_rows}")
        if self.orphans_found:
            print(f"  Осиротевших строк:     {self.orphans_found}")
        if self.backup_name:
            print(f"  Резервная копия:       {self.backup_name}")
        if self.sources:
            print("  Источники описаний:")
            for src, cnt in sorted(self.sources.items(), key=lambda x: -x[1]):
                print(f"    • {src}: {cnt}")
        total = sum(self.skipped.values())
        if total:
            print(f"  Пропущено:             {total}")
            for k, v in self.skipped.items():
                if v:
                    print(f"    • {k}: {v}")
            for s in self.skip_samples[:15]:
                print(f"    - {s}")
        print("  ───────────────────\n")


def log_skip(diag, reason, name):
    diag.skipped[reason] = diag.skipped.get(reason, 0) + 1
    if len(diag.skip_samples) < 30:
        diag.skip_samples.append(f"[{reason}] {name}")


# ============================================================
# УТИЛИТЫ СОПОСТАВЛЕНИЯ
# ============================================================

_STOP_WORDS = {"или", "как", "для", "при", "над", "под", "без", "про",
               "the", "and", "for", "with", "from", "that", "this",
               "его", "её", "ее", "их", "все", "весь", "себя", "это"}


def _normalize_for_match(text):
    if not text:
        return ""
    t = str(text).lower().replace("ё", "е")
    t = re.sub(r"[^\w\s]", " ", t, flags=re.UNICODE)
    return re.sub(r"\s+", " ", t).strip()


def _author_surname(author):
    if not author:
        return ""
    parts = [p for p in re.split(r"[,\s]+", author.strip()) if p]
    if not parts:
        return ""
    surname = max(parts, key=len).lower().replace("ё", "е")
    return surname if len(surname) > 3 else ""


def _title_phrase_matches(title, text):
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
# ПАРСИНГ ИМЕНИ ФАЙЛА
# ============================================================

SEPARATORS = [" - ", " — ", " – ", " –– "]


def _is_garbage_stem(stem):
    if not stem:
        return True
    s = stem.strip()
    if re.match(r"^fanfic_\d+$", s, re.IGNORECASE):
        return True
    if re.match(r"^[\d\.\s]+$", s):
        return True
    if re.match(r"^[\*\-_\.\s]+$", s):
        return True
    if any(ch in s for ch in _CP866_JUNK):
        return True
    if len(s) < 2:
        return True
    if not re.search(r"[A-Za-zА-Яа-яЁё]{2,}", s):
        return True
    return False


def parse_book_name(filename):
    stem = os.path.splitext(filename)[0].strip()
    if _is_garbage_stem(stem):
        return {"title": "", "author": ""}
    stem = re.sub(r"\s*\(\d+\)\s*$", "", stem)
    stem = re.sub(r"\s*\[.*?\]\s*", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip()

    author = ""
    title  = stem
    for sep in SEPARATORS:
        if sep in stem:
            head, tail = stem.split(sep, 1)
            head, tail = head.strip(), tail.strip()
            if head and tail:
                author = head
                title  = tail
                break
    if not author and "." in title:
        m = re.match(r"^([А-ЯA-Z][^.]{1,60}\.)\s+(.+)$", title)
        if m:
            author = m.group(1).strip()
            title  = m.group(2).strip()
    author = re.sub(r"\s+", " ", author).strip(" ,;")
    title  = re.sub(r"\s+", " ", title).strip(" ,;")
    return {"title": title, "author": author}


def parse_program_name(filename):
    stem = os.path.splitext(filename)[0].strip()
    stem = re.sub(r"\s+", " ", stem).strip()
    if not stem:
        return {"title": "", "version": ""}
    version = ""
    m = re.search(r"\b[vV]?(\d+(?:\.\d+){1,3})\b", stem)
    if m:
        version = m.group(1)
    return {"title": stem, "version": version}


def parse_music_name(filename):
    stem = os.path.splitext(filename)[0].strip()
    stem = re.sub(r"\s+", " ", stem).strip()
    result = {"title": stem, "artist": "", "album": "", "year": ""}
    if not stem:
        return result

    year_m = re.search(r"\b(19[5-9]\d|20[0-3]\d)\b", stem)
    if year_m:
        result["year"] = year_m.group(1)

    cleaned = re.sub(r"^\d{1,2}[\s\.\-–]+", "", stem).strip()
    parts = re.split(r"\s+[-–—]\s+", cleaned, maxsplit=1)
    if len(parts) == 2 and parts[0].strip() and parts[1].strip():
        result["artist"] = parts[0].strip()
        result["title"]  = parts[1].strip()
    else:
        result["title"] = cleaned

    result["title"] = re.sub(
        r"\s*\(?\b(19[5-9]\d|20[0-3]\d)\b\)?\s*", " ",
        result["title"]
    ).strip(" -–—")
    return result


def parse_universal_name(filename):
    stem = os.path.splitext(filename)[0].strip()
    stem = re.sub(r"\s+", " ", stem).strip()
    return {"title": stem}


# ============================================================
# ЯНДЕКС.ДИСК
# ============================================================

def get_public_key(public_url):
    m = re.search(r"/d/([A-Za-z0-9_-]+)", public_url)
    if not m:
        raise ValueError(f"Не удалось извлечь public_key из {public_url}")
    return m.group(1)


def _strip_disk_prefix(path):
    if path and path.startswith("disk:"):
        return path[len("disk:"):]
    return path or "/"


def is_public_root_link(public_url, start_path):
    """
    Проверяет, лежит ли папка start_path В КОРНЕ публичной ссылки.
    Используется только в режиме strip_prefix_mode='auto'.

    True  → ссылка ведёт в родительскую папку (префикс отрезать НЕ надо).
    False → ссылка ведёт внутрь start_path (префикс надо отрезать).
    """
    api = "https://cloud-api.yandex.net/v1/disk/public/resources"
    start_name = (start_path or "").strip("/")
    if not start_name:
        return False
    try:
        r = requests.get(
            api,
            params={"public_key": public_url, "path": "/", "limit": 100},
            headers=HEADERS, timeout=HTTP_TIMEOUT,
        )
        if r.status_code != 200:
            print(f"  [debug] is_public_root_link: HTTP {r.status_code} → True")
            return True
        items = r.json().get('_embedded', {}).get('items', [])
        names = [it.get('name') for it in items if it.get('type') == 'dir']
        print(f"  [debug] is_public_root_link: root-папки = {names}")
        for it in items:
            if it.get('name') == start_name and it.get('type') == 'dir':
                print(f"  [debug]   найден '{start_name}' в корне → True")
                return True
        print(f"  [debug]   '{start_name}' НЕ найден в корне → False")
        return False
    except Exception as e:
        print(f"  [debug] is_public_root_link: исключение {e} → True")
        return True


def _to_public_path(full_path, start_path, strip_prefix):
    if not full_path:
        return ""
    if not strip_prefix:
        return full_path
    sp = (start_path or "/").strip("/")
    if not sp:
        return full_path
    prefix = "/" + sp + "/"
    if full_path.startswith(prefix):
        return "/" + full_path[len(prefix):]
    return full_path


def list_public_files_recursive(client, public_key, path="/", depth=0,
                                max_depth=30):
    result = []
    if depth > max_depth:
        return result
    if depth == 0:
        time.sleep(SLEEP_BEFORE_LISTDIR)
    try:
        items = list(client.listdir(path, public_key=public_key))
    except Exception as e:
        print(f"    [!] Ошибка listdir({path}): {e}")
        return result
    for item in items:
        try:
            if item.type == "dir":
                result.extend(
                    list_public_files_recursive(
                        client, public_key, item.path, depth + 1, max_depth
                    )
                )
            elif item.type == "file":
                ext = os.path.splitext(item.name)[1].lower()
                clean_path = _strip_disk_prefix(item.path)
                result.append({
                    "name":      item.name,
                    "path":      clean_path,
                    "full_path": clean_path,
                    "size":      getattr(item, "size", 0) or 0,
                    "ext":       ext,
                })
        except Exception:
            continue
    return result


def print_folder_diagnostics(client, public_key, path="/"):
    try:
        time.sleep(SLEEP_BEFORE_LISTDIR)
        items = list(client.listdir(path, public_key=public_key))
    except Exception as e:
        print(f"  [!] Не удалось получить содержимое '{path}': {e}")
        return
    print(f"  Содержимое '{path}' ({len(items)} элементов):")
    for it in items[:15]:
        kind = "DIR " if it.type == "dir" else "FILE"
        print(f"    [{kind}] {it.name}")
    if len(items) > 15:
        print(f"    ... и ещё {len(items) - 15}")


def make_download_link(public_url, public_path):
    return f"{public_url}?path={quote(public_path)}"


def get_download_url(public_url, public_path):
    dl_api = "https://cloud-api.yandex.net/v1/disk/public/resources/download"
    params = {"public_key": public_url, "path": public_path}
    try:
        r = requests.get(dl_api, params=params, headers=HEADERS,
                         timeout=HTTP_TIMEOUT)
        if r.status_code == 200:
            return r.json().get('href')
    except Exception as e:
        print(f"    [!] download URL для {public_path}: {e}")
    return None


def _extract_folder(full_path, start_path):
    sp = (start_path or "/").strip("/")
    fp = (full_path or "").strip("/")
    if sp and fp.startswith(sp + "/"):
        rel = fp[len(sp) + 1:]
    elif sp and fp == sp:
        rel = ""
    else:
        rel = fp
    if not rel:
        return ""
    parts = rel.split("/")
    if len(parts) >= 2:
        return parts[0]
    return ""


def _extract_folder_full(full_path, start_path):
    sp = (start_path or "/").strip("/")
    fp = (full_path or "").strip("/")
    if sp and fp.startswith(sp + "/"):
        rel = fp[len(sp) + 1:]
    elif sp and fp == sp:
        rel = ""
    else:
        rel = fp
    if not rel:
        return ""
    parts = rel.split("/")
    if len(parts) < 2:
        return ""
    return " / ".join(parts[:-1])


# ============================================================
# СКАЧИВАНИЕ ФАЙЛА ДЛЯ ПАРСИНГА
# ============================================================

def _fetch_file_content(client, f, public_url, public_path, max_bytes=None):
    dl_url = get_download_url(public_url, public_path)
    if dl_url:
        try:
            if max_bytes:
                r = requests.get(dl_url, headers=HEADERS, timeout=60,
                                 stream=True)
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
                r = requests.get(dl_url, headers=HEADERS, timeout=60)
                if r.status_code == 200:
                    return r.content
        except Exception:
            pass

    if client is not None and f.get('path'):
        try:
            buf = io.BytesIO()
            if max_bytes:
                client.download(f['path'], buf,
                                byte_range=(0, max_bytes - 1))
            else:
                client.download(f['path'], buf)
            return buf.getvalue()
        except Exception:
            pass
    return b""


# ============================================================
# ПАРСИНГ FB2
# ============================================================

def parse_fb2(content_bytes):
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
                            content_type = b_el.attrib.get(
                                'content-type', 'image/jpeg')
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
                            except Exception:
                                pass
                            break
                    if result['cover_data']:
                        break
    return result


# ============================================================
# ПАРСИНГ TXT
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
    result = {'title': '', 'author': '', 'description': ''}
    head = content_bytes[:MAX_TXT_HEAD]
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
# ВНЕШНИЕ ИСТОЧНИКИ
# ============================================================

def _fantlab_lookup(title, author):
    if not title:
        return {}
    def _search(q, limit=5):
        try:
            r = requests.get(FANTLAB_SEARCH, params={"q": q, "onlymatches": 1},
                             headers=HEADERS, timeout=15)
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
    time.sleep(SLEEP_BETWEEN)
    if not matches:
        matches = _search(title)
        time.sleep(SLEEP_BETWEEN)
    if not matches:
        return {}

    title_l = (title or "").lower()
    surname = _author_surname(author)
    best, best_score = None, -1
    for m in matches:
        score = 0
        names = " ".join(filter(None, [
            m.get("rusname", ""), m.get("name", ""), m.get("fullname", ""),
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
        r = requests.get(FANTLAB_WORK.format(id=work_id),
                         headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return {}
        work = r.json() or {}
    except Exception:
        return {}
    time.sleep(SLEEP_BETWEEN)

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
    def _req(params):
        try:
            r = requests.get(WIKI_API, params=params, headers=HEADERS, timeout=15)
            return r.json() if r.status_code == 200 else {}
        except Exception:
            return {}

    def _search_page(q):
        data = _req({"action": "query", "format": "json", "list": "search",
                     "srsearch": q, "srlimit": 1, "srnamespace": 0})
        hits = (data.get("query") or {}).get("search") or []
        if not hits:
            return None, None
        return hits[0].get("pageid"), hits[0].get("title")

    def _extract(pageid):
        data = _req({"action": "query", "format": "json", "prop": "extracts",
                     "pageids": pageid, "explaintext": 1, "exintro": 1,
                     "redirects": 1})
        pages = (data.get("query") or {}).get("pages") or {}
        for _, p in pages.items():
            t = (p.get("extract") or "").strip()
            if t:
                return t
        return ""

    queries = [f'"{title}" роман', f'"{title}" книга', f'"{title}" повесть',
               f'{title} (роман)', f'{title} (книга)', f'{title} (повесть)']
    if author:
        queries.append(f'{title} {author} роман')

    seen = set()
    for q in queries:
        pageid, page_title = _search_page(q)
        if not pageid or pageid in seen:
            time.sleep(SLEEP_BETWEEN)
            continue
        seen.add(pageid)
        pt = (page_title or "").strip()
        if re.match(r"^[А-ЯЁ][а-яё]+\s*,\s*[А-ЯЁ][а-яё]+", pt):
            time.sleep(SLEEP_BETWEEN)
            continue
        text = _extract(pageid)
        if not text or len(text) < 150:
            time.sleep(SLEEP_BETWEEN)
            continue
        if not _title_phrase_matches(title, text):
            time.sleep(SLEEP_BETWEEN)
            continue
        paras = [p.strip() for p in text.split("\n") if p.strip()][:3]
        desc = "\n\n".join(paras)[:MAX_DESC_LEN]
        if desc:
            return {"description": desc, "cover": "", "source": "wikipedia"}
        time.sleep(SLEEP_BETWEEN)
    return {}


def _openlibrary_lookup(title, author):
    if not title:
        return {}
    params = {"title": title, "limit": 5}
    if author:
        params["author"] = author
    try:
        r = requests.get(OPENLIBRARY_API, params=params,
                         headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return {}
        docs = (r.json() or {}).get("docs") or []
    except Exception:
        return {}
    for d in docs:
        found = d.get("title", "") or ""
        if not _title_phrase_matches(title, found):
            continue
        desc = ""
        fs = d.get("first_sentence")
        if isinstance(fs, list) and fs:
            desc = fs[0]
        elif isinstance(fs, str):
            desc = fs
        if not desc and d.get("subtitle"):
            desc = d["subtitle"]
        if not desc:
            continue
        if len(desc) > MAX_DESC_LEN:
            desc = desc[:MAX_DESC_LEN] + "..."
        cover_id = d.get("cover_i") or 0
        cover = (f"https://covers.openlibrary.org/b/id/{cover_id}-L.jpg"
                 if cover_id else "")
        return {"description": desc.strip(), "cover": cover,
                "source": "openlibrary"}
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
        r = requests.get(GOOGLE_BOOKS_API, params=params,
                         headers=HEADERS, timeout=15)
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


def _llm_lookup(title, author):
    if os.environ.get("ENABLE_LLM_FALLBACK", "0") != "1":
        return {}
    provider = os.environ.get("LLM_PROVIDER", "yandexgpt").lower()
    prompt = ("Ты библиотекарь. Кратко опиши именно книгу (не автора!) "
              "в 2–3 предложениях, без спойлеров и без вступления.\n"
              f"Автор: {author or 'неизвестен'}\nНазвание: {title}")
    if provider == "openai":
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            return {}
        try:
            r = requests.post(OPENAI_URL,
                headers={"Authorization": f"Bearer {api_key}",
                         "Content-Type": "application/json"},
                json={"model": "gpt-4o-mini",
                      "messages": [
                          {"role": "system", "content": "Ты библиотекарь."},
                          {"role": "user", "content": prompt}],
                      "temperature": 0.3, "max_tokens": 300},
                timeout=60)
            if r.status_code != 200:
                return {}
            text = r.json()["choices"][0]["message"]["content"].strip()
            return {"description": text, "cover": "", "source": "openai"}
        except Exception:
            return {}
    api_key   = os.environ.get("YANDEX_GPT_API_KEY")
    folder_id = os.environ.get("YANDEX_FOLDER_ID")
    if not api_key or not folder_id:
        return {}
    try:
        r = requests.post(YANDEX_GPT_URL,
            headers={"Authorization": f"Api-Key {api_key}",
                     "Content-Type": "application/json"},
            json={"modelUri": f"gpt://{folder_id}/yandexgpt-lite/latest",
                  "completionOptions": {"temperature": 0.3, "maxTokens": 300},
                  "messages": [
                      {"role": "system", "text": "Ты библиотекарь."},
                      {"role": "user", "text": prompt}]},
            timeout=60)
        if r.status_code != 200:
            return {}
        text = r.json()["result"]["alternatives"][0]["message"]["text"].strip()
        return {"description": text, "cover": "", "source": "yandexgpt"}
    except Exception:
        return {}


def enrich_book(title, author, cache, diag):
    # Fast mode: не ходим во внешние API
    if SKIP_ENRICHMENT:
        return {}

    ck = hashlib.sha1(
        f"{(title or '').lower()}|{(author or '').lower()}".encode()
    ).hexdigest()
    if ck in cache:
        meta = cache[ck]
        src = meta.get("source", "none")
        diag.sources[src] = diag.sources.get(src, 0) + 1
        return meta

    meta = {}
    for fn in (_fantlab_lookup, _wikipedia_lookup,
               _openlibrary_lookup, _google_books_lookup, _llm_lookup):
        try:
            res = fn(title, author)
            if res and res.get("description"):
                meta = res
                break
        except Exception:
            pass
        time.sleep(SLEEP_BETWEEN)
    cache[ck] = meta
    if meta:
        src = meta.get("source", "none")
        diag.sources[src] = diag.sources.get(src, 0) + 1
    return meta


# ============================================================
# ОБЛОЖКИ
# ============================================================

def upload_cover_to_supabase(cover_data, ext, book_title):
    if not SUPABASE_SERVICE_KEY:
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
        resp = requests.put(url, headers=headers, data=cover_data,
                            timeout=HTTP_TIMEOUT)
        if resp.status_code in (200, 201):
            return f"{SUPABASE_URL}/storage/v1/object/public/{COVERS_BUCKET}/{filename}"
        print(f"    [!] Обложка: HTTP {resp.status_code} для '{book_title[:40]}'")
    except Exception as e:
        print(f"    [!] Обложка: {e}")
    return None


# ============================================================
# SUPABASE
# ============================================================

def _get_supabase_client():
    """Возвращает клиент Supabase или None, если не настроен."""
    if not supa_create_client:
        return None
    if not SUPABASE_SERVICE_KEY:
        return None
    try:
        return supa_create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)
    except Exception as e:
        print(f"[!] Ошибка создания Supabase-клиента: {e}")
        return None


def load_sections_from_supabase():
    """
    Возвращает список активных разделов site_sections,
    включая поля manual_override и strip_prefix_mode.
    """
    client = _get_supabase_client()
    if client is None:
        print("[!] Supabase-клиент недоступен — не могу загрузить разделы.")
        return []
    try:
        resp = (
            client.table("site_sections")
            .select("key,label,icon,handler_type,yandex_url,yandex_path,"
                    "sheet_name,json_path,container,columns,folderable,"
                    "is_active,manual_override,sort_order,strip_prefix_mode")
            .eq("is_active", True)
            .order("sort_order")
            .execute()
        )
        return resp.data or []
    except Exception as e:
        # Возможно, ещё нет колонки strip_prefix_mode — пробуем без неё
        print(f"[!] Ошибка загрузки разделов (с strip_prefix_mode): {e}")
        try:
            resp = (
                client.table("site_sections")
                .select("key,label,icon,handler_type,yandex_url,yandex_path,"
                        "sheet_name,json_path,container,columns,folderable,"
                        "is_active,manual_override,sort_order")
                .eq("is_active", True)
                .order("sort_order")
                .execute()
            )
            print("[i] Загружено без strip_prefix_mode (fallback на env).")
            return resp.data or []
        except Exception as e2:
            print(f"[!] Повторная ошибка загрузки разделов: {e2}")
            return []


def save_orphans_to_supabase(section_key, section_label,
                             sheet_name, orphans, headers=None):
    """
    Upsert в sync_orphans: одна строка на раздел.
    Дополнительно сохраняем headers (для diff-view).
    """
    client = _get_supabase_client()
    if client is None:
        return
    try:
        payload = {
            "section_key":   section_key,
            "section_label": section_label,
            "sheet_name":    sheet_name,
            "orphans":       orphans or [],
            "updated_at":    datetime.datetime.utcnow().isoformat() + "Z",
        }
        if headers is not None:
            payload["headers"] = headers
        client.table(SYNC_ORPHANS_TABLE).upsert(payload).execute()
    except Exception as e:
        print(f"    [!] Не удалось сохранить orphans в Supabase: {e}")


# ============================================================
# GOOGLE SHEETS
# ============================================================

def get_gspread_client():
    if gspread is None or Credentials is None:
        raise RuntimeError("gspread / google-auth не установлены")
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise RuntimeError("GOOGLE_CREDENTIALS_JSON не задан")
    creds_dict = json.loads(creds_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def open_spreadsheet(gs_client):
    if SPREADSHEET_ID:
        print(f"Открываю таблицу по ID: {SPREADSHEET_ID}")
        return gs_client.open_by_key(SPREADSHEET_ID)
    return gs_client.open(SPREADSHEET_NAME)


def get_or_create_sheet(sh, name):
    try:
        return sh.worksheet(name)
    except gspread.WorksheetNotFound:
        print(f"  [!] Лист '{name}' не найден, создаю...")
        return sh.add_worksheet(title=name, rows=2000, cols=12)


def _col_letter(n):
    """Индекс колонки (0-based) → буквенное обозначение A, B, ..., Z, AA, AB."""
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def ensure_headers(sheet, headers):
    """
    Обновляет заголовки листа. Полностью перезаписывает первую строку
    до максимальной ширины листа, чтобы не осталось «хвостов»
    от старых версий схемы (иначе gspread.get_all_records()
    ругается на дубликаты).
    """
    try:
        current = sheet.row_values(1)
    except Exception:
        current = []

    # Нужно ли обновлять?
    need_update = False
    if len(current) != len(headers):
        need_update = True
    else:
        for i, h in enumerate(headers):
            if (current[i] or "").strip() != h.strip():
                need_update = True
                break
    if not need_update:
        return

    # Определяем максимальную ширину листа (чтобы перезаписать «хвост»)
    try:
        all_values = sheet.get_all_values()
        max_cols = max((len(r) for r in all_values), default=len(headers))
    except Exception:
        max_cols = len(headers)
    max_cols = max(max_cols, len(headers))

    # Формируем строку: заголовки + пустые ячейки для «хвоста»
    new_row = list(headers) + [""] * (max_cols - len(headers))

    end_col_letter = _col_letter(max_cols - 1)
    range_a1 = f"A1:{end_col_letter}1"
    try:
        sheet.update(values=[new_row], range_name=range_a1,
                     value_input_option="USER_ENTERED")
        print(f"    [+] Обновлены заголовки: {headers} "
              f"(очищено до {max_cols} колонок)")
    except Exception as e:
        print(f"    [!] Заголовки: {e}")


def load_sheet_snapshot(sheet):
    """
    Возвращает (all_values, existing_index).

      all_values     — list[list[str]] — все значения листа (вкл. заголовок).
      existing_index — {download_link: {"row": int, "values": list[str]}}.

    download_link ищется по первой подходящей колонке.
    """
    try:
        rows = sheet.get_all_values()
    except Exception as e:
        print(f"    [!] Не удалось прочитать лист: {e}")
        return [], {}

    if not rows:
        return [], {}

    header = [h.strip().lower() for h in rows[0]]
    idx = None
    for cand in ("ссылка для скачивания", "ссылка",
                 "download_link", "link"):
        if cand in header:
            idx = header.index(cand)
            break

    existing = {}
    if idx is not None and len(rows) >= 2:
        for i, r in enumerate(rows[1:], start=2):
            if len(r) > idx and r[idx].strip():
                existing[r[idx].strip()] = {
                    "row":    i,
                    "values": list(r),
                }
    return rows, existing


def merge_row(old_values, new_values, headers, always_update=None):
    """
    Объединяет старую (уже сохранённую в Sheets) и новую
    (сгенерированную из Яндекс.Диска) строки.

    Правила:
      • Колонки из always_update → всегда значение из new_values.
      • Остальные: если в old_values непустое значение — сохраняем его
        (ручная правка), иначе берём значение из new_values.

    Возвращает список значений длиной len(headers).
    """
    if always_update is None:
        always_update = ALWAYS_UPDATE_HEADERS

    if not PRESERVE_USER_EDITS:
        return list(new_values)

    result = []
    n = len(headers)
    for i in range(n):
        header = headers[i]
        old_val = old_values[i] if i < len(old_values) else ""
        new_val = new_values[i] if i < len(new_values) else ""

        old_s = str(old_val).strip() if old_val is not None else ""
        new_s = str(new_val).strip() if new_val is not None else ""

        if header in always_update:
            result.append(new_s if new_s else old_s)
        else:
            result.append(old_val if old_s else new_val)
    return result


def backup_sheet(sh, worksheet, section_key, all_values,
                 max_backups=None):
    """
    Создаёт резервную копию листа с именем _backup_<section>_<ts>.
    Удаляет самые старые копии, оставляя max_backups.

    Возвращает имя созданного листа или None.
    """
    if not all_values:
        print("    [b] Backup: лист пуст, копия не нужна.")
        return None
    if max_backups is None:
        max_backups = BACKUP_KEEP_COUNT

    ts = datetime.datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    backup_name = f"_backup_{section_key}_{ts}"

    max_cols = max((len(r) for r in all_values), default=1)
    max_rows = len(all_values)

    try:
        backup = sh.add_worksheet(
            title=backup_name,
            rows=max(max_rows + 10, 100),
            cols=max(max_cols + 2, 10),
        )
        backup.update(values=all_values, range_name="A1",
                      value_input_option="USER_ENTERED")
        print(f"    [b] Backup создан: {backup_name} "
              f"({max_rows} строк × {max_cols} колонок)")
    except Exception as e:
        print(f"    [!] Backup: ошибка создания: {e}")
        return None

    # Чистим старые резервные копии этого раздела
    try:
        prefix = f"_backup_{section_key}_"
        backups = [ws for ws in sh.worksheets()
                   if ws.title.startswith(prefix)]
        backups.sort(key=lambda w: w.title, reverse=True)
        for old in backups[max_backups:]:
            try:
                sh.del_worksheet(old)
                print(f"    [b] Удалён старый backup: {old.title}")
            except Exception:
                pass
    except Exception as e:
        print(f"    [!] Backup cleanup: {e}")

    return backup_name


def batch_update_rows(sheet, updates):
    if not updates:
        return 0
    total = len(updates)
    updated = 0
    chunk = 100
    for i in range(0, total, chunk):
        part = updates[i:i + chunk]
        try:
            sheet.batch_update(part, value_input_option="USER_ENTERED")
            updated += len(part)
            print(f"    [+] Обновлено {updated}/{total}")
        except Exception as e:
            print(f"    [!] batch_update {i//chunk}: {e}")
        time.sleep(0.5)
    return updated


def append_rows_safe(sheet, rows, batch_size=200):
    added = 0
    failed = 0
    total = len(rows)
    for i in range(0, total, batch_size):
        batch = rows[i:i + batch_size]
        for attempt in range(3):
            try:
                sheet.append_rows(batch, value_input_option="USER_ENTERED")
                added += len(batch)
                print(f"    [+] Записано {added}/{total}")
                break
            except Exception as e:
                print(f"    [!] Попытка {attempt+1}: {e}")
                time.sleep(2 * (attempt + 1))
        else:
            failed += len(batch)
        time.sleep(1)
    return added, failed


# ============================================================
# СБОРКА СТРОК ПО ТИПУ ОБРАБОТЧИКА
# ============================================================

def _row_from_record(record, headers):
    """Собирает строку таблицы в порядке русских заголовков."""
    row = []
    for ru in headers:
        en = RU_TO_EN.get(ru, ru)
        row.append(record.get(en, ""))
    return row


def build_rows_books(client, files, public_url, start_path, strip_prefix,
                     headers, cache, diag):
    """Обработчик «books» — парсит fb2/txt, обогащает, тянет обложки."""
    rows = []
    seen_links = set()
    total = len(files)

    for idx, f in enumerate(files, 1):
        name = f["name"]
        ext  = f["ext"]

        if ALLOWED_EXTS and ext not in ALLOWED_EXTS:
            log_skip(diag, "bad_ext", name)
            continue

        meta0 = parse_book_name(name)
        title0, author0 = meta0["title"], meta0["author"]
        if not title0:
            log_skip(diag, "garbage", name)
            continue

        public_path = _to_public_path(f["full_path"], start_path, strip_prefix)
        link = make_download_link(public_url, public_path)
        if link in seen_links:
            log_skip(diag, "dup_link", name)
            continue
        seen_links.add(link)

        file_data = {}
        if ext in PARSEABLE_EXTS:
            max_b = MAX_TXT_HEAD if ext == ".txt" else None
            content = _fetch_file_content(client, f, public_url,
                                          public_path, max_bytes=max_b)
            if content:
                file_data = parse_fb2(content) if ext == ".fb2" \
                    else parse_txt(content)

        title  = file_data.get("title")  or title0
        author = file_data.get("author") or author0
        description = file_data.get("description", "")

        if not author:
            author = _extract_folder(f["full_path"], start_path) or ""

        if description:
            diag.from_file += 1
            diag.sources["file"] = diag.sources.get("file", 0) + 1

        if not description:
            ext_meta = enrich_book(title, author, cache, diag)
            if ext_meta.get("description"):
                description = ext_meta["description"]

        cover = ""
        if file_data.get("cover_data"):
            cover = upload_cover_to_supabase(
                file_data["cover_data"],
                file_data.get("cover_ext", "jpg"),
                title
            ) or ""
        if not cover:
            ck = hashlib.sha1(
                f"{(title or '').lower()}|{(author or '').lower()}".encode()
            ).hexdigest()
            meta = cache.get(ck)
            if meta and meta.get("cover"):
                cover = meta["cover"]

        fmt = ext.lstrip(".")
        size_mb = round(f["size"] / (1024 * 1024), 1) if f["size"] else 0
        folder = _extract_folder(f["full_path"], start_path)

        record = {
            "title":         title,
            "author":        author,
            "format":        fmt,
            "size":          str(size_mb),
            "download_link": link,
            "cover":         cover,
            "folder":        folder,
            "description":   description,
            "version":       "",
        }
        rows.append(_row_from_record(record, headers))

        if idx % 50 == 0:
            print(f"    ... обработано {idx}/{total}, записей: {len(rows)}, "
                  f"из файлов: {diag.from_file}")

    return rows


def build_rows_music(files, public_url, start_path, strip_prefix,
                     headers, diag):
    rows = []
    seen_links = set()

    for f in files:
        name = f["name"]
        ext  = f["ext"]

        if ext not in MUSIC_EXTS:
            log_skip(diag, "wrong_ext", name)
            continue

        parsed = parse_music_name(name)
        title, artist, year = parsed["title"], parsed["artist"], parsed["year"]
        if not title:
            log_skip(diag, "garbage", name)
            continue

        public_path = _to_public_path(f["full_path"], start_path, strip_prefix)
        link = make_download_link(public_url, public_path)
        if link in seen_links:
            log_skip(diag, "dup_link", name)
            continue
        seen_links.add(link)

        folder = _extract_folder_full(f["full_path"], start_path)
        if not artist and folder:
            artist = folder.split(" / ")[0]
        if not year and folder:
            ym = re.search(r"\b(19[5-9]\d|20[0-3]\d)\b", folder)
            if ym:
                year = ym.group(1)

        size_mb = round(f["size"] / (1024 * 1024), 1) if f["size"] else 0

        record = {
            "title":         title,
            "artist":        artist,
            "year":          year,
            "size":          str(size_mb),
            "download_link": link,
            "folder":        folder,
        }
        rows.append(_row_from_record(record, headers))

    return rows


def build_rows_programs(files, public_url, start_path, strip_prefix,
                        headers, diag):
    rows = []
    seen_links = set()

    for f in files:
        name = f["name"]
        ext  = f["ext"]

        if ext not in PROGRAM_EXTS:
            log_skip(diag, "wrong_ext", name)
            continue

        meta = parse_program_name(name)
        title, version = meta["title"], meta["version"]
        if not title:
            log_skip(diag, "garbage", name)
            continue

        public_path = _to_public_path(f["full_path"], start_path, strip_prefix)
        link = make_download_link(public_url, public_path)
        if link in seen_links:
            log_skip(diag, "dup_link", name)
            continue
        seen_links.add(link)

        size_mb = round(f["size"] / (1024 * 1024), 1) if f["size"] else 0
        folder = _extract_folder(f["full_path"], start_path)

        record = {
            "title":         title,
            "description":   "",
            "version":       version,
            "size":          str(size_mb),
            "download_link": link,
            "folder":        folder,
        }
        rows.append(_row_from_record(record, headers))

    return rows


def build_rows_universal(files, public_url, start_path, strip_prefix,
                         headers, diag):
    """Универсальный обработчик — берёт все файлы."""
    rows = []
    seen_links = set()

    for f in files:
        name = f["name"]

        meta = parse_universal_name(name)
        title = meta["title"]
        if not title:
            log_skip(diag, "garbage", name)
            continue

        public_path = _to_public_path(f["full_path"], start_path, strip_prefix)
        link = make_download_link(public_url, public_path)
        if link in seen_links:
            log_skip(diag, "dup_link", name)
            continue
        seen_links.add(link)

        size_mb = round(f["size"] / (1024 * 1024), 1) if f["size"] else 0
        folder = _extract_folder_full(f["full_path"], start_path)

        record = {
            "title":         title,
            "description":   "",
            "size":          str(size_mb),
            "download_link": link,
            "folder":        folder,
        }
        rows.append(_row_from_record(record, headers))

    return rows


# ============================================================
# ЗАГОЛОВКИ ИЗ columns
# ============================================================

def headers_from_columns(columns):
    """
    Преобразует список ключей (['title','description','download_link'])
    в список русских заголовков для Google Sheets.
    """
    headers = []
    for key in columns or []:
        ru = EN_TO_RU.get(key)
        if ru:
            headers.append(ru)
    # Гарантируем наличие «Ссылка для скачивания»
    if "Ссылка для скачивания" not in headers:
        headers.append("Ссылка для скачивания")
    return headers


# ============================================================
# ДЕТЕКТ ОСИРОТЕВШИХ СТРОК
# ============================================================

def find_orphan_rows(existing_index, rows, headers, link_idx):
    """
    Возвращает список осиротевших строк:
    строки Sheets, download_link которых отсутствует в свежесобранных rows.
    Дополнительно сохраняем ПОЛНЫЕ значения строки (для diff-view).
    """
    new_links = set()
    for row in rows:
        if link_idx < len(row):
            v = str(row[link_idx]).strip()
            if v:
                new_links.add(v)

    # Ищем колонку с названием
    title_idx = None
    for cand in TITLE_HEADER_CANDIDATES:
        if cand in headers:
            title_idx = headers.index(cand)
            break

    orphans = []
    for link, info in existing_index.items():
        if link in new_links:
            continue
        values = info["values"]
        title = ""
        if title_idx is not None and title_idx < len(values):
            title = str(values[title_idx]).strip()
        # Дополняем до длины headers, чтобы diff-view мог сопоставить
        padded = list(values) + [""] * max(0, len(headers) - len(values))
        orphans.append({
            "row":    info["row"],
            "title":  title,
            "link":   link,
            "values": padded[:len(headers)],
        })
    return orphans


# ============================================================
# СИНХРОНИЗАЦИЯ РАЗДЕЛА
# ============================================================

def sync_section(section, gs_client, cache):
    """section — строка из таблицы site_sections."""
    label = section.get("label") or section.get("key")
    section_key = section.get("key")
    yandex_url  = section.get("yandex_url") or ""
    start_path  = section.get("yandex_path") or "/"
    sheet_name  = section.get("sheet_name") or label
    handler_type = section.get("handler_type") or "universal"
    columns     = section.get("columns") or ["title", "description", "download_link"]
    manual_override = bool(section.get("manual_override"))

    # Режим отрезания префикса: сначала per-section, потом env-fallback.
    strip_mode = (
        section.get("strip_prefix_mode")
        or STRIP_PREFIX_MODE_DEFAULT
        or "always"
    ).lower()
    if strip_mode not in ("auto", "always", "never"):
        strip_mode = "always"

    print(f"\n=== Раздел: {label} ({section_key}) ===")
    print(f"Handler:         {handler_type}")
    print(f"Источник:        {yandex_url or '(не задан)'}")
    print(f"Подпапка:        {start_path}")
    print(f"Лист Sheets:     {sheet_name}")
    print(f"manual_override: {manual_override}")
    print(f"strip_prefix:    {strip_mode}")

    if not yandex_url:
        print("  [!] yandex_url не задан — раздел пропускается.")
        return

    if yadisk is None:
        print("  [!] yadisk не установлен.")
        return

    public_key = get_public_key(yandex_url)
    token = os.environ.get("YADISK_TOKEN")
    if not token:
        print("  [!] YADISK_TOKEN не задан.")
        return

    # ─── Определение режима отрезания префикса ───
    sp = (start_path or "").strip("/")

    if not sp:
        strip_prefix = False
        is_root = False
    elif strip_mode == "always":
        strip_prefix = True
        is_root = False
    elif strip_mode == "never":
        strip_prefix = False
        is_root = True
    else:  # auto
        is_root = is_public_root_link(yandex_url, start_path)
        strip_prefix = not is_root

    print(f"  [i] strip_prefix_mode='{strip_mode}', start_path='{start_path}' → "
          f"{'ОТРЕЗАЕМ' if strip_prefix else 'НЕ отрезаем'} префикс "
          f"(is_root={is_root})")

    diag = Diag(label)

    with yadisk.Client(token=token) as client:
        try:
            if not client.check_token():
                print("  [!] Неверный YADISK_TOKEN")
                return
        except Exception as e:
            print(f"  [!] Ошибка проверки токена: {e}")
            return

        print_folder_diagnostics(client, public_key, start_path)
        files = list_public_files_recursive(client, public_key, path=start_path)

        diag.found = len(files)
        print(f"  Найдено файлов: {len(files)}")

        headers = headers_from_columns(columns)

        if handler_type == "books":
            rows = build_rows_books(client, files, yandex_url, start_path,
                                    strip_prefix, headers, cache, diag)
        elif handler_type == "music":
            rows = build_rows_music(files, yandex_url, start_path,
                                    strip_prefix, headers, diag)
        elif handler_type == "programs":
            rows = build_rows_programs(files, yandex_url, start_path,
                                       strip_prefix, headers, diag)
        else:  # universal
            rows = build_rows_universal(files, yandex_url, start_path,
                                        strip_prefix, headers, diag)

    diag.kept = len(rows)
    print(f"\n  Сгенерировано строк: {len(rows)}")

    # --- Открываем таблицу ---
    try:
        sh = open_spreadsheet(gs_client)
    except Exception as e:
        print(f"  [!] Не удалось открыть таблицу: {e}")
        return

    sheet = get_or_create_sheet(sh, sheet_name)
    ensure_headers(sheet, headers)

    # --- Снимок текущего состояния ---
    all_values, existing = load_sheet_snapshot(sheet)
    print(f"  Существующих строк: {len(existing)}")

    link_ru = "Ссылка для скачивания"
    link_idx = headers.index(link_ru) if link_ru in headers else 0
    n_cols = len(headers)
    end_col_letter = _col_letter(n_cols - 1)

    # --- Детект сирот ---
    orphans = find_orphan_rows(existing, rows, headers, link_idx)
    diag.orphans_found = len(orphans)
    if orphans:
        print(f"  Осиротевших строк: {len(orphans)} "
              f"(файлы исчезли с Диска)")
        save_orphans_to_supabase(section_key, label, sheet_name,
                                 orphans, headers)
    else:
        # Очищаем запись — сирот нет
        save_orphans_to_supabase(section_key, label, sheet_name,
                                 [], headers)

    # --- Формируем изменения ---
    updates = []
    to_add = []
    preserved_rows = 0
    unchanged_rows = 0

    for row in rows:
        link = str(row[link_idx]).strip() if link_idx < len(row) else ""

        if not link or link not in existing:
            to_add.append(row)
            continue

        info = existing[link]
        row_num = info["row"]
        old_values = info["values"]

        # manual_override → существующие строки вообще не трогаем
        if manual_override:
            unchanged_rows += 1
            continue

        merged = merge_row(old_values, row, headers)

        # Сравниваем только первые n_cols столбцов
        old_cmp = [str(v) for v in old_values[:n_cols]]
        new_cmp = [str(v) for v in merged]

        if old_cmp == new_cmp:
            unchanged_rows += 1
            continue

        # Считаем, что была сохранена ручная правка:
        # контентная колонка непустая в old и отличается от new
        if PRESERVE_USER_EDITS:
            for j, h in enumerate(headers):
                if h in ALWAYS_UPDATE_HEADERS:
                    continue
                ov = (old_values[j] if j < len(old_values) else "").strip()
                nv = (row[j] if j < len(row) else "").strip()
                if ov and ov != nv:
                    preserved_rows += 1
                    break

        rng = f"A{row_num}:{end_col_letter}{row_num}"
        updates.append({"range": rng, "values": [merged]})

    print(f"\n    К обновлению:      {len(updates)}")
    print(f"    К добавлению:      {len(to_add)}")
    if preserved_rows:
        print(f"    Сохранено правок:  {preserved_rows}")
    if unchanged_rows:
        print(f"    Без изменений:     {unchanged_rows}")
    if manual_override and unchanged_rows:
        print(f"    [i] manual_override=True — существующие строки "
              f"не перезаписаны")
    if not PRESERVE_USER_EDITS:
        print("    [i] PRESERVE_USER_EDITS=0 — ручные правки перезаписываются")

    diag.preserved_edits = preserved_rows
    diag.unchanged_rows  = unchanged_rows

    # --- Backup перед записью ---
    if (updates or to_add) and BACKUP_BEFORE_SYNC:
        diag.backup_name = backup_sheet(sh, sheet, section_key,
                                        all_values) or ""

    # --- Применяем ---
    if updates:
        diag.updated = batch_update_rows(sheet, updates)
    if to_add:
        added, failed = append_rows_safe(sheet, to_add)
        diag.added = added
        if failed:
            print(f"    [!] Не удалось записать: {failed}")

    diag.report()


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("yandex_disk_sync.py — старт")
    print("=" * 60)
    print(f"PRESERVE_USER_EDITS  = {PRESERVE_USER_EDITS}")
    print(f"BACKUP_BEFORE_SYNC   = {BACKUP_BEFORE_SYNC} "
          f"(keep last {BACKUP_KEEP_COUNT})")
    print(f"STRIP_PREFIX_DEFAULT = {STRIP_PREFIX_MODE_DEFAULT}")
    print(f"SKIP_ENRICHMENT      = {SKIP_ENRICHMENT}")
    if SYNC_SECTIONS_FILTER:
        print(f"SYNC_SECTIONS_FILTER = {SYNC_SECTIONS_FILTER}")
    else:
        print("SYNC_SECTIONS_FILTER = (все активные разделы)")

    if yadisk is None:
        print("[!] yadisk не установлен: pip install -r requirements.txt")
        return

    print("\nЗагрузка разделов из Supabase...")
    sections = load_sections_from_supabase()
    if not sections:
        print("[!] Не удалось получить ни одного раздела — завершаю.")
        return

    # ─── Фильтр по SYNC_SECTIONS ───
    if SYNC_SECTIONS_FILTER:
        before = len(sections)
        sections = [s for s in sections
                    if s.get("key") in SYNC_SECTIONS_FILTER]
        print(f"  [i] SYNC_SECTIONS отфильтровал {len(sections)} "
              f"из {before} разделов")
        if not sections:
            print("[!] Ни один раздел не попал под фильтр — завершаю.")
            return

    print(f"  Получено разделов: {len(sections)}")
    for s in sections:
        flags = []
        if s.get("manual_override"):
            flags.append("manual_override")
        if s.get("strip_prefix_mode"):
            flags.append(f"strip={s.get('strip_prefix_mode')}")
        suffix = f" [{', '.join(flags)}]" if flags else ""
        print(f"    • {s.get('key')}: {s.get('label')} "
              f"({s.get('handler_type') or 'universal'}){suffix}")

    print("\nПодключение к Google Sheets...")
    try:
        gs_client = get_gspread_client()
    except Exception as e:
        print(f"[!] Google Sheets: {e}")
        return
    print("Клиент создан.")

    if not SPREADSHEET_ID:
        print("[!] SPREADSHEET_ID не задан.")

    cache = {}
    for section in sections:
        try:
            sync_section(section, gs_client, cache)
        except Exception as e:
            print(f"\n[!!!] Ошибка в разделе {section.get('key')}: {e}")
            traceback.print_exc()

    print("\n" + "=" * 60)
    print("Готово!")


if __name__ == "__main__":
    main()
