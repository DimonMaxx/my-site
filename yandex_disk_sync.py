#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
yandex_disk_sync.py
Синхронизация Яндекс.Диска → Google Sheets с диагностикой пропусков
и бесплатным обогащением описаний книг.

Источники описаний (в порядке приоритета, все бесплатные):
    1. Google Books API
    2. Wikipedia API (ru.wikipedia.org)
    3. Open Library API
    4. LLM (YandexGPT / OpenAI) — опционально, включается через env

Переменные окружения:
    GOOGLE_CREDENTIALS_JSON   — JSON сервисного аккаунта Google (обязательно)
    YADISK_TOKEN              — OAuth-токен Яндекс.Диска (обязательно)
    SPREADSHEET_NAME          — имя Google-таблицы (по умолчанию "Content")
    SUPABASE_SERVICE_ROLE_KEY — ключ Supabase для загрузки обложек (опционально)
    GOOGLE_BOOKS_API_KEY      — ключ Google Books (опционально, повышает квоту)
    WIKI_LANG                 — язык Wikipedia (по умолчанию "ru")
    ENABLE_LLM_FALLBACK       — "1" чтобы включить LLM-обогащение (опционально)
    LLM_PROVIDER              — "yandexgpt" или "openai"
    YANDEX_GPT_API_KEY        — если включён fallback через YandexGPT
    YANDEX_FOLDER_ID          — если включён fallback через YandexGPT
    OPENAI_API_KEY            — если включён fallback через OpenAI
"""

import os
import re
import json
import time
import hashlib
import traceback
from urllib.parse import quote

import requests

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

SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME", "Content")

# Разделы: имя листа → публичная ссылка Яндекс.Диска
SECTIONS = {
    "Книги":     "https://disk.yandex.ru/d/zMxF4nXHPkIVCQ",
    "Программы": "https://disk.yandex.ru/d/EjUHvm6mUcgVMw",
    "Музыка":    "",
    "Игры":      "",
    "Статьи":    "",
    "Фильмы":    "",
    "Разное":    "",
    "Новости":   "",
}

# Колонки для каждого листа
SHEET_HEADERS = {
    "Книги":     ["title", "author", "format", "size",
                  "download_link", "cover", "folder", "description"],
    "Программы": ["title", "description", "version", "size",
                  "download_link", "folder"],
}

# Разрешённые расширения для книг
ALLOWED_EXTS = {".fb2", ".epub", ".pdf", ".txt", ".djvu", ".mobi", ".azw3"}

# Кэш обогащения
CACHE_DIR  = "_cache"
CACHE_FILE = os.path.join(CACHE_DIR, "enrichment_cache.json")

# Supabase (для загрузки обложек)
SUPABASE_URL    = "https://rmoonebbvpmvthvpcmpt.supabase.co"
SUPABASE_BUCKET = "covers"

# Внешние API
GOOGLE_BOOKS_API = "https://www.googleapis.com/books/v1/volumes"
OPENLIBRARY_API  = "https://openlibrary.org/search.json"
WIKI_LANG        = os.environ.get("WIKI_LANG", "ru")
WIKI_API         = f"https://{WIKI_LANG}.wikipedia.org/w/api.php"

# LLM
YANDEX_GPT_URL = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion"
OPENAI_URL     = "https://api.openai.com/v1/chat/completions"

# Таймауты и паузы
HTTP_TIMEOUT = 20
SLEEP_BETWEEN_REQUESTS = 0.4

# User-Agent (Wikipedia требует явный)
USER_AGENT = "ContentSyncBot/1.0 (https://github.com/your/repo)"


# ============================================================
# ДИАГНОСТИКА
# ============================================================

class Diag:
    """Счётчики пропусков для отчёта."""
    def __init__(self):
        self.found    = 0
        self.kept     = 0
        self.enriched = 0
        self.sources  = {}     # сколько описаний дал каждый источник
        self.skipped  = {
            "empty_title": 0,
            "empty_link":  0,
            "dup_title":   0,
            "bad_ext":     0,
            "bad_name":    0,
            "other":       0,
        }
        self.skip_samples = []

    def report(self):
        total = sum(self.skipped.values())
        print("\n  ── ДИАГНОСТИКА ──")
        print(f"  Найдено файлов:      {self.found}")
        print(f"  Оставлено к записи:  {self.kept}")
        print(f"  Обогащено описаний:  {self.enriched}")
        if self.sources:
            print(f"  Источники описаний:")
            for src, cnt in sorted(self.sources.items(), key=lambda x: -x[1]):
                print(f"    • {src}: {cnt}")
        print(f"  Пропущено всего:     {total}")
        for k, v in self.skipped.items():
            if v:
                print(f"    • {k}: {v}")
        if self.skip_samples:
            print("  Примеры пропущенных:")
            for s in self.skip_samples[:20]:
                print(f"    - {s}")
        print("  ───────────────────\n")


def log_skip(diag: Diag, reason: str, name: str, extra: str = ""):
    diag.skipped[reason] = diag.skipped.get(reason, 0) + 1
    if len(diag.skip_samples) < 50:
        diag.skip_samples.append(f"[{reason}] {name} {extra}")


# ============================================================
# КЭШ ОБОГАЩЕНИЯ
# ============================================================

def load_cache() -> dict:
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"  [!] Не удалось прочитать кэш: {e}")
    return {}


def save_cache(cache: dict):
    os.makedirs(CACHE_DIR, exist_ok=True)
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"  [!] Не удалось сохранить кэш: {e}")


def cache_key(title: str, author: str) -> str:
    raw = f"{(title or '').strip().lower()}|{(author or '').strip().lower()}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


# ============================================================
# ЯНДЕКС.ДИСК
# ============================================================

def get_public_key(public_url: str) -> str:
    m = re.search(r"/d/([A-Za-z0-9_-]+)", public_url)
    if not m:
        raise ValueError(f"Не удалось извлечь public_key из {public_url}")
    return m.group(1)


def list_public_files_recursive(client, public_key: str,
                                path: str = "/",
                                depth: int = 0,
                                max_depth: int = 30) -> list:
    """Рекурсивный обход публичной папки Яндекс.Диска."""
    result = []
    if depth > max_depth:
        print(f"    [!] Достигнута максимальная глубина {max_depth} в {path}")
        return result

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
                result.append({
                    "name":      item.name,
                    "path":      item.path,
                    "full_path": item.path,
                    "size":      getattr(item, "size", 0) or 0,
                    "ext":       ext,
                    "modified":  getattr(item, "modified", "") or "",
                })
            else:
                print(f"    [?] Неизвестный тип: {item.type} {item.name}")
        except Exception as e:
            print(f"    [!] Ошибка обработки элемента "
                  f"{getattr(item, 'name', '?')}: {e}")
            continue
    return result


def make_download_link(public_key: str, full_path: str) -> str:
    return f"https://disk.yandex.ru/d/{public_key}?path={quote(full_path)}"


# ============================================================
# ПАРСИНГ ИМЕНИ ФАЙЛА → МЕТАДАННЫЕ
# ============================================================

SEPARATORS = [" - ", " — ", " – ", " –– "]

def parse_book_name(filename: str) -> dict:
    """
    Разбирает имя файла вида:
      'АРКАДИЙ СТРУГАЦКИЙ, БОРИС СТРУГАЦКИЙ - ЗА МИЛЛИАРД ЛЕТ ДО КОНЦА СВЕТА.txt'
    Возвращает {title, author}.
    """
    stem = os.path.splitext(filename)[0].strip()

    # Чистим мусорные суффиксы
    stem = re.sub(r"\s*\(\d+\)\s*$", "", stem)
    stem = re.sub(r"\s*\[.*?\]\s*", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip()

    author = ""
    title  = stem

    for sep in SEPARATORS:
        if sep in stem:
            head, tail = stem.split(sep, 1)
            head = head.strip()
            tail = tail.strip()
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


# ============================================================
# ИСТОЧНИК 1: GOOGLE BOOKS (бесплатно)
# ============================================================

def _google_books_query(title: str, author: str, api_key: str = "") -> dict:
    q_parts = []
    if title:
        q_parts.append(f'intitle:"{title}"')
    if author:
        q_parts.append(f'inauthor:"{author}"')
    if not q_parts:
        return {}

    params = {
        "q": " ".join(q_parts),
        "maxResults": 3,
        "printType": "books",
    }
    if api_key:
        params["key"] = api_key

    try:
        r = requests.get(GOOGLE_BOOKS_API, params=params,
                         headers={"User-Agent": USER_AGENT},
                         timeout=HTTP_TIMEOUT)
        if r.status_code != 200:
            return {}
        data = r.json()
    except Exception as e:
        print(f"    [!] Google Books ошибка: {e}")
        return {}

    items = data.get("items") or []
    for it in items:
        info = it.get("volumeInfo", {})
        desc = (info.get("description") or "").strip()
        if not desc:
            continue
        return {
            "description": desc,
            "cover":       (info.get("imageLinks") or {}).get("thumbnail", ""),
            "source":      "google_books",
        }
    return {}


def google_books_lookup(title: str, author: str, api_key: str = "") -> dict:
    res = _google_books_query(title, author, api_key)
    if res:
        return res
    time.sleep(SLEEP_BETWEEN_REQUESTS)
    return _google_books_query(title, "", api_key)


# ============================================================
# ИСТОЧНИК 2: WIKIPEDIA (бесплатно, без ключа)
# ============================================================

def _wiki_request(params: dict) -> dict:
    try:
        r = requests.get(
            WIKI_API,
            params=params,
            headers={"User-Agent": USER_AGENT},
            timeout=HTTP_TIMEOUT,
        )
        if r.status_code != 200:
            return {}
        return r.json()
    except Exception as e:
        print(f"    [!] Wikipedia ошибка: {e}")
        return {}


def _wiki_search_page(query: str) -> tuple:
    """
    Ищет страницу по запросу. Возвращает (pageid, title) или (None, None).
    """
    data = _wiki_request({
        "action":   "query",
        "format":   "json",
        "list":     "search",
        "srsearch": query,
        "srlimit":  1,
        "srnamespace": 0,
    })
    hits = (data.get("query") or {}).get("search") or []
    if not hits:
        return None, None
    return hits[0].get("pageid"), hits[0].get("title")


def _wiki_get_extract(pageid: int) -> str:
    """Возвращает первые абзацы страницы."""
    data = _wiki_request({
        "action":      "query",
        "format":      "json",
        "prop":        "extracts",
        "pageids":     pageid,
        "explaintext": 1,
        "exintro":     1,
        "redirects":   1,
    })
    pages = (data.get("query") or {}).get("pages") or {}
    for _, page in pages.items():
        extract = (page.get("extract") or "").strip()
        if extract:
            return extract
    return ""


def _wiki_first_paragraphs(text: str, max_chars: int = 800) -> str:
    """Берёт первые 2–3 абзаца и режет по длине."""
    if not text:
        return ""
    paragraphs = [p.strip() for p in text.split("\n") if p.strip()]
    result = []
    total = 0
    for p in paragraphs[:3]:
        if total + len(p) > max_chars:
            break
        result.append(p)
        total += len(p)
    joined = "\n\n".join(result) if result else text[:max_chars]
    return joined.strip()


def wikipedia_lookup(title: str, author: str) -> dict:
    """
    Ищет страницу Wikipedia по названию книги.
    Пробует несколько вариантов запроса.
    """
    if not title:
        return {}

    queries = [
        f"{title} {author}".strip() if author else title,
        f"{title} (книга)",
        f"{title} (роман)",
        f"{title} (повесть)",
        f"{title} (рассказ)",
        title,
    ]

    seen_pages = set()
    for q in queries:
        if not q:
            continue
        pageid, page_title = _wiki_search_page(q)
        if not pageid or pageid in seen_pages:
            time.sleep(SLEEP_BETWEEN_REQUESTS)
            continue
        seen_pages.add(pageid)

        # Проверяем, что найденная страница — про произведение,
        # а не про что-то постороннее
        extract = _wiki_get_extract(pageid)
        if not extract:
            time.sleep(SLEEP_BETWEEN_REQUESTS)
            continue

        # Отсеиваем страницы-неоднозначности и слишком короткие
        if len(extract) < 150:
            time.sleep(SLEEP_BETWEEN_REQUESTS)
            continue

        description = _wiki_first_paragraphs(extract, max_chars=900)
        if description:
            return {
                "description": description,
                "cover":       "",
                "source":      "wikipedia",
            }
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    return {}


# ============================================================
# ИСТОЧНИК 3: OPEN LIBRARY (бесплатно, без ключа)
# ============================================================

def openlibrary_lookup(title: str, author: str) -> dict:
    params = {"title": title, "limit": 3}
    if author:
        params["author"] = author

    try:
        r = requests.get(OPENLIBRARY_API, params=params,
                         headers={"User-Agent": USER_AGENT},
                         timeout=HTTP_TIMEOUT)
        if r.status_code != 200:
            return {}
        data = r.json()
    except Exception as e:
        print(f"    [!] Open Library ошибка: {e}")
        return {}

    docs = data.get("docs") or []
    for d in docs:
        desc = ""
        fs = d.get("first_sentence")
        if isinstance(fs, list) and fs:
            desc = fs[0]
        elif isinstance(fs, str):
            desc = fs
        if not desc and d.get("subtitle"):
            desc = d["subtitle"]
        if desc:
            cover_id = d.get("cover_i") or 0
            cover = (f"https://covers.openlibrary.org/b/id/{cover_id}-L.jpg"
                     if cover_id else "")
            return {
                "description": desc.strip(),
                "cover":       cover,
                "source":      "openlibrary",
            }
    return {}


# ============================================================
# ИСТОЧНИК 4: LLM (опционально, платно)
# ============================================================

def llm_lookup(title: str, author: str) -> dict:
    if os.environ.get("ENABLE_LLM_FALLBACK", "0") != "1":
        return {}

    provider = os.environ.get("LLM_PROVIDER", "yandexgpt").lower()
    prompt = (
        "Ты библиотекарь. Кратко опиши книгу в 2–3 предложениях "
        "без спойлеров и без вступления.\n"
        f"Автор: {author or 'неизвестен'}\n"
        f"Название: {title}"
    )

    if provider == "openai":
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            return {}
        try:
            r = requests.post(
                OPENAI_URL,
                headers={
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type":  "application/json",
                },
                json={
                    "model": "gpt-4o-mini",
                    "messages": [
                        {"role": "system", "content": "Ты библиотекарь."},
                        {"role": "user",   "content": prompt},
                    ],
                    "temperature": 0.3,
                    "max_tokens": 300,
                },
                timeout=60,
            )
            if r.status_code != 200:
                print(f"    [!] OpenAI {r.status_code}: {r.text[:200]}")
                return {}
            text = r.json()["choices"][0]["message"]["content"].strip()
            return {"description": text, "cover": "", "source": "openai"}
        except Exception as e:
            print(f"    [!] OpenAI ошибка: {e}")
            return {}

    # yandexgpt по умолчанию
    api_key   = os.environ.get("YANDEX_GPT_API_KEY")
    folder_id = os.environ.get("YANDEX_FOLDER_ID")
    if not api_key or not folder_id:
        return {}
    try:
        r = requests.post(
            YANDEX_GPT_URL,
            headers={
                "Authorization": f"Api-Key {api_key}",
                "Content-Type":  "application/json",
            },
            json={
                "modelUri": f"gpt://{folder_id}/yandexgpt-lite/latest",
                "completionOptions": {"temperature": 0.3, "maxTokens": 300},
                "messages": [
                    {"role": "system", "text": "Ты библиотекарь."},
                    {"role": "user",   "text": prompt},
                ],
            },
            timeout=60,
        )
        if r.status_code != 200:
            print(f"    [!] YandexGPT {r.status_code}: {r.text[:200]}")
            return {}
        text = r.json()["result"]["alternatives"][0]["message"]["text"].strip()
        return {"description": text, "cover": "", "source": "yandexgpt"}
    except Exception as e:
        print(f"    [!] YandexGPT ошибка: {e}")
        return {}


# ============================================================
# ОБОГАЩЕНИЕ: ОРКЕСТРАТОР
# ============================================================

def enrich_book(title: str, author: str, cache: dict, diag: Diag) -> dict:
    """
    Порядок источников:
      1. Google Books    (бесплатно)
      2. Wikipedia       (бесплатно)
      3. Open Library    (бесплатно)
      4. LLM             (опционально)
    Результат кэшируется.
    """
    ck = cache_key(title, author)
    if ck in cache:
        return cache[ck]

    gb_key = os.environ.get("GOOGLE_BOOKS_API_KEY", "")
    meta = {}

    # 1. Google Books
    try:
        meta = google_books_lookup(title, author, gb_key)
    except Exception as e:
        print(f"    [!] Google Books fallback: {e}")
    time.sleep(SLEEP_BETWEEN_REQUESTS)

    # 2. Wikipedia
    if not meta.get("description"):
        try:
            meta = wikipedia_lookup(title, author)
        except Exception as e:
            print(f"    [!] Wikipedia fallback: {e}")
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    # 3. Open Library
    if not meta.get("description"):
        try:
            meta = openlibrary_lookup(title, author)
        except Exception as e:
            print(f"    [!] Open Library fallback: {e}")
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    # 4. LLM
    if not meta.get("description"):
        try:
            meta = llm_lookup(title, author)
        except Exception as e:
            print(f"    [!] LLM fallback: {e}")
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    if not meta:
        meta = {"description": "", "cover": "", "source": "none"}

    # Учёт источника
    src = meta.get("source", "none")
    diag.sources[src] = diag.sources.get(src, 0) + 1

    cache[ck] = meta
    return meta


# ============================================================
# ЗАГРУЗКА ОБЛОЖЕК В SUPABASE
# ============================================================

def upload_cover_to_supabase(cover_url: str, supabase) -> str:
    if not cover_url or not supabase:
        return ""
    try:
        r = requests.get(cover_url,
                         headers={"User-Agent": USER_AGENT},
                         timeout=HTTP_TIMEOUT)
        if r.status_code != 200:
            return ""
        data = r.content
        ext = "png" if cover_url.lower().endswith(".png") else "jpg"
        name = hashlib.sha1(cover_url.encode()).hexdigest() + "." + ext

        try:
            supabase.storage.from_(SUPABASE_BUCKET).upload(
                name, data, {"content-type": f"image/{ext}"}
            )
        except Exception:
            # Возможно, уже загружено — это нормально
            pass
        public_url = supabase.storage.from_(SUPABASE_BUCKET).get_public_url(name)
        return public_url
    except Exception as e:
        print(f"    [!] Supabase upload ошибка: {e}")
        return ""


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


def get_or_create_sheet(sh, name: str, headers: list):
    try:
        return sh.worksheet(name)
    except gspread.WorksheetNotFound:
        print(f"  [!] Лист '{name}' не найден, создаю...")
        sheet = sh.add_worksheet(title=name, rows=2000, cols=len(headers) + 2)
        sheet.append_row(headers)
        return sheet


def load_existing_keys(sheet) -> set:
    try:
        rows = sheet.get_all_values()
    except Exception:
        return set()
    if not rows:
        return set()
    header = [h.strip().lower() for h in rows[0]]
    idx = None
    for cand in ("download_link", "ссылка", "link"):
        if cand in header:
            idx = header.index(cand)
            break
    if idx is None:
        return set()
    keys = set()
    for r in rows[1:]:
        if len(r) > idx and r[idx].strip():
            keys.add(r[idx].strip())
    return keys


def append_rows_safe(sheet, rows: list, batch_size: int = 200):
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
                print(f"    [!] Попытка {attempt+1} батча "
                      f"{i//batch_size} не удалась: {e}")
                time.sleep(2 * (attempt + 1))
        else:
            failed += len(batch)
        time.sleep(1)
    return added, failed


# ============================================================
# СБОРКА СТРОК ДЛЯ ЛИСТА
# ============================================================

def build_rows_for_section(section: str, files: list, public_key: str,
                           cache: dict, supabase, diag: Diag) -> list:
    rows = []
    seen = set()

    for f in files:
        name = f["name"]
        ext  = f["ext"]

        # 1. Фильтр по расширению
        if section == "Книги" and ALLOWED_EXTS and ext not in ALLOWED_EXTS:
            log_skip(diag, "bad_ext", name, f"(ext={ext})")
            continue

        # 2. Парсинг имени
        meta = parse_book_name(name)
        title  = meta["title"]
        author = meta["author"]

        if not title:
            log_skip(diag, "empty_title", name)
            continue

        # 3. Ссылка
        link = make_download_link(public_key, f["full_path"])
        if not link:
            log_skip(diag, "empty_link", name)
            continue

        # 4. Дедупликация
        key = (title.lower(), author.lower())
        if key in seen:
            log_skip(diag, "dup_title", name, f"(title={title})")
            continue
        seen.add(key)

        # 5. Обогащение
        enriched = {"description": "", "cover": "", "source": "none"}
        if section == "Книги":
            enriched = enrich_book(title, author, cache, diag)
            if enriched.get("description"):
                diag.enriched += 1

        # 6. Обложка через Supabase
        cover = enriched.get("cover", "")
        if cover and supabase:
            cover = upload_cover_to_supabase(cover, supabase) or cover

        # 7. Метаданные
        fmt = ext.lstrip(".")
        size_mb = round(f["size"] / (1024 * 1024), 1) if f["size"] else 0
        parts = f["full_path"].strip("/").split("/")
        folder = parts[0] if len(parts) >= 2 else ""

        record = {
            "title":         title,
            "author":        author,
            "format":        fmt,
            "size":          str(size_mb),
            "download_link": link,
            "cover":         cover,
            "folder":        folder,
            "description":   enriched.get("description", ""),
            "version":       "",
        }
        headers = SHEET_HEADERS.get(section, SHEET_HEADERS["Книги"])
        rows.append([record.get(h, "") for h in headers])

    return rows


# ============================================================
# СИНХРОНИЗАЦИЯ ОДНОГО РАЗДЕЛА
# ============================================================

def sync_section(section: str, public_url: str, gs_client,
                 cache: dict, supabase, diag: Diag):
    print(f"\n=== Раздел: {section} ===")
    if not public_url:
        print("  Пустая ссылка, пропускаю.")
        return
    print(f"Источник: {public_url}")

    if yadisk is None:
        print("  [!] yadisk не установлен, пропускаю.")
        return

    public_key = get_public_key(public_url)
    token = os.environ.get("YADISK_TOKEN")
    if not token:
        print("  [!] YADISK_TOKEN не задан, пропускаю.")
        return

    with yadisk.Client(token=token) as client:
        try:
            if not client.check_token():
                print("  [!] Неверный YADISK_TOKEN")
                return
        except Exception as e:
            print(f"  [!] Ошибка проверки токена: {e}")
            return

        files = list_public_files_recursive(client, public_key)

    diag.found = len(files)
    print(f"  Найдено файлов: {len(files)}")

    rows = build_rows_for_section(section, files, public_key,
                                  cache, supabase, diag)
    diag.kept = len(rows)
    diag.report()

    if not rows:
        print("  Нет строк для записи.")
        return

    try:
        sh = gs_client.open(SPREADSHEET_NAME)
    except Exception as e:
        print(f"  [!] Не удалось открыть таблицу '{SPREADSHEET_NAME}': {e}")
        return

    headers = SHEET_HEADERS.get(section, SHEET_HEADERS["Книги"])
    sheet = get_or_create_sheet(sh, section, headers)

    existing = load_existing_keys(sheet)
    print(f"  Существующих строк с download_link: {len(existing)}")

    link_idx = headers.index("download_link")
    new_rows = []
    upd_count = 0
    for row in rows:
        if row[link_idx] in existing:
            upd_count += 1
            continue
        new_rows.append(row)

    print(f"    Совпало с существующими: {upd_count}")
    print(f"    Новых строк к записи:     {len(new_rows)}")

    added, failed = append_rows_safe(sheet, new_rows)
    print(f"    Добавлено новых строк:    {added}")
    if failed:
        print(f"    [!] Не удалось записать:  {failed}")


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("yandex_disk_sync.py — старт")
    print("=" * 60)

    if yadisk is None:
        print("[!] Модуль yadisk не установлен: pip install -r requirements.txt")
        return

    print("Подключение к Google Sheets...")
    try:
        gs_client = get_gspread_client()
    except Exception as e:
        print(f"[!] Не удалось подключиться к Google Sheets: {e}")
        return
    print("Клиент создан.")

    print("Загрузка кэша обогащения...")
    cache = load_cache()
    print(f"  Записей в кэше: {len(cache)}")

    supabase = None
    if supa_create_client:
        key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
        if key:
            try:
                supabase = supa_create_client(SUPABASE_URL, key)
                print("Supabase подключён.")
            except Exception as e:
                print(f"[!] Supabase недоступен: {e}")

    if os.environ.get("ENABLE_LLM_FALLBACK", "0") == "1":
        print(f"LLM fallback включён (провайдер: "
              f"{os.environ.get('LLM_PROVIDER', 'yandexgpt')})")

    grand_found    = 0
    grand_kept     = 0
    grand_enriched = 0

    for section, url in SECTIONS.items():
        diag = Diag()
        try:
            sync_section(section, url, gs_client, cache, supabase, diag)
        except Exception as e:
            print(f"\n[!!!] Ошибка в разделе {section}: {e}")
            traceback.print_exc()
        grand_found    += diag.found
        grand_kept     += diag.kept
        grand_enriched += diag.enriched

    print("\nСохранение кэша обогащения...")
    save_cache(cache)
    print(f"  Кэш сохранён: {len(cache)} записей")

    print("\n" + "=" * 60)
    print("ИТОГО по всем разделам:")
    print(f"  Найдено файлов:  {grand_found}")
    print(f"  Записано:        {grand_kept}")
    print(f"  Обогащено:       {grand_enriched}")
    print("=" * 60)
    print("Готово!")


if __name__ == "__main__":
    main()
