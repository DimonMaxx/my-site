#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
yandex_disk_sync.py
Синхронизация Яндекс.Диска → Google Sheets с обогащением описаний книг.

Источники описаний (в порядке приоритета):
    1. Google Books API
    2. FantLab API (fantlab.ru)
    3. Wikipedia API
    4. Open Library API
    5. LLM (YandexGPT / OpenAI) — опционально
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

SPREADSHEET_ID   = os.environ.get("SPREADSHEET_ID", "")
SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME", "НаполнениеСайта")


# ВАЖНО: если публичная ссылка ведёт на КОРЕНЬ диска, укажите
# относительный путь до нужной папки в поле "path".
# Например: {"url": "...", "path": "/Книги"}
#
# Если ссылка ведёт уже на нужную папку — оставьте "/".
SECTIONS = {
    "Книги": {
        "url":  "https://disk.yandex.ru/d/zMxF4nXHPkIVCQ",
        "path": "/Книги",
    },
    "Программы": {
        "url":  "https://disk.yandex.ru/d/EjUHvm6mUcgVMw",
        "path": "/",
    },
    "Музыка":  {"url": "", "path": "/"},
    "Игры":    {"url": "", "path": "/"},
    "Статьи":  {"url": "", "path": "/"},
    "Фильмы":  {"url": "", "path": "/"},
    "Разное":  {"url": "", "path": "/"},
    "Новости": {"url": "", "path": "/"},
}


# РУССКИЕ заголовки — их видит пользователь в Google Sheets,
# и их читает generate_from_sheets.py через COLUMN_MAPPING.
SHEET_HEADERS = {
    "Книги":     ["Название", "Автор", "Формат", "Размер (МБ)",
                  "Ссылка для скачивания", "Обложка", "Папка", "Описание"],
    "Программы": ["Название", "Описание", "Версия", "Размер (МБ)",
                  "Ссылка для скачивания", "Папка"],
}


# Обратное соответствие (RU → EN) — для сборки строк
RU_TO_EN = {
    "Название":              "title",
    "Автор":                 "author",
    "Описание":              "description",
    "Формат":                "format",
    "Размер (МБ)":           "size",
    "Размер":                "size",
    "Ссылка для скачивания": "download_link",
    "Ссылка":                "download_link",
    "Обложка":               "cover",
    "Папка":                 "folder",
    "Версия":                "version",
    "Дата":                  "date",
    "Категория":             "category",
    "Теги":                  "tags",
}


ALLOWED_EXTS = {".fb2", ".epub", ".pdf", ".txt", ".djvu", ".mobi", ".azw3"}

CACHE_DIR  = "_cache"
CACHE_FILE = os.path.join(CACHE_DIR, "enrichment_cache.json")

SUPABASE_URL    = "https://rmoonebbvpmvthvpcmpt.supabase.co"
SUPABASE_BUCKET = "covers"

GOOGLE_BOOKS_API = "https://www.googleapis.com/books/v1/volumes"

FANTLAB_API        = "https://api.fantlab.ru"
FANTLAB_SEARCH_URL = f"{FANTLAB_API}/search-works"
FANTLAB_WORK_URL   = f"{FANTLAB_API}/work/{{work_id}}/extended"

WIKI_LANG = os.environ.get("WIKI_LANG", "ru")
WIKI_API  = f"https://{WIKI_LANG}.wikipedia.org/w/api.php"

OPENLIBRARY_API = "https://openlibrary.org/search.json"

YANDEX_GPT_URL = "https://llm.api.cloud.yandex.net/foundationModels/v1/completion"
OPENAI_URL     = "https://api.openai.com/v1/chat/completions"

HTTP_TIMEOUT = 20
SLEEP_BETWEEN_REQUESTS = 0.4
SLEEP_BEFORE_LISTDIR   = 1.5

USER_AGENT = "ContentSyncBot/1.0 (https://github.com/DimonMaxx/my-site)"

BOOK_MARKERS = (
    "книга", "роман", "повесть", "рассказ", "произведение",
    "автор", "писател", "фантаст", "литератур", "трилогия",
    "novel", "book", "story", "written by", "author",
    "сборник", "эпопея", "цикл",
)

# Символы CP866-«псевдографики». Буквы Ё здесь быть не должно!
_CP866_JUNK = set("╞░─┘╦╪╟┌┐└┴┬├┤│╫╬═║╔╗╚╝")


# ============================================================
# ДИАГНОСТИКА
# ============================================================

class Diag:
    def __init__(self):
        self.found    = 0
        self.kept     = 0
        self.enriched = 0
        self.sources  = {}
        self.skipped  = {
            "empty_title": 0,
            "empty_link":  0,
            "dup_title":   0,
            "bad_ext":     0,
            "bad_name":    0,
            "garbage":     0,
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
            print("  Источники описаний:")
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
# КЭШ
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
# ВАЛИДАЦИЯ ОБОГАЩЕНИЯ
# ============================================================

def _looks_like_book_article(desc: str, title: str, author: str) -> bool:
    text = (desc or "").lower()
    if not text:
        return False
    if any(m in text for m in BOOK_MARKERS):
        return True
    if author:
        surname = author.strip().split()[-1].lower()
        if len(surname) > 3 and surname in text:
            return True
    if title and title.lower() in text[:400]:
        return True
    return False


def _title_words_overlap(q_title: str, found_title: str) -> bool:
    q = set(re.findall(r"\w+", (q_title or "").lower()))
    f = set(re.findall(r"\w+", (found_title or "").lower()))
    q = {w for w in q if len(w) > 3}
    if not q:
        return False
    return len(q & f) / len(q) >= 0.5


def _author_mentioned(author: str, text: str) -> bool:
    if not author:
        return False
    surname = author.strip().split()[-1].lower()
    return len(surname) > 3 and surname in (text or "").lower()


def _result_is_acceptable(meta: dict, title: str, author: str) -> bool:
    if not isinstance(meta, dict):
        return False
    desc = (meta.get("description") or "").strip()
    if not desc:
        return True

    src = (meta.get("source") or "none").lower()

    if src in ("google_books", "fantlab", "openai", "yandexgpt", "none"):
        return True

    if src.startswith("wikipedia"):
        return _looks_like_book_article(desc, title, author)

    if src == "openlibrary":
        return (
            _title_words_overlap(title, meta.get("title", ""))
            or _author_mentioned(author, desc)
        )

    return True


# ============================================================
# ЯНДЕКС.ДИСК
# ============================================================

def get_public_key(public_url: str) -> str:
    m = re.search(r"/d/([A-Za-z0-9_-]+)", public_url)
    if not m:
        raise ValueError(f"Не удалось извлечь public_key из {public_url}")
    return m.group(1)


def _strip_disk_prefix(path: str) -> str:
    """Убирает префикс 'disk:' из пути, который возвращает yadisk."""
    if path and path.startswith("disk:"):
        return path[len("disk:"):]
    return path or "/"


def list_public_files_recursive(client, public_key: str,
                                path: str = "/",
                                depth: int = 0,
                                max_depth: int = 30) -> list:
    result = []
    if depth > max_depth:
        print(f"    [!] Достигнута максимальная глубина {max_depth} в {path}")
        return result

    # Принудительная пауза на верхнем уровне — чтобы Яндекс.Диск
    # не отдал закэшированный результат.
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


def print_root_diagnostics(client, public_key: str, path: str = "/"):
    """Печатает первые 20 элементов указанной папки — чтобы понять структуру."""
    try:
        time.sleep(SLEEP_BEFORE_LISTDIR)
        items = list(client.listdir(path, public_key=public_key))
    except Exception as e:
        print(f"  [!] Не удалось получить содержимое {path}: {e}")
        return
    print(f"  Содержимое '{path}' ({len(items)} элементов):")
    for it in items[:20]:
        kind = "DIR " if it.type == "dir" else "FILE"
        print(f"    [{kind}] {it.name}")
    if len(items) > 20:
        print(f"    ... и ещё {len(items) - 20}")


# ============================================================
# ПАРСИНГ ИМЕНИ
# ============================================================

SEPARATORS = [" - ", " — ", " – ", " –– "]


def _is_garbage_stem(stem: str) -> bool:
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


def parse_book_name(filename: str) -> dict:
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
# ИСТОЧНИК 1: GOOGLE BOOKS
# ============================================================

def _google_books_query(title: str, author: str, api_key: str = "") -> dict:
    q_parts = []
    if title:
        q_parts.append(f'intitle:"{title}"')
    if author:
        q_parts.append(f'inauthor:"{author}"')
    if not q_parts:
        return {}

    params = {"q": " ".join(q_parts), "maxResults": 3, "printType": "books"}
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
            "title":       info.get("title", ""),
        }
    return {}


def google_books_lookup(title: str, author: str, api_key: str = "") -> dict:
    res = _google_books_query(title, author, api_key)
    if res:
        return res
    time.sleep(SLEEP_BETWEEN_REQUESTS)
    return _google_books_query(title, "", api_key)


# ============================================================
# ИСТОЧНИК 2: FANTLAB
# ============================================================

def _fantlab_search_works(query: str, limit: int = 5) -> list:
    params = {"q": query, "onlymatches": 1}
    try:
        r = requests.get(
            FANTLAB_SEARCH_URL,
            params=params,
            headers={"User-Agent": USER_AGENT},
            timeout=HTTP_TIMEOUT,
        )
        if r.status_code != 200:
            return []
        data = r.json()
    except Exception as e:
        print(f"    [!] FantLab search ошибка: {e}")
        return []

    if isinstance(data, dict):
        matches = data.get("matches") or []
    elif isinstance(data, list):
        matches = data
    else:
        matches = []

    return matches[:limit]


def _fantlab_get_work(work_id: int) -> dict:
    try:
        r = requests.get(
            FANTLAB_WORK_URL.format(work_id=work_id),
            headers={"User-Agent": USER_AGENT},
            timeout=HTTP_TIMEOUT,
        )
        if r.status_code != 200:
            return {}
        return r.json() or {}
    except Exception as e:
        print(f"    [!] FantLab work ошибка: {e}")
        return {}


def _fantlab_pick_best(matches: list, title: str, author: str) -> dict:
    title_l = (title or "").lower()
    author_l = (author or "").lower()

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
        elif title_l and any(
            w in names for w in title_l.split() if len(w) > 3
        ):
            score += 1

        authors_str = " ".join(filter(None, [
            m.get("all_autor_rusname", ""),
            m.get("autor1_rusname", ""),
            m.get("autor2_rusname", ""),
            m.get("autor3_rusname", ""),
        ])).lower()

        if author_l:
            surname = author_l.split()[-1]
            if len(surname) > 3 and surname in authors_str:
                score += 3

        if score > best_score:
            best_score = score
            best = m

    if best is None and matches:
        best = max(matches, key=lambda x: x.get("weight", 0) or 0)

    return best or {}


def _clean_fantlab_text(text: str) -> str:
    if not text:
        return ""
    text = re.sub(r"<a[^>]*>(.*?)</a>", r"\1", text, flags=re.DOTALL)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\[/?[a-zA-Z_]+\]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def fantlab_lookup(title: str, author: str) -> dict:
    if not title:
        return {}

    query_parts = [title]
    if author:
        query_parts.append(author)
    query = " ".join(query_parts)

    matches = _fantlab_search_works(query, limit=5)
    time.sleep(SLEEP_BETWEEN_REQUESTS)

    if not matches:
        matches = _fantlab_search_works(title, limit=5)
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    if not matches:
        return {}

    best = _fantlab_pick_best(matches, title, author)
    work_id = best.get("work_id")
    if not work_id:
        return {}

    work = _fantlab_get_work(int(work_id))
    time.sleep(SLEEP_BETWEEN_REQUESTS)

    if not work:
        return {}

    desc = _clean_fantlab_text(
        work.get("work_description") or work.get("work_description_author") or ""
    )
    if not desc or len(desc) < 50:
        return {}

    cover = ""
    img = work.get("image") or {}
    if isinstance(img, dict):
        cover = img.get("url") or ""
        if cover and not cover.startswith("http"):
            cover = "https://fantlab.ru" + cover

    return {
        "description": desc,
        "cover":       cover,
        "source":      "fantlab",
        "title":       work.get("work_name_orig") or work.get("work_name") or "",
    }


# ============================================================
# ИСТОЧНИК 3: WIKIPEDIA
# ============================================================

def _wiki_request(params: dict) -> dict:
    try:
        r = requests.get(
            WIKI_API, params=params,
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
    data = _wiki_request({
        "action": "query", "format": "json", "list": "search",
        "srsearch": query, "srlimit": 1, "srnamespace": 0,
    })
    hits = (data.get("query") or {}).get("search") or []
    if not hits:
        return None, None
    return hits[0].get("pageid"), hits[0].get("title")


def _wiki_get_extract(pageid: int) -> str:
    data = _wiki_request({
        "action": "query", "format": "json", "prop": "extracts",
        "pageids": pageid, "explaintext": 1, "exintro": 1, "redirects": 1,
    })
    pages = (data.get("query") or {}).get("pages") or {}
    for _, page in pages.items():
        extract = (page.get("extract") or "").strip()
        if extract:
            return extract
    return ""


def _wiki_first_paragraphs(text: str, max_chars: int = 800) -> str:
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
    if not title:
        return {}

    queries = []
    if author:
        queries.append(f"{title} {author} роман")
    queries += [
        f"{title} (роман)", f"{title} (книга)", f"{title} (повесть)",
        f"{title} (рассказ)", f"{title} (литературное произведение)",
    ]
    if author:
        queries.append(f"{title} {author}")

    seen_pages = set()
    for q in queries:
        if not q:
            continue
        pageid, page_title = _wiki_search_page(q)
        if not pageid or pageid in seen_pages:
            time.sleep(SLEEP_BETWEEN_REQUESTS)
            continue
        seen_pages.add(pageid)

        extract = _wiki_get_extract(pageid)
        if not extract or len(extract) < 150:
            time.sleep(SLEEP_BETWEEN_REQUESTS)
            continue

        if not _looks_like_book_article(extract, title, author):
            time.sleep(SLEEP_BETWEEN_REQUESTS)
            continue

        description = _wiki_first_paragraphs(extract, max_chars=900)
        if description:
            return {
                "description": description,
                "cover":       "",
                "source":      "wikipedia",
                "title":       page_title or "",
            }
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    return {}


# ============================================================
# ИСТОЧНИК 4: OPEN LIBRARY
# ============================================================

def openlibrary_lookup(title: str, author: str) -> dict:
    params = {"title": title, "limit": 5}
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
        found_title = d.get("title", "") or ""
        if not _title_words_overlap(title, found_title):
            authors_list = d.get("author_name") or []
            if not any(_author_mentioned(author, " ".join(authors_list)) for _ in [0]):
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

        cover_id = d.get("cover_i") or 0
        cover = (f"https://covers.openlibrary.org/b/id/{cover_id}-L.jpg"
                 if cover_id else "")
        return {
            "description": desc.strip(),
            "cover":       cover,
            "source":      "openlibrary",
            "title":       found_title,
        }
    return {}


# ============================================================
# ИСТОЧНИК 5: LLM
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
# ОБОГАЩЕНИЕ
# ============================================================

def enrich_book(title: str, author: str, cache: dict, diag: Diag) -> dict:
    ck = cache_key(title, author)

    cached = cache.get(ck)
    if cached is not None and _result_is_acceptable(cached, title, author):
        src = cached.get("source", "none")
        diag.sources[src] = diag.sources.get(src, 0) + 1
        return cached

    gb_key = os.environ.get("GOOGLE_BOOKS_API_KEY", "")
    meta = {}

    try:
        meta = google_books_lookup(title, author, gb_key)
    except Exception as e:
        print(f"    [!] Google Books fallback: {e}")
    time.sleep(SLEEP_BETWEEN_REQUESTS)

    if not _result_is_acceptable(meta, title, author):
        try:
            meta = fantlab_lookup(title, author)
        except Exception as e:
            print(f"    [!] FantLab fallback: {e}")
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    if not _result_is_acceptable(meta, title, author):
        try:
            meta = wikipedia_lookup(title, author)
        except Exception as e:
            print(f"    [!] Wikipedia fallback: {e}")
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    if not _result_is_acceptable(meta, title, author):
        try:
            meta = openlibrary_lookup(title, author)
        except Exception as e:
            print(f"    [!] Open Library fallback: {e}")
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    if not _result_is_acceptable(meta, title, author):
        try:
            meta = llm_lookup(title, author)
        except Exception as e:
            print(f"    [!] LLM fallback: {e}")
        time.sleep(SLEEP_BETWEEN_REQUESTS)

    if meta and not _result_is_acceptable(meta, title, author):
        meta = {}

    if not meta:
        meta = {"description": "", "cover": "", "source": "none"}

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


def open_spreadsheet(gs_client):
    if SPREADSHEET_ID:
        print(f"Открываю таблицу по ID: {SPREADSHEET_ID}")
        return gs_client.open_by_key(SPREADSHEET_ID)
    print(f"SPREADSHEET_ID не задан, открываю по имени: {SPREADSHEET_NAME}")
    return gs_client.open(SPREADSHEET_NAME)


def get_or_create_sheet(sh, name: str):
    try:
        return sh.worksheet(name)
    except gspread.WorksheetNotFound:
        print(f"  [!] Лист '{name}' не найден, создаю...")
        sheet = sh.add_worksheet(title=name, rows=2000, cols=12)
        return sheet


def ensure_headers(sheet, headers: list):
    """
    Проверяет, что первая строка листа = headers.
    Если заголовки отличаются — перезаписывает первую строку.
    Данные ниже не трогает.
    """
    try:
        current = sheet.row_values(1)
    except Exception:
        current = []
    current_clean = [c.strip() for c in current if c is not None]
    headers_clean = [h.strip() for h in headers]

    if current_clean[:len(headers_clean)] == headers_clean:
        return

    end_col_letter = chr(ord('A') + len(headers) - 1) if len(headers) <= 26 else "Z"
    range_a1 = f"A1:{end_col_letter}1"
    try:
        # Новый синтаксис gspread 6.x: сначала values, потом range_name
        sheet.update(
            values=[headers],
            range_name=range_a1,
            value_input_option="USER_ENTERED",
        )
        print(f"    [+] Обновлены заголовки: {headers}")
    except Exception as e:
        print(f"    [!] Не удалось обновить заголовки: {e}")


def load_existing_keys(sheet) -> set:
    try:
        rows = sheet.get_all_values()
    except Exception:
        return set()
    if not rows:
        return set()

    header = [h.strip().lower() for h in rows[0]]
    idx = None
    for cand in ("ссылка для скачивания", "ссылка", "download_link", "link"):
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
# СБОРКА СТРОК
# ============================================================

def build_rows_for_section(section: str, files: list, public_key: str,
                           cache: dict, supabase, diag: Diag) -> list:
    headers = SHEET_HEADERS.get(section)
    if headers is None:
        print(f"  [!] Нет заголовков для раздела '{section}', пропускаю.")
        return []

    rows = []
    seen = set()

    for f in files:
        name = f["name"]
        ext  = f["ext"]

        if section == "Книги" and ALLOWED_EXTS and ext not in ALLOWED_EXTS:
            log_skip(diag, "bad_ext", name, f"(ext={ext})")
            continue

        meta = parse_book_name(name)
        title  = meta["title"]
        author = meta["author"]

        if not title:
            log_skip(diag, "garbage", name)
            continue

        link = make_download_link(public_key, f["full_path"])
        if not link:
            log_skip(diag, "empty_link", name)
            continue

        key = (title.lower(), author.lower())
        if key in seen:
            log_skip(diag, "dup_title", name, f"(title={title})")
            continue
        seen.add(key)

        enriched = {"description": "", "cover": "", "source": "none"}
        if section == "Книги":
            enriched = enrich_book(title, author, cache, diag)
            if enriched.get("description"):
                diag.enriched += 1

        cover = enriched.get("cover", "")
        if cover and supabase:
            cover = upload_cover_to_supabase(cover, supabase) or cover

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
            "date":          "",
            "category":      "",
            "tags":          "",
        }

        row = []
        for ru in headers:
            en = RU_TO_EN.get(ru, ru)
            row.append(record.get(en, ""))
        rows.append(row)

    return rows


# ============================================================
# СИНХРОНИЗАЦИЯ РАЗДЕЛА
# ============================================================

def sync_section(section: str, section_cfg: dict, gs_client,
                 cache: dict, supabase, diag: Diag):
    print(f"\n=== Раздел: {section} ===")

    public_url = section_cfg.get("url", "")
    start_path = section_cfg.get("path", "/") or "/"

    if not public_url:
        print("  Пустая ссылка, пропускаю.")
        return

    if section not in SHEET_HEADERS:
        print(f"  Раздел '{section}' не настроен, пропускаю.")
        return

    print(f"Источник:   {public_url}")
    print(f"Подпапка:   {start_path}")

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

        # Диагностика: что лежит на верхнем уровне
        print_root_diagnostics(client, public_key, "/")

        # Если задан start_path != "/" — диагностируем и его
        if start_path != "/":
            print_root_diagnostics(client, public_key, start_path)

        files = list_public_files_recursive(client, public_key, path=start_path)

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
        sh = open_spreadsheet(gs_client)
    except Exception as e:
        print(f"  [!] Не удалось открыть таблицу: {e}")
        return

    sheet = get_or_create_sheet(sh, section)
    headers = SHEET_HEADERS[section]
    ensure_headers(sheet, headers)

    existing = load_existing_keys(sheet)
    print(f"  Существующих строк с ссылкой: {len(existing)}")

    link_ru = "Ссылка для скачивания"
    link_idx = headers.index(link_ru) if link_ru in headers else 0

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

    if not SPREADSHEET_ID:
        print("[!] SPREADSHEET_ID не задан — открытие по имени может не работать.")

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

    for section, cfg in SECTIONS.items():
        diag = Diag()
        try:
            sync_section(section, cfg, gs_client, cache, supabase, diag)
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
