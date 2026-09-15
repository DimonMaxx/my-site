# common.py
# Общие константы и функции для всех Python-скриптов MyFiles

import os
import re
import sys
import json
import hashlib

import gspread

# ============ Константы ============
SPREADSHEET_ID = "1kcG0TG4GZtSM2mypjgvNDUpIbLfIvcmW80_hBKA11nw"

SECTION_TO_SHEET = {
    "programs": "Программы",
    "books":    "Книги",
    "news":     "Новости",
    "articles": "Статьи",
    "movies":   "Фильмы",
    "music":    "Музыка",
    "games":    "Игры",
    "misc":     "Разное",
}

COLUMN_MAPPING = {
    "Название": "title",
    "Описание": "description",
    "Версия": "version",
    "Размер (МБ)": "size",
    "Ссылка для скачивания": "download_link",
    "Автор": "author",
    "Формат": "format",
    "Год": "year",
    "Платформа": "platform",
    "Текст": "body",
    "Обложка": "cover",
    "Папка": "folder",
}

SHEET_CONFIG = {
    "Новости":   {"json": "_content/news.json",     "folder": "_content/news"},
    "Программы": {"json": "_content/programs.json", "folder": "_content/programs"},
    "Книги":     {"json": "_content/books.json",    "folder": "_content/books"},
    "Музыка":    {"json": "_content/music.json",    "folder": "_content/music"},
    "Игры":      {"json": "_content/games.json",    "folder": "_content/games"},
    "Статьи":    {"json": "_content/articles.json", "folder": "_content/articles"},
    "Фильмы":    {"json": "_content/movies.json",   "folder": "_content/movies"},
    "Разное":    {"json": "_content/misc.json",     "folder": "_content/misc"},
}

# Лимит длины slug в символах (не байтах — на безопасной стороне).
# Реальное имя файла = slug + ".md" + возможный префикс, поэтому 80 символов с запасом.
MAX_SLUG_LENGTH = 80


# ============ Google Sheets ============
def get_gspread_client():
    """Возвращает клиент gspread. Использует GOOGLE_CREDENTIALS_JSON или credentials.json."""
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


# ============ Утилиты ============
def normalize(name):
    """Нормализует название: убирает расширение, скобки, регистр, лишние символы."""
    if not name:
        return ''
    name = os.path.splitext(name)[0]
    name = re.sub(r'\s*\([^)]*\)\s*$', '', name)
    name = name.strip().lower()
    name = re.sub(r'\s+', ' ', name)
    name = re.sub(r'[—–]', '-', name)
    name = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', name)
    name = re.sub(r'[^\w\s\-]', ' ', name)
    return re.sub(r'\s+', ' ', name).strip()


def slugify(title):
    """
    Преобразует название в slug для имени .md-файла.
    - Транслитерирует через unicodedata (если получится).
    - Обрезает до MAX_SLUG_LENGTH символов.
    - Если обрезали — добавляет короткий хеш от оригинала, чтобы избежать коллизий.
    """
    if not title:
        return 'untitled'

    # Базовая транслитерация кириллицы в латиницу
    TRANSLIT = {
        'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e',
        'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'i', 'к': 'k', 'л': 'l', 'м': 'm',
        'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
        'ф': 'f', 'х': 'h', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'shch',
        'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya',
        'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Е': 'E', 'Ё': 'E',
        'Ж': 'Zh', 'З': 'Z', 'И': 'I', 'Й': 'I', 'К': 'K', 'Л': 'L', 'М': 'M',
        'Н': 'N', 'О': 'O', 'П': 'P', 'Р': 'R', 'С': 'S', 'Т': 'T', 'У': 'U',
        'Ф': 'F', 'Х': 'H', 'Ц': 'Ts', 'Ч': 'Ch', 'Ш': 'Sh', 'Щ': 'Shch',
        'Ъ': '', 'Ы': 'Y', 'Ь': '', 'Э': 'E', 'Ю': 'Yu', 'Я': 'Ya',
    }

    slug = ''
    for ch in title:
        if ch.lower() in TRANSLIT:
            slug += TRANSLIT.get(ch, ch)
        else:
            slug += ch

    # Приводим к нижнему регистру, оставляем только буквы/цифры/дефисы
    slug = slug.lower()
    slug = re.sub(r'[^a-z0-9\s\-]', '', slug)
    slug = re.sub(r'[\s\-]+', '-', slug).strip('-')

    if not slug:
        slug = 'item'

    # Обрезка
    if len(slug) > MAX_SLUG_LENGTH:
        suffix = hashlib.md5(title.encode('utf-8')).hexdigest()[:8]
        slug = slug[:MAX_SLUG_LENGTH - 9].rstrip('-') + '-' + suffix

    return slug
