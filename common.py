# common.py
# Общие константы и функции для всех Python-скриптов MyFiles

import os
import re
import sys
import json

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
}

SHEET_CONFIG = {
    "Новости":  {"json": "_content/news.json",     "folder": "_content/news"},
    "Программы": {"json": "_content/programs.json", "folder": "_content/programs"},
    "Книги":    {"json": "_content/books.json",    "folder": "_content/books"},
    "Музыка":   {"json": "_content/music.json",    "folder": "_content/music"},
    "Игры":     {"json": "_content/games.json",    "folder": "_content/games"},
    "Статьи":   {"json": "_content/articles.json", "folder": "_content/articles"},
    "Фильмы":   {"json": "_content/movies.json",   "folder": "_content/movies"},
    "Разное":   {"json": "_content/misc.json",     "folder": "_content/misc"},
}


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
    """Преобразует название в slug для имени .md-файла."""
    slug = re.sub(r'[^\w\s-]', '', title).strip().lower()
    slug = re.sub(r'[-\s]+', '-', slug)
    return slug
