# common.py
# Общие константы и функции для всех скриптов проекта.

import os
import re
import json
import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# КОНСТАНТЫ
# ============================================================

SPREADSHEET_ID = os.environ.get(
    "SPREADSHEET_ID",
    "1kcG0TG4GZtSM2mypjgvNDUpIbLfIvcmW80_hBKA11nw",   # ← замените на реальный ID
)

SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME", "НаполнениеСайта")

CREDENTIALS_FILE = os.environ.get("GOOGLE_CREDENTIALS_FILE", "credentials.json")


SHEET_CONFIG = {
    "Книги":     {"json": "_content/books.json"},
    "Программы": {"json": "_content/programs.json"},
    "Музыка":    {"json": "_content/music.json"},
    "Игры":      {"json": "_content/games.json"},
    "Статьи":    {"json": "_content/articles.json"},
    "Фильмы":    {"json": "_content/movies.json"},
    "Разное":    {"json": "_content/misc.json"},
    "Новости":   {"json": "_content/news.json"},
}


# Соответствие русских заголовков колонок в Sheets → английским ключам в JSON.
# Поддерживаем несколько вариантов для одного ключа, чтобы старые таблицы
# не ломались.
COLUMN_MAPPING = {
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
    "Текст":                 "text",
    "Дата":                  "date",
    "Категория":             "category",
    "Теги":                  "tags",
}


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


# ============================================================
# УТИЛИТЫ
# ============================================================

def normalize(s) -> str:
    """Нормализация строки для сравнения названий."""
    if not s:
        return ""
    return (
        str(s).strip().lower()
        .replace("ё", "е")
        .replace(" ", "")
        .replace("-", "")
        .replace("_", "")
    )


# ============================================================
# АВТОРИЗАЦИЯ GOOGLE
# ============================================================

def _load_credentials_dict() -> dict:
    raw = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if raw:
        try:
            return json.loads(raw)
        except json.JSONDecodeError as e:
            raise RuntimeError(
                f"GOOGLE_CREDENTIALS_JSON содержит невалидный JSON: {e}"
            )

    if os.path.exists(CREDENTIALS_FILE):
        with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)

    raise RuntimeError(
        "Не найдены credentials ни в GOOGLE_CREDENTIALS_JSON, "
        f"ни в файле '{CREDENTIALS_FILE}'."
    )


def get_gspread_client():
    creds_dict = _load_credentials_dict()
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)
