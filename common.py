# common.py
# Общие константы и функции для всех скриптов проекта.

import os
import json
import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# КОНСТАНТЫ
# ============================================================

# ID Google-таблицы (из URL: /spreadsheets/d/<ID>/edit)
SPREADSHEET_ID = os.environ.get(
    "SPREADSHEET_ID",
    "1AbCdEfGhIjKlMnOpQrStUvWxYz1234567890",   # ← замените на ваш реальный ID
)

# Путь к JSON сервисного аккаунта (fallback, если нет env)
CREDENTIALS_FILE = os.environ.get("GOOGLE_CREDENTIALS_FILE", "credentials.json")

# Конфиг листов: имя листа → путь к JSON-файлу и правило
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

# Соответствие русских заголовков колонок в Sheets → английским ключам в JSON
COLUMN_MAPPING = {
    "Название":       "title",
    "Автор":          "author",
    "Описание":       "description",
    "Формат":         "format",
    "Размер":         "size",
    "Ссылка":         "download_link",
    "Обложка":        "cover",
    "Папка":          "folder",
    "Версия":         "version",
    "Текст":          "text",
    "Дата":           "date",
    "Категория":      "category",
    "Теги":           "tags",
}


# ============================================================
# АВТОРИЗАЦИЯ GOOGLE
# ============================================================

def _load_credentials_dict() -> dict:
    """
    Возвращает словарь с данными сервисного аккаунта.
    Приоритет:
        1) env GOOGLE_CREDENTIALS_JSON  (используется в GitHub Actions)
        2) файл credentials.json        (используется локально)
    """
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
        f"ни в файле '{CREDENTIALS_FILE}'. "
        "Задайте переменную окружения или положите файл рядом со скриптом."
    )


def get_gspread_client():
    """
    Создаёт авторизованный клиент gspread.
    Работает и локально (через credentials.json),
    и в GitHub Actions (через GOOGLE_CREDENTIALS_JSON).
    """
    creds_dict = _load_credentials_dict()
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)
