import gspread
import pandas as pd
import os
import re
import json
import sys

# ========== НАСТРОЙКИ ==========
SPREADSHEET_ID = "1kcG0TG4GZtSM2mypjgvNDUpIbLfIvcmW80_hBKA11nw"

# Соответствие: имя листа -> имя выходного JSON-файла (в папке _content/)
SHEET_TO_JSON = {
    "Новости": "_content/news.json",
    "Программы": "_content/programs.json",
    "Книги": "_content/books.json",
    "Музыка": "_content/music.json",
    "Игры": "_content/games.json",
    "Статьи": "_content/articles.json",
    "Фильмы": "_content/movies.json",
    "Разное": "_content/misc.json",
}

# Сопоставление русских названий колонок -> ключи для JSON
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
}
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
            print("Файл credentials.json не найден. Убедитесь, что он есть при локальном запуске.")
            sys.exit(1)

def generate_json_from_sheet(worksheet, json_path):
    print(f"Обработка листа: {worksheet.title}")
    records = worksheet.get_all_records()
    if not records:
        print(f"  Лист '{worksheet.title}' пуст, создаём пустой JSON.")
        data = []
    else:
        data = []
        for row in records:
            item = {}
            for ru_col, en_key in COLUMN_MAPPING.items():
                value = row.get(ru_col)
                if pd.isna(value) or value == "":
                    continue
                if ru_col == "Текст":
                    item[en_key] = str(value)
                else:
                    item[en_key] = str(value).strip()
            if 'title' in item and item['title']:
                data.append(item)
            else:
                print(f"  Пропущена строка без названия: {row}")

    # Записываем JSON
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  Создан JSON: {json_path} ({len(data)} записей)")

def main():
    print("Подключение к Google Sheets...")
    gc = get_gspread_client()
    print("Клиент создан.")

    try:
        sh = gc.open_by_key(SPREADSHEET_ID)
        print("Таблица найдена.")
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"ОШИБКА: Таблица с ID '{SPREADSHEET_ID}' не найдена.")
        return

    for sheet_title, json_path in SHEET_TO_JSON.items():
        try:
            worksheet = sh.worksheet(sheet_title)
            generate_json_from_sheet(worksheet, json_path)
        except gspread.exceptions.WorksheetNotFound:
            print(f"ПРЕДУПРЕЖДЕНИЕ: Лист '{sheet_title}' не найден. Пропускаем.")

    print("Готово! JSON-файлы созданы.")

if __name__ == "__main__":
    main()
