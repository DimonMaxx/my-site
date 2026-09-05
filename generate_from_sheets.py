import gspread
import pandas as pd
import os
import re
import json
import sys

# ========== НАСТРОЙКИ ==========
SPREADSHEET_ID = "1kcG0TG4GZtSM2mypjgvNDUpIbLfIvcmW80_hBKA11nw"

SHEET_TO_FOLDER = {
    "Новости": "_content/news",
    "Программы": "_content/programs",
    "Книги": "_content/books",
    "Музыка": "_content/music",
    "Игры": "_content/games",
    "Статьи": "_content/articles",
    "Фильмы": "_content/movies",
    "Разное": "_content/misc",
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
}
# ==============================

def get_gspread_client():
    # Если переменная окружения задана (в GitHub Actions), используем её
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if creds_json:
        try:
            creds_dict = json.loads(creds_json)
            return gspread.service_account_from_dict(creds_dict)
        except Exception as e:
            print(f"Ошибка парсинга GOOGLE_CREDENTIALS_JSON: {e}")
            sys.exit(1)
    else:
        # Локально: читаем из файла
        try:
            return gspread.service_account(filename="credentials.json")
        except FileNotFoundError:
            print("Файл credentials.json не найден. Убедитесь, что он есть при локальном запуске.")
            sys.exit(1)

def slugify(title):
    slug = re.sub(r'[^\w\s-]', '', title).strip().lower()
    slug = re.sub(r'[-\s]+', '-', slug)
    return slug

def generate_md_files_from_sheet(worksheet, folder):
    print(f"Обработка листа: {worksheet.title}")
    records = worksheet.get_all_records()
    if not records:
        print(f"  Лист '{worksheet.title}' пуст, пропускаем.")
        return

    os.makedirs(folder, exist_ok=True)

    for row in records:
        title = row.get("Название", "").strip()
        if not title:
            print(f"  Пропущена строка без названия: {row}")
            continue

        filename = f"{slugify(title)}.md"
        filepath = os.path.join(folder, filename)

        front_matter = "---\n"
        for ru_col, en_key in COLUMN_MAPPING.items():
            value = row.get(ru_col)
            if pd.isna(value) or value == "":
                continue
            safe_value = str(value).replace('"', '\\"')
            front_matter += f'{en_key}: "{safe_value}"\n'
        front_matter += "---\n\n"

        body = row.get("Текст", "")
        if pd.isna(body):
            body = ""

        full_content = front_matter + body
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(full_content)

        print(f"  Создан файл: {filepath}")

def main():
    print("Подключение к Google Sheets...")
    gc = get_gspread_client()
    print("Клиент создан.")

    try:
        sh = gc.open_by_key(SPREADSHEET_ID)
        print("Таблица найдена.")
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"ОШИБКА: Таблица с ID '{SPREADSHEET_ID}' не найдена.")
        print("Проверьте ID и права доступа для сервисного аккаунта.")
        return

    for sheet_title, folder in SHEET_TO_FOLDER.items():
        try:
            worksheet = sh.worksheet(sheet_title)
            generate_md_files_from_sheet(worksheet, folder)
        except gspread.exceptions.WorksheetNotFound:
            print(f"ПРЕДУПРЕЖДЕНИЕ: Лист '{sheet_title}' не найден. Пропускаем.")

    print("Готово! Файлы созданы.")

if __name__ == "__main__":
    main()
