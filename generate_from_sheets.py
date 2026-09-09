import gspread
import pandas as pd
import os
import json
import sys
import re

SPREADSHEET_ID = "1kcG0TG4GZtSM2mypjgvNDUpIbLfIvcmW80_hBKA11nw"

# Сопоставление листов с JSON-файлом и папкой для Markdown
SHEET_CONFIG = {
    "Новости": {"json": "_content/news.json", "folder": "_content/news"},
    "Программы": {"json": "_content/programs.json", "folder": "_content/programs"},
    "Книги": {"json": "_content/books.json", "folder": "_content/books"},
    "Музыка": {"json": "_content/music.json", "folder": "_content/music"},
    "Игры": {"json": "_content/games.json", "folder": "_content/games"},
    "Статьи": {"json": "_content/articles.json", "folder": "_content/articles"},
    "Фильмы": {"json": "_content/movies.json", "folder": "_content/movies"},
    "Разное": {"json": "_content/misc.json", "folder": "_content/misc"},
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

def slugify(title):
    slug = re.sub(r'[^\w\s-]', '', title).strip().lower()
    slug = re.sub(r'[-\s]+', '-', slug)
    return slug

def generate_json_and_md(worksheet, json_path, folder_path):
    print(f"Обработка листа: {worksheet.title}")
    records = worksheet.get_all_records()
    if not records:
        print(f"  Лист '{worksheet.title}' пуст, создаём пустые файлы.")
        data = []
        # Создаём пустой JSON
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        # Создаём пустую папку (файлы не создаются)
        os.makedirs(folder_path, exist_ok=True)
        return

    # Данные для JSON
    json_data = []
    # Создаём папку для Markdown, если её нет
    os.makedirs(folder_path, exist_ok=True)

    for row in records:
        # Формируем объект для JSON
        item = {}
        for ru_col, en_key in COLUMN_MAPPING.items():
            value = row.get(ru_col)
            if pd.isna(value) or value == "":
                continue
            if ru_col == "Текст":
                item[en_key] = str(value)
            else:
                item[en_key] = str(value).strip()
        # Проверяем наличие обязательного поля 'title'
        if 'title' in item and item['title']:
            json_data.append(item)

            # Генерируем Markdown файл для админки
            title = item['title']
            slug = slugify(title)
            md_filename = f"{slug}.md"
            md_path = os.path.join(folder_path, md_filename)

            # Формируем front matter (YAML)
            front_matter = "---\n"
            for key, value in item.items():
                # Экранируем кавычки в значениях
                safe_value = str(value).replace('"', '\\"')
                front_matter += f'{key}: "{safe_value}"\n'
            front_matter += "---\n\n"

            # Тело (если есть)
            body = item.get('body', '')
            full_content = front_matter + body

            # Записываем Markdown
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(full_content)
            print(f"  Создан Markdown: {md_path}")
        else:
            print(f"  Пропущена строка без названия: {row}")

    # Записываем JSON
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)
    print(f"  Создан JSON: {json_path} ({len(json_data)} записей)")

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

    for sheet_title, config in SHEET_CONFIG.items():
        try:
            worksheet = sh.worksheet(sheet_title)
            generate_json_and_md(worksheet, config["json"], config["folder"])
        except gspread.exceptions.WorksheetNotFound:
            print(f"ПРЕДУПРЕЖДЕНИЕ: Лист '{sheet_title}' не найден. Пропускаем.")

    print("Готово! JSON и Markdown файлы созданы.")

if __name__ == "__main__":
    main()
