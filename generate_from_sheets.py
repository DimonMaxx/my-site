# generate_from_sheets.py
# Google Sheets → JSON + MD
# Общие константы и функции — в common.py

import os
import json

import pandas as pd
import gspread

from common import (
    SPREADSHEET_ID,
    SHEET_CONFIG,
    COLUMN_MAPPING,
    get_gspread_client,
    slugify,
)


def generate_json_and_md(worksheet, json_path, folder_path):
    print(f"Обработка листа: {worksheet.title}")
    records = worksheet.get_all_records()
    if not records:
        print(f"  Лист '{worksheet.title}' пуст, создаём пустые файлы.")
        data = []
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.makedirs(folder_path, exist_ok=True)
        return

    json_data = []
    os.makedirs(folder_path, exist_ok=True)

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
            json_data.append(item)
            title = item['title']
            slug = slugify(title)
            md_filename = f"{slug}.md"
            md_path = os.path.join(folder_path, md_filename)
            front_matter = "---\n"
            for key, value in item.items():
                safe_value = str(value).replace('"', '\\"')
                front_matter += f'{key}: "{safe_value}"\n'
            front_matter += "---\n\n"
            body = item.get('body', '')
            full_content = front_matter + body
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(full_content)
            print(f"  Создан Markdown: {md_path}")
        else:
            print(f"  Пропущена строка без названия: {row}")

    try:
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        print(f"  Создан JSON: {json_path} ({len(json_data)} записей)")
        if json_data:
            print(f"  Пример данных: {json_data[0]}")
    except Exception as e:
        print(f"  ОШИБКА записи JSON: {e}")


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
