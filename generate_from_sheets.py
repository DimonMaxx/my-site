# generate_from_sheets.py
# Google Sheets → JSON (без Markdown)

import os
import json

import pandas as pd
import gspread

from common import (
    SPREADSHEET_ID,
    SHEET_CONFIG,
    COLUMN_MAPPING,
    get_gspread_client,
)


def generate_json(worksheet, json_path):
    print(f"Обработка листа: {worksheet.title}")

    # Создаём папку, если её нет
    os.makedirs(os.path.dirname(json_path) or ".", exist_ok=True)

    records = worksheet.get_all_records()
    if not records:
        print(f"  Лист '{worksheet.title}' пуст, создаём пустой JSON.")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump([], f, ensure_ascii=False, indent=2)
        return

    json_data = []

    for row in records:
        item = {}
        for ru_col, en_key in COLUMN_MAPPING.items():
            value = row.get(ru_col)
            if value is None:
                continue
            # pandas ставит NaN для пустых ячеек — отсеиваем
            try:
                if pd.isna(value):
                    continue
            except (TypeError, ValueError):
                pass
            if value == "":
                continue

            if ru_col == "Текст":
                item[en_key] = str(value)
            else:
                item[en_key] = str(value).strip()

        # Пропускаем строки без названия
        if "title" in item and item["title"]:
            json_data.append(item)

    try:
        with open(json_path, "w", encoding="utf-8") as f:
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
        print("  Проверьте, что сервисному аккаунту дан доступ (Share → Editor).")
        return

    for sheet_title, config in SHEET_CONFIG.items():
        try:
            worksheet = sh.worksheet(sheet_title)
            generate_json(worksheet, config["json"])
        except gspread.exceptions.WorksheetNotFound:
            print(f"ПРЕДУПРЕЖДЕНИЕ: Лист '{sheet_title}' не найден. Пропускаем.")

    print("Готово! JSON-файлы созданы.")


if __name__ == "__main__":
    main()
