# delete_content.py
# Удаление строк из Google Sheets по названию
# Общие константы и функции — в common.py

import os
import json

import gspread

from common import (
    SPREADSHEET_ID,
    SECTION_TO_SHEET,
    get_gspread_client,
    normalize,
)


def main():
    delete_list_json = os.environ.get('DELETE_LIST', '[]')
    try:
        delete_list = json.loads(delete_list_json)
    except Exception as e:
        print(f"Ошибка парсинга DELETE_LIST: {e}")
        return

    if not delete_list:
        print("Список удаляемых записей пуст.")
        return

    print(f"Получено на удаление: {len(delete_list)} записей")

    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    # Группируем по листам
    by_sheet = {}
    for item in delete_list:
        section = item.get('section')
        title = item.get('title')
        if not section or not title:
            continue
        sheet_name = SECTION_TO_SHEET.get(section)
        if not sheet_name:
            continue
        by_sheet.setdefault(sheet_name, []).append(title)

    # Удаляем
    for sheet_name, titles in by_sheet.items():
        print(f"\nЛист '{sheet_name}': удаляем {len(titles)} записей")
        try:
            worksheet = sh.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            print(f"  Лист '{sheet_name}' не найден, пропускаем.")
            continue

        all_values = worksheet.get_all_values()
        if not all_values:
            print("  Лист пустой.")
            continue

        headers = all_values[0]
        try:
            col_title = headers.index("Название")
        except ValueError:
            print(f"  На листе '{sheet_name}' нет колонки 'Название'")
            continue

        rows_to_delete = []
        titles_norm = [normalize(t) for t in titles]
        for i, row in enumerate(all_values[1:], start=2):
            if col_title < len(row):
                row_title_norm = normalize(row[col_title])
                if row_title_norm in titles_norm:
                    rows_to_delete.append(i)
                    titles_norm.remove(row_title_norm)

        if not rows_to_delete:
            print(f"  Ничего не найдено для удаления.")
            continue

        print(f"  Удаляем {len(rows_to_delete)} строк: {sorted(rows_to_delete, reverse=True)[:5]}...")
        for row_num in sorted(rows_to_delete, reverse=True):
            worksheet.delete_rows(row_num)

        print(f"  ✓ Удалено {len(rows_to_delete)} строк.")

    print("\nГотово!")


if __name__ == "__main__":
    main()
