import gspread
import os
import json
import sys
import re

SPREADSHEET_ID = "1kcG0TG4GZtSM2mypjgvNDUpIbLfIvcmW80_hBKA11nw"

SECTION_TO_SHEET = {
    "programs": "Программы",
    "books": "Книги",
    "news": "Новости",
    "articles": "Статьи",
    "movies": "Фильмы",
    "music": "Музыка",
    "games": "Игры",
    "misc": "Разное",
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
            print("Файл credentials.json не найден.")
            sys.exit(1)

def normalize(name):
    if not name:
        return ''
    name = os.path.splitext(name)[0]
    name = re.sub(r'\s*\([^)]*\)\s*$', '', name)
    name = name.strip().lower()
    name = re.sub(r'\s+', ' ', name)
    name = re.sub(r'[—–]', '-', name)
    name = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', name)
    name = re.sub(r'[^\w\s\-]', ' ', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name

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

        # Собираем номера строк для удаления (1-based)
        rows_to_delete = []
        titles_norm = [normalize(t) for t in titles]
        for i, row in enumerate(all_values[1:], start=2):
            if col_title < len(row):
                row_title_norm = normalize(row[col_title])
                if row_title_norm in titles_norm:
                    rows_to_delete.append(i)
                    titles_norm.remove(row_title_norm)  # чтобы не удалить повторно

        if not rows_to_delete:
            print(f"  Ничего не найдено для удаления.")
            continue

        # Удаляем строки в обратном порядке (снизу вверх), чтобы не сбить индексы
        print(f"  Удаляем {len(rows_to_delete)} строк: {sorted(rows_to_delete, reverse=True)[:5]}...")
        for row_num in sorted(rows_to_delete, reverse=True):
            worksheet.delete_rows(row_num)

        print(f"  ✓ Удалено {len(rows_to_delete)} строк.")

    print("\nГотово!")

if __name__ == "__main__":
    main()
