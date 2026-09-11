import gspread
import os
import json
import sys
import re
import requests

# ========== НАСТРОЙКИ ==========
SPREADSHEET_ID = "1kcG0TG4GZtSM2mypjgvNDUpIbLfIvcmW80_hBKA11nw"

# Публичная ссылка на папку Яндекс.Диска, например: https://disk.yandex.ru/d/XXXXX
# Оставьте пустым, если не используете автосинхронизацию
YANDEX_PUBLIC_FOLDER_URL = "https://disk.yandex.ru/d/zMxF4nXHPkIVCQ"  # например: "https://disk.yandex.ru/d/abc123"

# Какие поля добавлять для новых файлов
# Ключи должны совпадать с названиями колонок в Google Sheets
NEW_FILE_TEMPLATE = {
    "Название": "",           # будет взято из имени файла (без расширения)
    "Автор": "",
    "Описание": "",
    "Формат": "",             # будет взято из расширения
    "Ссылка для скачивания": "",
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
            print("Файл credentials.json не найден.")
            sys.exit(1)

def get_yandex_public_files(public_url, limit=1000):
    """Получает список файлов из публичной папки Яндекс.Диска."""
    # Извлекаем public_key из URL
    # URL вида https://disk.yandex.ru/d/XXXXX или https://yadi.sk/d/XXXXX
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', public_url)
    if not match:
        print(f"Не удалось извлечь public_key из URL: {public_url}")
        return []
    public_key = match.group(1)

    # Запрашиваем список файлов
    api_url = "https://cloud-api.yandex.net/v1/disk/public/resources"
    params = {
        "public_key": public_key,
        "limit": limit,
        "sort": "name"
    }
    try:
        response = requests.get(api_url, params=params, timeout=30)
        if response.status_code != 200:
            print(f"Ошибка Яндекс.Диска: {response.status_code} — {response.text}")
            return []
        data = response.json()
        items = data.get('_embedded', {}).get('items', [])
        print(f"Получено {len(items)} файлов из Яндекс.Диска.")
        return items
    except Exception as e:
        print(f"Ошибка при запросе к Яндекс.Диску: {e}")
        return []

def get_existing_names(worksheet, name_column_index):
    """Возвращает множество названий, которые уже есть в таблице."""
    try:
        all_values = worksheet.get_all_values()
        if len(all_values) < 2:
            return set()
        # Пропускаем заголовок
        names = set()
        for row in all_values[1:]:
            if name_column_index < len(row):
                name = row[name_column_index].strip()
                if name:
                    names.add(name.lower())
        return names
    except Exception as e:
        print(f"Ошибка при чтении существующих названий: {e}")
        return set()

def main():
    if not YANDEX_PUBLIC_FOLDER_URL:
        print("YANDEX_PUBLIC_FOLDER_URL не задан. Синхронизация с Яндекс.Диском пропущена.")
        return

    print("Подключение к Google Sheets...")
    gc = get_gspread_client()
    print("Клиент создан.")

    try:
        sh = gc.open_by_key(SPREADSHEET_ID)
        print("Таблица найдена.")
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"ОШИБКА: Таблица с ID '{SPREADSHEET_ID}' не найдена.")
        return

    # Лист "Книги" — здесь хранятся книги
    sheet_title = "Книги"
    try:
        worksheet = sh.worksheet(sheet_title)
    except gspread.exceptions.WorksheetNotFound:
        print(f"ПРЕДУПРЕЖДЕНИЕ: Лист '{sheet_title}' не найден. Пропускаем синхронизацию.")
        return

    # Получаем заголовки
    headers = worksheet.row_values(1)
    if not headers:
        print("ПРЕДУПРЕЖДЕНИЕ: На листе 'Книги' нет заголовков.")
        return
    print(f"Заголовки: {headers}")

    # Индекс колонки "Название"
    try:
        name_col_index = headers.index("Название")
    except ValueError:
        print("ПРЕДУПРЕЖДЕНИЕ: На листе 'Книги' нет колонки 'Название'. Синхронизация невозможна.")
        return

    # Существующие названия (в нижнем регистре для сравнения)
    existing_names = get_existing_names(worksheet, name_col_index)
    print(f"В таблице уже есть {len(existing_names)} названий.")

    # Получаем файлы из Яндекс.Диска
    files = get_yandex_public_files(YANDEX_PUBLIC_FOLDER_URL)
    if not files:
        print("Не удалось получить список файлов. Синхронизация завершена.")
        return

    # Формируем новые записи
    new_rows = []
    for item in files:
        if item.get('type') != 'file':
            continue
        file_name = item.get('name', '')
        if not file_name:
            continue
        # Убираем расширение
        name_without_ext = os.path.splitext(file_name)[0]
        ext = os.path.splitext(file_name)[1].lstrip('.').lower()

        # Проверяем, есть ли уже такая книга в таблице
        if name_without_ext.lower() in existing_names:
            continue

        # Формируем строку по заголовкам таблицы
        row = []
        for header in headers:
            if header == "Название":
                row.append(name_without_ext)
            elif header == "Формат":
                row.append(ext)
            elif header == "Ссылка для скачивания":
                # Можно взять публичную ссылку на файл, если знаем
                # Но обычно Яндекс.Диск даёт ссылку только на папку.
                # Оставим пустым — пользователь заполнит вручную или через другой механизм.
                row.append("")
            else:
                row.append("")
        new_rows.append(row)
        print(f"  Новая запись: {name_without_ext} ({ext})")

    if not new_rows:
        print("Нет новых файлов для добавления.")
        return

    # Добавляем строки в конец таблицы
    print(f"Добавляем {len(new_rows)} новых записей...")
    for row in new_rows:
        worksheet.append_row(row, value_input_option='USER_ENTERED')

    print(f"Готово! Добавлено {len(new_rows)} новых записей.")

if __name__ == "__main__":
    main()
