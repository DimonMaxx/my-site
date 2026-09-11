import gspread
import os
import json
import sys
import re
import time
import requests

# ========== НАСТРОЙКИ ==========
SPREADSHEET_ID = "1kcG0TG4GZtSM2mypjgvNDUpIbLfIvcmW80_hBKA11nw"

YANDEX_PUBLIC_FOLDER_URL = "https://disk.yandex.ru/d/zMxF4nXHPkIVCQ"

NEW_FILE_TEMPLATE = {
    "Название": "",
    "Автор": "",
    "Описание": "",
    "Формат": "",
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
    api_url = "https://cloud-api.yandex.net/v1/disk/public/resources"
    params = {
        "public_key": public_url,
        "limit": limit,
        "sort": "name"
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    }
    try:
        response = requests.get(api_url, params=params, headers=headers, timeout=30)
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

def normalize_name(name):
    """Нормализует название для сравнения:
    - убирает расширение
    - убирает суффикс в скобках в конце (например, '(fb2)')
    - приводит к нижнему регистру
    - убирает лишние пробелы
    """
    if not name:
        return ''
    # Убираем расширение
    name = os.path.splitext(name)[0]
    # Убираем суффикс в скобках в конце: " ... (fb2)" или " ... (pdf)"
    name = re.sub(r'\s*\([^)]*\)\s*$', '', name)
    # Приводим к нижнему регистру и убираем лишние пробелы
    name = name.strip().lower()
    name = re.sub(r'\s+', ' ', name)
    return name

def get_existing_names(worksheet, name_column_index):
    """Возвращает множество нормализованных названий, которые уже есть в таблице."""
    try:
        all_values = worksheet.get_all_values()
        if len(all_values) < 2:
            print(f"  В таблице только заголовок или пусто (строк: {len(all_values)})")
            return set()
        names = set()
        for row_idx, row in enumerate(all_values[1:], start=2):
            if name_column_index < len(row):
                name = row[name_column_index].strip()
                if name:
                    names.add(normalize_name(name))
        print(f"  Прочитано {len(all_values) - 1} строк, уникальных названий: {len(names)}")
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

    sheet_title = "Книги"
    try:
        worksheet = sh.worksheet(sheet_title)
    except gspread.exceptions.WorksheetNotFound:
        print(f"ПРЕДУПРЕЖДЕНИЕ: Лист '{sheet_title}' не найден. Пропускаем синхронизацию.")
        return

    headers = worksheet.row_values(1)
    if not headers:
        print("ПРЕДУПРЕЖДЕНИЕ: На листе 'Книги' нет заголовков.")
        return
    print(f"Заголовки: {headers}")

    try:
        name_col_index = headers.index("Название")
    except ValueError:
        print("ПРЕДУПРЕЖДЕНИЕ: На листе 'Книги' нет колонки 'Название'. Синхронизация невозможна.")
        return

    print("Чтение существующих названий...")
    existing_names = get_existing_names(worksheet, name_col_index)
    print(f"В таблице уже есть {len(existing_names)} названий.")

    files = get_yandex_public_files(YANDEX_PUBLIC_FOLDER_URL)
    if not files:
        print("Не удалось получить список файлов. Синхронизация завершена.")
        return

    new_rows = []
    skipped = 0
    for item in files:
        if item.get('type') != 'file':
            continue
        file_name = item.get('name', '')
        if not file_name:
            continue
        # Имя без расширения
        name_without_ext = os.path.splitext(file_name)[0]
        # Убираем суффикс "(fb2)" и т.п. из имени
        name_clean = re.sub(r'\s*\([^)]*\)\s*$', '', name_without_ext).strip()
        ext = os.path.splitext(file_name)[1].lstrip('.').lower()

        # Проверяем, есть ли уже такая книга (нормализованное сравнение)
        norm = normalize_name(file_name)
        if norm in existing_names:
            skipped += 1
            continue

        # Формируем строку по заголовкам таблицы
        row = []
        for header in headers:
            if header == "Название":
                row.append(name_clean)
            elif header == "Формат":
                row.append(ext)
            else:
                row.append("")
        new_rows.append(row)
        print(f"  Новая запись: {name_clean} ({ext})")

    print(f"\nПропущено (уже есть в таблице): {skipped}")
    print(f"Новых записей для добавления: {len(new_rows)}")

    if not new_rows:
        print("Нет новых файлов для добавления.")
        return

    # Добавляем все строки ОДНИМ запросом, чтобы не превысить квоту
    print(f"Добавляем {len(new_rows)} записей одним батчем...")
    try:
        worksheet.append_rows(new_rows, value_input_option='USER_ENTERED')
        print(f"Готово! Добавлено {len(new_rows)} новых записей.")
    except gspread.exceptions.APIError as e:
        print(f"Ошибка при добавлении строк: {e}")
        # Если всё-таки превышена квота — попробуем ещё раз через 60 секунд
        if '429' in str(e):
            print("Превышена квота. Ждём 60 секунд и пробуем ещё раз...")
            time.sleep(60)
            worksheet.append_rows(new_rows, value_input_option='USER_ENTERED')
            print(f"Готово! Добавлено {len(new_rows)} новых записей (со второй попытки).")
        else:
            raise

if __name__ == "__main__":
    main()
