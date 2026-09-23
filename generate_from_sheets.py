# generate_from_sheets.py
# Google Sheets → JSON.
# Разделы читаются из Supabase (site_sections, is_active = true),
# поэтому новые разделы, добавленные через админ-панель,
# генерируются автоматически без правки этого файла.

import os
import sys
import json
import traceback

import pandas as pd
import gspread

from common import (
    SPREADSHEET_ID,
    COLUMN_MAPPING as BASE_COLUMN_MAPPING,
    get_gspread_client,
)

# ------------------------------------------------------------
# Supabase — опциональная зависимость (не критично для чтения
# из Sheets, но критично для получения списка разделов).
# ------------------------------------------------------------
try:
    from supabase import create_client as supa_create_client
except ImportError:
    supa_create_client = None


# ============================================================
# КОНФИГУРАЦИЯ
# ============================================================

SUPABASE_URL = os.environ.get(
    "SUPABASE_URL",
    "https://rmoonebbvpmvthvpcmpt.supabase.co",
)
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")


# ============================================================
# РАСШИРЕННЫЙ МАППИНГ RU → EN
# ============================================================
# Берём базовый маппинг из common.py и дополняем ключами,
# которые реально используются в новых разделах (music, games,
# articles, news и любые пользовательские).
#
# Если в common.py появятся новые ключи — они автоматически
# подхватятся, т.к. мы делаем merge, а не замену.

_EXTRA_RU_TO_EN = {
    # Музыка
    "Исполнитель": "artist",
    "Год":         "year",
    # Игры
    "Платформа":   "platform",
    # Статьи / Новости / Разное
    "Дата":        "date",
    "Категория":   "category",
    "Теги":        "tags",
    "Текст":       "text",
    # Служебные (на случай если кто-то положит в Sheets логи)
    "Файл":        "file_name",
    "Пользователь": "username",
    "Дата и время": "downloaded_at",
}

# Итоговый RU → EN
RU_TO_EN = dict(BASE_COLUMN_MAPPING)
for ru, en in _EXTRA_RU_TO_EN.items():
    RU_TO_EN.setdefault(ru, en)

# Обратный маппинг: EN → [RU, RU, ...]
# (у одного EN-ключа может быть несколько русских вариантов)
EN_TO_RU = {}
for ru, en in RU_TO_EN.items():
    EN_TO_RU.setdefault(en, []).append(ru)

# Псевдонимы: некоторые EN-ключи пишутся в один и тот же RU-заголовок.
# Пример: frontend ожидает "body", а common.py использует "text"
# для того же "Текст". Поддерживаем оба.
EN_TO_RU.setdefault("body", []).extend(EN_TO_RU.get("text", []))


# ============================================================
# ЗАГРУЗКА РАЗДЕЛОВ ИЗ SUPABASE
# ============================================================

def load_active_sections():
    """
    Возвращает список активных разделов из таблицы site_sections,
    отсортированных по sort_order.
    """
    if supa_create_client is None:
        print("[!] supabase-py не установлен (pip install supabase).")
        return []

    if not SUPABASE_SERVICE_KEY:
        print("[!] SUPABASE_SERVICE_ROLE_KEY не задан — не могу загрузить разделы.")
        return []

    try:
        client = supa_create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)
        resp = (
            client.table("site_sections")
            .select("key,label,icon,handler_type,yandex_url,yandex_path,"
                    "sheet_name,json_path,container,columns,folderable,"
                    "is_active,sort_order")
            .eq("is_active", True)
            .order("sort_order")
            .execute()
        )
        return resp.data or []
    except Exception as e:
        print(f"[!] Ошибка загрузки разделов из Supabase: {e}")
        traceback.print_exc()
        return []


# ============================================================
# УТИЛИТЫ ПРЕОБРАЗОВАНИЯ
# ============================================================

def _is_empty(value):
    """Пустая ли ячейка (None / NaN / пустая строка / пробелы)."""
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(value, str) and not value.strip():
        return True
    return False


def row_to_item(row, columns=None):
    """
    Преобразует одну строку Google Sheets (dict {RU-заголовок: значение})
    в JSON-объект {EN-ключ: значение}.

    :param row:     dict из worksheet.get_all_records()
    :param columns: список EN-ключей из site_sections.columns.
                    Если не задан — выводятся все известные EN-ключи.
    :return:        dict (может быть пустым)
    """
    item = {}

    # Какие EN-ключи выводить
    target_keys = [k for k in (columns or []) if k]
    if not target_keys:
        target_keys = list(EN_TO_RU.keys())

    for en_key in target_keys:
        value = None

        # 1. Ищем по русским вариантам
        for ru_col in EN_TO_RU.get(en_key, []):
            if ru_col in row and not _is_empty(row[ru_col]):
                value = row[ru_col]
                break

        # 2. Fallback: возможно, в листе колонка названа EN-ключом
        if _is_empty(value) and en_key in row and not _is_empty(row[en_key]):
            value = row[en_key]

        if _is_empty(value):
            continue

        # Многострочные поля оставляем как есть, остальные — strip()
        if en_key in ("text", "body", "description"):
            item[en_key] = str(value)
        else:
            item[en_key] = str(value).strip()

    return item


# ============================================================
# ГЕНЕРАЦИЯ JSON ДЛЯ ОДНОГО РАЗДЕЛА
# ============================================================

def generate_json_for_section(worksheet, json_path, columns=None):
    """
    Читает worksheet и пишет JSON-массив в json_path.
    Возвращает количество записей (int) или -1 при ошибке.
    """
    print(f"  Лист:    {worksheet.title}")
    print(f"  JSON:    {json_path}")

    # Создаём папку, если её нет
    dir_part = os.path.dirname(json_path) or "."
    os.makedirs(dir_part, exist_ok=True)

    # --- Чтение ---
    try:
        records = worksheet.get_all_records()
    except Exception as e:
        print(f"  ОШИБКА чтения листа: {e}")
        return -1

    # --- Пустой лист ---
    if not records:
        print("  Лист пуст → записываем []")
        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump([], f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"  ОШИБКА записи JSON: {e}")
            return -1
        return 0

    # --- Преобразование строк ---
    json_data = []
    skipped = 0
    for row in records:
        item = row_to_item(row, columns)
        if not item.get("title"):
            skipped += 1
            continue
        json_data.append(item)

    # --- Запись ---
    try:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"  ОШИБКА записи JSON: {e}")
        return -1

    # --- Отчёт ---
    msg = f"  ✓ Записано: {len(json_data)} записей"
    if skipped:
        msg += f" (пропущено без title: {skipped})"
    print(msg)

    if json_data:
        sample_keys = list(json_data[0].keys())
        preview = ", ".join(sample_keys[:8])
        if len(sample_keys) > 8:
            preview += ", ..."
        print(f"    Ключи: {preview}")

    return len(json_data)


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("generate_from_sheets.py — старт")
    print("=" * 60)

    if gspread is None:
        print("[!] gspread не установлен.")
        sys.exit(1)

    # --- 1. Разделы из Supabase ---
    print("\nЗагрузка разделов из Supabase (site_sections)...")
    sections = load_active_sections()
    if not sections:
        print("[!] Нет активных разделов — завершаю.")
        sys.exit(0)

    print(f"  Получено разделов: {len(sections)}")
    for s in sections:
        print(
            f"    • {s.get('key'):<12} "
            f"→ лист «{s.get('sheet_name') or '—'}», "
            f"JSON «{s.get('json_path') or '—'}»"
        )

    # --- 2. Google Sheets ---
    print("\nПодключение к Google Sheets...")
    try:
        gc = get_gspread_client()
    except Exception as e:
        print(f"[!] Не удалось авторизоваться: {e}")
        traceback.print_exc()
        sys.exit(1)

    try:
        sh = gc.open_by_key(SPREADSHEET_ID)
        print(f"  Таблица открыта: {sh.title}")
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"[!] Таблица с ID '{SPREADSHEET_ID}' не найдена.")
        print("    Проверьте, что сервисному аккаунту дан доступ "
              "(Share → Editor).")
        sys.exit(1)
    except Exception as e:
        print(f"[!] Ошибка открытия таблицы: {e}")
        traceback.print_exc()
        sys.exit(1)

    # --- 3. Обход разделов ---
    total_ok = 0
    total_records = 0
    failed = []   # [(key, reason), ...]

    for section in sections:
        key = section.get("key") or "?"
        label = section.get("label") or key
        sheet_name = section.get("sheet_name") or label
        json_path = section.get("json_path")
        columns = section.get("columns") or []

        print(f"\n=== Раздел: {label} ({key}) ===")

        if not json_path:
            print("  [!] json_path не задан → пропуск.")
            failed.append((key, "нет json_path"))
            continue

        # Получаем лист
        try:
            worksheet = sh.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            print(f"  [!] Лист «{sheet_name}» не найден → пропуск.")
            failed.append((key, f"лист «{sheet_name}» не найден"))
            continue
        except Exception as e:
            print(f"  [!] Ошибка доступа к листу: {e}")
            failed.append((key, str(e)))
            continue

        count = generate_json_for_section(worksheet, json_path, columns)
        if count >= 0:
            total_ok += 1
            total_records += count
        else:
            failed.append((key, "ошибка генерации"))

    # --- 4. Итоги ---
    print("\n" + "=" * 60)
    print(f"Обработано разделов:  {total_ok} из {len(sections)}")
    print(f"Всего записей:        {total_records}")
    if failed:
        print(f"\nПроблемные разделы ({len(failed)}):")
        for key, reason in failed:
            print(f"  • {key}: {reason}")
    print("=" * 60)

    # Возвращаем ненулевой код, только если вообще ничего не сгенерировано
    if total_ok == 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
