# generate_from_sheets.py
# Google Sheets → JSON.
# Разделы читаются из Supabase (site_sections, is_active = true),
# поэтому новые разделы, добавленные через админ-панель,
# генерируются автоматически без правки этого файла.
#
# Поддерживает SYNC_SECTIONS=programs,music — генерирует JSON только для
# перечисленных разделов (для селективной синхронизации).
#
# Устойчив к «дубликатам заголовков» и «хвостам» старых колонок:
# не использует get_all_records(), а читает значения напрямую.

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

# SYNC_SECTIONS — список ключей разделов через запятую.
# Пусто → генерировать JSON для всех активных разделов.
_raw_sync_sections = os.environ.get("SYNC_SECTIONS", "").strip()
SYNC_SECTIONS_FILTER = [
    s.strip() for s in _raw_sync_sections.split(",") if s.strip()
] if _raw_sync_sections else []


# ============================================================
# РАСШИРЕННЫЙ МАППИНГ RU → EN
# ============================================================

_EXTRA_RU_TO_EN = {
    "Исполнитель":  "artist",
    "Год":          "year",
    "Платформа":    "platform",
    "Дата":         "date",
    "Категория":    "category",
    "Теги":         "tags",
    "Текст":        "text",
    "Файл":         "file_name",
    "Пользователь": "username",
    "Дата и время": "downloaded_at",
}

RU_TO_EN = dict(BASE_COLUMN_MAPPING)
for ru, en in _EXTRA_RU_TO_EN.items():
    RU_TO_EN.setdefault(ru, en)

EN_TO_RU = {}
for ru, en in RU_TO_EN.items():
    EN_TO_RU.setdefault(en, []).append(ru)

EN_TO_RU.setdefault("body", []).extend(EN_TO_RU.get("text", []))


# ============================================================
# SUPABASE
# ============================================================

def load_active_sections():
    if supa_create_client is None:
        print("[!] supabase-py не установлен.")
        return []
    if not SUPABASE_SERVICE_KEY:
        print("[!] SUPABASE_SERVICE_ROLE_KEY не задан.")
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
# УТИЛИТЫ
# ============================================================

def _is_empty(value):
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


def _read_records_safe(worksheet, headers):
    """
    Читает лист построчно БЕЗ get_all_records().
    Возвращает список dict {RU-заголовок: значение}.

    Устойчив к:
      • дубликатам заголовков;
      • «хвостам» старых колонок;
      • пустым заголовкам.

    :param worksheet: gspread.Worksheet
    :param headers:   список ожидаемых RU-заголовков (в нужном порядке)
    """
    try:
        all_values = worksheet.get_all_values()
    except Exception as e:
        print(f"  ОШИБКА чтения листа: {e}")
        return None  # сигнал об ошибке

    if not all_values:
        return []

    raw_header = all_values[0]

    # Сопоставляем ожидаемые заголовки с позициями в листе.
    # При дубликатах берём ПЕРВОЕ совпадение.
    positions = {}
    for want_idx, want in enumerate(headers):
        want_clean = want.strip()
        pos = None
        # 1. Точное совпадение (регистронезависимо, без пробелов)
        for i, h in enumerate(raw_header):
            if (h or "").strip().lower() == want_clean.lower():
                pos = i
                break
        # 2. Частичное совпадение (например, «Размер» vs «Размер (МБ)»)
        if pos is None:
            for i, h in enumerate(raw_header):
                if want_clean.lower() in (h or "").strip().lower():
                    pos = i
                    break
        positions[want] = pos

    records = []
    for row in all_values[1:]:
        rec = {}
        for want, pos in positions.items():
            if pos is None or pos >= len(row):
                rec[want] = ""
            else:
                rec[want] = row[pos]
        records.append(rec)
    return records


def row_to_item(row, columns=None):
    item = {}
    target_keys = [k for k in (columns or []) if k]
    if not target_keys:
        target_keys = list(EN_TO_RU.keys())

    for en_key in target_keys:
        value = None
        for ru_col in EN_TO_RU.get(en_key, []):
            if ru_col in row and not _is_empty(row[ru_col]):
                value = row[ru_col]
                break
        if _is_empty(value) and en_key in row and not _is_empty(row[en_key]):
            value = row[en_key]
        if _is_empty(value):
            continue
        if en_key in ("text", "body", "description"):
            item[en_key] = str(value)
        else:
            item[en_key] = str(value).strip()
    return item


# ============================================================
# ГЕНЕРАЦИЯ JSON
# ============================================================

def generate_json_for_section(worksheet, json_path, columns=None, headers=None):
    """
    :param worksheet: gspread.Worksheet
    :param json_path: путь к JSON
    :param columns:   список EN-ключей из site_sections.columns
    :param headers:   список RU-заголовков (в порядке columns) —
                      если не задан, вычисляется автоматически
    """
    print(f"  Лист:    {worksheet.title}")
    print(f"  JSON:    {json_path}")

    dir_part = os.path.dirname(json_path) or "."
    os.makedirs(dir_part, exist_ok=True)

    # Определяем ожидаемые RU-заголовки
    if headers is None:
        headers = []
        for key in columns or []:
            variants = EN_TO_RU.get(key, [])
            if variants:
                headers.append(variants[0])
            else:
                headers.append(key)
        if "Ссылка для скачивания" not in headers:
            headers.append("Ссылка для скачивания")

    records = _read_records_safe(worksheet, headers)
    if records is None:
        return -1
    if not records:
        print("  Лист пуст → записываем []")
        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump([], f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"  ОШИБКА записи JSON: {e}")
            return -1
        return 0

    json_data = []
    skipped = 0
    for row in records:
        item = row_to_item(row, columns)
        if not item.get("title"):
            skipped += 1
            continue
        json_data.append(item)

    try:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"  ОШИБКА записи JSON: {e}")
        return -1

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

    if SYNC_SECTIONS_FILTER:
        print(f"SYNC_SECTIONS_FILTER = {SYNC_SECTIONS_FILTER}")
    else:
        print("SYNC_SECTIONS_FILTER = (все активные разделы)")

    if gspread is None:
        print("[!] gspread не установлен.")
        sys.exit(1)

    print("\nЗагрузка разделов из Supabase (site_sections)...")
    sections = load_active_sections()
    if not sections:
        print("[!] Нет активных разделов — завершаю.")
        sys.exit(0)

    # ─── Фильтр по SYNC_SECTIONS ───
    if SYNC_SECTIONS_FILTER:
        before = len(sections)
        sections = [s for s in sections
                    if s.get("key") in SYNC_SECTIONS_FILTER]
        print(f"  [i] SYNC_SECTIONS отфильтровал {len(sections)} "
              f"из {before} разделов")
        if not sections:
            print("[!] Ни один раздел не попал под фильтр — завершаю.")
            sys.exit(0)

    print(f"  Получено разделов: {len(sections)}")
    for s in sections:
        print(
            f"    • {s.get('key'):<12} "
            f"→ лист «{s.get('sheet_name') or '—'}», "
            f"JSON «{s.get('json_path') or '—'}»"
        )

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
        sys.exit(1)
    except Exception as e:
        print(f"[!] Ошибка открытия таблицы: {e}")
        traceback.print_exc()
        sys.exit(1)

    total_ok = 0
    total_records = 0
    failed = []

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

    print("\n" + "=" * 60)
    print(f"Обработано разделов:  {total_ok} из {len(sections)}")
    print(f"Всего записей:        {total_records}")
    if failed:
        print(f"\nПроблемные разделы ({len(failed)}):")
        for key, reason in failed:
            print(f"  • {key}: {reason}")
    print("=" * 60)

    if total_ok == 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
