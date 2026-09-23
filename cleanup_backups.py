#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cleanup_backups.py
Удаляет листы вида `_backup_<section>_<YYYYMMDD_HHMMSS>` старше N дней.
N задаётся переменной окружения CLEANUP_DAYS (по умолчанию 60).
"""

import os
import re
import json
import datetime
import traceback

import gspread
from google.oauth2.service_account import Credentials


SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
CLEANUP_DAYS   = int(os.environ.get("CLEANUP_DAYS", "60"))

BACKUP_PREFIX = "_backup_"
BACKUP_RE = re.compile(r"^_backup_(.+?)_(\d{8})_(\d{6})$")


def get_gspread_client():
    raw = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not raw:
        raise RuntimeError("GOOGLE_CREDENTIALS_JSON не задан")
    creds_dict = json.loads(raw)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def parse_backup_date(title):
    """
    Извлекает дату из имени листа вида:
      _backup_music_20260923_031500
    Возвращает datetime или None.
    """
    m = BACKUP_RE.match(title)
    if not m:
        return None
    date_str = m.group(2) + m.group(3)   # YYYYMMDD + HHMMSS
    try:
        return datetime.datetime.strptime(date_str, "%Y%m%d%H%M%S")
    except ValueError:
        return None


def main():
    print("=" * 60)
    print("cleanup_backups.py — старт")
    print(f"Удаляем backup'ы старше {CLEANUP_DAYS} дней")
    print("=" * 60)

    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    print(f"Таблица: {sh.title}")

    threshold = datetime.datetime.utcnow() - datetime.timedelta(days=CLEANUP_DAYS)
    print(f"Порог:    {threshold.isoformat()}")

    to_delete = []
    for ws in sh.worksheets():
        if not ws.title.startswith(BACKUP_PREFIX):
            continue
        dt = parse_backup_date(ws.title)
        if dt is None:
            print(f"  [i] Пропуск (нестандартное имя): {ws.title}")
            continue
        if dt < threshold:
            age = (datetime.datetime.utcnow() - dt).days
            to_delete.append((ws, dt, age))

    if not to_delete:
        print("\nНечего удалять.")
        return

    print(f"\nК удалению: {len(to_delete)}")
    for ws, dt, age in sorted(to_delete, key=lambda x: x[1]):
        print(f"  • {ws.title}  (возраст: {age} дн.)")

    deleted = 0
    for ws, _, _ in to_delete:
        try:
            sh.del_worksheet(ws)
            deleted += 1
            print(f"  ✓ Удалён: {ws.title}")
        except Exception as e:
            print(f"  [!] Ошибка удаления {ws.title}: {e}")

    print(f"\nИтого удалено: {deleted} из {len(to_delete)}")
    print("Готово!")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[!] Критическая ошибка: {e}")
        traceback.print_exc()
        raise
