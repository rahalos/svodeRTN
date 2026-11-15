"""
menu.py — интерактивное консольное меню: выбор папки, МТУ, этапов.
"""

from __future__ import annotations

import os
import sys

from koap_pkg.config import REGIONS, CACHE_FILE


def load_last_folder() -> str | None:
    if os.path.isfile(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                folder = f.read().strip()
            if os.path.isdir(folder):
                return folder
        except OSError:
            pass
    return None


def save_last_folder(folder: str) -> None:
    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            f.write(folder)
    except OSError:
        pass


def menu_choose_folder() -> str:
    print("\n" + "═" * 60)
    print("  Выбор папки с файлами")
    print("═" * 60)
    last = load_last_folder()
    if last:
        print(f"  Последняя папка: {last}")
        if input("  Использовать её? (д/н): ").strip().lower() == 'д':
            return last
    while True:
        folder = input("  Введите путь к папке: ").strip().strip('"').strip("'")
        if os.path.isdir(folder):
            save_last_folder(folder)
            return folder
        print(f"  ⚠ Папка не найдена: {folder}")


def menu_choose_region() -> str | None:
    print("\n" + "═" * 60)
    print("  Выбор МТУ")
    print("═" * 60)
    for i, r in enumerate(REGIONS, 1):
        print(f"  {i}. {r}")
    print("  0. Выйти")
    while True:
        raw = input("\n  Введите номер: ").strip()
        if raw == '0':
            return None
        try:
            idx = int(raw) - 1
            if 0 <= idx < len(REGIONS):
                return REGIONS[idx]
        except ValueError:
            pass
        print("  ⚠ Неверный выбор.")


def menu_choose_stages() -> set[str]:
    """Выбор этапов для выполнения (можно несколько)."""
    print("\n" + "─" * 60)
    print("  Выберите этапы (через запятую, например: 1,2,3)")
    print("  1 — КоАП пайплайн (4 надзора + ГАДН → таблица результатов)")
    print("  2 — Суммирование статей (testsym)")
    print("  3 — ОПД (ГАДН → шаблон ОПД)")
    print("  0 — Все этапы подряд")
    raw = input("  Выбор: ").strip()
    if raw == '0':
        return {'1', '2', '3'}
    return {s.strip() for s in raw.split(',') if s.strip() in ('1', '2', '3')}


def menu_next_action() -> str:
    print("\n" + "─" * 60)
    print("  Что делаем дальше?")
    print("  1. Обработать другой МТУ")
    print("  2. Сменить рабочую папку")
    print("  0. Выйти")
    while True:
        choice = input("  Выбор: ").strip()
        if choice in ('0', '1', '2'):
            return choice
        print("  ⚠ Введите 0, 1 или 2.")


def exit_program() -> None:
    print("\n  До свидания!")
    sys.exit(0)
