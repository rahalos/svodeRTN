"""
runner.py — главный цикл: связывает меню, поиск файлов и три этапа обработки.
"""

from __future__ import annotations

import os

from koap_pkg.filefinder import ask_file, resolve_koap_files, resolve_opd_files
from koap_pkg.menu import (
    menu_choose_folder,
    menu_choose_region,
    menu_choose_stages,
    menu_next_action,
    exit_program,
)
from koap_pkg.stage1_koap_pipeline import run_koap_pipeline
from koap_pkg.stage2_testsym import run_testsym
from koap_pkg.stage3_opd import run_opd
from koap_pkg.stage4_all_mtu import run_all_mtu


def stage1(folder: str, region: str) -> None:
    files = resolve_koap_files(folder, region)
    if files is None:
        print("  Этап 1 пропущен.")
        return
    output = os.path.join(folder, f"Результат КоАП {region}.xlsx")
    try:
        run_koap_pipeline(*files, output)
    except Exception as e:
        print(f"  ✖ Ошибка этапа 1: {e}")


def stage2(folder: str, region: str) -> None:
    # Если этап 1 уже создал файл — используем его
    auto_file = os.path.join(folder, f"Результат КоАП {region}.xlsx")
    if os.path.isfile(auto_file):
        filepath2 = auto_file
        print(f"\n  Этап 2: используем файл из этапа 1 — "
              f"{os.path.basename(auto_file)}")
    else:
        filepath2 = ask_file("\n  Путь к файлу с 4 закладками для суммирования: ")
    try:
        run_testsym(filepath2)
    except Exception as e:
        print(f"  ✖ Ошибка этапа 2: {e}")


def stage3(folder: str, region: str) -> None:
    files_opd = resolve_opd_files(folder, region)
    if files_opd is None:
        print("  Этап 3 пропущен.")
        return
    try:
        run_opd(*files_opd, region)
    except Exception as e:
        print(f"  ✖ Ошибка этапа 3: {e}")


def stage4(folder: str, region: str) -> None:
    # Этап 4 не зависит от выбранного МТУ — считает сводку по всем 8 спискам
    try:
        run_all_mtu(folder)
    except Exception as e:
        print(f"  ✖ Ошибка этапа 4: {e}")


STAGES = {'1': stage1, '2': stage2, '3': stage3, '4': stage4}


def run() -> None:
    print("\n" + "═" * 60)
    print("  КоАП — Единый пайплайн обработки")
    print("═" * 60)

    folder = menu_choose_folder()

    while True:
        region = menu_choose_region()
        if region is None:
            exit_program()

        print(f"\n  МТУ: {region}")
        stages = menu_choose_stages()

        if not stages:
            print("  ⚠ Этапы не выбраны, пропускаем.")
        else:
            for stage_id in sorted(stages):
                STAGES[stage_id](folder, region)

        action = menu_next_action()
        if action == '0':
            print("\n  До свидания!")
            break
        elif action == '2':
            folder = menu_choose_folder()
