"""
filefinder.py — поиск входных файлов в рабочей папке по ключевым словам.
"""

from __future__ import annotations

import os


def _contains_all(filename: str, keywords: list) -> bool:
    name_up = filename.upper()
    return all(kw.upper() in name_up for kw in keywords)


def find_file(folder: str, keywords: list) -> str | None:
    """Ищет первый .xlsx в папке, содержащий все ключевые слова в имени."""
    try:
        for fname in os.listdir(folder):
            if fname.lower().endswith('.xlsx') and _contains_all(fname, keywords):
                return os.path.join(folder, fname)
    except OSError as e:
        print(f"  ⚠ Ошибка чтения папки: {e}")
    return None


def ask_file(prompt: str) -> str:
    """Запрашивает у пользователя путь к существующему файлу."""
    while True:
        path = input(prompt).strip().strip('"').strip("'")
        if os.path.isfile(path):
            return path
        print(f"  ⚠ Файл не найден: {path}")


def resolve_koap_files(folder: str, region: str) -> tuple | None:
    """Ищет файлы КоАП (4 надзора) и ГАДН для этапа 1."""
    path_koap = find_file(folder, [region, 'КоАП'])
    if not path_koap:
        print(f"  ⚠ Файл КоАП для {region} не найден автоматически.")
        if input("    Ввести путь вручную? (д/н): ").strip().lower() != 'д':
            return None
        path_koap = ask_file("    Путь к файлу КоАП 4 надзора: ")
    else:
        print(f"  ✔ КоАП: {os.path.basename(path_koap)}")

    path_gadn = find_file(folder, [region, 'ГАДН'])
    if not path_gadn:
        print(f"  ⚠ Файл ГАДН для {region} не найден автоматически.")
        if input("    Ввести путь вручную? (д/н): ").strip().lower() != 'д':
            return None
        path_gadn = ask_file("    Путь к файлу ГАДН: ")
    else:
        print(f"  ✔ ГАДН: {os.path.basename(path_gadn)}")

    return path_koap, path_gadn


def resolve_opd_files(folder: str, region: str) -> tuple | None:
    """Ищет ГАДН и шаблон ОПД для этапа 3."""
    path_gadn = find_file(folder, [region, 'ГАДН'])
    if not path_gadn:
        print(f"  ⚠ Файл ГАДН для {region} не найден автоматически.")
        if input("    Ввести путь вручную? (д/н): ").strip().lower() != 'д':
            return None
        path_gadn = ask_file("    Путь к файлу ГАДН: ")
    else:
        print(f"  ✔ ГАДН: {os.path.basename(path_gadn)}")

    path_opd = find_file(folder, ['ОПД'])
    if not path_opd:
        print("  ⚠ Шаблон ОПД не найден автоматически.")
        if input("    Ввести путь вручную? (д/н): ").strip().lower() != 'д':
            return None
        path_opd = ask_file("    Путь к шаблону ОПД: ")
    else:
        print(f"  ✔ ОПД: {os.path.basename(path_opd)}")

    return path_gadn, path_opd
