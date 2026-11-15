"""
opd_svod.py  (бывший opd_t_3.py)
════════════════════════════════════════════════════════════════════
Свод ОПД ГАДН: читает файл КоАП ГАДН, извлекает показатели
по строкам ВСЕГО: / ИТОГО:, выводит в консоль.

Меню позволяет выбрать файлы из указанной папки по номеру.
════════════════════════════════════════════════════════════════════
"""

import os
from openpyxl import load_workbook

# ══════════════════════════════════════════════════════════════════
# НАСТРОЙКИ
# ══════════════════════════════════════════════════════════════════
DEFAULT_FOLDER = r"E:\pyton\svod_ ter\КОАП"


# ══════════════════════════════════════════════════════════════════
# УТИЛИТЫ МЕНЮ
# ══════════════════════════════════════════════════════════════════

def sep(char="═", w=60):
    print(char * w)


def get_xlsx_files(folder: str) -> list[str]:
    try:
        return sorted(
            f for f in os.listdir(folder)
            if f.lower().endswith(".xlsx")
            and os.path.isfile(os.path.join(folder, f))
        )
    except FileNotFoundError:
        print(f"  ⚠ Папка не найдена: {folder}")
        return []


def show_files(files: list[str]):
    sep("─")
    for i, name in enumerate(files, 1):
        print(f"  [{i:>2}]  {name}")
    sep("─")


def pick_file(files: list[str], prompt: str) -> str | None:
    while True:
        raw = input(f"  {prompt} (номер / Enter — пропустить): ").strip()
        if raw == "":
            return None
        if raw.isdigit() and 1 <= int(raw) <= len(files):
            return files[int(raw) - 1]
        print(f"  ⚠ Введите число от 1 до {len(files)}")


def choose_folder() -> str:
    sep()
    print(f"  Папка по умолчанию: {DEFAULT_FOLDER}")
    raw = input("  Другой путь (или Enter): ").strip().strip('"').strip("'")
    folder = raw if raw else DEFAULT_FOLDER
    if not os.path.isdir(folder):
        print(f"  ⚠ Папка не найдена: {folder}")
        return ""
    return folder


# ══════════════════════════════════════════════════════════════════
# ЛОГИКА ЧТЕНИЯ ДАННЫХ
# ══════════════════════════════════════════════════════════════════

def load_dict_with_duplicates(ws) -> dict[str, list[float]]:
    """
    Читает лист в словарь.
    Повторяющиеся ключи получают суффикс _2, _3 и т.д.,
    чтобы не терять строки ВСЕГО: и ИТОГО: если их несколько.
    """
    result:      dict[str, list[float]] = {}
    key_counter: dict[str, int]         = {}

    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        if not any(v is not None for v in row):
            continue

        key = str(row[0]).strip()
        values: list[float] = []
        for val in row[1:]:
            try:
                values.append(float(val) if val is not None else 0.0)
            except (ValueError, TypeError):
                values.append(0.0)

        if key in result:
            # дубль — добавляем суффикс
            key_counter[key] = key_counter.get(key, 2)
            unique_key = f"{key}_{key_counter[key]}"
            key_counter[key] += 1
        else:
            unique_key = key

        result[unique_key] = values

    return result


def safe_get(d: dict, key: str, idx: int):
    """Безопасно возвращает d[key][idx] или None."""
    row = d.get(key, [None] * 50)
    if idx < len(row):
        return row[idx]
    return None


# ══════════════════════════════════════════════════════════════════
# ВЫВОД ПОКАЗАТЕЛЕЙ ОПД
# ══════════════════════════════════════════════════════════════════

def print_opd(d: dict[str, list[float]]):
    """
    Выводит показатели ОПД ГАДН в консоль.
    Каждый показатель — строка из словаря по ключу ВСЕГО: или ИТОГО:
    с нужным индексом столбца.
    """
    V = 'ВСЕГО:'
    I = 'ИТОГО:'

    # Сумма столбцов 25 + 25.1 + 27 + 27.1 (индексы 27,28,30,31)
    s_vals = [safe_get(d, I, idx) or 0.0 for idx in [27, 28, 30, 31]]
    sum_auto = sum(s_vals)

    # Таблица: (метка, значение)
    rows = [
        ("Возбуждено административных дел, ед.",
            safe_get(d, V, 0)),

        ("Вынесено постановлений о привлечении к адм. ответственности, ед.",
            safe_get(d, I, 4)),

        ("из них в виде предупреждения, ед.",
            safe_get(d, I, 6)),

        ("Исполнено постановлений, вынесенных сотрудниками Ространснадзора, ед.",
            safe_get(d, I, 32)),

        ("Отменено постановлений в результате обжалования, ед.",
            safe_get(d, I, 14)),

        ("Привлечено к адм. ответственности, ед.",
            safe_get(d, I, 35)),

        ("Общая сумма наложенных штрафов — всего, тыс. руб.",
            safe_get(d, I, 15)),

        ("из них обжалованных, тыс. руб.",
            safe_get(d, I, 17)),

        ("Сумма штрафов за вычетом обжалованных, тыс. руб.",
            safe_get(d, I, 18)),

        ("Сумма взысканных (уплаченных) штрафов (столбец 23), тыс. руб.",
            safe_get(d, I, 25)),

        ("Сумма взысканных (авт. режим: 25 + 25.1 + 27 + 27.1), тыс. руб.",
            sum_auto),

        ("Сумма взысканных (уплаченных) штрафов (столбец 21), тыс. руб.",
            safe_get(d, I, 19)),
    ]

    sep("─")
    print("  ПОКАЗАТЕЛИ ОПД ГАДН")
    sep("─")
    for i, (label, val) in enumerate(rows, 1):
        val_str = str(val) if val is not None else "—"
        print(f"  {i:>2}. {label}")
        print(f"       → {val_str}")
    sep("─")


# ══════════════════════════════════════════════════════════════════
# ОСНОВНАЯ ФУНКЦИЯ СВОДА
# ══════════════════════════════════════════════════════════════════

def run_svod(folder: str, files: list[str]):
    sep()
    print("  СВОД ОПД ГАДН — выбор файлов")
    sep()
    show_files(files)

    # Файл КоАП ГАДН (источник данных)
    f_gadn = pick_file(files, "КоАП ГАДН (источник данных)")
    if not f_gadn:
        print("  Файл не выбран — отменено.")
        return

    print(f"\n  КоАП ГАДН : {f_gadn}")

    ws_gadn = load_workbook(
        os.path.join(folder, f_gadn), data_only=True
    ).active

    d = load_dict_with_duplicates(ws_gadn)
    print(f"  Строк прочитано: {len(d)}")

    print_opd(d)
    print("\n  ✔ ОПД ГАДН выведен.")


# ══════════════════════════════════════════════════════════════════
# МЕНЮ
# ══════════════════════════════════════════════════════════════════

def main():
    sep()
    print("  СВОД ОПД ГАДН")
    sep()

    folder = choose_folder()
    if not folder:
        input("  Нажмите Enter для выхода...")
        return

    files = get_xlsx_files(folder)
    if not files:
        print("  В папке нет .xlsx файлов.")
        input("  Нажмите Enter для выхода...")
        return

    print(f"  Найдено файлов: {len(files)}")

    while True:
        sep()
        print("  [1] Запустить свод ОПД ГАДН")
        print("  [2] Сменить папку")
        print("  [0] Выход")
        sep("─")
        choice = input("  Выбор: ").strip()

        if choice == "1":
            run_svod(folder, files)

        elif choice == "2":
            folder = choose_folder()
            if folder:
                files = get_xlsx_files(folder)
                print(f"  Найдено файлов: {len(files)}")

        elif choice == "0":
            print("  До свидания!")
            break
        else:
            print("  ⚠ Введите 0, 1 или 2")

        input("\n  Нажмите Enter для возврата в меню...")


if __name__ == "__main__":
    main()
