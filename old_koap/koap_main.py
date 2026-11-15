"""
koap_main.py
════════════════════════════════════════════════════════════════════
Единый пайплайн обработки КоАП.

Объединяет три модуля:
  1. koap_pipeline  — загрузка КоАП + ГАДН, создание таблицы результатов
  2. testsym        — суммирование Прямых и Судебных статей
  3. opd_t_3        — перенос данных ГАДН в шаблон ОПД

Общие утилиты (чтение листа, суммирование словарей, запись ячеек)
вынесены в один блок и используются всеми тремя этапами.
════════════════════════════════════════════════════════════════════
"""

import os
import sys
from openpyxl import load_workbook, Workbook

# ══════════════════════════════════════════════════════════════════
# КОНСТАНТЫ
# ══════════════════════════════════════════════════════════════════

REGIONS = ['ДФО', 'ПФО', 'СЗФО', 'СКФО', 'УФО', 'ЦФО', 'ЮФО']

KEY_VSEGO = 'ВСЕГО:'
KEY_ITOGO = 'ИТОГО:'

AUTO_ARTICLES = [
    '12.21.1 ч.1 (в автоматическом режиме)',
    '12.21.1 ч.2 (в автоматическом режиме)',
    '12.21.1 ч.3 (в автоматическом режиме)',
    '12.21.1 ч.4 (в автоматическом режиме)',
    '12.21.1 ч.5 (в автоматическом режиме)',
    '12.21.1 ч.6 (в автоматическом режиме)',
    '12.21.3 ч.1 (в автоматическом режиме)',
    '12.21.3 ч.2 (в автоматическом режиме)',
    '12.21.4 ч.1',
    '12.21.4 ч.2',
]

ROW_WIDTH  = 40   # фиксированная ширина строки значений (этапы 1–2)
CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.opd_last_folder')


# ══════════════════════════════════════════════════════════════════
# ОБЩИЕ УТИЛИТЫ  (используются всеми тремя этапами)
# ══════════════════════════════════════════════════════════════════

def ws_to_dict(ws, allow_duplicates: bool = False) -> dict[str, list[float]]:
    """
    Читает лист ws в словарь за один проход.

    Ключ   — первая ячейка строки (str).
    Значение — список float остальных ячеек (None/нечисловое → 0.0).

    allow_duplicates=True: при повторяющихся ключах добавляет суффикс _2, _3, …
                           (нужно для opd_t_3, где ИТОГО: встречается дважды).
    allow_duplicates=False: повторный ключ перезаписывает предыдущий
                            (поведение koap_pipeline / testsym).
    """
    result:  dict[str, list[float]] = {}
    counter: dict[str, int]         = {}

    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        key = str(row[0]).strip()

        values: list[float] = []
        for val in row[1:]:
            try:
                values.append(float(val) if val is not None else 0.0)
            except (ValueError, TypeError):
                values.append(0.0)

        if allow_duplicates and key in result:
            counter[key] = counter.get(key, 2)
            unique_key   = f"{key}_{counter[key]}"
            counter[key] += 1
        else:
            unique_key = key

        result[unique_key] = values

    return result


def merge_dicts(dict_a: dict[str, list[float]],
                dict_b: dict[str, list[float]]) -> dict[str, list[float]]:
    """
    Объединяет два словаря статей:
      • Статья есть в обоих  — поэлементно складывает значения.
      • Статья только в одном — берёт как есть.

    Списки разной длины выравниваются нулями.
    """
    all_keys = set(dict_a) | set(dict_b)
    result: dict[str, list[float]] = {}

    for key in all_keys:
        list_a = dict_a.get(key, [])
        list_b = dict_b.get(key, [])
        max_len = max(len(list_a), len(list_b), 1)
        list_a  = list_a + [0.0] * (max_len - len(list_a))
        list_b  = list_b + [0.0] * (max_len - len(list_b))
        result[key] = [a + b for a, b in zip(list_a, list_b)]

    return result


def safe_get(d: dict, key: str, col: int) -> float:
    """Безопасно возвращает d[key][col] как float; при ошибке — 0.0."""
    row = d.get(key)
    if row is None:
        return 0.0
    try:
        val = row[col]
        return float(val) if val is not None else 0.0
    except (IndexError, TypeError, ValueError):
        return 0.0


def sum_col(d: dict, keys: list, col: int) -> float:
    """Суммирует значение одного столбца по списку статей."""
    return sum(safe_get(d, k, col) for k in keys)


def write_cells(ws, mapping: dict) -> None:
    """Записывает значения в ячейки шаблона по словарю {ячейка: значение}."""
    for cell, value in mapping.items():
        ws[cell] = value


def write_dict_to_sheet(ws, data: dict[str, list[float]], width: int = ROW_WIDTH) -> None:
    """
    Записывает словарь статей на лист ws.
    Строка: [ключ, val1, …, valN] — ровно width значений, ключи отсортированы.
    """
    for key in sorted(data):
        vals = (data[key] + [0.0] * width)[:width]
        ws.append([key] + vals)


def reset_sheet(wb, name: str):
    """Удаляет лист name (если есть) и создаёт новый пустой."""
    if name in wb.sheetnames:
        del wb[name]
    return wb.create_sheet(name)


# ══════════════════════════════════════════════════════════════════
# ЭТАП 1: koap_pipeline — загрузка блоков и создание таблицы
# ══════════════════════════════════════════════════════════════════

def collect_block(ws_source, marker_start: str) -> list[tuple]:
    """
    Извлекает строки одного блока из листа ws_source.

    Ищет строку-маркер marker_start, затем собирает строки
    до строки ИТОГО: (включительно).
    """
    rows_out: list[tuple] = []
    collecting = False

    for row in ws_source.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        txt = str(row[0]).strip()

        if txt == marker_start:
            collecting = True
            print(f"    ✔ Начало блока: «{marker_start}»")
            continue

        if collecting:
            rows_out.append(row)
            if txt.upper().startswith("ИТОГО:"):
                print(f"    ✔ Конец блока: {txt}")
                return rows_out

    if collecting:
        print(f"    ⚠ Блок «{marker_start}» не закрыт строкой ИТОГО:")
    else:
        print(f"    ⚠ Блок «{marker_start}» не найден")

    return rows_out


def run_koap_pipeline(filepath_koap: str, filepath_gadn: str, output_file: str) -> None:
    """
    Этап 1: читает файлы 4 надзора и ГАДН,
            суммирует блоки Прямых и Судебных статей,
            сохраняет Excel с 6 закладками.
    """
    print("\n" + "═" * 60)
    print("  ЭТАП 1: Пайплайн КоАП (4 надзора + ГАДН)")
    print("═" * 60)

    wb_koap = load_workbook(filepath_koap, data_only=True)
    wb_gadn = load_workbook(filepath_gadn,  data_only=True)
    ws_koap = wb_koap.active
    ws_gadn = wb_gadn.active

    print("  4 надзора:")
    direct   = collect_block(ws_koap, "Статьи прямого действия")
    judicial = collect_block(ws_koap, "Судебные статьи")

    print("  ГАДН:")
    gadn_d = collect_block(ws_gadn, "Статьи прямого действия")
    gadn_j = collect_block(ws_gadn, "Судебные статьи")

    # Суммируем через общие утилиты
    dict_direct   = ws_to_dict_from_rows(direct)
    dict_judicial = ws_to_dict_from_rows(judicial)
    dict_gadn_d   = ws_to_dict_from_rows(gadn_d)
    dict_gadn_j   = ws_to_dict_from_rows(gadn_j)

    sum_direct   = merge_dicts(dict_direct,   dict_gadn_d)
    sum_judicial = merge_dicts(dict_judicial, dict_gadn_j)

    print(f"\n  Прямых статей итого:   {len(sum_direct)}")
    print(f"  Судебных статей итого: {len(sum_judicial)}")

    # Запись в файл
    wb = Workbook()
    sheets_raw = [
        ("Прямые статьи",        direct),
        ("Судебные статьи",      judicial),
        ("Прямые статьи ГАДН",   gadn_d),
        ("Судебные статьи ГАДН", gadn_j),
    ]

    first = True
    for title, rows in sheets_raw:
        ws = wb.active if first else wb.create_sheet(title)
        if first:
            ws.title = title
            first = False
        for row in rows:
            ws.append(row)
        print(f"  Лист «{title}»: {ws.max_row} строк")

    for title, data in [("Сумма Прямых", sum_direct), ("Суммы Судебных", sum_judicial)]:
        ws = wb.create_sheet(title)
        write_dict_to_sheet(ws, data)
        print(f"  Лист «{title}»: {ws.max_row} строк")

    wb.save(output_file)
    print(f"\n  ✔ Сохранено: {output_file}")


def ws_to_dict_from_rows(rows: list[tuple]) -> dict[str, list[float]]:
    """
    Вспомогательная функция: превращает список кортежей (сырых строк)
    в словарь, как это делает ws_to_dict, — без чтения листа.
    """
    result: dict[str, list[float]] = {}
    for row in rows:
        if not row or row[0] is None:
            continue
        key = str(row[0]).strip()
        values: list[float] = []
        for val in row[1:]:
            try:
                values.append(float(val) if val is not None else 0.0)
            except (ValueError, TypeError):
                values.append(0.0)
        result[key] = values
    return result


# ══════════════════════════════════════════════════════════════════
# ЭТАП 2: testsym — суммирование Прямых и Судебных по 4 закладкам
# ══════════════════════════════════════════════════════════════════

def run_testsym(filepath: str) -> None:
    """
    Этап 2: читает файл с 4 закладками (Прямые / Судебные × 4надзора / ГАДН),
            суммирует попарно, добавляет 2 итоговых листа в тот же файл.
    """
    print("\n" + "═" * 60)
    print("  ЭТАП 2: Суммирование статей КоАП (testsym)")
    print("═" * 60)

    wb    = load_workbook(filepath, data_only=True)
    names = wb.sheetnames
    print(f"  Закладки: {names}")

    dict_direct   = ws_to_dict(wb[names[0]])
    dict_judicial = ws_to_dict(wb[names[1]])
    dict_gadn_d   = ws_to_dict(wb[names[2]])
    dict_gadn_j   = ws_to_dict(wb[names[3]])

    print(f"  [{names[0]}]: {len(dict_direct)} статей")
    print(f"  [{names[1]}]: {len(dict_judicial)} статей")
    print(f"  [{names[2]}]: {len(dict_gadn_d)} статей")
    print(f"  [{names[3]}]: {len(dict_gadn_j)} статей")

    sum_direct   = merge_dicts(dict_direct,   dict_gadn_d)
    sum_judicial = merge_dicts(dict_judicial, dict_gadn_j)

    print(f"\n  Прямых статей итого:   {len(sum_direct)}")
    print(f"  Судебных статей итого: {len(sum_judicial)}")

    ws_p = reset_sheet(wb, 'Сумма Прямых')
    write_dict_to_sheet(ws_p, sum_direct)
    print(f"  Лист 'Сумма Прямых': {ws_p.max_row} строк")

    ws_s = reset_sheet(wb, 'Суммы Судебных')
    write_dict_to_sheet(ws_s, sum_judicial)
    print(f"  Лист 'Суммы Судебных': {ws_s.max_row} строк")

    wb.save(filepath)
    print(f"\n  ✔ Сохранено: {filepath}")
    print(f"  Закладки: {wb.sheetnames}")


# ══════════════════════════════════════════════════════════════════
# ЭТАП 3: opd_t_3 — перенос ГАДН → шаблон ОПД
# ══════════════════════════════════════════════════════════════════

def run_opd(filepath_gadn: str, filepath_opd: str, region: str) -> None:
    """
    Этап 3: читает ГАДН, вычисляет показатели административной практики
            и автоматизированной фиксации, записывает в шаблон ОПД.
    """
    print("\n" + "═" * 60)
    print(f"  ЭТАП 3: ОПД — {region}")
    print("═" * 60)
    print(f"  ГАДН : {os.path.basename(filepath_gadn)}")
    print(f"  ОПД  : {os.path.basename(filepath_opd)}")

    wbgadn = load_workbook(filepath_gadn, data_only=True)
    wbopd  = load_workbook(filepath_opd)          # без data_only — сохраняем формулы
    d      = ws_to_dict(wbgadn.active, allow_duplicates=True)
    ws_opd = wbopd.active

    print(f"  Статей загружено: {len(d)}")

    # ── Административная практика ──────────────────────────────────
    val_1       = safe_get(d, KEY_VSEGO, 0)
    val_2       = safe_get(d, KEY_ITOGO, 4)
    val_3       = safe_get(d, KEY_ITOGO, 6)
    val_4       = safe_get(d, KEY_ITOGO, 32)
    val_5       = safe_get(d, KEY_ITOGO, 14)
    val_6       = safe_get(d, KEY_ITOGO, 35)
    val_7       = safe_get(d, KEY_ITOGO, 15)
    val_8       = safe_get(d, KEY_ITOGO, 17)
    val_9       = safe_get(d, KEY_ITOGO, 18)
    val_10      = safe_get(d, KEY_ITOGO, 25)
    val_11      = safe_get(d, KEY_ITOGO, 19)
    sum_val_11  = (safe_get(d, KEY_ITOGO, 27) + safe_get(d, KEY_ITOGO, 28) +
                   safe_get(d, KEY_ITOGO, 30) + safe_get(d, KEY_ITOGO, 31))

    # ── Автоматизированная фиксация ────────────────────────────────
    sum_auto_6          = sum_col(d, AUTO_ARTICLES, 4)
    val_12214_ch2_col16 = safe_get(d, '12.21.4 ч.2', 14)
    sum_auto_28         = sum_col(d, AUTO_ARTICLES, 32)
    sum_auto_16         = sum_col(d, AUTO_ARTICLES, 14)
    sum_auto_31         = sum_col(d, AUTO_ARTICLES, 35)
    sum_auto_17         = sum_col(d, AUTO_ARTICLES, 15)
    sum_auto_19         = sum_col(d, AUTO_ARTICLES, 17)
    sum_auto_20         = sum_col(d, AUTO_ARTICLES, 18)
    sum_auto_23         = sum_col(d, AUTO_ARTICLES, 25)
    sum_auto_21         = sum_col(d, AUTO_ARTICLES, 19)
    sum_auto_multi      = (sum_col(d, AUTO_ARTICLES, 27) + sum_col(d, AUTO_ARTICLES, 28) +
                           sum_col(d, AUTO_ARTICLES, 30) + sum_col(d, AUTO_ARTICLES, 31))

    # ── Запись в шаблон ────────────────────────────────────────────
    write_cells(ws_opd, {
        'C67': val_1,  'C68': val_2,      'C69': val_3,
        'C70': val_4,  'C71': val_5,      'C72': val_6,
        'C73': val_7,  'C74': val_8,      'C75': val_9,
        'C76': val_10, 'C77': sum_val_11, 'C78': val_11,
        'C81': sum_auto_6,
        'C82': val_12214_ch2_col16,
        'C83': sum_auto_28,
        'C84': sum_auto_16,
        'C85': sum_auto_31,
        'C86': sum_auto_17,
        'C87': sum_auto_19,
        'C88': sum_auto_20,
        'C89': sum_auto_23,
        'C90': sum_auto_multi,
        'C91': sum_auto_21,
    })

    out_name = os.path.join(os.path.dirname(filepath_opd), f"ОПД_{region}_результат.xlsx")
    wbopd.save(out_name)
    print(f"\n  ✔ Сохранено: {out_name}")


# ══════════════════════════════════════════════════════════════════
# ПОИСК ФАЙЛОВ
# ══════════════════════════════════════════════════════════════════

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
        print(f"  ⚠ Шаблон ОПД не найден автоматически.")
        if input("    Ввести путь вручную? (д/н): ").strip().lower() != 'д':
            return None
        path_opd = ask_file("    Путь к шаблону ОПД: ")
    else:
        print(f"  ✔ ОПД: {os.path.basename(path_opd)}")

    return path_gadn, path_opd


# ══════════════════════════════════════════════════════════════════
# ИНТЕРАКТИВНОЕ МЕНЮ
# ══════════════════════════════════════════════════════════════════

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


# ══════════════════════════════════════════════════════════════════
# ГЛАВНЫЙ ЦИКЛ
# ══════════════════════════════════════════════════════════════════

def run() -> None:
    print("\n" + "═" * 60)
    print("  КоАП — Единый пайплайн обработки")
    print("═" * 60)

    folder = menu_choose_folder()

    while True:
        region = menu_choose_region()
        if region is None:
            print("\n  До свидания!")
            sys.exit(0)

        print(f"\n  МТУ: {region}")
        stages = menu_choose_stages()

        if not stages:
            print("  ⚠ Этапы не выбраны, пропускаем.")
        else:
            # ── Этап 1: KoAP Pipeline ────────────────────────────────
            if '1' in stages:
                files = resolve_koap_files(folder, region)
                if files is None:
                    print("  Этап 1 пропущен.")
                else:
                    output = os.path.join(folder, f"Результат КоАП {region}.xlsx")
                    try:
                        run_koap_pipeline(*files, output)
                    except Exception as e:
                        print(f"  ✖ Ошибка этапа 1: {e}")

            # ── Этап 2: testsym ──────────────────────────────────────
            if '2' in stages:
                # Если этап 1 уже создал файл — используем его
                auto_file = os.path.join(folder, f"Результат КоАП {region}.xlsx")
                if os.path.isfile(auto_file):
                    filepath2 = auto_file
                    print(f"\n  Этап 2: используем файл из этапа 1 — {os.path.basename(auto_file)}")
                else:
                    filepath2 = ask_file(f"\n  Путь к файлу с 4 закладками для суммирования: ")
                try:
                    run_testsym(filepath2)
                except Exception as e:
                    print(f"  ✖ Ошибка этапа 2: {e}")

            # ── Этап 3: OPD ──────────────────────────────────────────
            if '3' in stages:
                files_opd = resolve_opd_files(folder, region)
                if files_opd is None:
                    print("  Этап 3 пропущен.")
                else:
                    try:
                        run_opd(*files_opd, region)
                    except Exception as e:
                        print(f"  ✖ Ошибка этапа 3: {e}")

        action = menu_next_action()
        if action == '0':
            print("\n  До свидания!")
            break
        elif action == '2':
            folder = menu_choose_folder()


if __name__ == "__main__":
    run()
