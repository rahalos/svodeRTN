"""
opd_pipeline.py
════════════════════════════════════════════════════════════════════
Пайплайн переноса значений КоАП ГАДН → шаблон ОПД.

Оптимизации по сравнению с opd_t_3.py:
  • Вся логика «взять N значений из списка статей» заменена на
    одну универсальную функцию sum_col().
  • Повторяющиеся блоки fs_val_XX убраны — вместо 100+ переменных
    один вызов sum_col(dict_row, AUTO_ARTICLES, col_idx).
  • Интерактивное контекстное меню: выбор папки → выбор МТУ →
    автопоиск файлов → выполнение → продолжить/выйти.
════════════════════════════════════════════════════════════════════
"""

import os
import glob

from openpyxl import load_workbook

# ══════════════════════════════════════════════════════════════════
# КОНСТАНТЫ
# ══════════════════════════════════════════════════════════════════

REGIONS = ['ДФО', 'ПФО', 'СЗФО', 'СКФО', 'УФО', 'ЦФО', 'ЮФО']

# Статьи «автоматического режима» — одинаковы для всех столбцов
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

# Ключевые строки итогов
KEY_VSEGO = 'ВСЕГО:'
KEY_ITOGO = 'ИТОГО:'


# ══════════════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ══════════════════════════════════════════════════════════════════

def ws_to_dict(ws) -> dict[str, list[float]]:
    """
    Читает лист ws → словарь {строка_ячейки_A: [float, ...]}.
    Повторяющиеся ключи получают суффикс _2, _3, ...
    """
    result: dict[str, list[float]] = {}
    counter: dict[str, int] = {}

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

        if key in result:
            counter[key] = counter.get(key, 2)
            unique_key = f"{key}_{counter[key]}"
            counter[key] += 1
        else:
            unique_key = key

        result[unique_key] = values

    return result


def get_val(d: dict, key: str, col: int, default: float = 0.0) -> float:
    """Безопасно берёт значение по ключу и индексу столбца."""
    row = d.get(key)
    if row is None:
        return default
    try:
        return float(row[col]) if row[col] is not None else default
    except (IndexError, TypeError, ValueError):
        return default


def sum_col(d: dict, keys: list[str], col: int) -> float:
    """Суммирует значение одного столбца по списку статей."""
    return sum(get_val(d, k, col) for k in keys)


# ══════════════════════════════════════════════════════════════════
# ОСНОВНАЯ ЛОГИКА ОБРАБОТКИ
# ══════════════════════════════════════════════════════════════════

def process(filepath_gadn: str, filepath_opd: str, region: str) -> None:
    """
    Читает файл ГАДН, вычисляет все показатели,
    записывает в шаблон ОПД и сохраняет результат.
    """
    print(f"\n  Загрузка: {os.path.basename(filepath_gadn)}")
    wb_gadn = load_workbook(filepath_gadn, data_only=True)
    wb_opd  = load_workbook(filepath_opd,  data_only=True)

    ws_gadn = wb_gadn.active
    ws_opd  = wb_opd.active

    d = ws_to_dict(ws_gadn)
    print(f"  Статей загружено: {len(d)}")

    # ── Блок: Административная практика ──────────────────────────
    # col — индекс (0-based) в списке значений, т.е. столбец Excel минус 2
    # Столбец 2  → col 0
    # Столбец 6  → col 4
    # Столбец 8  → col 6   и т.д.

    val_1  = get_val(d, KEY_VSEGO, 0)   # кол-во возбуждённых дел (столбец 2, ВСЕГО)
    val_2  = get_val(d, KEY_ITOGO, 4)   # постановлений о привлечении (столбец 6)
    val_3  = get_val(d, KEY_ITOGO, 6)   # из них предупреждений (столбец 8)
    val_4  = get_val(d, KEY_ITOGO, 32)  # исполненных постановлений (столбец 28→34? → 32)
    val_5  = get_val(d, KEY_ITOGO, 14)  # отменённых при обжаловании (столбец 16)
    val_6  = get_val(d, KEY_ITOGO, 35)  # привлечено к отв-ти (столбец 31→36? → 35)
    val_7  = get_val(d, KEY_ITOGO, 15)  # сумма наложенных штрафов (столбец 17)
    val_8  = get_val(d, KEY_ITOGO, 17)  # из них обжалованных (столбец 19)
    val_9  = get_val(d, KEY_ITOGO, 18)  # штрафов за вычетом обжалованных (столбец 20)
    val_10 = get_val(d, KEY_ITOGO, 25)  # взысканных штрафов (столбец 23→26? → 25)
    val_11 = get_val(d, KEY_ITOGO, 19)  # столбец 21

    # Сумма столбцов 25 + 25.1 + 27 + 27.1  → индексы 27, 28, 30, 31
    sum_val_11 = (get_val(d, KEY_ITOGO, 27) + get_val(d, KEY_ITOGO, 28) +
                  get_val(d, KEY_ITOGO, 30) + get_val(d, KEY_ITOGO, 31))

    # ── Блок: Автоматизированная фиксация ────────────────────────
    # Используем sum_col — заменяет 10 отдельных переменных на каждый столбец

    sum_auto_col6  = sum_col(d, AUTO_ARTICLES, 4)   # столбец 6
    val_1221_4_2   = get_val(d, '12.21.4 ч.2', 14)  # статья 12.21.4 ч.2, столбец 16

    sum_auto_col28 = sum_col(d, AUTO_ARTICLES, 32)  # столбец 28
    sum_auto_col16 = sum_col(d, AUTO_ARTICLES, 14)  # столбец 16
    sum_auto_col31 = sum_col(d, AUTO_ARTICLES, 35)  # столбец 31
    sum_auto_col17 = sum_col(d, AUTO_ARTICLES, 15)  # столбец 17
    sum_auto_col19 = sum_col(d, AUTO_ARTICLES, 17)  # столбец 19
    sum_auto_col20 = sum_col(d, AUTO_ARTICLES, 18)  # столбец 20
    sum_auto_col23 = sum_col(d, AUTO_ARTICLES, 25)  # столбец 23

    # Сумма 25 + 25.1 + 27 + 27.1 по авто-статьям
    sum_auto_2525_1_27_271 = (sum_col(d, AUTO_ARTICLES, 27) +
                               sum_col(d, AUTO_ARTICLES, 28) +
                               sum_col(d, AUTO_ARTICLES, 30) +
                               sum_col(d, AUTO_ARTICLES, 31))

    sum_auto_col21 = sum_col(d, AUTO_ARTICLES, 19)  # столбец 21

    # ── Диагностический вывод ─────────────────────────────────────
    print("\n  ── Административная практика ──")
    print(f"  C67 Возбуждённых дел:              {val_1}")
    print(f"  C68 Постановлений о привлечении:   {val_2}")
    print(f"  C69 Из них предупреждений:         {val_3}")
    print(f"  C70 Исполненных постановлений:     {val_4}")
    print(f"  C71 Отменённых при обжаловании:    {val_5}")
    print(f"  C72 Привлечено к ответственности:  {val_6}")
    print(f"  C73 Сумма наложенных штрафов:      {val_7}")
    print(f"  C74 Из них обжалованных:           {val_8}")
    print(f"  C75 За вычетом обжалованных:       {val_9}")
    print(f"  C76 Взысканных штрафов:            {val_10}")
    print(f"  C77 25+25.1+27+27.1:               {sum_val_11}")
    print(f"  C78 Столбец 21:                    {val_11}")

    print("\n  ── Автоматизированная фиксация ──")
    print(f"  C81 Авто стол.6:   {sum_auto_col6}")
    print(f"  C82 12.21.4ч.2/16: {val_1221_4_2}")
    print(f"  C83 Авто стол.28:  {sum_auto_col28}")
    print(f"  C84 Авто стол.16:  {sum_auto_col16}")
    print(f"  C85 Авто стол.31:  {sum_auto_col31}")
    print(f"  C86 Авто стол.17:  {sum_auto_col17}")
    print(f"  C87 Авто стол.19:  {sum_auto_col19}")
    print(f"  C88 Авто стол.20:  {sum_auto_col20}")
    print(f"  C89 Авто стол.23:  {sum_auto_col23}")
    print(f"  C90 25+25.1+27+27.1: {sum_auto_2525_1_27_271}")
    print(f"  C91 Авто стол.21:  {sum_auto_col21}")

    # ── Запись в шаблон ОПД ───────────────────────────────────────
    ws_opd['C67'] = val_1
    ws_opd['C68'] = val_2
    ws_opd['C69'] = val_3
    ws_opd['C70'] = val_4
    ws_opd['C71'] = val_5
    ws_opd['C72'] = val_6
    ws_opd['C73'] = val_7
    ws_opd['C74'] = val_8
    ws_opd['C75'] = val_9
    ws_opd['C76'] = val_10
    ws_opd['C77'] = sum_val_11
    ws_opd['C78'] = val_11

    ws_opd['C81'] = sum_auto_col6
    ws_opd['C82'] = val_1221_4_2
    ws_opd['C83'] = sum_auto_col28
    ws_opd['C84'] = sum_auto_col16
    ws_opd['C85'] = sum_auto_col31
    ws_opd['C86'] = sum_auto_col17
    ws_opd['C87'] = sum_auto_col19
    ws_opd['C88'] = sum_auto_col20
    ws_opd['C89'] = sum_auto_col23
    ws_opd['C90'] = sum_auto_2525_1_27_271
    ws_opd['C91'] = sum_auto_col21

    # ── Сохранение ────────────────────────────────────────────────
    out_name = os.path.join(
        os.path.dirname(filepath_opd),
        f"ОПД_{region}_результат.xlsx"
    )
    wb_opd.save(out_name)
    print(f"\n  ✔ Сохранено: {out_name}")


# ══════════════════════════════════════════════════════════════════
# ПОИСК ФАЙЛОВ
# ══════════════════════════════════════════════════════════════════

def find_file(folder: str, region: str, kind: str) -> str | None:
    """
    Ищет файл в папке folder по маске вида:
      kind='gadn' → *{region}*ГАДН*.xlsx  или  *ГАДН*{region}*.xlsx
      kind='opd'  → *ОПД*.xlsx  (шаблон — не зависит от МТУ)

    Возвращает путь к первому найденному файлу или None.
    """
    region_up = region.upper()
    patterns: list[str] = []

    if kind == 'gadn':
        patterns = [
            f"*{region_up}*ГАДН*.xlsx",
            f"*ГАДН*{region_up}*.xlsx",
            f"*{region_up}ГАДН*.xlsx",
            f"*МТУ*{region_up}*.xlsx",
        ]
    elif kind == 'opd':
        patterns = [
            "*ОПД*шаблон*.xlsx",
            "*ОПД*.xlsx",
            "*opd*.xlsx",
        ]

    for pat in patterns:
        matches = glob.glob(os.path.join(folder, pat), flags=0)
        # glob не поддерживает кириллицу на Windows через flags, используем listdir
        matches = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.endswith('.xlsx') and _match_pattern(f.upper(), pat.upper().replace('*.XLSX', ''))
        ]
        if matches:
            return matches[0]

    return None


def _match_pattern(filename: str, keywords: str) -> bool:
    """Простая проверка: все непустые фрагменты шаблона (между *) есть в имени."""
    parts = [p for p in keywords.split('*') if p]
    return all(p in filename for p in parts)


def choose_file_manual(prompt: str) -> str:
    """Запрашивает путь вручную, пока файл не будет найден."""
    while True:
        path = input(prompt).strip().strip('"').strip("'")
        if os.path.isfile(path):
            return path
        print(f"  ⚠ Файл не найден: {path}. Попробуйте снова.")


def resolve_files(folder: str, region: str) -> tuple[str, str] | None:
    """
    Пытается найти файл ГАДН и шаблон ОПД автоматически.
    При неудаче предлагает ввести вручную.
    Возвращает (path_gadn, path_opd) или None при отказе.
    """
    print(f"\n  Поиск файлов в папке: {folder}")

    path_gadn = find_file(folder, region, 'gadn')
    path_opd  = find_file(folder, region, 'opd')

    if path_gadn:
        print(f"  ✔ ГАДН:   {os.path.basename(path_gadn)}")
    else:
        print(f"  ⚠ Файл ГАДН для {region} не найден автоматически.")
        choice = input("    Ввести путь вручную? (д/н): ").strip().lower()
        if choice != 'д':
            return None
        path_gadn = choose_file_manual("    Путь к файлу ГАДН: ")

    if path_opd:
        print(f"  ✔ ОПД:    {os.path.basename(path_opd)}")
    else:
        print(f"  ⚠ Шаблон ОПД не найден автоматически.")
        choice = input("    Ввести путь вручную? (д/н): ").strip().lower()
        if choice != 'д':
            return None
        path_opd = choose_file_manual("    Путь к шаблону ОПД: ")

    return path_gadn, path_opd


# ══════════════════════════════════════════════════════════════════
# ИНТЕРАКТИВНОЕ МЕНЮ
# ══════════════════════════════════════════════════════════════════

def menu_choose_folder() -> str:
    """
    Шаг 1: выбор рабочей папки.
    Предлагает последнюю использованную (если есть) или ввод вручную.
    """
    print("\n" + "═" * 55)
    print("  Шаг 1: Выбор папки с файлами")
    print("═" * 55)

    # Простой «кэш» последней папки в файле .last_folder
    cache_file = os.path.join(os.path.dirname(__file__), '.last_folder')
    last_folder = None
    if os.path.isfile(cache_file):
        with open(cache_file, 'r', encoding='utf-8') as f:
            last_folder = f.read().strip()

    if last_folder and os.path.isdir(last_folder):
        print(f"  Последняя папка: {last_folder}")
        choice = input("  Использовать её? (д/н): ").strip().lower()
        if choice == 'д':
            return last_folder

    while True:
        folder = input("  Введите путь к папке: ").strip().strip('"').strip("'")
        if os.path.isdir(folder):
            # Сохраняем в кэш
            with open(cache_file, 'w', encoding='utf-8') as f:
                f.write(folder)
            return folder
        print(f"  ⚠ Папка не найдена: {folder}")


def menu_choose_region() -> str | None:
    """
    Шаг 2: выбор МТУ (региона).
    Показывает нумерованный список, возвращает выбранный регион или None для выхода.
    """
    print("\n" + "═" * 55)
    print("  Шаг 2: Выбор МТУ")
    print("═" * 55)
    for i, r in enumerate(REGIONS, 1):
        print(f"  {i}. {r}")
    print("  0. Выйти из программы")

    while True:
        raw = input("\n  Введите номер МТУ: ").strip()
        if raw == '0':
            return None
        try:
            idx = int(raw) - 1
            if 0 <= idx < len(REGIONS):
                return REGIONS[idx]
        except ValueError:
            pass
        print("  ⚠ Неверный выбор. Введите число из списка.")


def run_interactive() -> None:
    """Главный цикл интерактивного меню."""
    print("\n" + "═" * 55)
    print("  ОПД Пайплайн: перенос значений КоАП ГАДН → ОПД")
    print("═" * 55)

    # ── Шаг 1: выбор папки (один раз за сеанс) ──
    folder = menu_choose_folder()

    while True:
        # ── Шаг 2: выбор МТУ ──
        region = menu_choose_region()
        if region is None:
            print("\n  До свидания!")
            break

        print(f"\n  Выбран МТУ: {region}")

        # ── Шаг 3: поиск файлов ──
        files = resolve_files(folder, region)
        if files is None:
            print("  Пропускаем этот МТУ.")
        else:
            path_gadn, path_opd = files

            # ── Шаг 4: выполнение ──
            try:
                process(path_gadn, path_opd, region)
                print(f"\n  ✔ МТУ {region} — обработан успешно.")
            except Exception as e:
                print(f"\n  ✖ Ошибка при обработке {region}: {e}")

        # ── Шаг 5: продолжить или выйти ──
        print("\n" + "─" * 55)
        print("  Что делаем дальше?")
        print("  1. Обработать другой МТУ")
        print("  2. Сменить рабочую папку")
        print("  0. Выйти")

        nxt = input("  Выбор: ").strip()
        if nxt == '0':
            print("\n  До свидания!")
            break
        elif nxt == '2':
            folder = menu_choose_folder()
        # nxt == '1' или любой другой → возврат к выбору МТУ


# ══════════════════════════════════════════════════════════════════
# ТОЧКА ВХОДА
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    run_interactive()
