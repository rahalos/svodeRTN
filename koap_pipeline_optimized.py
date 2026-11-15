"""
koap_pipeline.py
════════════════════════════════════════════════════════════════════
Пайплайн обработки КоАП: объединяет koap_p.py и testsym_v2.py.

Этап 1 (extract):  читает два исходных файла (4 надзора + ГАДН),
                   извлекает блоки «Прямые» и «Судебные» статей.

Этап 2 (sum):      строит словари, складывает значения по совпадающим
                   статьям, уникальные статьи берёт как есть.

Результат:         один Excel-файл с 6 закладками.
════════════════════════════════════════════════════════════════════
"""

from openpyxl import load_workbook, Workbook

# ══════════════════════════════════════════════════════════════════
# НАСТРОЙКИ  — меняйте только этот блок
# ══════════════════════════════════════════════════════════════════
FILEPATH_KOAP = r"E:\pyton\svod_ ter\КОАП\ДФО КоАП 4 надзора МАРТ..xlsx"
FILEPATH_GADN = r"E:\pyton\svod_ ter\КОАП\МТУ по ДФО ГАДН МАРТ.xlsx"
REGION        = "ДФО"           # подставляется в имя выходного файла
OUTPUT_FILE   = f"Результат КоАП {REGION}.xlsx"
ROW_WIDTH     = 40              # фиксированная ширина строки (колонки значений)


# ══════════════════════════════════════════════════════════════════
# ЭТАП 1: ИЗВЛЕЧЕНИЕ БЛОКОВ ИЗ ИСХОДНЫХ ФАЙЛОВ
# ══════════════════════════════════════════════════════════════════

def collect_block(ws_source, marker_start: str) -> list[tuple]:
    """
    Извлекает строки одного блока из листа ws_source.

    Логика:
      - Ищет строку, где первая ячейка == marker_start  → начало блока.
      - Собирает все последующие строки.
      - Останавливается (включительно) на строке с ИТОГО:.
      - Если ИТОГО: не найдено — предупреждает и возвращает собранное.

    Возвращает список кортежей (строк), без строки-маркера.
    """
    rows_out: list[tuple] = []
    collecting = False

    for row in ws_source.iter_rows(values_only=True):
        # Пропускаем полностью пустые строки и строки с пустой первой ячейкой
        if not row or row[0] is None:
            continue
        txt = str(row[0]).strip()

        if txt == marker_start:
            collecting = True
            print(f"    ✔ Начало блока: «{marker_start}»")
            continue  # строку-заголовок блока не включаем в результат

        if collecting:
            rows_out.append(row)
            if txt.upper().startswith("ИТОГО:"):
                print(f"    ✔ Конец блока: {txt}")
                return rows_out  # нормальное завершение

    if collecting:
        print(f"    ⚠ Блок «{marker_start}» не закрыт строкой ИТОГО:")
    else:
        print(f"    ⚠ Блок «{marker_start}» не найден в источнике")

    return rows_out


def extract_all_blocks(filepath_koap: str, filepath_gadn: str) -> dict[str, list[tuple]]:
    """
    Читает оба исходных файла и возвращает словарь из 4 списков строк:
      'direct'   — Прямые статьи 4 надзора
      'judicial' — Судебные статьи 4 надзора
      'gadn_d'   — Прямые статьи ГАДН
      'gadn_j'   — Судебные статьи ГАДН
    """
    print("\n─── Этап 1: извлечение блоков ───")

    wb_koap = load_workbook(filepath_koap, data_only=True)
    wb_gadn = load_workbook(filepath_gadn, data_only=True)
    ws_koap = wb_koap.active
    ws_gadn = wb_gadn.active

    print("  4 надзора:")
    direct   = collect_block(ws_koap, "Статьи прямого действия")
    judicial = collect_block(ws_koap, "Судебные статьи")

    print("  ГАДН:")
    gadn_d = collect_block(ws_gadn, "Статьи прямого действия")
    gadn_j = collect_block(ws_gadn, "Судебные статьи")

    return {
        "direct":   direct,
        "judicial": judicial,
        "gadn_d":   gadn_d,
        "gadn_j":   gadn_j,
    }


# ══════════════════════════════════════════════════════════════════
# ЭТАП 2: ПОСТРОЕНИЕ СЛОВАРЕЙ И СУММИРОВАНИЕ
# ══════════════════════════════════════════════════════════════════

def rows_to_dict(rows: list[tuple]) -> dict[str, list[float]]:
    """
    Преобразует список строк (кортежей) в словарь:
      ключ   — строковое значение первой ячейки (номер статьи)
      значение — список float остальных ячеек (0.0 если пусто/нечисловое)

    Строки с пустым или None-ключом пропускаются.
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


def sum_dicts(dict_a: dict[str, list[float]],
              dict_b: dict[str, list[float]]) -> dict[str, list[float]]:
    """
    Объединяет два словаря статей:
      - Если статья есть в обоих — поэлементно складывает значения.
      - Если только в одном — берёт как есть.

    Списки разной длины выравниваются нулями до длины большего.
    """
    all_keys = set(dict_a) | set(dict_b)   # объединение всех ключей
    result: dict[str, list[float]] = {}

    for key in all_keys:
        list_a = dict_a.get(key, [])
        list_b = dict_b.get(key, [])

        # Выравниваем длину
        max_len = max(len(list_a), len(list_b))
        list_a = list_a + [0.0] * (max_len - len(list_a))
        list_b = list_b + [0.0] * (max_len - len(list_b))

        result[key] = [a + b for a, b in zip(list_a, list_b)]

    return result


def build_sums(blocks: dict[str, list[tuple]]) -> tuple[
    dict[str, list[float]], dict[str, list[float]]
]:
    """
    Принимает 4 блока строк, строит 4 словаря, суммирует попарно.
    Возвращает (sum_direct, sum_judicial) — итоговые словари.
    """
    print("\n─── Этап 2: суммирование ───")

    dict_direct   = rows_to_dict(blocks["direct"])
    dict_judicial = rows_to_dict(blocks["judicial"])
    dict_gadn_d   = rows_to_dict(blocks["gadn_d"])
    dict_gadn_j   = rows_to_dict(blocks["gadn_j"])

    print(f"  Прямые 4 надзора: {len(dict_direct)} статей")
    print(f"  Судебные 4 надзора: {len(dict_judicial)} статей")
    print(f"  Прямые ГАДН: {len(dict_gadn_d)} статей")
    print(f"  Судебные ГАДН: {len(dict_gadn_j)} статей")

    sum_direct   = sum_dicts(dict_direct, dict_gadn_d)
    sum_judicial = sum_dicts(dict_judicial, dict_gadn_j)

    print(f"  Итого уникальных Прямых статей: {len(sum_direct)}")
    print(f"  Итого уникальных Судебных статей: {len(sum_judicial)}")

    return sum_direct, sum_judicial


# ══════════════════════════════════════════════════════════════════
# ЭТАП 3: ЗАПИСЬ В EXCEL
# ══════════════════════════════════════════════════════════════════

def write_sheet(ws, data_dict: dict[str, list[float]], width: int = ROW_WIDTH):
    """
    Записывает словарь статей на лист ws.
    Каждая строка: [ключ, val1, val2, ..., valN] — ровно width значений.
    Ключи сортируются для стабильного порядка строк.
    """
    for key in sorted(data_dict):
        vals = data_dict[key]
        # Приводим к фиксированной длине
        if len(vals) < width:
            vals = vals + [0.0] * (width - len(vals))
        else:
            vals = vals[:width]
        ws.append([key] + vals)


def save_result(blocks: dict[str, list[tuple]],
                sum_direct: dict[str, list[float]],
                sum_judicial: dict[str, list[float]],
                output_path: str):
    """
    Создаёт выходной Excel-файл с 6 закладками:
      1. Прямые статьи        — исходные данные 4 надзора
      2. Судебные статьи      — исходные данные 4 надзора
      3. Прямые статьи ГАДН   — исходные данные ГАДН
      4. Судебные статьи ГАДН — исходные данные ГАДН
      5. Сумма Прямых         — итоговая сумма прямых
      6. Суммы Судебных       — итоговая сумма судебных
    """
    print("\n─── Этап 3: запись результатов ───")
    wb = Workbook()

    # ── Исходные данные (4 блока сырых строк) ──
    sheets_raw = [
        ("Прямые статьи",        blocks["direct"]),
        ("Судебные статьи",      blocks["judicial"]),
        ("Прямые статьи ГАДН",   blocks["gadn_d"]),
        ("Судебные статьи ГАДН", blocks["gadn_j"]),
    ]

    first = True
    for title, rows in sheets_raw:
        if first:
            ws = wb.active
            ws.title = title
            first = False
        else:
            ws = wb.create_sheet(title)
        for row in rows:
            ws.append(row)
        print(f"  Записан лист «{title}»: {ws.max_row} строк")

    # ── Суммарные листы ──
    for title, data in [("Сумма Прямых", sum_direct), ("Суммы Судебных", sum_judicial)]:
        ws = wb.create_sheet(title)
        write_sheet(ws, data)
        print(f"  Записан лист «{title}»: {ws.max_row} строк")

    wb.save(output_path)
    print(f"\n✔ Файл сохранён: {output_path}")
    print(f"  Закладки: {wb.sheetnames}")


# ══════════════════════════════════════════════════════════════════
# ТОЧКА ВХОДА
# ══════════════════════════════════════════════════════════════════

def main():
    print("═" * 60)
    print("  Пайплайн обработки КоАП")
    print("═" * 60)

    # Шаг 1: извлекаем блоки строк из двух исходных файлов
    blocks = extract_all_blocks(FILEPATH_KOAP, FILEPATH_GADN)

    # Шаг 2: строим словари и суммируем
    sum_direct, sum_judicial = build_sums(blocks)

    # Шаг 3: пишем результат в Excel
    save_result(blocks, sum_direct, sum_judicial, OUTPUT_FILE)

    print("\n✔ Готово.")


if __name__ == "__main__":
    main()
