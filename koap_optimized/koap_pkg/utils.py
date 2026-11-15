"""
utils.py — общие утилиты работы с Excel-данными.

Используются всеми тремя этапами пайплайна:
  • чтение листа / строк в словарь статей
  • поэлементное сложение словарей
  • безопасный доступ к ячейкам и суммирование по столбцам
  • запись словарей на листы, сброс листов
"""

from __future__ import annotations

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from koap_pkg.config import ROW_WIDTH


def _row_to_key_values(row) -> tuple[str, list[float]] | None:
    """Разбирает кортеж строки на (ключ, список float). None — если строка пустая."""
    if not row or row[0] is None:
        return None
    key = str(row[0]).strip()
    values = [
        _to_float(val) for val in row[1:]
    ]
    return key, values


def _to_float(val) -> float:
    """Приводит значение ячейки к float; всё нечисловое → 0.0."""
    try:
        return float(val) if val is not None else 0.0
    except (ValueError, TypeError):
        return 0.0


def ws_to_dict(ws: Worksheet, allow_duplicates: bool = False) -> dict[str, list[float]]:
    """
    Читает лист ws в словарь за один проход.

    Ключ     — первая ячейка строки (str).
    Значение — список float остальных ячеек (None/нечисловое → 0.0).

    allow_duplicates=True: при повторяющихся ключах добавляет суффикс _2, _3, …
                           (нужно для этапа ОПД, где ИТОГО: встречается дважды).
    allow_duplicates=False: повторный ключ перезаписывает предыдущий.
    """
    result: dict[str, list[float]] = {}
    counter: dict[str, int] = {}

    for row in ws.iter_rows(values_only=True):
        parsed = _row_to_key_values(row)
        if parsed is None:
            continue
        key, values = parsed

        if allow_duplicates and key in result:
            counter[key] = counter.get(key, 2)
            unique_key = f"{key}_{counter[key]}"
            counter[key] += 1
        else:
            unique_key = key

        result[unique_key] = values

    return result


def rows_to_dict(rows: list[tuple]) -> dict[str, list[float]]:
    """
    Превращает список кортежей (сырых строк) в словарь статей —
    без чтения листа (то же поведение, что ws_to_dict с allow_duplicates=False).
    """
    result: dict[str, list[float]] = {}
    for row in rows:
        parsed = _row_to_key_values(row)
        if parsed is None:
            continue
        key, values = parsed
        result[key] = values
    return result


def merge_dicts(dict_a: dict[str, list[float]],
                dict_b: dict[str, list[float]]) -> dict[str, list[float]]:
    """
    Объединяет два словаря статей:
      • Статья есть в обоих  — поэлементно складывает значения.
      • Статья только в одном — берёт как есть.

    Списки разной длины выравниваются нулями.
    """
    result: dict[str, list[float]] = {}

    for key in dict_a.keys() | dict_b.keys():
        list_a = dict_a.get(key, [])
        list_b = dict_b.get(key, [])
        max_len = max(len(list_a), len(list_b), 1)
        list_a = list_a + [0.0] * (max_len - len(list_a))
        list_b = list_b + [0.0] * (max_len - len(list_b))
        result[key] = [a + b for a, b in zip(list_a, list_b)]

    return result


def safe_get(d: dict, key: str, col: int) -> float:
    """Безопасно возвращает d[key][col] как float; при ошибке — 0.0."""
    row = d.get(key)
    if row is None:
        return 0.0
    try:
        return _to_float(row[col])
    except IndexError:
        return 0.0


def sum_col(d: dict, keys: list, col: int) -> float:
    """Суммирует значение одного столбца по списку статей."""
    return sum(safe_get(d, k, col) for k in keys)


def write_cells(ws: Worksheet, mapping: dict) -> None:
    """Записывает значения в ячейки шаблона по словарю {ячейка: значение}."""
    for cell, value in mapping.items():
        ws[cell] = value


def write_dict_to_sheet(ws: Worksheet, data: dict[str, list[float]],
                        width: int = ROW_WIDTH) -> None:
    """
    Записывает словарь статей на лист ws.
    Строка: [ключ, val1, …, valN] — ровно width значений, ключи отсортированы.
    """
    for key in sorted(data):
        vals = (data[key] + [0.0] * width)[:width]
        ws.append([key] + vals)


def reset_sheet(wb: Workbook, name: str) -> Worksheet:
    """Удаляет лист name (если есть) и создаёт новый пустой."""
    if name in wb.sheetnames:
        del wb[name]
    return wb.create_sheet(name)
