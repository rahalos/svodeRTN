"""
stage4_all_mtu.py — ЭТАП 4: сводный подсчёт значений по всем МТУ.

Модуль проходит по всем территориальным управлениям (8 списков:
7 МТУ + «Центральный» вариант ЦФО), для каждого читает файлы
КоАП (4 надзора) и ГАДН, суммирует блоки Прямых и Судебных статей
и складывает результаты ВСЕХ МТУ в общую сводку.

Результат сохраняется отдельной закладкой «Номера статей», где:
  • строка            — номер статьи КоАП;
  • столбцы B..I      — значения по каждому МТУ;
  • последний столбец — «ВСЕГО МТУ» (сумма по всем управлениям).

Использование:
    python -m koap_pkg.stage4_all_mtu <папка с файлами> [выходной.xlsx]
"""

from __future__ import annotations

import os

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment

from koap_pkg.config import REGIONS, ALL_MTU, ROW_WIDTH
from koap_pkg.filefinder import find_file
from koap_pkg.utils import merge_dicts, rows_to_dict


ALL_MTU_LISTS = ALL_MTU               # все 8 МТУ (7 округов + Центральный)

SHEET_NAME = 'Номера статей'          # имя итоговой закладки
SUM_COLUMN = 'ВСЕГО МТУ'              # имя сводного столбца

# Имена листов-источников внутри файла результатов (этап 1/2)
SOURCE_SHEETS = ('Сумма Прямых', 'Суммы Судебных')


def _find_region_files(folder: str, region: str) -> tuple | None:
    """Ищет пару файлов (КоАП, ГАДН) для одного МТУ без интерактива."""
    path_koap = find_file(folder, [region, 'КоАП'])
    path_gadn = find_file(folder, [region, 'ГАДН'])
    if path_koap and path_gadn:
        return path_koap, path_gadn
    return None


def collect_region_totals(folder: str, region: str) -> dict[str, list[float]]:
    """
    Считает сводные значения статей для одного МТУ:
    4 надзора + ГАДН (прямые и судебные статьи вместе).
    """
    from koap_pkg.stage1_koap_pipeline import collect_block

    files = _find_region_files(folder, region)
    if files is None:
        print(f"  ⚠ {region}: файлы не найдены, пропуск.")
        return {}

    path_koap, path_gadn = files
    print(f"  {region}:")
    print(f"    ✔ КоАП: {os.path.basename(path_koap)}")
    print(f"    ✔ ГАДН: {os.path.basename(path_gadn)}")

    wb_koap = load_workbook(path_koap, data_only=True)
    wb_gadn = load_workbook(path_gadn, data_only=True)

    total: dict[str, list[float]] = {}
    for ws, label in ((wb_koap.active, '4 надзора'), (wb_gadn.active, 'ГАДН')):
        for marker in ("Статьи прямого действия", "Судебные статьи"):
            block = [r for r in collect_block(ws, marker)
                     if not str(r[0]).strip().upper().startswith("ИТОГО:")]
            total = merge_dicts(total, rows_to_dict(block))
        print(f"    [{label}] статей после суммирования: {len(total)}")

    return total


def collect_from_result_file(filepath: str) -> dict[str, list[float]]:
    """
    Альтернативный источник: читает уже готовый файл результатов
    («Результат КоАП <МТУ>.xlsx») и возвращает объединённые суммы
    Прямых и Судебных статей.
    """
    from koap_pkg.utils import ws_to_dict

    wb = load_workbook(filepath, data_only=True)
    result: dict[str, list[float]] = {}
    for name in SOURCE_SHEETS:
        if name in wb.sheetnames:
            result = merge_dicts(result, ws_to_dict(wb[name]))
    return result


def write_summary_sheet(ws, mtu_data: dict[str, dict[str, list[float]]]) -> None:
    """
    Пишет сводную таблицу на лист ws:
      A — номер статьи, далее по одному столбцу на МТУ,
      последний столбец — сумма по всем МТУ (ВСЕГО МТУ).
    """
    regions = list(mtu_data.keys())
    header = ['Номер статьи'] + regions + [SUM_COLUMN]
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

    all_articles = sorted({art for data in mtu_data.values() for art in data})

    for article in all_articles:
        row: list = [article]
        region_totals: list[float] = []
        for region in regions:
            vals = mtu_data[region].get(article, [])
            # Берём первую числовую колонку (основной показатель);
            # если нужно суммировать все колонки — используем sum(vals)
            value = round(sum(vals), 2) if vals else 0.0
            row.append(value)
            region_totals.append(value)
        row.append(round(sum(region_totals), 2))
        ws.append(row)

    # Автоширина первой колонки
    ws.column_dimensions['A'].width = 45


def run_all_mtu(folder: str, output_file: str | None = None) -> str | None:
    """
    Этап 4: сводный подсчёт по всем МТУ → закладка «Номера статей».

    Возвращает путь к созданному файлу либо None, если данных нет.
    """
    print("\n" + "═" * 60)
    print("  ЭТАП 4: Сводный подсчёт по всем МТУ")
    print("═" * 60)

    if output_file is None:
        output_file = os.path.join(folder, "Сводка по всем МТУ.xlsx")

    mtu_data: dict[str, dict[str, list[float]]] = {}

    for region in ALL_MTU_LISTS:
        # Приоритет: готовые файлы результатов этапа 1/2
        result_file = os.path.join(folder, f"Результат КоАП {region}.xlsx")
        if os.path.isfile(result_file):
            print(f"  {region}: используется готовый результат "
                  f"«{os.path.basename(result_file)}»")
            data = collect_from_result_file(result_file)
        else:
            data = collect_region_totals(folder, region)
        if data:
            mtu_data[region] = data

    if not mtu_data:
        print("  ✖ Ни одного набора данных по МТУ не найдено — этап пропущен.")
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    write_summary_sheet(ws, mtu_data)

    wb.save(output_file)
    n_articles = ws.max_row - 1
    print(f"\n  ✔ Сохранено: {output_file}")
    print(f"  Закладка «{SHEET_NAME}»: {n_articles} номеров статей, "
          f"{len(mtu_data)} МТУ + столбец «{SUM_COLUMN}»")
    return output_file


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Использование: python -m koap_pkg.stage4_all_mtu "
              "<папка с файлами> [выходной.xlsx]")
        raise SystemExit(1)
    out = sys.argv[2] if len(sys.argv) > 2 else None
    run_all_mtu(sys.argv[1], out)
