"""
stage1_koap_pipeline.py — ЭТАП 1: загрузка КоАП (4 надзора) + ГАДН,
суммирование блоков и создание таблицы результатов с 6 закладками.
"""

from __future__ import annotations

from openpyxl import load_workbook, Workbook

from koap_pkg.utils import merge_dicts, rows_to_dict, write_dict_to_sheet


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
    wb_gadn = load_workbook(filepath_gadn, data_only=True)
    ws_koap = wb_koap.active
    ws_gadn = wb_gadn.active

    print("  4 надзора:")
    direct   = collect_block(ws_koap, "Статьи прямого действия")
    judicial = collect_block(ws_koap, "Судебные статьи")

    print("  ГАДН:")
    gadn_d = collect_block(ws_gadn, "Статьи прямого действия")
    gadn_j = collect_block(ws_gadn, "Судебные статьи")

    # Суммируем через общие утилиты
    sum_direct   = merge_dicts(rows_to_dict(direct),   rows_to_dict(gadn_d))
    sum_judicial = merge_dicts(rows_to_dict(judicial), rows_to_dict(gadn_j))

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


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 4:
        print("Использование: python -m koap_pkg.stage1_koap_pipeline "
              "<КоАП.xlsx> <ГАДН.xlsx> <Результат.xlsx>")
        raise SystemExit(1)
    run_koap_pipeline(*sys.argv[1:3], sys.argv[3])
