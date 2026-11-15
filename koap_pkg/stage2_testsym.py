"""
stage2_testsym.py — ЭТАП 2: суммирование Прямых и Судебных статей
по 4 закладкам результата (КоАП + ГАДН), добавление итоговых листов.
"""

from __future__ import annotations

from openpyxl import load_workbook

from koap_pkg.utils import ws_to_dict, merge_dicts, write_dict_to_sheet, reset_sheet


def run_testsym(filepath: str) -> None:
    """
    Этап 2: читает файл с 4 закладками (Прямые / Судебные × 4надзора / ГАДН),
            суммирует попарно, добавляет 2 итоговых листа в тот же файл.
    """
    print("\n" + "═" * 60)
    print("  ЭТАП 2: Суммирование статей КоАП (testsym)")
    print("═" * 60)

    wb = load_workbook(filepath, data_only=True)
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


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Использование: python -m koap_pkg.stage2_testsym <Результат_КоАП.xlsx>")
        raise SystemExit(1)
    run_testsym(sys.argv[1])
