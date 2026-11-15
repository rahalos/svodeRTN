"""
stage3_opd.py — ЭТАП 3: перенос показателей ГАДН в шаблон ОПД.

Расчёт показателей административной практики и автоматизированной
фиксации, запись результатов в шаблон.

Оптимизация: вместо десятков однотипных строк используется таблица
соответствия «ячейка шаблона → столбец источника» и суммирование
по группам статей одним проходом.
"""

from __future__ import annotations

import os

from openpyxl import load_workbook

from koap_pkg.config import KEY_VSEGO, KEY_ITOGO, AUTO_ARTICLES
from koap_pkg.utils import ws_to_dict, safe_get, sum_col, write_cells

# ── Административная практика: ячейка → (ключ строки, столбец) ──────────
ADMIN_PRACTICE_CELLS: dict[str, tuple[str, int]] = {
    'C67': (KEY_VSEGO, 0),
    'C68': (KEY_ITOGO, 4),
    'C69': (KEY_ITOGO, 6),
    'C70': (KEY_ITOGO, 32),
    'C71': (KEY_ITOGO, 14),
    'C72': (KEY_ITOGO, 35),
    'C73': (KEY_ITOGO, 15),
    'C74': (KEY_ITOGO, 17),
    'C75': (KEY_ITOGO, 18),
    'C76': (KEY_ITOGO, 25),
    'C78': (KEY_ITOGO, 19),
}

# Ячейка C77 — сумма столбцов 27, 28, 30, 31 строки ИТОГО:
ADMIN_SUM_COLS: tuple[int, ...] = (27, 28, 30, 31)

# ── Автоматизированная фиксация: ячейка → столбец (сумма по AUTO_ARTICLES) ─
AUTO_FIXATION_CELLS: dict[str, int] = {
    'C81': 4,
    'C83': 32,
    'C84': 14,
    'C85': 35,
    'C86': 15,
    'C87': 17,
    'C88': 18,
    'C89': 25,
    'C91': 19,
}

# Ячейка C90 — сумма столбцов 27, 28, 30, 31 по AUTO_ARTICLES
AUTO_SUM_COLS: tuple[int, ...] = (27, 28, 30, 31)

# Исключение из общего правила: C82 — отдельная статья, свой столбец
AUTO_SPECIAL_CELL = ('C82', '12.21.4 ч.2', 14)


def build_opd_mapping(data: dict) -> dict[str, float]:
    """Формирует словарь {ячейка шаблона: значение} по данным ГАДН."""
    mapping: dict[str, float] = {}

    # Административная практика
    for cell, (key, col) in ADMIN_PRACTICE_CELLS.items():
        mapping[cell] = safe_get(data, key, col)
    mapping['C77'] = sum(safe_get(data, KEY_ITOGO, c) for c in ADMIN_SUM_COLS)

    # Автоматизированная фиксация
    for cell, col in AUTO_FIXATION_CELLS.items():
        mapping[cell] = sum_col(data, AUTO_ARTICLES, col)
    mapping['C90'] = sum(sum_col(data, AUTO_ARTICLES, c) for c in AUTO_SUM_COLS)

    cell, art_key, col = AUTO_SPECIAL_CELL
    mapping[cell] = safe_get(data, art_key, col)

    return mapping


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
    wbopd = load_workbook(filepath_opd)   # без data_only — сохраняем формулы
    # allow_duplicates=True: в ГАДН строка «ИТОГО:» встречается дважды
    data = ws_to_dict(wbgadn.active, allow_duplicates=True)
    ws_opd = wbopd.active

    print(f"  Статей загружено: {len(data)}")

    write_cells(ws_opd, build_opd_mapping(data))

    out_name = os.path.join(os.path.dirname(filepath_opd),
                            f"ОПД_{region}_результат.xlsx")
    wbopd.save(out_name)
    print(f"\n  ✔ Сохранено: {out_name}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 4:
        print("Использование: python -m koap_pkg.stage3_opd "
              "<ГАДН.xlsx> <ОПД_шаблон.xlsx> <МТУ>")
        raise SystemExit(1)
    run_opd(*sys.argv[1:4])
