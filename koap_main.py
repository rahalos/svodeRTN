"""
koap_main.py — точка входа.

Вся логика вынесена в пакет koap_pkg:
  config.py                — константы
  utils.py                 — общие утилиты Excel
  filefinder.py            — поиск входных файлов
  stage1_koap_pipeline.py  — этап 1 (КоАП + ГАДН → таблица результатов)
  stage2_testsym.py        — этап 2 (суммирование закладок)
  stage3_opd.py            — этап 3 (ГАДН → шаблон ОПД)
  menu.py                  — интерактивное меню
  runner.py                — главный цикл
"""

from koap_pkg import run

if __name__ == "__main__":
    run()
