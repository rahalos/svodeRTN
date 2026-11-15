from openpyxl import load_workbook, Workbook
import os

print('Работа по своду КоАП')
search_opd = ['ДФО', 'ПФО', 'СЗФО', 'СКФО', 'УФО', 'ЦФО', 'ЮФО']

print('Загрузка КоАП в разрезе МТУ')

# ─── Пути к файлам ────────────────────────────────────────────────────────────
filepath_koap      = r"E:\pyton\svod_ ter\КОАП\ДФО КоАП 4 надзора МАРТ..xlsx"
filepath_gadn_koap = r"E:\pyton\svod_ ter\КОАП\МТУ по ДФО ГАДН МАРТ.xlsx"

# ─── Загрузка книг ────────────────────────────────────────────────────────────
wbkoap = load_workbook(filepath_koap, data_only=True)
wbgadn = load_workbook(filepath_gadn_koap, data_only=True)
wskoap = wbkoap.active
wsgadn = wbgadn.active

# ─── Создаём новую книгу и закладки ──────────────────────────────────────────
new_koap = Workbook()
ws_nkoap       = new_koap.active
ws_nkoap.title = "Прямые статьи"
ws_nkoap_s     = new_koap.create_sheet("Судебные статьи")
ws_gkoap       = new_koap.create_sheet("Прямые статьи ГАДН")
ws_gkoap_s     = new_koap.create_sheet("Судебные статьи ГАДН")
ws_sum_p       = new_koap.create_sheet("Сумма Прямых")      # заполняется в testsym
ws_sum_s       = new_koap.create_sheet("Суммы Судебных")    # заполняется в testsym


# ─── ИСПРАВЛЕНИЕ: вынесена универсальная функция сбора блока ─────────────────
# БАГ оригинала: is_collecting НЕ сбрасывался между циклами —
# если блок не закрылся строкой ИТОГО:, следующий цикл продолжал писать.
# Решение: каждый вызов функции начинает с is_collecting = False.

def collect_block(ws_source, marker_start: str, ws_target):
    """
    Читает ws_source построчно.
    Начинает запись в ws_target при нахождении строки marker_start.
    Останавливается (включительно) на строке, начинающейся с 'ИТОГО:'.

    Параметры:
        ws_source    — лист-источник openpyxl
        marker_start — точная строка первой ячейки, открывающей блок
        ws_target    — лист назначения, куда пишем строки блока
    """
    is_collecting = False  # FIX: локальная переменная — нет утечки между вызовами

    for row in ws_source.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        txt = str(row[0]).strip()

        # Открываем блок
        if txt == marker_start:
            is_collecting = True
            print(f"  Найден блок: «{marker_start}»")
            continue  # саму строку-маркер не пишем

        if is_collecting:
            ws_target.append(row)

            # Закрываем блок строкой ИТОГО (включаем её в результат)
            if txt.upper().startswith("ИТОГО:"):
                print(f"  Конец блока: {txt}")
                is_collecting = False
                return  # блок закончен, дальше не читаем

    # Если дошли до конца файла и блок не закрылся — предупреждаем
    if is_collecting:
        print(f"  ВНИМАНИЕ: блок «{marker_start}» не закрыт строкой ИТОГО:")


# ─── Вызовы сбора данных ──────────────────────────────────────────────────────
print("\nОбработка 4 надзора...")
collect_block(wskoap, "Статьи прямого действия", ws_nkoap)
collect_block(wskoap, "Судебные статьи",          ws_nkoap_s)

print("\nОбработка ГАДН...")
collect_block(wsgadn, "Статьи прямого действия", ws_gkoap)
collect_block(wsgadn, "Судебные статьи",          ws_gkoap_s)

# ─── Сохранение ───────────────────────────────────────────────────────────────
out_name = f"Результат КоАП {search_opd[0]}.xlsx"
new_koap.save(out_name)
print(f"\nФайл сохранён: {out_name}")
print(f"Закладки: {new_koap.sheetnames}")
