"""
Вспомогательный скрипт для детального просмотра структуры Excel-файлов:
- исходной карточки счета 51
- листа 'реестр' в файле УУ2025

Запуск:
    python inspect_excel.py
"""

import pandas as pd


SOURCE_FILE = "Карточка счета 51 за 01.01.2026 - 04.02.2026.xlsx"
TARGET_FILE = "УУ2025 - оптим.xlsx"
TARGET_SHEET = "реестр"


def show_source():
    print("=" * 80)
    print("КАРТОЧКА СЧЕТА 51 (первые строки, без заголовков pandas)")
    print("=" * 80)
    df = pd.read_excel(SOURCE_FILE, sheet_name=0, header=None)
    print("Размер:", df.shape)
    # Покажем первые 30 строк и 15 столбцов
    print(df.iloc[:30, :15])

    # Попробуем показать предполагаемую строку заголовков таблицы проводок
    print("\nСтрока 7 (возможные заголовки таблицы):")
    if df.shape[0] > 7:
        row = df.iloc[7]
        for i, v in enumerate(row):
            print(f"  Колонка {i}: {repr(v)}")

    # Покажем первую строку с датой проводки (строка 8)
    print("\nСтрока 8 (первая проводка):")
    if df.shape[0] > 8:
        row8 = df.iloc[8]
        for i, v in enumerate(row8):
            print(f"  Колонка {i}: {repr(v)}")


def show_target():
    print("=" * 80)
    print("ЛИСТ 'реестр' В ФАЙЛЕ УУ2025 (первые строки, без заголовков pandas)")
    print("=" * 80)
    df = pd.read_excel(TARGET_FILE, sheet_name=TARGET_SHEET, header=None)
    print("Размер:", df.shape)
    # Покажем первые 15 строк и 40 столбцов
    print(df.iloc[:15, :40])

    # Печать строки заголовков (скорее всего строка 1)
    print("\nСтрока 1 (заголовки реестра):")
    if df.shape[0] > 1:
        row = df.iloc[1]
        for i, v in enumerate(row):
            print(f"  Колонка {i}: {repr(v)}")


def main():
    show_source()
    show_target()


if __name__ == "__main__":
    main()

