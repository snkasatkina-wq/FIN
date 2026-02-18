"""
Скрипт для автоматического добавления новых проводок
из файла «Карточка счета 51 ...» в лист «реестр» файла «УУ2025 - оптим.xlsx».

Логика:
- Определяем максимальную дату проводки в листе «реестр» (колонка D / индекс 3).
- Из карточки счета берем все строки с датой (колонка A) строго позже этой даты.
- Для каждой такой проводки формируем строку в формате реестра и
  дописываем ее в конец листа, заполняя только колонки D–N.

Колонки листа «реестр» (по данным файла):
    0: ID
    1: Месяц
    2: Год
    3: Период (дата)
    4: Документ
    5: Аналитика Дт
    6: Аналитика Кт
    7: Дебет
    8: Сумма П
    9: 1
    10: Кредит
    11: Сумма С
    12: 2
    13: Назначение платежа
    14: Дт1
    15: Кт1 / Ст1
    16: Комментарий / Ст2
    17: Проект (если есть)

Колонки карточки счета 51 (по данным файла):
    0..10: Вся полезная информация по проводке
           (дата, документ, реквизиты, счета, суммы, признак Д/К и пр.)
    11   : Последняя колонка – сальдо (остаток по счету) – ее НЕ переносим.

Если правила сопоставления полей нужно будет скорректировать,
проще всего поменять функцию build_reestr_row() ниже.
"""

from __future__ import annotations

import datetime as _dt
from typing import Optional, Tuple, List, Any

import pandas as pd
from openpyxl import load_workbook
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
TARGET_FILE = SCRIPT_DIR / "УУ2025 - оптим.xlsx"
TARGET_SHEET = "реестр"


def find_latest_card_file() -> Path:
    """
    Ищет в папке скрипта (и подпапках) последний по дате изменения файл
    с именем «Карточка счета 51*» или «Анализ счета 51*»
    и расширением .xlsx или .xls.
    """
    prefixes = ("Карточка счета 51", "Анализ счета 51")
    candidates = []
    for ext in ("*.xlsx", "*.xls"):
        for path in SCRIPT_DIR.glob(ext):
            if path.name.startswith(prefixes):
                candidates.append(path)
    candidates = sorted(
        candidates,
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        existing = list(SCRIPT_DIR.glob("*.xlsx")) + list(SCRIPT_DIR.glob("*.xls"))
        hint = f"\nНайдены Excel в папке: {', '.join(f.name for f in existing)}" if existing else ""
        raise FileNotFoundError(
            f"Не найден файл 'Карточка счета 51*' или 'Анализ счета 51*' (.xlsx или .xls) "
            f"в папке: {SCRIPT_DIR}{hint}"
        )
    return candidates[0]


def load_card_operations() -> pd.DataFrame:
    """
    Читает файл карточки счета и возвращает dataframe только с реальными проводками.

    - Колонка 0 интерпретируется как дата (dd.mm.yyyy).
    - Строки, где дата не распознана, отбрасываются (заголовки, итоги и т.п.).
    - Для надежности дата также записывается в отдельную колонку 'date'.
    """
    source_path = find_latest_card_file()
    print(f"Используем файл карточки: {source_path.name}")

    engine = "xlrd" if source_path.suffix.lower() == ".xls" else None
    df = pd.read_excel(source_path, sheet_name=0, header=None, engine=engine)

    # Преобразуем первую колонку в даты (формат дд.мм.гггг).
    # Используем СЫРЫЕ значения: если в ячейке не дата (заголовок, итог сальдо),
    # to_datetime вернет NaT, и такая строка будет исключена.
    dates_raw = pd.to_datetime(df.iloc[:, 0], dayfirst=True, errors="coerce")

    # Для дальнейшей работы достаточно хранить распознанную дату как есть,
    # без протягивания вниз – каждая проводка в карточке занимает одну строку.
    df["date"] = dates_raw

    # Фильтр только реальных проводок: есть распознанная дата в первой колонке.
    ops = df[df["date"].notna()].copy()

    # Сортируем по дате на всякий случай.
    ops.sort_values("date", inplace=True)
    ops.reset_index(drop=True, inplace=True)
    return ops


def get_max_reestr_state() -> Tuple[Optional[_dt.date], int, int]:
    """
    Определяет состояние листа «реестр»:
    - максимальную дату проводки (колонка D),
    - индекс последней занятой строки,
    - максимальный ID проводки (колонка A).

    Возвращает:
        (максимальная_дата_или_None,
         номер_последней_занятой_строки_в_листе,
         максимальный_ID_или_0)
    """
    df = pd.read_excel(TARGET_FILE, sheet_name=TARGET_SHEET, header=None)

    # Колонка с датой – индекс 3 (D). Преобразуем в даты.
    dates = pd.to_datetime(df.iloc[:, 3], dayfirst=True, errors="coerce")
    max_ts = dates.max()

    max_date: Optional[_dt.date]
    if pd.isna(max_ts):
        max_date = None
    else:
        max_date = max_ts.date()

    # Последняя занятая строка (по dataframe) – длина df.
    last_row_index = len(df)  # номер строки в Excel будет last_row_index (т.к. первая строка = 1)

    # Максимальный ID по первой колонке (A).
    ids = pd.to_numeric(df.iloc[:, 0], errors="coerce")
    max_id_val = ids.max()
    max_id: int = int(max_id_val) if pd.notna(max_id_val) else 0

    return max_date, last_row_index, max_id


def build_reestr_row(src_row: pd.Series) -> List[Any]:
    """
    Формирует значения для одной строки реестра (только колонки D–N, индексы 3–13)
    из строки карточки счета.

    Требование: из карточки должна перенестись ВСЯ информация,
    кроме самой последней колонки (сальдо).

    Поэтому делаем простое копирование:
        D..N  <-  колонки 0..10 карточки (один в один, без перерасчетов).
    """
    # Просто берем первые 11 колонок исходной строки (0..10).
    return [src_row.get(i) for i in range(11)]


def append_operations_to_reestr(new_ops: pd.DataFrame, start_id: int) -> None:
    """
    Дописывает новые проводки в конец листа «реестр»,
    заполняя только колонки D–N.
    """
    if new_ops.empty:
        print("Новых проводок для добавления нет – реестр уже актуален.")
        return

    wb = load_workbook(TARGET_FILE)
    ws = wb[TARGET_SHEET]

    current_id = start_id

    for _, src_row in new_ops.iterrows():
        current_id += 1

        values_dn = build_reestr_row(src_row)

        # Следующая свободная строка в листе (openpyxl считает с 1).
        row_idx = ws.max_row + 1

        # Дата проводки для вычисления месяца и года.
        date_val = src_row.get("date")
        if pd.isna(date_val):
            month_val = None
            year_val = None
        else:
            month_val = int(date_val.month)
            year_val = int(date_val.year)

        # Колонки A–C: ID, месяц, год.
        ws.cell(row=row_idx, column=1, value=current_id)  # A: ID
        ws.cell(row=row_idx, column=2, value=month_val)   # B: месяц
        ws.cell(row=row_idx, column=3, value=year_val)    # C: год

        # Колонки D–N – это 4..14 в нумерации Excel.
        for offset, value in enumerate(values_dn):
            col_idx = 4 + offset
            ws.cell(row=row_idx, column=col_idx, value=value)

    wb.save(TARGET_FILE)
    print(f"Добавлено строк: {len(new_ops)}")


def fill_ct_and_project() -> None:
    """
    Заполняет поля Ст1, Ст2 и Проект по правилу:
    - Для строк, где Ст1 и Ст2 пусты (и Проект пуст),
      ищем выше строку с такой же аналитикой Дт/Кт,
      у которой Ст1/Ст2 заполнены и Проект пуст.
    - Берем ближайшую сверху подходящую строку и копируем Ст1/Ст2.
    - Если ничего не найдено — оставляем поля пустыми.

    Колонки (в листе «реестр»):
        Аналитика Дт : F (6)
        Аналитика Кт : G (7)
        Ст1          : O (15)
        Ст2          : P (16)
        Проект       : Q (17)
    """
    wb = load_workbook(TARGET_FILE)
    ws = wb[TARGET_SHEET]

    max_row = ws.max_row

    def norm(v: Any) -> Optional[str]:
        if v is None:
            return None
        s = str(v).strip()
        return s if s else None

    def same_analytic(a: Optional[str], b: Optional[str]) -> bool:
        """
        Сравнение аналитики Дт/Кт.
        Считаем совпадающими, если строки равны ИЛИ
        одна является подстрокой другой (на случай длинных описаний).
        """
        if a is None or b is None:
            return False
        if a == b:
            return True
        return a in b or b in a

    # Проходим по всем строкам с данными (начиная с 3-й строки, 1–2 — заголовки).
    for row in range(3, max_row + 1):
        anal_dt = norm(ws.cell(row=row, column=6).value)
        anal_kt = norm(ws.cell(row=row, column=7).value)

        ct1 = norm(ws.cell(row=row, column=15).value)
        ct2 = norm(ws.cell(row=row, column=16).value)
        project = norm(ws.cell(row=row, column=17).value)

        # Нас интересуют только строки, где Ст1 и Ст2 еще не заполнены
        # и проект не указан.
        if ct1 is not None or ct2 is not None or project is not None:
            continue
        if anal_dt is None and anal_kt is None:
            continue

        # Ищем образцовую строку выше (до строки 2, строка 1 — заголовки).
        for r2 in range(row - 1, 1, -1):
            anal_dt2 = norm(ws.cell(row=r2, column=6).value)
            anal_kt2 = norm(ws.cell(row=r2, column=7).value)

            if not same_analytic(anal_dt2, anal_dt) or not same_analytic(anal_kt2, anal_kt):
                continue

            ct1_2 = norm(ws.cell(row=r2, column=15).value)
            ct2_2 = norm(ws.cell(row=r2, column=16).value)
            project2 = norm(ws.cell(row=r2, column=17).value)

            # Образцовую строку можно использовать только если у нее
            # Ст1/Ст2 заполнены и Проект пуст.
            if (ct1_2 is None and ct2_2 is None) or project2 is not None:
                continue

            # Копируем значения Ст1/Ст2.
            ws.cell(row=row, column=15, value=ct1_2)
            ws.cell(row=row, column=16, value=ct2_2)
            break

    wb.save(TARGET_FILE)


def main() -> None:
    print("Загрузка карточки счета 51...")
    ops = load_card_operations()
    if ops.empty:
        print("Не удалось найти ни одной строки с датой в карточке счета.")
        return

    first_date = ops["date"].min().date()
    last_date = ops["date"].max().date()
    print(f"Диапазон дат в карточке: c {first_date} по {last_date}")

    print("Определение состояния реестра...")
    max_reestr_date, _, max_id = get_max_reestr_state()
    if max_reestr_date is None:
        print("В реестре еще нет дат (или все даты не распознаны). Будем добавлять все проводки.")
        new_ops = ops
    else:
        print(f"Максимальная дата в реестре: {max_reestr_date}")
        # Берем проводки строго ПОЗЖЕ максимальной даты в реестре.
        new_ops = ops[ops["date"].dt.date > max_reestr_date].copy()

    print(f"Найдено новых проводок для добавления: {len(new_ops)}")
    append_operations_to_reestr(new_ops, start_id=max_id)

    # После добавления/обновления проводок заполняем Ст1/Ст2/Проект по правилу.
    print("Заполнение полей Ст1/Ст2 по ранее заданным проводкам...")
    fill_ct_and_project()


if __name__ == "__main__":
    main()

