"""
Главный скрипт для автоматической обработки выписок из банка:
1. Загружает выписку из почты (mail.ru)
2. Обрабатывает PDF файл
3. Добавляет данные в реестр
4. Формирует отчет
"""

import sys
from pathlib import Path
from datetime import datetime
import datetime as dt
import pandas as pd

from email_fetcher import EmailFetcher, fetch_from_config
from pdf_processor import PDFStatementProcessor
from update_reestr_from_51 import (
    get_max_reestr_state,
    append_operations_to_reestr,
    fill_ct_and_project,
    TARGET_FILE,
    TARGET_SHEET
)
from report_generator import generate_report


def process_pdf_to_excel_format(pdf_ops: pd.DataFrame) -> pd.DataFrame:
    """
    Преобразует операции из PDF формата в формат Excel карточки счета.
    
    Args:
        pdf_ops: DataFrame с операциями из PDF
    
    Returns:
        DataFrame в формате Excel карточки (колонки 0-10 + 'date')
    """
    # Создаем список строк для DataFrame
    rows_list = []
    
    for idx, row in pdf_ops.iterrows():
        # Берем данные из колонок 0-10
        row_data = []
        
        # Если есть колонка 'data' (список), используем её
        if 'data' in row and isinstance(row['data'], list):
            row_data = row['data'][:11]  # Берем первые 11 элементов
        else:
            # Иначе берем первые 11 колонок DataFrame
            for i in range(11):
                if i < len(row):
                    row_data.append(row.iloc[i] if hasattr(row, 'iloc') else row[i])
                else:
                    row_data.append(None)
        
        # Убеждаемся, что у нас ровно 11 элементов
        while len(row_data) < 11:
            row_data.append(None)
        
        # Добавляем дату в первую колонку, если она есть
        date_val = row.get('date')
        if date_val and pd.notna(date_val):
            if isinstance(date_val, pd.Timestamp):
                row_data[0] = date_val.strftime("%d.%m.%Y")
            elif isinstance(date_val, dt.date):
                row_data[0] = date_val.strftime("%d.%m.%Y")
            else:
                row_data[0] = str(date_val)
        
        rows_list.append(row_data)
    
    # Создаем DataFrame с колонками 0-10
    excel_ops = pd.DataFrame(rows_list)
    
    # Добавляем колонку date для совместимости с append_operations_to_reestr
    if 'date' in pdf_ops.columns:
        excel_ops['date'] = pdf_ops['date'].values
    else:
        # Пытаемся извлечь дату из первой колонки
        dates = pd.to_datetime(excel_ops.iloc[:, 0], dayfirst=True, errors='coerce')
        excel_ops['date'] = dates
    
    return excel_ops


def main():
    """Основная функция."""
    print("=" * 80)
    print("АВТОМАТИЧЕСКАЯ ОБРАБОТКА ВЫПИСОК ИЗ БАНКА")
    print("=" * 80)
    print()
    
    # Шаг 1: Загрузка выписки из почты
    print("ШАГ 1: Загрузка выписки из почты...")
    print("-" * 80)
    
    config_path = Path("email_config.txt")
    if config_path.exists():
        pdf_path = fetch_from_config(config_path)
    else:
        print("Файл конфигурации не найден. Используем интерактивный режим.")
        email_type = input("Тип почты (mail.ru/gmail/outlook): ").strip() or "mail.ru"
        email_address = input("Email адрес: ").strip()
        password = input("Пароль: ").strip()
        
        fetcher = EmailFetcher(
            email_type=email_type,
            email_address=email_address,
            password=password
        )
        
        pdf_path = fetcher.fetch_bank_statement(
            subject_keywords=["выписка", "карточка", "счет"],
            days_back=7
        )
    
    if not pdf_path:
        print("✗ Не удалось загрузить выписку из почты")
        return
    
    print(f"✓ Выписка загружена: {pdf_path.name}")
    print()
    
    # Шаг 2: Обработка PDF
    print("ШАГ 2: Обработка PDF выписки...")
    print("-" * 80)
    
    try:
        processor = PDFStatementProcessor(pdf_path)
        pdf_ops = processor.process()
        
        if pdf_ops.empty:
            print("✗ Не удалось извлечь операции из PDF")
            return
        
        print(f"✓ Извлечено операций: {len(pdf_ops)}")
        
        if 'date' in pdf_ops.columns:
            first_date = pdf_ops['date'].min()
            last_date = pdf_ops['date'].max()
            print(f"  Диапазон дат: с {first_date} по {last_date}")
        print()
        
    except Exception as e:
        print(f"✗ Ошибка при обработке PDF: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # Шаг 3: Преобразование в формат Excel и добавление в реестр
    print("ШАГ 3: Добавление операций в реестр...")
    print("-" * 80)
    
    try:
        # Преобразуем в формат Excel
        excel_ops = process_pdf_to_excel_format(pdf_ops)
        
        # Определяем состояние реестра
        max_reestr_date, _, max_id = get_max_reestr_state()
        
        if max_reestr_date is None:
            print("В реестре еще нет дат. Будем добавлять все операции.")
            new_ops = excel_ops
        else:
            print(f"Максимальная дата в реестре: {max_reestr_date}")
            # Фильтруем операции строго позже максимальной даты
            if 'date' in excel_ops.columns:
                new_ops = excel_ops[excel_ops['date'].dt.date > max_reestr_date].copy()
            else:
                new_ops = excel_ops
        
        print(f"Найдено новых операций для добавления: {len(new_ops)}")
        
        if not new_ops.empty:
            # Добавляем операции в реестр
            append_operations_to_reestr(new_ops, start_id=max_id)
            
            # Заполняем Ст1/Ст2/Проект
            print("Заполнение полей Ст1/Ст2...")
            fill_ct_and_project()
            print("✓ Операции добавлены в реестр")
        else:
            print("Новых операций для добавления нет")
        print()
        
    except Exception as e:
        print(f"✗ Ошибка при добавлении в реестр: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # Шаг 4: Формирование отчета
    print("ШАГ 4: Формирование отчета...")
    print("-" * 80)
    
    try:
        report_path = generate_report()
        if report_path:
            print(f"✓ Отчет сформирован: {report_path}")
        else:
            print("✗ Не удалось сформировать отчет")
    except Exception as e:
        print(f"✗ Ошибка при формировании отчета: {e}")
        import traceback
        traceback.print_exc()
    
    print()
    print("=" * 80)
    print("ОБРАБОТКА ЗАВЕРШЕНА")
    print("=" * 80)


if __name__ == "__main__":
    main()
