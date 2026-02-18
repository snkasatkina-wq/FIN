"""
Модуль для формирования отчетов на основе данных из реестра.
"""

import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


TARGET_FILE = "УУ2025 - оптим.xlsx"
TARGET_SHEET = "реестр"


def load_reestr_data() -> pd.DataFrame:
    """
    Загрузка данных из реестра.
    
    Returns:
        DataFrame с данными реестра
    """
    df = pd.read_excel(TARGET_FILE, sheet_name=TARGET_SHEET, header=None)
    
    # Преобразуем даты
    dates = pd.to_datetime(df.iloc[:, 3], dayfirst=True, errors="coerce")
    df['date'] = dates
    
    # Преобразуем суммы
    df['debit'] = pd.to_numeric(df.iloc[:, 7], errors='coerce')
    df['credit'] = pd.to_numeric(df.iloc[:, 10], errors='coerce')
    
    return df


def generate_summary_report() -> dict:
    """
    Генерация сводного отчета.
    
    Returns:
        Словарь с данными отчета
    """
    df = load_reestr_data()
    
    # Фильтруем только валидные строки с датами
    valid_df = df[df['date'].notna()].copy()
    
    if valid_df.empty:
        return {
            'total_operations': 0,
            'total_debit': 0,
            'total_credit': 0,
            'balance': 0,
            'by_month': {}
        }
    
    # Общие суммы
    total_debit = valid_df['debit'].sum()
    total_credit = valid_df['credit'].sum()
    balance = total_debit - total_credit
    
    # По месяцам
    valid_df['year_month'] = valid_df['date'].dt.to_period('M')
    by_month = {}
    
    for period, group in valid_df.groupby('year_month'):
        month_debit = group['debit'].sum()
        month_credit = group['credit'].sum()
        by_month[str(period)] = {
            'operations': len(group),
            'debit': month_debit,
            'credit': month_credit,
            'balance': month_debit - month_credit
        }
    
    return {
        'total_operations': len(valid_df),
        'total_debit': total_debit,
        'total_credit': total_credit,
        'balance': balance,
        'by_month': by_month,
        'first_date': valid_df['date'].min().date(),
        'last_date': valid_df['date'].max().date()
    }


def create_excel_report(report_data: dict, output_path: Path = None) -> Path:
    """
    Создание Excel отчета.
    
    Args:
        report_data: Данные отчета
        output_path: Путь для сохранения (по умолчанию автоматически)
    
    Returns:
        Путь к созданному файлу
    """
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = Path(f"Отчет_{timestamp}.xlsx")
    
    # Создаем новую книгу
    wb = load_workbook()
    ws = wb.active
    ws.title = "Сводка"
    
    # Стили
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    title_font = Font(bold=True, size=14)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Заголовок
    ws['A1'] = "ОТЧЕТ ПО ОПЕРАЦИЯМ"
    ws['A1'].font = title_font
    ws.merge_cells('A1:D1')
    
    row = 3
    
    # Общая информация
    ws[f'A{row}'] = "Общая информация"
    ws[f'A{row}'].font = title_font
    row += 1
    
    ws[f'A{row}'] = "Период:"
    ws[f'B{row}'] = f"{report_data['first_date']} - {report_data['last_date']}"
    row += 1
    
    ws[f'A{row}'] = "Всего операций:"
    ws[f'B{row}'] = report_data['total_operations']
    row += 1
    
    ws[f'A{row}'] = "Общий дебет:"
    ws[f'B{row}'] = report_data['total_debit']
    ws[f'B{row}'].number_format = '#,##0.00'
    row += 1
    
    ws[f'A{row}'] = "Общий кредит:"
    ws[f'B{row}'] = report_data['total_credit']
    ws[f'B{row}'].number_format = '#,##0.00'
    row += 1
    
    ws[f'A{row}'] = "Сальдо:"
    ws[f'B{row}'] = report_data['balance']
    ws[f'B{row}'].number_format = '#,##0.00'
    row += 2
    
    # По месяцам
    ws[f'A{row}'] = "Детализация по месяцам"
    ws[f'A{row}'].font = title_font
    row += 1
    
    # Заголовки таблицы
    headers = ["Месяц", "Операций", "Дебет", "Кредит", "Сальдо"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center')
    row += 1
    
    # Данные по месяцам
    for month, data in sorted(report_data['by_month'].items()):
        ws.cell(row=row, column=1, value=month).border = border
        ws.cell(row=row, column=2, value=data['operations']).border = border
        ws.cell(row=row, column=3, value=data['debit']).border = border
        ws.cell(row=row, column=3).number_format = '#,##0.00'
        ws.cell(row=row, column=4, value=data['credit']).border = border
        ws.cell(row=row, column=4).number_format = '#,##0.00'
        ws.cell(row=row, column=5, value=data['balance']).border = border
        ws.cell(row=row, column=5).number_format = '#,##0.00'
        row += 1
    
    # Настройка ширины колонок
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    
    # Сохранение
    wb.save(output_path)
    return output_path


def generate_report(output_path: Path = None) -> Path:
    """
    Основная функция для генерации отчета.
    
    Args:
        output_path: Путь для сохранения
    
    Returns:
        Путь к созданному файлу
    """
    print("Формирование сводного отчета...")
    
    report_data = generate_summary_report()
    
    if report_data['total_operations'] == 0:
        print("Нет данных для отчета")
        return None
    
    report_path = create_excel_report(report_data, output_path)
    
    print(f"  Всего операций: {report_data['total_operations']}")
    print(f"  Период: {report_data['first_date']} - {report_data['last_date']}")
    print(f"  Сальдо: {report_data['balance']:,.2f}")
    
    return report_path
