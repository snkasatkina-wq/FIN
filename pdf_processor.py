"""
Модуль для обработки PDF выписок из банка.
Извлекает данные из PDF и преобразует их в формат для добавления в реестр.
"""

import pandas as pd
from pathlib import Path
from typing import List, Optional, Any
import datetime as dt
import re
import pdfplumber
import tabula


class PDFStatementProcessor:
    """Класс для обработки PDF выписок из банка."""
    
    def __init__(self, pdf_path: Path):
        """
        Инициализация процессора.
        
        Args:
            pdf_path: Путь к PDF файлу выписки
        """
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF файл не найден: {pdf_path}")
    
    def extract_tables_with_tabula(self) -> List[pd.DataFrame]:
        """
        Извлечение таблиц из PDF с помощью tabula-py.
        
        Returns:
            Список DataFrame с таблицами
        """
        try:
            tables = tabula.read_pdf(
                str(self.pdf_path),
                pages='all',
                multiple_tables=True,
                encoding='utf-8'
            )
            return [df for df in tables if df is not None and not df.empty]
        except Exception as e:
            print(f"Ошибка при извлечении таблиц через tabula: {e}")
            return []
    
    def extract_text_with_pdfplumber(self) -> str:
        """
        Извлечение текста из PDF с помощью pdfplumber.
        
        Returns:
            Весь текст из PDF
        """
        text = ""
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for page in pdf.pages:
                    text += page.extract_text() or ""
        except Exception as e:
            print(f"Ошибка при извлечении текста через pdfplumber: {e}")
        return text
    
    def parse_date(self, date_str: str) -> Optional[dt.date]:
        """
        Парсинг даты из строки в различных форматах.
        
        Args:
            date_str: Строка с датой
        
        Returns:
            Объект date или None
        """
        if not date_str or pd.isna(date_str):
            return None
        
        date_str = str(date_str).strip()
        
        # Различные форматы дат
        date_formats = [
            "%d.%m.%Y",
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d.%m.%y",
            "%d/%m/%y",
        ]
        
        for fmt in date_formats:
            try:
                return dt.datetime.strptime(date_str, fmt).date()
            except:
                continue
        
        # Попытка парсинга через pandas
        try:
            parsed = pd.to_datetime(date_str, dayfirst=True, errors='coerce')
            if pd.notna(parsed):
                return parsed.date()
        except:
            pass
        
        return None
    
    def clean_amount(self, amount_str: Any) -> Optional[float]:
        """
        Очистка и преобразование суммы в число.
        
        Args:
            amount_str: Строка или число с суммой
        
        Returns:
            Число или None
        """
        if pd.isna(amount_str):
            return None
        
        if isinstance(amount_str, (int, float)):
            return float(amount_str)
        
        # Удаляем пробелы, запятые, заменяем запятую на точку
        amount_str = str(amount_str).strip()
        amount_str = amount_str.replace(" ", "").replace(",", ".")
        
        # Удаляем все кроме цифр, точки и минуса
        amount_str = re.sub(r'[^\d.\-]', '', amount_str)
        
        try:
            return float(amount_str)
        except:
            return None
    
    def extract_operations_from_tables(self, tables: List[pd.DataFrame]) -> pd.DataFrame:
        """
        Извлечение операций из таблиц PDF.
        
        Args:
            tables: Список DataFrame с таблицами
        
        Returns:
            DataFrame с операциями
        """
        all_operations = []
        
        for table_idx, df in enumerate(tables):
            print(f"Обработка таблицы {table_idx + 1} ({len(df)} строк)...")
            
            # Ищем колонки с датами (обычно первая или вторая колонка)
            date_col_idx = None
            for col_idx in range(min(3, len(df.columns))):
                col = df.iloc[:, col_idx]
                # Проверяем, есть ли в колонке даты
                date_count = sum(1 for val in col if self.parse_date(str(val)) is not None)
                if date_count > len(col) * 0.3:  # Если больше 30% значений - даты
                    date_col_idx = col_idx
                    break
            
            if date_col_idx is None:
                print(f"  Не найдена колонка с датами в таблице {table_idx + 1}")
                continue
            
            # Обрабатываем каждую строку таблицы
            for row_idx, row in df.iterrows():
                date_val = row.iloc[date_col_idx] if date_col_idx < len(row) else None
                parsed_date = self.parse_date(str(date_val)) if date_val is not None else None
                
                if parsed_date is None:
                    continue
                
                # Формируем строку операции (аналогично формату из Excel)
                operation = {
                    'date': parsed_date,
                    'raw_row': row
                }
                
                # Извлекаем данные из всех колонок
                operation_data = []
                for col_idx in range(min(11, len(row))):  # Берем первые 11 колонок
                    val = row.iloc[col_idx] if col_idx < len(row) else None
                    operation_data.append(val)
                
                operation['data'] = operation_data
                all_operations.append(operation)
        
        if not all_operations:
            return pd.DataFrame()
        
        # Создаем DataFrame
        operations_list = []
        for op in all_operations:
            row_data = op['data']
            # Дополняем до 11 колонок
            while len(row_data) < 11:
                row_data.append(None)
            operations_list.append(row_data)
        
        df_ops = pd.DataFrame(operations_list)
        df_ops['date'] = [op['date'] for op in all_operations]
        
        # Сортируем по дате
        df_ops = df_ops.sort_values('date').reset_index(drop=True)
        
        return df_ops
    
    def process(self) -> pd.DataFrame:
        """
        Основной метод обработки PDF выписки.
        
        Returns:
            DataFrame с операциями в формате для реестра
        """
        print(f"Обработка PDF файла: {self.pdf_path.name}")
        
        # Сначала пробуем извлечь таблицы через tabula
        tables = self.extract_tables_with_tabula()
        
        if tables:
            print(f"Извлечено таблиц: {len(tables)}")
            operations = self.extract_operations_from_tables(tables)
            if not operations.empty:
                print(f"Извлечено операций из таблиц: {len(operations)}")
                return operations
        
        # Если таблицы не извлеклись, пробуем извлечь текст и парсить вручную
        print("Попытка извлечения данных из текста...")
        text = self.extract_text_with_pdfplumber()
        
        if text:
            operations = self.parse_text_operations(text)
            if not operations.empty:
                print(f"Извлечено операций из текста: {len(operations)}")
                return operations
        
        print("Не удалось извлечь операции из PDF")
        return pd.DataFrame()
    
    def parse_text_operations(self, text: str) -> pd.DataFrame:
        """
        Парсинг операций из текста PDF (резервный метод).
        
        Args:
            text: Текст из PDF
        
        Returns:
            DataFrame с операциями
        """
        operations = []
        lines = text.split('\n')
        
        # Ищем строки с датами
        date_pattern = re.compile(r'(\d{1,2}[./]\d{1,2}[./]\d{2,4})')
        
        current_operation = None
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Ищем дату в строке
            date_match = date_pattern.search(line)
            if date_match:
                date_str = date_match.group(1)
                parsed_date = self.parse_date(date_str)
                
                if parsed_date:
                    # Сохраняем предыдущую операцию
                    if current_operation:
                        operations.append(current_operation)
                    
                    # Начинаем новую операцию
                    parts = line.split()
                    current_operation = {
                        'date': parsed_date,
                        'data': parts[:11] if len(parts) >= 11 else parts + [None] * (11 - len(parts))
                    }
            elif current_operation:
                # Дополняем текущую операцию
                parts = line.split()
                current_operation['data'].extend(parts[:11])
                if len(current_operation['data']) >= 11:
                    current_operation['data'] = current_operation['data'][:11]
        
        # Добавляем последнюю операцию
        if current_operation:
            operations.append(current_operation)
        
        if not operations:
            return pd.DataFrame()
        
        # Формируем DataFrame
        ops_list = []
        for op in operations:
            row_data = op['data']
            while len(row_data) < 11:
                row_data.append(None)
            ops_list.append(row_data)
        
        df_ops = pd.DataFrame(ops_list)
        df_ops['date'] = [op['date'] for op in operations]
        df_ops = df_ops.sort_values('date').reset_index(drop=True)
        
        return df_ops


def process_pdf_statement(pdf_path: Path) -> pd.DataFrame:
    """
    Упрощенная функция для обработки PDF выписки.
    
    Args:
        pdf_path: Путь к PDF файлу
    
    Returns:
        DataFrame с операциями
    """
    processor = PDFStatementProcessor(pdf_path)
    return processor.process()
