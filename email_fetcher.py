"""
Модуль для загрузки файлов выписки из банка из почты.
Поддерживает Gmail, Outlook и Mail.ru (через IMAP).
"""

import imaplib
import email
from email.header import decode_header
from pathlib import Path
from typing import Optional, List
import re
from datetime import datetime


class EmailFetcher:
    """Класс для получения файлов из почты."""
    
    def __init__(self, email_type: str = "gmail", email_address: str = "", 
                 password: str = "", imap_server: str = "", imap_port: int = 993):
        """
        Инициализация подключения к почте.
        
        Args:
            email_type: Тип почты ("gmail" или "outlook")
            email_address: Email адрес
            password: Пароль или app password для Gmail
            imap_server: IMAP сервер (автоматически для gmail/outlook)
            imap_port: IMAP порт (по умолчанию 993)
        """
        self.email_address = email_address
        self.password = password
        
        if email_type.lower() == "gmail":
            self.imap_server = imap_server or "imap.gmail.com"
            self.imap_port = imap_port or 993
        elif email_type.lower() == "outlook":
            self.imap_server = imap_server or "outlook.office365.com"
            self.imap_port = imap_port or 993
        elif email_type.lower() in ["mail.ru", "mailru", "mail"]:
            # Определяем IMAP сервер по домену email
            if email_address and "@" in email_address:
                domain = email_address.split("@")[1].lower()
                # Для корпоративных доменов mail.ru может использоваться другой сервер
                if domain in ["mail.ru", "inbox.ru", "list.ru", "bk.ru"]:
                    self.imap_server = imap_server or "imap.mail.ru"
                else:
                    # Для других доменов пробуем стандартный mail.ru IMAP
                    self.imap_server = imap_server or "imap.mail.ru"
            else:
                self.imap_server = imap_server or "imap.mail.ru"
            self.imap_port = imap_port or 993
        else:
            self.imap_server = imap_server
            self.imap_port = imap_port
        
        self.mail = None
    
    def connect(self) -> bool:
        """Подключение к почтовому серверу."""
        try:
            if not self.email_address or not self.password:
                print("ОШИБКА: Email адрес или пароль не указаны в конфигурации")
                return False
            
            if not self.imap_server:
                print("ОШИБКА: IMAP сервер не указан")
                return False
            
            self.mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            self.mail.login(self.email_address, self.password)
            self.mail.select("INBOX")
            print(f"[OK] Подключение к почте успешно: {self.email_address}")
            return True
        except Exception as e:
            print(f"[ERROR] Ошибка подключения к почте: {e}")
            print(f"Проверьте:")
            print(f"  - Правильность email адреса: {self.email_address}")
            print(f"  - Правильность пароля")
            print(f"  - Включен ли IMAP в настройках mail.ru")
            print(f"  - IMAP сервер: {self.imap_server}:{self.imap_port}")
            return False
    
    def disconnect(self):
        """Закрытие соединения с почтой."""
        if self.mail:
            try:
                self.mail.close()
                self.mail.logout()
            except:
                pass
    
    def decode_mime_words(self, s: str) -> str:
        """Декодирование заголовков письма."""
        decoded_parts = decode_header(s)
        decoded_str = ""
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                if encoding:
                    decoded_str += part.decode(encoding)
                else:
                    decoded_str += part.decode('utf-8', errors='ignore')
            else:
                decoded_str += part
        return decoded_str
    
    def search_emails(self, subject_keywords: List[str] = None, 
                     sender_keywords: List[str] = None, 
                     days_back: int = 7) -> List[int]:
        """
        Поиск писем по критериям.
        
        Args:
            subject_keywords: Ключевые слова в теме письма
            sender_keywords: Ключевые слова в адресе отправителя
            days_back: За сколько дней назад искать (по умолчанию 7)
        
        Returns:
            Список ID писем
        """
        if not self.mail:
            return []
        
        # Поиск писем за последние N дней
        from datetime import timedelta
        date_criteria = (datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) 
                        - timedelta(days=days_back)).strftime("%d-%b-%Y")
        search_criteria = f'(SINCE {date_criteria})'
        
        # Добавляем фильтры по теме (используем OR для нескольких ключевых слов)
        if subject_keywords:
            if len(subject_keywords) == 1:
                subject_filter = f'SUBJECT "{subject_keywords[0]}"'
            else:
                # Для нескольких ключевых слов используем OR
                subject_parts = [f'SUBJECT "{kw}"' for kw in subject_keywords]
                subject_filter = f'({" OR ".join(subject_parts)})'
            search_criteria = f'({search_criteria} {subject_filter})'
        
        # Добавляем фильтры по отправителю (используем OR для нескольких ключевых слов)
        if sender_keywords:
            if len(sender_keywords) == 1:
                sender_filter = f'FROM "{sender_keywords[0]}"'
            else:
                # Для нескольких ключевых слов используем OR
                sender_parts = [f'FROM "{kw}"' for kw in sender_keywords]
                sender_filter = f'({" OR ".join(sender_parts)})'
            search_criteria = f'({search_criteria} {sender_filter})'
        
        try:
            status, messages = self.mail.search(None, search_criteria)
            if status == "OK":
                email_ids = messages[0].split()
                return [int(id) for id in email_ids]
        except Exception as e:
            print(f"Ошибка поиска писем: {e}")
        
        return []
    
    def download_attachments(self, email_id: int, 
                           filename_patterns: List[str] = None,
                           save_dir: Path = None) -> List[Path]:
        """
        Скачивание вложений из письма.
        
        Args:
            email_id: ID письма
            filename_patterns: Паттерны имен файлов для фильтрации (например, ["выписка", "карточка", ".xlsx"])
            save_dir: Директория для сохранения (по умолчанию текущая)
        
        Returns:
            Список путей к скачанным файлам
        """
        if not self.mail:
            return []
        
        if save_dir is None:
            save_dir = Path(".")
        else:
            save_dir = Path(save_dir)
            save_dir.mkdir(parents=True, exist_ok=True)
        
        downloaded_files = []
        
        try:
            status, msg_data = self.mail.fetch(str(email_id), "(RFC822)")
            if status != "OK":
                return []
            
            email_body = msg_data[0][1]
            email_message = email.message_from_bytes(email_body)
            
            # Проверяем вложения
            for part in email_message.walk():
                if part.get_content_disposition() == "attachment":
                    filename = part.get_filename()
                    if filename:
                        filename = self.decode_mime_words(filename)
                        
                        # Фильтрация по паттернам
                        if filename_patterns:
                            if not any(pattern.lower() in filename.lower() 
                                      for pattern in filename_patterns):
                                continue
                        
                        # Сохранение файла
                        filepath = save_dir / filename
                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        
                        downloaded_files.append(filepath)
                        print(f"[OK] Скачан файл: {filename}")
        
        except Exception as e:
            print(f"Ошибка при скачивании вложений из письма {email_id}: {e}")
        
        return downloaded_files
    
    def fetch_bank_statement(self, save_dir: Path = None, 
                            subject_keywords: List[str] = None,
                            sender_keywords: List[str] = None,
                            days_back: int = 7) -> Optional[Path]:
        """
        Основной метод для получения выписки из банка.
        
        Args:
            save_dir: Директория для сохранения
            subject_keywords: Ключевые слова в теме (по умолчанию ["выписка", "карточка", "счет"])
            sender_keywords: Ключевые слова в отправителе (например, ["bank.ru", "bank.com"])
            days_back: За сколько дней искать
        
        Returns:
            Путь к скачанному файлу или None
        """
        if subject_keywords is None:
            subject_keywords = ["выписка", "карточка", "счет", "51"]
        
        if not self.connect():
            return None
        
        try:
            # Поиск писем
            email_ids = self.search_emails(
                subject_keywords=subject_keywords,
                sender_keywords=sender_keywords,
                days_back=days_back
            )
            
            if not email_ids:
                print("Не найдено писем с выпиской за указанный период")
                return None
            
            # Сортируем по дате (новые первыми)
            email_ids.sort(reverse=True)
            
            # Ищем вложения в последних письмах
            filename_patterns = ["выписка", "карточка", "счет", "51", ".pdf", ".xlsx", ".xls"]
            
            for email_id in email_ids:
                files = self.download_attachments(
                    email_id,
                    filename_patterns=filename_patterns,
                    save_dir=save_dir
                )
                
                # Ищем файлы, похожие на выписку (PDF или Excel)
                for filepath in files:
                    filename_lower = filepath.name.lower()
                    if any(keyword in filename_lower for keyword in ["выписка", "карточка", "счет"]):
                        if filepath.suffix in [".pdf", ".xlsx", ".xls"]:
                            print(f"[OK] Найдена выписка: {filepath.name}")
                            return filepath
            
            # Если не нашли по ключевым словам, возвращаем первый PDF или Excel файл
            if files:
                return files[0]
            
        finally:
            self.disconnect()
        
        return None


def fetch_from_config(config_path: Path = Path("email_config.txt")) -> Optional[Path]:
    """
    Упрощенная функция для загрузки выписки с использованием конфигурационного файла.
    
    Формат email_config.txt:
        email_type=mail.ru
        email_address=your_email@mail.ru
        password=your_password
        subject_keywords=выписка,карточка,счет
        sender_keywords=bank.ru
        days_back=7
    """
    if not config_path.exists():
        print(f"Файл конфигурации не найден: {config_path}")
        print("Создайте файл email_config.txt с настройками почты")
        return None
    
    config = {}
    with open(config_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                key, value = line.split("=", 1)
                config[key.strip()] = value.strip()
    
    # Парсинг списков
    subject_keywords = [kw.strip() for kw in config.get("subject_keywords", "выписка,карточка,счет").split(",")]
    sender_keywords = [kw.strip() for kw in config.get("sender_keywords", "").split(",")] if config.get("sender_keywords") else None
    days_back = int(config.get("days_back", "7"))
    
    fetcher = EmailFetcher(
        email_type=config.get("email_type", "gmail"),
        email_address=config.get("email_address", ""),
        password=config.get("password", ""),
        imap_server=config.get("imap_server", ""),
        imap_port=int(config.get("imap_port", "993"))
    )
    
    return fetcher.fetch_bank_statement(
        subject_keywords=subject_keywords,
        sender_keywords=sender_keywords,
        days_back=days_back
    )
