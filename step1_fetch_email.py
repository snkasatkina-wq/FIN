"""
Шаг 1: Загрузка последнего письма от vtb-inform из mail.ru
и сохранение вложений в каталог "выписки ВТБ"
"""

from pathlib import Path
from email_fetcher import EmailFetcher


def main():
    """Загрузка последнего письма от vtb-inform."""
    print("=" * 80)
    print("ШАГ 1: Загрузка письма от vtb-inform из mail.ru")
    print("=" * 80)
    print()
    
    # Создаем каталог для выписок
    output_dir = Path("выписки ВТБ")
    output_dir.mkdir(exist_ok=True)
    print(f"Каталог для выписок: {output_dir.absolute()}")
    print()
    
    # Читаем конфигурацию
    config_path = Path("email_config.txt")
    if not config_path.exists():
        print("Файл конфигурации не найден. Введите данные вручную:")
        print()
        email_address = input("Email адрес (mail.ru): ").strip()
        password = input("Пароль: ").strip()
        
        fetcher = EmailFetcher(
            email_type="mail.ru",
            email_address=email_address,
            password=password
        )
    else:
        # Читаем конфигурацию из файла
        config = {}
        print(f"Чтение конфигурации из: {config_path.absolute()}")
        
        # Пробуем разные кодировки
        encodings = ['utf-8', 'utf-8-sig', 'cp1251', 'windows-1251']
        file_content = None
        
        for encoding in encodings:
            try:
                with open(config_path, "r", encoding=encoding) as f:
                    file_content = f.readlines()
                print(f"Файл прочитан в кодировке: {encoding}")
                break
            except Exception as e:
                continue
        
        if file_content is None:
            print("[ERROR] Не удалось прочитать файл конфигурации")
            return
        
        # Парсим конфигурацию
        for line_num, line in enumerate(file_content, 1):
            original_line = line
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip()
                config[key] = value
                # Отладочный вывод для важных полей
                if key in ['email_address', 'password']:
                    if key == 'password':
                        print(f"  Строка {line_num}: {key} = {'[ЗАПОЛНЕНО]' if value else '[ПУСТО]'} (длина: {len(value)})")
                    else:
                        print(f"  Строка {line_num}: {key} = {value if value else '[ПУСТО]'}")
        
        email_address = config.get("email_address", "").strip()
        password = config.get("password", "").strip()
        
        print()
        print(f"Загружена конфигурация:")
        print(f"  Email: {email_address if email_address else '[НЕ УКАЗАН]'}")
        print(f"  Пароль: {'[УКАЗАН]' if password else '[НЕ УКАЗАН]'} (длина: {len(password)})")
        print()
        
        if not email_address or not password:
            print("[ERROR] В конфигурации не указаны email_address или password")
            print("Проверьте файл email_config.txt и убедитесь, что:")
            print("  - email_address=ваш_email@mail.ru")
            print("  - password=ваш_пароль")
            print()
            print("Файл должен быть сохранен в кодировке UTF-8")
            return
        
        fetcher = EmailFetcher(
            email_type=config.get("email_type", "mail.ru"),
            email_address=email_address,
            password=password,
            imap_server=config.get("imap_server", ""),
            imap_port=int(config.get("imap_port", "993")) if config.get("imap_port") else 993
        )
    
    # Подключаемся к почте
    if not fetcher.connect():
        print("[ERROR] Не удалось подключиться к почте")
        return
    
    try:
        # Ищем письма от vtb-inform
        print("Поиск писем от vtb-inform...")
        # Пробуем разные варианты поиска отправителя
        email_ids = fetcher.search_emails(
            subject_keywords=None,  # Любая тема
            sender_keywords=["vtb-inform", "vtb-inform@", "@vtb.ru"],  # Разные варианты отправителя
            days_back=30  # За последние 30 дней
        )
        
        if not email_ids:
            print("✗ Не найдено писем от vtb-inform")
            print("Попробуйте проверить:")
            print("  - Правильность написания отправителя (vtb-inform)")
            print("  - Есть ли письма за последние 30 дней")
            return
        
        print(f"✓ Найдено писем: {len(email_ids)}")
        
        # Берем последнее письмо (самое новое)
        email_ids.sort(reverse=True)  # Сортируем по убыванию (новые первыми)
        latest_email_id = email_ids[0]
        
        # Получаем информацию о письме для отладки
        try:
            import email as email_lib
            if fetcher.mail:
                status, msg_data = fetcher.mail.fetch(str(latest_email_id), "(RFC822)")
                if status == "OK":
                    email_body = msg_data[0][1]
                    email_message = email_lib.message_from_bytes(email_body)
                    subject = fetcher.decode_mime_words(email_message["Subject"] or "")
                    from_addr = email_message["From"] or ""
                    date = email_message["Date"] or ""
                    print(f"\nПоследнее письмо:")
                    print(f"  От: {from_addr}")
                    print(f"  Тема: {subject}")
                    print(f"  Дата: {date}")
        except Exception as e:
            print(f"Не удалось получить информацию о письме: {e}")
        
        print(f"\nЗагрузка вложений из письма (ID: {latest_email_id})...")
        
        # Скачиваем все вложения из этого письма
        downloaded_files = fetcher.download_attachments(
            latest_email_id,
            filename_patterns=None,  # Все файлы
            save_dir=output_dir
        )
        
        if downloaded_files:
            print()
            print("[OK] Файлы успешно загружены:")
            for filepath in downloaded_files:
                print(f"  - {filepath.name}")
            print()
            print(f"Все файлы сохранены в: {output_dir.absolute()}")
        else:
            print("[ERROR] В письме не найдено вложений")
    
    finally:
        fetcher.disconnect()
    
    print()
    print("=" * 80)
    print("ШАГ 1 ЗАВЕРШЕН")
    print("=" * 80)


if __name__ == "__main__":
    main()
