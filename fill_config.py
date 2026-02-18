"""
Вспомогательный скрипт для заполнения email_config.txt
"""

from pathlib import Path

def main():
    config_path = Path("email_config.txt")
    
    print("=" * 80)
    print("Заполнение конфигурации email_config.txt")
    print("=" * 80)
    print()
    
    # Читаем текущий файл
    lines = []
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
    else:
        print("Файл email_config.txt не найден!")
        return
    
    # Запрашиваем данные
    print("Введите данные для подключения к mail.ru:")
    print()
    email = input("Email адрес (например, ivanov@mail.ru): ").strip()
    password = input("Пароль: ").strip()
    
    if not email or not password:
        print("Ошибка: Email и пароль не могут быть пустыми!")
        return
    
    # Обновляем строки
    new_lines = []
    for line in lines:
        if line.strip().startswith("email_address="):
            new_lines.append(f"email_address={email}\n")
        elif line.strip().startswith("password="):
            new_lines.append(f"password={password}\n")
        else:
            new_lines.append(line)
    
    # Сохраняем
    with open(config_path, "w", encoding="utf-8") as f:
        f.writelines(new_lines)
    
    print()
    print(f"[OK] Конфигурация сохранена в {config_path.absolute()}")
    print(f"Email: {email}")
    print(f"Пароль: {'*' * len(password)}")
    print()
    print("Теперь можно запустить: python step1_fetch_email.py")

if __name__ == "__main__":
    main()
