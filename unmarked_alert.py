"""
Модуль для проверки нерозмеченных по статьям операций в реестре
и формирования сообщения для Telegram через OpenAI (финансовый коуч).

Логика:
1. Находит строки, где Ст1 не заполнена
2. Разбивает на блоки: поступления (Сумма П) и списания (Сумма С)
3. Подсчитывает количество и сумму по каждому блоку
4. Передаёт данные в OpenAI для формирования сообщения в формате Telegram
5. Отправляет сообщение в Telegram через бота
"""

import json
import os
import urllib.request
from pathlib import Path
from typing import Optional

import pandas as pd
from openai import OpenAI

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DATA_DIR", str(SCRIPT_DIR)))
TARGET_FILE = DATA_DIR / "УУ2025 - оптим.xlsx"
TARGET_SHEET = "реестр"

# Индексы колонок листа «реестр»
COL_SUMMA_P = 8   # Сумма П (поступления)
COL_SUMMA_S = 11  # Сумма С (списания)
COL_ST1 = 15      # Ст1 (статья)


def _is_empty(val) -> bool:
    """Проверяет, пусто ли значение."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return True
    s = str(val).strip()
    return not s


def _to_float(val) -> float:
    """Преобразует значение в число."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def load_unmarked_rows() -> tuple[dict, dict, float, float]:
    """
    Загружает реестр и возвращает два блока нерозмеченных строк
    и общие суммы по всему файлу.

    Returns:
        (receipts, withdrawals, total_receipts_all, total_withdrawals_all)
    """
    if not TARGET_FILE.exists():
        raise FileNotFoundError(f"Файл не найден: {TARGET_FILE}")

    df = pd.read_excel(TARGET_FILE, sheet_name=TARGET_SHEET, header=None)

    # Сумма П и Сумма С как числа
    summa_p = pd.to_numeric(df.iloc[:, COL_SUMMA_P], errors="coerce").fillna(0)
    summa_s = pd.to_numeric(df.iloc[:, COL_SUMMA_S], errors="coerce").fillna(0)
    st1 = df.iloc[:, COL_ST1]

    # Общие суммы по всему файлу (все строки с данными)
    total_receipts_all = float(summa_p.sum())
    total_withdrawals_all = float(summa_s.sum())

    # Строки, где Ст1 пуста
    st1_empty = st1.apply(lambda v: _is_empty(v))

    receipts_rows = []
    withdrawals_rows = []

    for idx in df[st1_empty].index:
        row = df.iloc[idx]
        sp = _to_float(row.iloc[COL_SUMMA_P])
        ss = _to_float(row.iloc[COL_SUMMA_S])

        row_data = {
            "date": row.iloc[3],
            "doc": row.iloc[4],
            "purpose": row.iloc[13] if len(row) > 13 else "",
            "summa_p": sp,
            "summa_s": ss,
        }

        if sp > 0:
            receipts_rows.append(row_data)
        if ss > 0:
            withdrawals_rows.append(row_data)

    receipts = {
        "count": len(receipts_rows),
        "total_sum": sum(r["summa_p"] for r in receipts_rows),
        "rows": receipts_rows,
    }
    withdrawals = {
        "count": len(withdrawals_rows),
        "total_sum": sum(r["summa_s"] for r in withdrawals_rows),
        "rows": withdrawals_rows,
    }

    return receipts, withdrawals, total_receipts_all, total_withdrawals_all


def generate_telegram_message(
    receipts: dict,
    withdrawals: dict,
    total_receipts_all: float,
    total_withdrawals_all: float,
    api_key: Optional[str] = None,
) -> str:
    """
    Формирует сообщение для Telegram через OpenAI (финансовый консультант).

    Args:
        receipts: блок нерозмеченных поступлений (count, total_sum, rows)
        withdrawals: блок нерозмеченных списаний (count, total_sum, rows)
        total_receipts_all: общая сумма поступлений по всему файлу
        total_withdrawals_all: общая сумма списаний по всему файлу
        api_key: ключ OpenAI
    """
    client = OpenAI(api_key=api_key)

    pct_receipts = (receipts["total_sum"] / total_receipts_all * 100) if total_receipts_all > 0 else 0
    pct_withdrawals = (withdrawals["total_sum"] / total_withdrawals_all * 100) if total_withdrawals_all > 0 else 0

    prompt = f"""Ты — финансовый консультант. Сформируй сообщение для Telegram.

Данные по файлу управленческого учёта:

ОБЩИЕ СУММЫ ПО ВСЕМУ ФАЙЛУ:
- Всего поступлений: {total_receipts_all:,.2f} руб.
- Всего списаний: {total_withdrawals_all:,.2f} руб.

НЕРАЗМЕЧЕННЫЕ ПО СТАТЬЯМ ОПЕРАЦИИ:
- Списания: {withdrawals['count']} операций на {withdrawals['total_sum']:,.2f} руб. ({pct_withdrawals:.1f}% от общих списаний)
- Поступления: {receipts['count']} операций на {receipts['total_sum']:,.2f} руб. ({pct_receipts:.1f}% от общих поступлений)

Задачи:
1. Напиши краткое сообщение для Telegram (без markdown, без хештегов), которое начинается со слова «Внимание» и информирует о нерозмеченных доходах и расходах с указанием количества и сумм.
2. Дай оценку как финансовый консультант: насколько сильно то, что эти строки не заполнены, искажает общую управленческую отчётность. Укажи долю в процентах и краткий вывод (критично / существенно / незначительно).

Тон — деловой, лаконичный. Объём — до 400 слов."""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "Ты финансовый консультант. Даёшь оценки и пишешь деловые сообщения для Telegram."},
            {"role": "user", "content": prompt},
        ],
        max_tokens=500,
    )

    return response.choices[0].message.content.strip()


def get_telegram_chat_id(bot_token: str) -> Optional[str]:
    """
    Получает chat_id из последних обновлений бота.
    Напишите боту /start, затем вызовите эту функцию.
    """
    url = f"https://api.telegram.org/bot{bot_token}/getUpdates"
    try:
        with urllib.request.urlopen(url, timeout=10) as resp:
            data = json.loads(resp.read().decode())
            results = data.get("result", [])
            if results:
                last = results[-1]
                chat = last.get("message", {}).get("chat", {})
                return str(chat.get("id", ""))
    except Exception:
        pass
    return None


def send_telegram_message(text: str, bot_token: str, chat_id: str) -> bool:
    """
    Отправляет сообщение в Telegram через Bot API.

    Args:
        text: текст сообщения
        bot_token: токен бота (из @BotFather)
        chat_id: ID чата или канала для отправки

    Returns:
        True при успехе, False при ошибке
    """
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    data = json.dumps({"chat_id": chat_id, "text": text}).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read().decode())
            return result.get("ok", False)
    except Exception as e:
        print(f"Ошибка отправки в Telegram: {e}")
        return False


def run_and_send(
    api_key: Optional[str] = None,
    bot_token: Optional[str] = None,
    chat_id: Optional[str] = None,
) -> dict:
    """
    Полный цикл: загрузка данных, формирование сообщения, отправка в Telegram.
    Возвращает dict с ключами: success, message, sent, error.
    """
    import os
    from dotenv import load_dotenv
    load_dotenv()

    api_key = api_key or os.environ.get("OPENAI_API_KEY")
    bot_token = bot_token or os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = chat_id or os.environ.get("TELEGRAM_CHAT_ID")

    if not chat_id and bot_token:
        chat_id = get_telegram_chat_id(bot_token)

    result = {"success": False, "message": None, "sent": False, "error": None}

    if not api_key:
        result["error"] = "OPENAI_API_KEY не задан"
        return result
    if not bot_token:
        result["error"] = "TELEGRAM_BOT_TOKEN не задан"
        return result
    if not chat_id:
        result["error"] = "TELEGRAM_CHAT_ID не задан"
        return result

    try:
        receipts, withdrawals, total_receipts_all, total_withdrawals_all = load_unmarked_rows()
    except FileNotFoundError as e:
        result["error"] = str(e)
        return result
    except Exception as e:
        result["error"] = str(e)
        return result

    if receipts["count"] == 0 and withdrawals["count"] == 0:
        result["success"] = True
        result["message"] = "Нет нерозмеченных операций"
        return result

    try:
        msg = generate_telegram_message(
            receipts, withdrawals,
            total_receipts_all, total_withdrawals_all,
            api_key=api_key,
        )
    except Exception as e:
        result["error"] = str(e)
        return result

    result["message"] = msg
    result["success"] = True

    if send_telegram_message(msg, bot_token, chat_id):
        result["sent"] = True
    else:
        result["error"] = "Не удалось отправить в Telegram"

    return result


def run(api_key: Optional[str] = None) -> Optional[str]:
    """
    Основная функция: загружает данные, формирует сообщение через OpenAI.

    Returns:
        Текст сообщения для Telegram или None, если нет нерозмеченных операций
    """
    receipts, withdrawals, total_receipts_all, total_withdrawals_all = load_unmarked_rows()

    if receipts["count"] == 0 and withdrawals["count"] == 0:
        return None

    return generate_telegram_message(
        receipts, withdrawals,
        total_receipts_all, total_withdrawals_all,
        api_key=api_key,
    )


if __name__ == "__main__":
    import os
    from dotenv import load_dotenv

    load_dotenv()
    api_key = os.environ.get("OPENAI_API_KEY")
    bot_token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")

    if not api_key:
        print("Ошибка: задайте OPENAI_API_KEY в .env")
        exit(1)
    if not bot_token:
        print("Ошибка: задайте TELEGRAM_BOT_TOKEN в .env")
        exit(1)
    if not chat_id:
        print("TELEGRAM_CHAT_ID не задан. Пытаюсь получить из getUpdates...")
        chat_id = get_telegram_chat_id(bot_token)
        if chat_id:
            print(f"Найден chat_id: {chat_id}. Добавьте TELEGRAM_CHAT_ID={chat_id} в .env для следующих запусков.")
        else:
            print("Напишите боту /start в Telegram, затем запустите скрипт снова.")
            exit(1)

    print("Загрузка данных...")
    receipts, withdrawals, total_receipts_all, total_withdrawals_all = load_unmarked_rows()

    print(f"По всему файлу: поступления {total_receipts_all:,.2f} руб., списания {total_withdrawals_all:,.2f} руб.")
    print(f"Без статьи — списания: {withdrawals['count']} операций, {withdrawals['total_sum']:,.2f} руб.")
    print(f"Без статьи — поступления: {receipts['count']} операций, {receipts['total_sum']:,.2f} руб.")

    if receipts["count"] == 0 and withdrawals["count"] == 0:
        print("Нет нерозмеченных операций.")
        exit(0)

    print("\nФормирование сообщения через OpenAI...")
    msg = generate_telegram_message(
        receipts, withdrawals,
        total_receipts_all, total_withdrawals_all,
        api_key=api_key,
    )
    print("\n--- Сообщение ---")
    print(msg)
    print("---")

    print("\nОтправка в Telegram...")
    if send_telegram_message(msg, bot_token, chat_id):
        print("Сообщение отправлено.")
    else:
        print("Не удалось отправить сообщение.")
