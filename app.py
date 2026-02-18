"""
FIN — приложение проверки нерозмеченных операций.

- При запуске: выполняет проверку принудительно
- Ежедневно в 13:45 МСК: автоматический запуск
- HTTP API: GET /run — ручной запуск
"""

import logging
import os
from contextlib import asynccontextmanager

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from dotenv import load_dotenv
from fastapi import FastAPI

from unmarked_alert import run_and_send

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

scheduler = BackgroundScheduler()


def run_check():
    """Запуск проверки нерозмеченных операций."""
    logger.info("Запуск проверки нерозмеченных операций...")
    result = run_and_send()
    if result["error"]:
        logger.error("Ошибка: %s", result["error"])
    elif result["sent"]:
        logger.info("Сообщение отправлено в Telegram")
    elif result["message"] == "Нет нерозмеченных операций":
        logger.info("Нет нерозмеченных операций")
    else:
        logger.info("Проверка выполнена")
    return result


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Запуск при старте и принудительная проверка."""
    # Принудительный запуск при старте (с небольшой задержкой для монтирования volume)
    import asyncio
    await asyncio.sleep(2)
    try:
        run_check()
    except Exception as e:
        logger.exception("Ошибка при стартовой проверке: %s", e)

    # Планировщик: ежедневно 13:45 по Москве
    scheduler.add_job(
        run_check,
        CronTrigger(hour=13, minute=45, timezone="Europe/Moscow"),
    )
    scheduler.start()
    logger.info("Планировщик запущен: ежедневно в 13:45 МСК")

    yield

    scheduler.shutdown()


app = FastAPI(
    title="FIN — Проверка разметки",
    description="Проверка нерозмеченных операций в реестре, уведомления в Telegram",
    lifespan=lifespan,
)


@app.get("/")
def root():
    try:
        jobs = scheduler.get_jobs()
        next_run = jobs[0].next_run_time.isoformat() if jobs else None
    except Exception:
        next_run = None
    return {
        "service": "FIN",
        "status": "ok",
        "endpoints": ["/", "/health", "/run", "/status"],
        "next_run": next_run,
    }


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/status")
def status():
    """Статус планировщика и следующего запуска."""
    jobs = scheduler.get_jobs()
    next_run = jobs[0].next_run_time.isoformat() if jobs else None
    return {"scheduler": "running", "next_run": next_run}


@app.get("/run")
def run_manual():
    """Ручной запуск проверки."""
    result = run_check()
    return {
        "success": result["success"],
        "sent": result["sent"],
        "message": result.get("message", "")[:200] if result.get("message") else None,
        "error": result.get("error"),
    }
