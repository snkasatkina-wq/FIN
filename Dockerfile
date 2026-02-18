FROM python:3.11-slim

WORKDIR /app

# Зависимости приложения
COPY requirements-docker.txt .
RUN pip install --no-cache-dir -r requirements-docker.txt

# Код приложения
COPY app.py .
COPY unmarked_alert.py .

# Папка для данных (монтируется как volume)
RUN mkdir -p /data

ENV DATA_DIR=/data
ENV PYTHONUNBUFFERED=1

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
