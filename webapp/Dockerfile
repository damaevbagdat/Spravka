FROM python:3.11-slim

WORKDIR /app

# Установка зависимостей системы и шрифтов для PDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-dejavu \
    fonts-dejavu-core \
    fonts-dejavu-extra \
    && rm -rf /var/lib/apt/lists/*

# Копирование зависимостей
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копирование приложения
COPY . .

# num2text.py будет смонтирован через volume

# Создание директорий
RUN mkdir -p uploads generated

# Порт
EXPOSE 8000

# Запуск
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
