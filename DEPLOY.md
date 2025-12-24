# Деплой на VPS сервер

## Требования

- VPS с Ubuntu 20.04+ или Debian 11+
- Минимум 1GB RAM, 10GB диск
- Docker и Docker Compose
- Домен (опционально)

---

## Быстрый старт

### 1. Подключитесь к VPS

```bash
ssh root@your-server-ip
```

### 2. Установите Docker

```bash
# Обновление системы
apt update && apt upgrade -y

# Установка Docker
curl -fsSL https://get.docker.com | sh

# Установка Docker Compose
apt install docker-compose-plugin -y
```

### 3. Клонируйте проект

```bash
cd /opt
git clone https://github.com/YOUR_USERNAME/certificates-generator.git
cd certificates-generator
```

Или загрузите файлы вручную через SFTP.

### 4. Запустите приложение

```bash
docker compose up -d
```

### 5. Проверьте работу

Откройте в браузере: `http://your-server-ip:8000`

Логин: **Kirito**
Пароль: **Kirito**

---

## Настройка домена и SSL (опционально)

### 1. Укажите домен на IP сервера

В DNS настройках вашего домена создайте A-запись:
```
certificates.yourdomain.com -> your-server-ip
```

### 2. Установите Certbot

```bash
apt install certbot -y
```

### 3. Получите SSL сертификат

```bash
certbot certonly --standalone -d certificates.yourdomain.com
```

### 4. Создайте nginx.conf

```bash
cat > nginx.conf << 'EOF'
events {
    worker_connections 1024;
}

http {
    server {
        listen 80;
        server_name certificates.yourdomain.com;
        return 301 https://$server_name$request_uri;
    }

    server {
        listen 443 ssl;
        server_name certificates.yourdomain.com;

        ssl_certificate /etc/nginx/ssl/fullchain.pem;
        ssl_certificate_key /etc/nginx/ssl/privkey.pem;

        location / {
            proxy_pass http://webapp:8000;
            proxy_set_header Host $host;
            proxy_set_header X-Real-IP $remote_addr;
            proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
            proxy_set_header X-Forwarded-Proto $scheme;
        }
    }
}
EOF
```

### 5. Скопируйте сертификаты

```bash
mkdir -p ssl
cp /etc/letsencrypt/live/certificates.yourdomain.com/fullchain.pem ssl/
cp /etc/letsencrypt/live/certificates.yourdomain.com/privkey.pem ssl/
```

### 6. Запустите с nginx

```bash
docker compose --profile production up -d
```

---

## Изменение логина/пароля

Отредактируйте файл `webapp/app.py`:

```python
USERS = {
    "Kirito": "Kirito",           # Текущий пользователь
    "admin": "your_password",     # Добавьте нового
}
```

Перезапустите:
```bash
docker compose restart
```

---

## Полезные команды

```bash
# Просмотр логов
docker compose logs -f

# Перезапуск
docker compose restart

# Остановка
docker compose down

# Обновление кода
git pull
docker compose build --no-cache
docker compose up -d

# Очистка старых файлов (справок и загрузок)
rm -rf webapp/uploads/* webapp/generated/*
```

---

## Резервное копирование

```bash
# Создание бэкапа
tar -czf backup-$(date +%Y%m%d).tar.gz webapp/uploads webapp/generated

# Восстановление
tar -xzf backup-YYYYMMDD.tar.gz
```

---

## Troubleshooting

### Порт 8000 занят
```bash
# Найти процесс
lsof -i :8000
# Убить процесс
kill -9 PID
```

### Ошибка авторизации
Браузер кеширует учётные данные. Очистите кеш или откройте в режиме инкогнито.

### Не генерируются PDF
Проверьте логи:
```bash
docker compose logs webapp | grep -i error
```

---

## Архитектура

```
┌─────────────────────────────────────────┐
│               Internet                  │
└─────────────────┬───────────────────────┘
                  │
┌─────────────────▼───────────────────────┐
│            Nginx (443/80)               │
│         SSL termination                 │
└─────────────────┬───────────────────────┘
                  │
┌─────────────────▼───────────────────────┐
│         FastAPI App (8000)              │
│   - Авторизация                         │
│   - Загрузка Excel                      │
│   - Генерация справок                   │
│   - Предпросмотр                        │
│   - История                             │
└─────────────────────────────────────────┘
```

---

## Контакты

SwissCapital — Республика Казахстан
