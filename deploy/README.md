# Развёртывание на сервере

Эталонные конфигурации для production на Ubuntu 22.04 (SberCloud).

## Схема

```text
Браузер → Caddy (:443, HTTPS) → Gunicorn (127.0.0.1:8000) → Django → SQLite
                ↓
         /static/, /media/
```

- **Домен:** https://bykovogazzz.ru
- **Каталог проекта:** `/opt/BYKOVOGAZ`
- **Публичный IP:** `84.54.59.117`

## Файлы

| Файл | Назначение |
|------|------------|
| `Caddyfile` | Домен, static, media, reverse_proxy |
| `gunicorn.service.example` | Unit systemd для Gunicorn |

## Быстрая установка Caddyfile на сервере

```bash
sudo mkdir -p /etc/caddy
sudo cp /opt/BYKOVOGAZ/deploy/Caddyfile /etc/caddy/Caddyfile
# или скопируйте содержимое вручную, если папки deploy/ ещё нет на сервере
sudo caddy validate --config /etc/caddy/Caddyfile
sudo systemctl restart caddy
```

При `apt install caddy` на вопрос про Caddyfile отвечайте **N** (оставить свою версию).

## Gunicorn

```bash
sudo cp deploy/gunicorn.service.example /etc/systemd/system/gunicorn.service
sudo nano /etc/systemd/system/gunicorn.service   # DJANGO_SECRET_KEY из .env
sudo systemctl daemon-reload
sudo systemctl enable gunicorn && sudo systemctl restart gunicorn
```

## Переменные окружения

См. `.env.example` в корне проекта. Ключ из `.env` должен совпадать с `Environment=` в `gunicorn.service`.

## Полная инструкция с командами

`docs/Развертывание_команды_терминала.docx` — генерируется скриптом:

```powershell
pip install python-docx
python scripts/build_deploy_commands_docx.py
```
