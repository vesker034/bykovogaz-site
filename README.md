# Сайт «БыковоГаз»

Веб-сайт для компании «БыковоГаз» с информационными страницами и каталогом услуг.

**Стек:** Python, Django, HTML, CSS, Bootstrap 5, Docker.

## Локальный запуск

```bash
python -m venv .venv
python -m pip install -r requirements.txt
python manage.py migrate
python manage.py runserver
```

Переменные окружения задаются локально; секреты и база данных в репозиторий не входят.
