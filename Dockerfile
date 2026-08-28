FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libpq-dev \
    && rm -rf /var/lib/apt/lists/^*

RUN pip install --no-cache-dir \
    "Django>=5.0,<6.0" \
    "gunicorn>=21.2.0" \
    "whitenoise>=6.6.0" \
    "psycopg2-binary>=2.9.9" \
    "dj-database-url>=2.1.0" \
    "pillow>=10.2.0" \
    "django-environ>=0.11.2" \
    "openpyxl>=3.1.2"

COPY . .
