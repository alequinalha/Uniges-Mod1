FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV DEBUG=True

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libpq-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN python manage.py collectstatic --noinput || true

EXPOSE 8000
CMD ["sh", "-c", "python patch_db.py && python manage.py migrate --run-syncdb --noinput && python manage.py collectstatic --noinput && gunicorn core.wsgi:application --bind 0.0.0.0:8000 --workers 3"]
