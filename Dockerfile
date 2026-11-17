FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV SOH_DATA=/data SOH_BEHIND_PROXY=1 SOH_SECURE_COOKIE=1 TZ=Europe/Rome PORT=8000
VOLUME /data
EXPOSE 8000
# 1 worker + più thread: il controllo SLA gira in background nel processo; SQLite in modalità WAL
CMD gunicorn -w 1 --threads 8 -b 0.0.0.0:${PORT} --access-logfile - app:app
