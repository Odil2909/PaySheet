FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PAYROLL_DB_PATH=/data/payroll.sqlite3

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY database.py payroll_bot.py storage.py ./

CMD ["python", "payroll_bot.py"]
