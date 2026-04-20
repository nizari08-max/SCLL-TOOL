FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads outputs

ENV PORT=7860

EXPOSE 7860

CMD gunicorn app:app --workers 2 --threads 4 --timeout 120 --bind 0.0.0.0:7860
