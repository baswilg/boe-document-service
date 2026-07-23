FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py boe_document.py markdown_parser.py ./
COPY templates/ ./templates/

# Hosting-platforms geven de poort mee via $PORT; val terug op 8080.
ENV PORT=8080
EXPOSE 8080

CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT}"]
