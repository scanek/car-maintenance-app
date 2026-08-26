FROM python:3.11-slim

WORKDIR /app

# Prevent python from buffering stdout/stderr
ENV PYTHONUNBUFFERED=1

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080

CMD ["python", "app.py"]
