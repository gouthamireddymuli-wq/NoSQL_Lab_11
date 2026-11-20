# ── NoSQL Lab 11: MongoDB Sharding ────────────────────────────────────────────
FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt ./requirements.txt
COPY web/requirements.txt ./web_requirements.txt
RUN pip install --no-cache-dir -r requirements.txt -r web_requirements.txt

# Copy source and documentation
COPY config/          ./config/
COPY lab_01_shard_key_eval.py .
COPY lab_02_shard_key_hotspot.py .
COPY create_docx_report.py .
COPY Lab_11_Report.docx .
COPY Lab_11_Report.pdf .
COPY web/             ./web/

ENV PYTHONUNBUFFERED=1 \
    FLASK_APP=web/app.py \
    MONGO_URI=mongodb://mongo:27017/

EXPOSE 5000

CMD ["python", "web/app.py"]
