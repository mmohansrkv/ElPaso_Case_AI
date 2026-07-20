FROM python:3.11-slim

WORKDIR /app

# Install dependencies first (better layer caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the app
COPY . .

# Render sets $PORT at runtime; default to 8501 for local docker run
ENV PORT=8501
EXPOSE 8501

CMD streamlit run app.py --server.port $PORT --server.address 0.0.0.0
