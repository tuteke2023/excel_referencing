FROM python:3.9-slim

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY tb_gl_linker.py .
COPY quick_link.py .
COPY web_app.py .

# Create directories for file processing
RUN mkdir -p /app/uploads /app/outputs

# Expose port for web app
EXPOSE 8501

# Set environment variables
ENV PYTHONPATH=/app

# Default command (can be overridden)
CMD ["python", "tb_gl_linker.py", "--help"]