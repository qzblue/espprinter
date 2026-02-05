FROM python:3.9-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
# If requirements.txt doesn't exist, we install manually (fallback)
RUN pip install --no-cache-dir flask waitress pymysql requests openpyxl flask-caching

# Copy app code
COPY . .

# Environment variables (Defaults, override in K8s)
ENV SHARP_PRINTERS="http://10.64.48.120,http://10.96.48.109,http://10.32.48.155"
ENV DB_HOST="10.32.65.22"
ENV DB_USER="printer"
ENV DB_PASS="HDtAHFahLsdkNazm"
ENV DB_NAME="printer"

# Expose port
EXPOSE 5000

# Run with Waitress
CMD ["python", "webapp.py"]
