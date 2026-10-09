#!/bin/bash
# Fast Startup script for Azure App Service
export PORT="${PORT:-8000}"
echo "Starting DMB Performance Dashboard on port $PORT..."
exec gunicorn app:server \
    --workers 2 \
    --threads 4 \
    --worker-class gthread \
    --timeout 120 \
    --access-logfile - \
    --error-logfile - \
    --bind "0.0.0.0:${PORT}"
