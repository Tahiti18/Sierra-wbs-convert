#!/bin/bash

# Railway deployment start script for Flask Sierra-WBS converter
echo "🚀 Starting Sierra-WBS Flask Server..."
echo "📍 Working directory: $(pwd)"
echo "🐍 Python version: $(python --version)"
echo "📦 Flask version: $(python -c 'import flask; print(flask.__version__)')"
echo "🌐 PORT environment: ${PORT:-'Not set'}"

# Navigate to app directory if needed
if [ ! -f "src/main.py" ]; then
    echo "❌ main.py not found in src/, checking current directory..."
    ls -la
    exit 1
fi

echo "✅ Starting Flask server..."
exec python src/main.py