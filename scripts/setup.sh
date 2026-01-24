#!/bin/bash

# PDFCon Setup Script
# This script helps users set up the development environment

set -e

echo "🚀 PDFCon - Government Financial Report Analysis System"
echo "======================================================"
echo ""

# Check if Docker is installed
if ! command -v docker &> /dev/null; then
    echo "❌ Docker is not installed. Please install Docker first."
    exit 1
fi

# Check if Docker Compose is installed
if ! command -v docker-compose &> /dev/null; then
    echo "❌ Docker Compose is not installed. Please install Docker Compose first."
    exit 1
fi

echo "✅ Docker and Docker Compose are installed"
echo ""

# Create .env file if it doesn't exist
if [ ! -f .env ]; then
    echo "📝 Creating .env file from .env.example..."
    cp .env.example .env
    echo "✅ .env file created"
    echo ""
    echo "⚠️  IMPORTANT: Please edit .env file and add your OPENAI_API_KEY"
    echo "   You can get one at: https://platform.openai.com/api-keys"
    echo ""
else
    echo "✅ .env file already exists"
fi

# Create necessary directories
echo "📁 Creating necessary directories..."
mkdir -p backend/uploads
mkdir -p backend/logs
mkdir -p data/norway_fund
mkdir -p data/benchmarks
mkdir -p data/templates
echo "✅ Directories created"
echo ""

# Build and start services
echo "🔨 Building Docker images..."
docker-compose build
echo ""

echo "🚀 Starting services..."
docker-compose up -d
echo ""

echo "⏳ Waiting for services to be ready..."
sleep 10

echo "✅ Setup complete!"
echo ""
echo "📊 Access Points:"
echo "   Frontend:        http://localhost:3000"
echo "   Backend API:     http://localhost:8000"
echo "   API Docs:        http://localhost:8000/api/docs"
echo "   Celery Flower:   http://localhost:5555"
echo "   MinIO Console:   http://localhost:9001"
echo ""
echo "📚 Next steps:"
echo "   1. Edit .env and add your OPENAI_API_KEY"
echo "   2. Visit http://localhost:3000 to use the application"
echo "   3. Check logs with: docker-compose logs -f"
echo ""
echo "🛑 To stop services: docker-compose down"
echo "🧹 To clean everything: docker-compose down -v"