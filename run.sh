#!/bin/bash
echo "======================================="
echo " 🏢 Real Estate Intelligence Map"
echo "======================================="

cd backend

# Install dependencies
echo "📦 Installing Python dependencies..."
pip install -r requirements.txt --quiet

# Create uploads folder
mkdir -p uploads

echo ""
echo "📁 Put your PDF/DOCX files in: backend/uploads/"
echo ""
echo "🚀 Starting server on http://localhost:5000"
echo "🌐 Open frontend/index.html in your browser"
echo ""

python app.py