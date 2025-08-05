#!/bin/bash
# Setup script for JSON to Tabular Converter

echo "🔧 Setting up JSON to Tabular Converter..."

# Make scripts executable
chmod +x main.py
chmod +x utils/json_to_excel.py
chmod +x utils/test_excel_functionality.py
chmod +x examples/demo_excel.py

# Activate virtual environment and install dependencies
if [ -d "venv" ]; then
    echo "📦 Installing dependencies..."
    ./venv/bin/pip install -r requirements.txt
    echo "✅ Dependencies installed!"
else
    echo "❌ Virtual environment not found. Please create one first:"
    echo "python3 -m venv venv"
    echo "source venv/bin/activate"
    echo "pip install -r requirements.txt"
fi

echo ""
echo "🚀 Setup complete! To run the application:"
echo "   ./main.py                    # Run GUI application"
echo "   python utils/json_to_excel.py input.json output.xlsx  # Command line"
echo "   python examples/demo_excel.py    # Demo with all samples"
echo ""
