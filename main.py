#!/usr/bin/env python3
"""
JSON to Tabular Converter - Main Launcher
Launch the GUI application from any directory
"""

import sys
import os

# Add the src directory to Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.join(current_dir, 'src')
sys.path.insert(0, src_dir)

# Import and run the main application
from src.json_converter import JSONToTabularConverter
import tkinter as tk

if __name__ == "__main__":
    print("🔄 Starting JSON to Tabular Converter...")
    root = tk.Tk()
    app = JSONToTabularConverter(root)
    print("✅ Application initialized. Opening GUI window...")
    root.mainloop()
    print("👋 Application closed.") 
    print("🎉 JSON to Tabular Converter is ready to use!")