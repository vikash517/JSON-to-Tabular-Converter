#!/usr/bin/env python3
"""
Demo script to showcase Excel export functionality
Converts all sample JSON files to Excel format
"""

import os
import sys
# Add utils directory to path
utils_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'utils')
sys.path.append(utils_dir)
from utils.json_to_excel import json_to_excel

def demo_excel_exports():
    """Convert all sample JSON files to Excel for demonstration"""
    
    sample_dir = "sample_data"
    output_dir = "excel_outputs"
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Find all JSON files in sample_data
    json_files = []
    if os.path.exists(sample_dir):
        json_files = [f for f in os.listdir(sample_dir) if f.endswith('.json')]
    
    if not json_files:
        print("❌ No JSON files found in sample_data directory")
        return False
    
    print("🔄 JSON to Excel Demo")
    print("=" * 50)
    print(f"Found {len(json_files)} JSON files to convert\n")
    
    results = []
    
    for json_file in json_files:
        input_path = os.path.join(sample_dir, json_file)
        output_name = json_file.replace('.json', '.xlsx')
        output_path = os.path.join(output_dir, output_name)
        
        print(f"🔄 Converting: {json_file}")
        
        # Test with different separators for variety
        separator = "_"
        if "nested" in json_file:
            separator = "."
        elif "employee" in json_file:
            separator = "-"
        
        success = json_to_excel(input_path, output_path, separator)
        
        if success:
            print(f"✅ Created: {output_path}")
            results.append({"file": json_file, "output": output_path, "status": "success"})
        else:
            print(f"❌ Failed: {json_file}")
            results.append({"file": json_file, "output": output_path, "status": "failed"})
        
        print("")
    
    # Summary
    print("📊 Conversion Summary")
    print("=" * 50)
    successful = [r for r in results if r["status"] == "success"]
    failed = [r for r in results if r["status"] == "failed"]
    
    print(f"✅ Successful: {len(successful)}")
    print(f"❌ Failed: {len(failed)}")
    print(f"📁 Output directory: {output_dir}")
    
    if successful:
        print("\n🎉 Successfully created Excel files:")
        for result in successful:
            print(f"   • {result['output']}")
    
    if failed:
        print(f"\n⚠️  Failed conversions:")
        for result in failed:
            print(f"   • {result['file']}")
    
    print(f"\n💡 Open the '{output_dir}' folder to view the Excel files!")
    
    return len(failed) == 0

if __name__ == "__main__":
    success = demo_excel_exports()
    sys.exit(0 if success else 1)
