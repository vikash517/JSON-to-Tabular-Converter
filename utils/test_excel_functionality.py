#!/usr/bin/env python3
"""
Test script for Excel export functionality
"""

import os
import sys
import json
import pandas as pd

# Add parent directory to path to import json_to_excel
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from json_to_excel import json_to_excel

def test_excel_exports():
    """Test Excel export with all sample files"""
    sample_dir = "sample_data"
    output_dir = "test_exports"
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    print("🧪 Testing Excel Export Functionality")
    print("=" * 50)
    
    # Get all JSON files in sample_data
    json_files = [f for f in os.listdir(sample_dir) if f.endswith('.json')]
    
    successful = 0
    failed = 0
    
    for json_file in json_files:
        input_path = os.path.join(sample_dir, json_file)
        output_path = os.path.join(output_dir, f"{os.path.splitext(json_file)[0]}_test.xlsx")
        
        print(f"\n📁 Testing: {json_file}")
        print(f"   Input:  {input_path}")
        print(f"   Output: {output_path}")
        
        try:
            success = json_to_excel(input_path, output_path)
            if success:
                file_size = os.path.getsize(output_path)
                print(f"   ✅ Success! File size: {file_size:,} bytes")
                successful += 1
            else:
                print(f"   ❌ Failed!")
                failed += 1
        except Exception as e:
            print(f"   ❌ Error: {str(e)}")
            failed += 1
    
    print("\n" + "=" * 50)
    print(f"📊 Test Results:")
    print(f"   ✅ Successful: {successful}")
    print(f"   ❌ Failed: {failed}")
    print(f"   📁 Output directory: {output_dir}")
    
    if successful > 0:
        print("\n🎉 Excel export functionality is working!")
        print(f"Check the '{output_dir}' directory for generated Excel files.")
    else:
        print("\n⚠️  All tests failed. Check your setup.")
    
    return successful, failed

def validate_excel_files():
    """Validate that generated Excel files can be read back"""
    output_dir = "test_exports"
    
    if not os.path.exists(output_dir):
        print("❌ No test exports found. Run test_excel_exports() first.")
        return
    
    excel_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx')]
    
    print(f"\n🔍 Validating {len(excel_files)} Excel files...")
    
    for excel_file in excel_files:
        file_path = os.path.join(output_dir, excel_file)
        try:
            # Try to read the Excel file
            sheets = pd.read_excel(file_path, sheet_name=None)
            print(f"   ✅ {excel_file}: {len(sheets)} sheets")
            
            for sheet_name, df in sheets.items():
                print(f"      📊 {sheet_name}: {df.shape[0]} rows × {df.shape[1]} columns")
        
        except Exception as e:
            print(f"   ❌ {excel_file}: {str(e)}")

if __name__ == "__main__":
    # Run tests
    successful, failed = test_excel_exports()
    
    # Validate if any were successful
    if successful > 0:
        validate_excel_files()
    
    sys.exit(0 if failed == 0 else 1)
