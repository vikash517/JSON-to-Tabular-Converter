#!/usr/bin/env python3
"""
Command-line utility to convert JSON files directly to Excel format
Usage: python json_to_excel.py input.json output.xlsx
"""

import json
import pandas as pd
import sys
import os
from pandas import json_normalize
from datetime import datetime

def json_to_excel(input_file, output_file, separator="_", max_level=None):
    """
    Convert JSON file to Excel with enhanced formatting
    
    Args:
        input_file (str): Path to input JSON file
        output_file (str): Path to output Excel file
        separator (str): Separator for nested keys (default: "_")
        max_level (int): Maximum nesting level to flatten (default: None - all levels)
    """
    try:
        # Load JSON data
        print(f"Loading JSON file: {input_file}")
        with open(input_file, 'r', encoding='utf-8') as file:
            json_data = json.load(file)
        
        print(f"Converting JSON to tabular format...")
        
        # Convert JSON to DataFrame
        if isinstance(json_data, list):
            # Handle array of objects
            df = json_normalize(json_data, sep=separator, max_level=max_level)
        elif isinstance(json_data, dict):
            # Handle single object
            df = json_normalize([json_data], sep=separator, max_level=max_level)
        else:
            raise ValueError("JSON data must be an object or array of objects")
        
        # Convert any remaining list/dict columns to strings
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].apply(lambda x: str(x) if isinstance(x, (list, dict)) else x)
        
        print(f"Data shape: {df.shape[0]} rows, {df.shape[1]} columns")
        
        # Export to Excel with formatting
        print(f"Exporting to Excel: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Write main data to 'Data' sheet
            df.to_excel(writer, sheet_name='Data', index=False)
            
            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Data']
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Add header formatting
            try:
                from openpyxl.styles import Font, PatternFill, Alignment
                header_font = Font(bold=True, color="FFFFFF")
                header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                header_alignment = Alignment(horizontal="center", vertical="center")
                
                for cell in worksheet[1]:  # First row (headers)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_alignment
            except ImportError:
                print("Warning: openpyxl styling features not available. Basic export only.")
            
            # Create summary sheet
            summary_data = {
                'Metric': [
                    'Source File',
                    'Total Rows',
                    'Total Columns', 
                    'Missing Values',
                    'Complete Rows',
                    'Memory Usage (MB)',
                    'Conversion Date',
                    'Separator Used',
                    'Max Level Used'
                ],
                'Value': [
                    os.path.basename(input_file),
                    len(df),
                    len(df.columns),
                    df.isnull().sum().sum(),
                    len(df.dropna()),
                    round(df.memory_usage(deep=True).sum() / 1024 / 1024, 2),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    separator,
                    str(max_level) if max_level else "All levels"
                ]
            }
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Column details sheet
            col_details = []
            for col in df.columns:
                col_info = {
                    'Column_Name': col,
                    'Data_Type': str(df[col].dtype),
                    'Non_Null_Count': df[col].count(),
                    'Null_Count': df[col].isnull().sum(),
                    'Unique_Values': df[col].nunique(),
                    'Sample_Value': str(df[col].dropna().iloc[0]) if not df[col].dropna().empty else 'N/A'
                }
                col_details.append(col_info)
            
            details_df = pd.DataFrame(col_details)
            details_df.to_excel(writer, sheet_name='Column_Details', index=False)
        
        print(f"✅ Successfully exported to: {output_file}")
        print(f"📊 Sheets created: Data, Summary, Column_Details")
        print(f"📈 Data: {len(df)} rows × {len(df.columns)} columns")
        
        return True
        
    except FileNotFoundError:
        print(f"❌ Error: File not found: {input_file}")
        return False
    except json.JSONDecodeError as e:
        print(f"❌ Error: Invalid JSON format: {e}")
        return False
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return False

def main():
    """Main function for command-line usage"""
    if len(sys.argv) < 3:
        print("Usage: python json_to_excel.py <input.json> <output.xlsx> [separator] [max_level]")
        print("Example: python json_to_excel.py data.json output.xlsx")
        print("Example: python json_to_excel.py data.json output.xlsx . 3")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    separator = sys.argv[3] if len(sys.argv) > 3 else "_"
    max_level = int(sys.argv[4]) if len(sys.argv) > 4 and sys.argv[4].isdigit() else None
    
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'
    
    success = json_to_excel(input_file, output_file, separator, max_level)
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
