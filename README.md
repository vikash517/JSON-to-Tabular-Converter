# 🔄 JSON to Tabular Converter

A powerful Python desktop application that transforms nested JSON data into flat, structured tabular format for easy analysis and processing.

## 🎯 Project Overview

This application solves the common problem of converting complex, nested JSON structures into usable tabular data. Built with Python and Tkinter, it provides an intuitive GUI for importing JSON files, configuring conversion options, and exporting results to CSV or Excel formats.

## ✨ Features

### Core Functionality
- **JSON File Import**: Load JSON files with drag-and-drop interface
- **Intelligent Flattening**: Automatically flatten nested JSON structures
- **Configurable Options**: Customize separator characters, nesting levels, and data handling
- **Multiple Export Formats**: Save results as CSV or Excel files with advanced formatting
- **Real-time Preview**: View original JSON and converted tabular data side-by-side

### Advanced Excel Export Features
- **Formatted Headers**: Professional styling with bold fonts and colored backgrounds
- **Auto-sized Columns**: Intelligent column width adjustment based on content
- **Multiple Sheets**: 
  - **Data Sheet**: Main converted tabular data
  - **Summary Sheet**: Conversion statistics and metadata
  - **Column Details**: Detailed analysis of each column
- **Advanced Excel Export**: Category-based sheet separation for complex datasets

### Conversion Options
- **Custom Separators**: Choose how nested keys are joined (default: underscore)
- **Nesting Level Control**: Limit flattening depth for complex structures
- **Array Handling**: Convert JSON arrays into separate table rows
- **Data Cleaning**: Remove null/empty values automatically
- **Memory Optimization**: Efficient processing of large JSON files

### User Interface
- **Modern GUI**: Clean, professional interface with tabbed navigation
- **Status Updates**: Real-time feedback during conversion process
- **Data Summary**: Comprehensive statistics about converted data
- **Error Handling**: User-friendly error messages and validation

## 📁 Project Structure

```
json-to-tabular-converter/
├── 📄 main.py                    # Main launcher script
├── 📄 setup.sh                   # Automated setup script
├── 📄 requirements.txt           # Python dependencies
├── 📄 README.md                  # Project documentation
├── 📄 PROJECT_STRUCTURE.md       # Detailed structure guide
├── 📄 .gitignore                 # Git ignore rules
├── 🗂️ src/
│   └── 📄 json_converter.py      # Main GUI application
├── 🗂️ utils/
│   ├── 📄 json_to_excel.py       # Command-line Excel converter
│   └── 📄 test_excel_functionality.py  # Test suite
├── 🗂️ examples/
│   ├── 📄 demo_excel.py          # Demo script
│   └── 🗂️ sample_data/          # Sample JSON files for testing
│       ├── simple_object.json
│       ├── simple_array.json
│       ├── nested_object.json
│       ├── complex_nested_array.json
│       ├── employee_records.json
│       ├── deeply_nested_sales.json
│       ├── mixed_data_types.json
│       └── README.md
├── 🗂️ docs/
│   └── 📄 EXCEL_INTEGRATION_SUMMARY.md  # Technical documentation
└── 🗂️ venv/                      # Python virtual environment
```

### 📋 Directory Descriptions

- **`src/`** - Core application source code
- **`utils/`** - Command-line utilities and testing tools
- **`examples/`** - Demo scripts and sample data for testing
- **`docs/`** - Documentation and technical guides
- **`main.py`** - Primary entry point for the application
- **`setup.sh`** - One-command project setup script

## 🚀 Getting Started

### Prerequisites

Ensure you have Python 3.7+ installed with the following packages:

```bash
pip install -r requirements.txt
```

### Installation

1. Clone or download the project files
2. Navigate to the project directory
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Quick Start Options

#### Option 1: Automated Setup (Recommended)
```bash
chmod +x setup.sh
./setup.sh
./main.py
```

#### Option 2: Manual Launch
```bash
# Method 1: Use the main launcher
python main.py

# Method 2: Run directly from src
python src/json_converter.py

# Method 3: Virtual environment
./venv/bin/python main.py
```

#### Option 3: Command Line Only
```bash
python utils/json_to_excel.py examples/sample_data/employee_records.json output.xlsx
```

## 💻 Usage Guide

### Basic Workflow

1. **Launch Application**
   ```bash
   # Using main launcher
   python main.py
   
   # Or directly from source
   python src/json_converter.py
   ```

2. **Import JSON File**
   - Click "Choose JSON File" button
   - Select your JSON file from the file dialog
   - Try sample files from `examples/sample_data/` directory
   - The original JSON will display in the "Original JSON" tab

3. **Configure Conversion Options**
   - **Separator**: Character used to join nested keys (e.g., `_`, `.` , `-`)
   - **Max Level**: Maximum nesting depth to flatten (leave empty for all levels)
   - **Handle Arrays**: Convert array elements to separate rows
   - **Remove Nulls**: Automatically remove empty/null columns

4. **Convert Data**
   - Click "Convert to Tabular Format"
   - View results in the "Tabular Data" tab
   - Check conversion statistics in the "Summary" tab

5. **Export Results**
   - **Export as CSV**: Save as comma-separated values format
   - **Export as Excel**: Save as formatted Excel spreadsheet with:
     - Styled headers with bold fonts and colored backgrounds
     - Auto-adjusted column widths
     - Multiple sheets (Data, Summary, Column Details)
   - **Advanced Excel**: Multi-sheet export with category-based data separation

### Command-Line Excel Export

For batch processing or automation, use the command-line utility:

```bash
python utils/json_to_excel.py input.json output.xlsx [separator] [max_level]
```

**Examples:**
```bash
# Basic export
python utils/json_to_excel.py examples/sample_data/employee_records.json output.xlsx

# Custom separator and max levels
python utils/json_to_excel.py examples/sample_data/nested_object.json output.xlsx "." 3

# Using different separators
python utils/json_to_excel.py examples/sample_data/complex_nested_array.json output.xlsx "-"
```

**Command-line Features:**
- Automatic multi-sheet Excel generation
- Data summary and column analysis
- Progress feedback and error reporting
- Batch processing support

### Batch Processing

The application now supports batch processing for multiple JSON files:

#### GUI Batch Processing
1. Launch the application: `python main.py`
2. Click "📊 Batch Convert to Excel" button
3. Select multiple JSON files
4. Choose output directory
5. Monitor progress and view results

#### Command-Line Batch Processing
```bash
# Process all JSON files in a directory
for file in examples/sample_data/*.json; do
    python utils/json_to_excel.py "$file" "output/$(basename "$file" .json).xlsx"
done
```

#### Demo Script
Run the demo to convert all sample files:
```bash
python examples/demo_excel.py
```

## 💡 Tips for Effective Use

- **Explore Sample Data**: Use the provided sample JSON files to familiarize yourself with the converter's capabilities.
- **Start with Defaults**: Begin with default settings for separator and nesting levels; adjust as needed based on your data.
- **Check Output**: Always review the exported CSV/Excel files to ensure data integrity and correct formatting.
- **Use Batch Processing**: For large projects, utilize batch processing to convert multiple files efficiently.
- **Monitor Performance**: Keep an eye on the status bar for real-time feedback on conversion progress and any potential issues.

## 🛠️ Technical Details

### Architecture
- **Main Class**: `JSONToTabularConverter` - Handles application logic and GUI
- **JSON Processing**: Uses `pandas.json_normalize()` for efficient flattening
- **GUI Framework**: Tkinter with modern styling and responsive design
- **Data Handling**: Pandas DataFrames for robust data manipulation

### Key Methods
- `load_json_file()`: Imports and validates JSON data
- `convert_json_to_tabular()`: Performs the flattening operation
- `display_tabular_data()`: Shows converted results in table format
- `export_to_csv()` / `export_to_excel()`: Handles data export

### Conversion Strategy
- **Nested Objects**: Flattened using configurable separators
- **Arrays**: Can be exploded into separate rows or kept as strings
- **Data Types**: Automatically inferred and preserved where possible
- **Missing Values**: Handled gracefully with user-configurable options

## 🎨 Interface Features

### Color Scheme
- **Primary**: Professional blue (#2563eb)
- **Success**: Green for successful operations (#10b981)
- **Warning**: Orange for warnings (#f59e0b)
- **Clean Layout**: Light backgrounds with clear visual hierarchy

### Responsive Design
- **Tabbed Interface**: Organized display of data and options
- **Scrollable Text Areas**: Handle large datasets efficiently
- **Status Bar**: Real-time feedback on operations
- **Error Dialogs**: Clear error messaging and recovery suggestions

## 🔧 Customization

### Adding New Export Formats
To support additional export formats:

1. Add new export method to the `JSONToTabularConverter` class
2. Create corresponding button in `create_export_section()`
3. Handle file type in the export dialog

### Modifying Conversion Logic
Customize flattening behavior by:
- Updating parameters in `convert_json_to_tabular()`
- Adding preprocessing steps for specific JSON structures
- Implementing custom normalization functions

## 🐛 Troubleshooting

### Common Issues

**Large File Performance**
- For very large JSON files (>100MB), consider chunked processing
- Increase system memory if encountering memory errors
- Use nesting level limits to reduce output complexity

**Invalid JSON Format**
- Ensure JSON files are properly formatted
- Check for trailing commas or syntax errors
- Use online JSON validators for verification

**Export Errors**
- Verify write permissions in target directory
- Ensure sufficient disk space for output files
- Check that target applications can handle the column count

### Error Messages
- **"JSON data must be an object or array"**: Input must be valid JSON structure
- **"Failed to convert JSON"**: Check JSON syntax and format
- **"No data to export"**: Perform conversion before attempting export

## 🤝 Contributing

This project is part of the GUVI Data Science program. Contributions and improvements are welcome!

### Development Setup
1. Fork the repository
2. Create a virtual environment
3. Install development dependencies
4. Make changes and test thoroughly
5. Submit pull request with clear description

## 📄 License

This project is created for educational purposes as part of the GUVI Data Science curriculum.

## 🎯 Future Enhancements

- [ ] Support for JSONL (JSON Lines) format
- [ ] Batch processing for multiple JSON files
- [ ] Custom column renaming and reordering
- [ ] Data type inference and conversion options
- [ ] Integration with database systems
- [ ] Command-line interface for automation
- [ ] "" filtering and transformation rules
- [ ] Support for XML to tabular conversion

---

**Project**: GUVI Data Science - JSON to Tabular Converter  
**Created**: August 2025  
**Technology Stack**: Python, Pandas, Tkinter, JSON Processing

# JSON-Tabular-Converter

## 📊 Excel Export Features

The enhanced Excel export functionality provides professional-grade spreadsheet output with advanced formatting and multiple analysis sheets.

### Standard Excel Export
- **Formatted Headers**: Bold white text on blue background
- **Auto-sizing**: Intelligent column width adjustment (max 50 characters)
- **Data Types**: Proper handling of numbers, dates, and text
- **Summary Sheet**: Conversion statistics and metadata
- **Column Details**: Analysis of data types, null counts, and sample values

### Advanced Excel Export
- **Multi-sheet Organization**: Separate sheets by data categories
- **Category Detection**: Automatic grouping by type/category columns
- **Comprehensive Analysis**: Detailed breakdown of data structure
- **Professional Formatting**: Consistent styling across all sheets

### Excel Export Benefits
- **Business Ready**: Professional formatting suitable for reports
- **Data Analysis**: Multiple sheets enable different analytical perspectives  
- **Preservation**: Maintains data integrity during conversion
- **Accessibility**: Easy to share and collaborate using Excel

### Command-Line Excel Tool
The `json_to_excel.py` utility provides:
- **Batch Processing**: Convert multiple files programmatically
- **Custom Configuration**: Flexible separator and level options
- **Progress Tracking**: Real-time conversion status
- **Error Handling**: Detailed error reporting and recovery

**Usage Examples:**
```bash
# Convert with default settings
python json_to_excel.py employee_data.json report.xlsx

# Use dot separator with 2-level nesting
python json_to_excel.py nested_data.json analysis.xlsx "." 2

# Process sample data
python json_to_excel.py sample_data/complex_nested_array.json products.xlsx
```