# Test JSON Data Files

This directory contains various JSON test files to thoroughly test the JSON to Tabular Converter application.

## Test Files Overview

### 1. `simple_object.json`
- **Purpose**: Tests basic single object conversion
- **Structure**: Simple flat object with basic data types
- **Expected Result**: Single row table with 5 columns

### 2. `simple_array.json`
- **Purpose**: Tests array of simple objects
- **Structure**: Array containing 3 employee objects
- **Expected Result**: 3-row table with consistent columns

### 3. `nested_object.json`
- **Purpose**: Tests deeply nested single object
- **Structure**: Complex nested object with multiple levels
- **Expected Result**: Single row with flattened column names (using separator)

### 4. `complex_nested_array.json`
- **Purpose**: Tests array with nested objects and arrays
- **Structure**: Product catalog with nested specifications and reviews
- **Expected Result**: Complex flattened table with array handling

### 5. `employee_records.json`
- **Purpose**: Tests realistic employee data with mixed nesting
- **Structure**: Employee records with personal, employment, and project data
- **Expected Result**: Employee table with flattened address, manager, and project info

### 6. `deeply_nested_sales.json`
- **Purpose**: Tests very deep nesting (4+ levels)
- **Structure**: Sales data with regions → countries → cities hierarchy
- **Expected Result**: Heavily flattened table with deep column names

### 7. `mixed_data_types.json`
- **Purpose**: Tests handling of null values, empty arrays, and empty objects
- **Structure**: Products with various null/empty combinations
- **Expected Result**: Table with proper null handling (depending on settings)

## Testing Scenarios

### Basic Functionality
- Load each file and verify successful parsing
- Test conversion with default settings
- Verify output structure and data integrity

### "" Features
- Test different separator characters (`_`, `.`, `-`)
- Test max nesting level limits
- Test null value removal option
- Test array handling options

### Edge Cases
- Empty arrays and objects
- Null values in various positions
- Very deep nesting levels
- Mixed data types

## Usage Tips

1. Start with `simple_object.json` and `simple_array.json` for basic testing
2. Use `nested_object.json` to test separator functionality
3. Test array handling with `complex_nested_array.json`
4. Use `mixed_data_types.json` to test null value handling
5. Challenge the converter with `deeply_nested_sales.json` for maximum nesting

## Expected Behavior

- All files should load without errors
- Conversion should produce readable tabular data
- Column names should reflect the nested structure using the chosen separator
- Array elements should be handled according to the selected options
- Null values should be processed based on the removal setting
