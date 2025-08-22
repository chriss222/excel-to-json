Excel to JSON Converter
A powerful Node.js utility that converts Excel files (.xlsx, .xls) to JSON format with automatic ID generation, camelCase column names, and flexible output options.
✨ Features

🔄 Convert Excel to JSON - Support for .xlsx and .xls files
🏷️ Auto ID Generation - Automatically adds sequential ID field (0, 1, 2, 3...)
🐪 CamelCase Columns - Converts column names to camelCase format ("premiu turneu" → "premiuTurneu")
📋 Multi-sheet Support - Convert specific sheets or all sheets at once
🎨 Pretty Formatting - Optional JSON pretty printing with indentation
⚙️ Flexible Options - Comprehensive command-line interface
📦 Programmatic API - Use as a module in your own projects

🚀 Quick Start
Installation
bash# Clone or download the project
git clone <repository-url>
cd excel-to-json

# Install dependencies
npm install xlsx fs-extra yargs
Basic Usage
bash# Convert Excel file to JSON
node excel-to-json.js data.xlsx

# Convert with pretty formatting
node excel-to-json.js data.xlsx --pretty
📖 Usage Guide
Basic Commands
bash# Convert Excel to JSON with default settings
node excel-to-json.js input.xlsx

# Convert with custom output file
node excel-to-json.js input.xlsx --output output.json

# Convert with pretty formatting
node excel-to-json.js input.xlsx --pretty
Working with Sheets
bash# List all sheets in Excel file
node excel-to-json.js data.xlsx --list-sheets

# Convert specific sheet
node excel-to-json.js data.xlsx --sheet "Sales Data"

# Convert all sheets to single JSON file
node excel-to-json.js data.xlsx --all-sheets --pretty
Column and Data Options
bash# Keep original column names (disable camelCase)
node excel-to-json.js data.xlsx --no-camel-case

# Skip adding ID field
node excel-to-json.js data.xlsx --no-id

# Combine options
node excel-to-json.js data.xlsx --no-id --no-camel-case --pretty
🛠️ Command Line Options
OptionAliasTypeDescription--output-ostringOutput JSON file path--sheet-sstringSpecific sheet name to convert--all-sheets-abooleanConvert all sheets--pretty-pbooleanPretty print JSON with indentation--list-sheets-lbooleanList all sheet names in Excel file--no-idbooleanSkip adding sequential ID field--no-camel-casebooleanSkip converting column names to camelCase--header-hnumberRow to use as header (0-indexed, default: 0)--helpbooleanShow help information
📋 Examples
Real-world Usage Examples
bash# Convert sales report with pretty formatting
node excel-to-json.js sales-report.xlsx --pretty --output sales.json

# Convert specific quarterly data sheet
node excel-to-json.js annual-report.xlsx --sheet "Q1 Data" --pretty

# Convert all sheets from a workbook
node excel-to-json.js workbook.xlsx --all-sheets --output combined-data.json

# Convert keeping original formatting (no camelCase, no ID)
node excel-to-json.js raw-data.xlsx --no-camel-case --no-id --pretty

# Quick preview of sheet names
node excel-to-json.js data.xlsx -l
🔄 Data Transformations
Column Name Conversion (CamelCase)
The converter automatically transforms column names to camelCase:
Original Excel ColumnConverted JSON Key"premiu turneu""premiuTurneu""First Name""firstName""user_email""userEmail""product-category""productCategory""Order   Date""orderDate""total.amount""totalAmount"
Sample Input/Output
Excel Data:
| premiu turneu | phone      | cash | tesla   |
|---------------|------------|------|---------|
| Premium Gold  | 1234567890 | 500  | Model 3 |
| Basic Silver  | 9876543210 | 200  | Model Y |
JSON Output (default settings):
json[
  {
    "id": 0,
    "premiuTurneu": "Premium Gold",
    "phone": "1234567890",
    "cash": 500,
    "tesla": "Model 3"
  },
  {
    "id": 1,
    "premiuTurneu": "Basic Silver", 
    "phone": "9876543210",
    "cash": 200,
    "tesla": "Model Y"
  }
]
JSON Output (with --no-camel-case --no-id):
json[
  {
    "premiu turneu": "Premium Gold",
    "phone": "1234567890",
    "cash": 500,
    "tesla": "Model 3"
  },
  {
    "premiu turneu": "Basic Silver",
    "phone": "9876543210", 
    "cash": 200,
    "tesla": "Model Y"
  }
]
🔧 Programmatic Usage
You can also use the converter as a module in your Node.js projects:
javascriptconst { convertExcelToJson, quickConvert } = require('./excel-to-json.js');

// Quick conversion
try {
  const data = quickConvert('data.xlsx', 'output.json', { 
    pretty: true 
  });
  console.log('Conversion successful!');
} catch (error) {
  console.error('Conversion failed:', error.message);
}

// Advanced usage with options
const result = convertExcelToJson('data.xlsx', {
  sheet: 'Sales Data',
  allSheets: false,
  noId: false,
  noCamelCase: false
});

if (result.success) {
  console.log('Data:', result.data);
  console.log('Available sheets:', result.sheets);
} else {
  console.error('Error:', result.error);
}
🗂️ File Structure
After setup, your project should look like:
excel-to-json/
├── excel-to-json.js          # Main converter script
├── package.json              # Project dependencies
├── package-lock.json         # Locked dependencies
├── README.md                 # This file
├── node_modules/             # Installed packages (auto-generated)
└── .gitignore               # Git ignore rules
⚡ Performance Notes

Fast Processing: Optimized for large Excel files
Memory Efficient: Processes data in streams where possible
Error Handling: Comprehensive error messages and validation
Date Support: Automatically handles Excel date formats
Empty Data: Filters out completely empty rows

🐛 Troubleshooting
Common Issues
"File not found" error:
bash# Make sure the file path is correct
node excel-to-json.js /full/path/to/your/file.xlsx
"Sheet not found" error:
bash# List available sheets first
node excel-to-json.js data.xlsx --list-sheets

# Then use exact sheet name
node excel-to-json.js data.xlsx --sheet "Exact Sheet Name"
Permission errors:
bash# Make sure you have write permissions for output directory
node excel-to-json.js data.xlsx --output ~/Desktop/output.json
Dependencies Issues
bash# Reinstall dependencies if needed
rm -rf node_modules package-lock.json
npm install xlsx fs-extra yargs
📋 Requirements

Node.js 14.0.0 or higher
npm 6.0.0 or higher

Dependencies

xlsx - Excel file parsing and processing
fs-extra - Enhanced file system operations
yargs - Command-line argument parsing