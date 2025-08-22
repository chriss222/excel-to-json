#!/usr/bin/env node

/**
 * Excel to JSON Converter (Node.js)
 *
 * This script converts Excel files (.xlsx, .xls) to JSON format using Node.js
 *
 * Requirements:
 *   npm install xlsx fs-extra yargs
 *
 * Usage:
 *   node excel-to-json.js input.xlsx
 *   node excel-to-json.js input.xlsx --output output.json
 *   node excel-to-json.js input.xlsx --sheet "Sheet1" --pretty
 *   node excel-to-json.js input.xlsx --all-sheets
 */

const XLSX = require("xlsx");
const fs = require("fs-extra");
const path = require("path");
const yargs = require("yargs/yargs");
const { hideBin } = require("yargs/helpers");

// Configure command line arguments
const argv = yargs(hideBin(process.argv))
  .usage("Usage: $0 <excel-file> [options]")
  .command("$0 <excelFile>", "Convert Excel file to JSON", (yargs) => {
    yargs.positional("excelFile", {
      describe: "Path to Excel file",
      type: "string",
    });
  })
  .option("output", {
    alias: "o",
    type: "string",
    description: "Output JSON file path",
  })
  .option("sheet", {
    alias: "s",
    type: "string",
    description: "Specific sheet name to convert",
  })
  .option("all-sheets", {
    alias: "a",
    type: "boolean",
    description: "Convert all sheets",
    default: false,
  })
  .option("pretty", {
    alias: "p",
    type: "boolean",
    description: "Pretty print JSON with indentation",
    default: false,
  })
  .option("list-sheets", {
    alias: "l",
    type: "boolean",
    description: "List all sheet names",
    default: false,
  })
  .option("header", {
    alias: "h",
    type: "number",
    description: "Row to use as header (0-indexed)",
    default: 0,
  })
  .option("no-id", {
    type: "boolean",
    description: "Skip adding sequential ID field to each row",
    default: false,
  })
  .option("no-camel-case", {
    type: "boolean",
    description: "Skip converting column names to camelCase",
    default: false,
  })
  .example("$0 data.xlsx", "Convert Excel to JSON")
  .example("$0 data.xlsx --pretty", "Convert with pretty formatting")
  .example('$0 data.xlsx --sheet "Sales" -o sales.json', "Convert specific sheet")
  .example("$0 data.xlsx --no-camel-case", "Keep original column names")
  .example("$0 data.xlsx --no-id", "Convert without adding ID field")
  .example("$0 data.xlsx --all-sheets --pretty", "Convert all sheets")
  .help().argv;

/**
 * Convert string to camelCase
 * @param {string} str - String to convert
 * @returns {string} camelCased string
 */
function toCamelCase(str) {
  return (
    str
      .toString()
      .trim()
      // Replace multiple spaces with single space
      .replace(/\s+/g, " ")
      // Split by spaces, hyphens, underscores, and other separators
      .split(/[\s\-_\.]+/)
      // Filter out empty strings
      .filter((word) => word.length > 0)
      // Convert to camelCase
      .map((word, index) => {
        // First word stays lowercase, others get capitalized
        if (index === 0) {
          return word.toLowerCase();
        }
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
      })
      .join("")
  );
}

/**
 * Clean and process sheet data
 * @param {Array} data - Raw sheet data
 * @param {boolean} addId - Whether to add sequential ID field
 * @param {boolean} camelCase - Whether to convert column names to camelCase
 * @returns {Array} Processed data
 */
function processSheetData(data, addId = true, camelCase = true) {
  if (!data || data.length === 0) return [];

  // Remove empty rows
  const filteredData = data.filter((row) => {
    return Object.values(row).some((cell) => cell !== null && cell !== undefined && cell !== "");
  });

  // Clean up column names and cell values, add ID if requested
  return filteredData.map((row, index) => {
    const cleanRow = {};

    // Add ID field first if requested
    if (addId) {
      cleanRow.id = index;
    }

    for (let key in row) {
      // Clean and format column names
      let cleanKey;
      if (camelCase) {
        cleanKey = toCamelCase(key);
      } else {
        // Just clean up spaces if camelCase is disabled
        cleanKey = key.toString().trim().replace(/\s+/g, " ");
      }

      // Skip empty column names
      if (!cleanKey) continue;

      // Convert cell values to appropriate types
      let value = row[key];

      // Handle dates
      if (value instanceof Date) {
        value = value.toISOString().split("T")[0]; // YYYY-MM-DD format
      }
      // Handle empty values
      else if (value === null || value === undefined || value === "") {
        value = null;
      }

      cleanRow[cleanKey] = value;
    }
    return cleanRow;
  });
}

/**
 * Convert Excel file to JSON
 * @param {string} filePath - Path to Excel file
 * @param {Object} options - Conversion options
 * @returns {Object} Conversion result
 */
function convertExcelToJson(filePath, options = {}) {
  try {
    // Check if file exists
    if (!fs.existsSync(filePath)) {
      throw new Error(`File not found: ${filePath}`);
    }

    // Read Excel file
    const workbook = XLSX.readFile(filePath);
    const sheetNames = workbook.SheetNames;

    if (options.listSheets) {
      return {
        success: true,
        sheets: sheetNames,
        message: `Found ${sheetNames.length} sheet(s)`,
      };
    }

    let result = {};
    const addId = !options.noId; // Add ID unless specifically disabled
    const camelCase = !options.noCamelCase; // Use camelCase unless specifically disabled

    if (options.allSheets) {
      // Convert all sheets
      sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, {
          header: options.header,
          defval: null,
        });
        result[sheetName] = processSheetData(rawData, addId, camelCase);
      });
    } else {
      // Convert specific sheet or first sheet
      const targetSheet = options.sheet || sheetNames[0];

      if (!sheetNames.includes(targetSheet)) {
        throw new Error(`Sheet "${targetSheet}" not found. Available sheets: ${sheetNames.join(", ")}`);
      }

      const worksheet = workbook.Sheets[targetSheet];
      const rawData = XLSX.utils.sheet_to_json(worksheet, {
        header: options.header,
        defval: null,
      });
      result = processSheetData(rawData, addId, camelCase);
    }

    return {
      success: true,
      data: result,
      sheets: sheetNames,
      message: "Conversion successful",
    };
  } catch (error) {
    return {
      success: false,
      error: error.message,
    };
  }
}

/**
 * Save JSON data to file
 * @param {Object} data - Data to save
 * @param {string} outputPath - Output file path
 * @param {boolean} pretty - Pretty print option
 */
async function saveJsonFile(data, outputPath, pretty = false) {
  try {
    const jsonString = pretty ? JSON.stringify(data, null, 2) : JSON.stringify(data);

    await fs.writeFile(outputPath, jsonString, "utf8");
    return { success: true, message: `Saved to ${outputPath}` };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Main function
 */
async function main() {
  const {
    excelFile,
    output,
    sheet,
    "all-sheets": allSheets,
    pretty,
    "list-sheets": listSheets,
    header,
    "no-id": noId,
    "no-camel-case": noCamelCase,
  } = argv;

  // Convert Excel to JSON
  const result = convertExcelToJson(excelFile, {
    sheet,
    allSheets,
    listSheets,
    header,
    noId,
    noCamelCase,
  });

  if (!result.success) {
    console.error("❌ Error:", result.error);
    process.exit(1);
  }

  // Handle list sheets option
  if (listSheets) {
    console.log(`📋 Sheets in "${excelFile}":`);
    result.sheets.forEach((sheetName, index) => {
      console.log(`   ${index + 1}. ${sheetName}`);
    });
    return;
  }

  // Determine output file
  let outputFile = output;
  if (!outputFile) {
    const parsedPath = path.parse(excelFile);
    outputFile = path.join(parsedPath.dir, `${parsedPath.name}.json`);
  }

  // Save to file
  const saveResult = await saveJsonFile(result.data, outputFile, pretty);

  if (saveResult.success) {
    console.log("✅ Success!");
    console.log(`📁 Input: ${excelFile}`);
    console.log(`📄 Output: ${outputFile}`);
    console.log(`📊 Sheets processed: ${result.sheets.join(", ")}`);

    // Show data preview
    if (allSheets) {
      Object.keys(result.data).forEach((sheetName) => {
        console.log(`   ${sheetName}: ${result.data[sheetName].length} rows`);
      });
    } else {
      console.log(`📈 Rows: ${result.data.length}`);
    }
  } else {
    console.error("❌ Save Error:", saveResult.error);
    process.exit(1);
  }
}

// Quick conversion functions for programmatic use
function quickConvert(excelFile, outputFile = null, options = {}) {
  const result = convertExcelToJson(excelFile, options);
  if (!result.success) {
    throw new Error(result.error);
  }

  if (outputFile) {
    const jsonString = options.pretty ? JSON.stringify(result.data, null, 2) : JSON.stringify(result.data);
    fs.writeFileSync(outputFile, jsonString);
    return `Converted to ${outputFile}`;
  }

  return result.data;
}

// Export for use as module
module.exports = {
  convertExcelToJson,
  processSheetData,
  quickConvert,
};

// Run if called directly
if (require.main === module) {
  main().catch((error) => {
    console.error("❌ Fatal Error:", error.message);
    process.exit(1);
  });
}
