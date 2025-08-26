const XLSX = require('xlsx');
const fs = require('fs');

// Step 1: Load Excel and convert to CSV
const workbook = XLSX.readFile('db.xlsx');
const sheetName = workbook.SheetNames[0];
const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);

// Step 2: Write CSV to temp file
fs.writeFileSync('db.csv', csv);
