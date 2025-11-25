// generate-business-unit.js
// Run with: node generate-business-unit.js 87200

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// CONFIG — your Excel file
const EXCEL_FILE = path.join(__dirname, 'sheets', '872-UO.xlsx');
const EXCEL_PATH = path.resolve(__dirname, EXCEL_FILE);

// Get entity code from command line
const args = process.argv.slice(2);
if (args.length === 0) {
    console.log(`
Usage:
  node generate-business-unit.js <ENTITY_CODE>
File: ${EXCEL_FILE}
Generates INSERT ... ON CONFLICT for business_unit table
`);
    process.exit(1);
}

const ENTITY_CODE = args[0].trim();

// Check file exists
if (!fs.existsSync(EXCEL_PATH)) {
    console.error(`ERROR: File not found: ${EXCEL_PATH}`);
    process.exit(1);
}

console.log(`Generating business_unit INSERTs for entity: ${ENTITY_CODE}`);
console.log(`Reading: ${EXCEL_FILE}\n`);

// Read workbook
const workbook = XLSX.readFile(EXCEL_PATH);
const firstSheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[firstSheetName];

// Read as array of arrays
const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

// Skip header row
const dataRows = rawData.slice(1);

let sql = '';
let count = 0;

for (const row of dataRows) {
    if (!row || row.length < 3) continue;

    const codeUO = String(row[0]).trim();     // Column A: CODE UO (AGR, PRO, etc.)
    const libelle = String(row[1]).trim();     // Column B: LIBELLE
    const typeStr = String(row[2]).trim();     // Column C: TYPE UO

    if (!codeUO || !libelle || !typeStr) continue;

    const type = parseInt(typeStr, 10);
    if (isNaN(type)) continue;

    const fullCode = `${ENTITY_CODE}:${codeUO}`;

    sql += `INSERT INTO business_unit\n`;
    sql += `(entity_code, "type", code, descr, business_unit_code, "rank", address_email)\n`;
    sql += `VALUES ('${ENTITY_CODE}', ${type}, '${fullCode}', '${libelle}', '${codeUO}', NULL, NULL) ON CONFLICT (code) DO UPDATE SET descr = EXCLUDED.descr, type = EXCLUDED.type, rank = EXCLUDED.rank; \n`;
    count++;
}

const outputFile = `insert_uo_${ENTITY_CODE}.sql`;
fs.writeFileSync(outputFile, sql.trim() + '\n');

console.log('✅ SUCCESS!');
console.log(`Generated: ${outputFile}`);
console.log(`Total INSERTs: ${count}`);