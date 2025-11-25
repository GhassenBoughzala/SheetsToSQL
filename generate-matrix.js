import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Your file
const EXCEL_FILE = './sheets/872-Matrix.xlsx';
const EXCEL_PATH = path.resolve(__dirname, EXCEL_FILE);

// Get entity code
const args = process.argv.slice(2);
if (args.length === 0) {
    console.log(`
Usage:
  node generate-matrix-sql.js <ENTITY_CODE>

Reading file: ${EXCEL_FILE}
Using first sheet automatically
`);
    process.exit(1);
}

const ENTITY_CODE = args[0].trim();

// Check file exists
if (!fs.existsSync(EXCEL_PATH)) {
    console.error(`File not found: ${EXCEL_PATH}`);
    console.error(`Make sure the file exists at: ${EXCEL_FILE}`);
    process.exit(1);
}

console.log(`Generating SQL for entity: ${ENTITY_CODE}`);
console.log(`Reading: ${EXCEL_FILE}\n`);

// Read workbook
const workbook = XLSX.readFile(EXCEL_PATH);

// Use the FIRST sheet automatically
const firstSheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[firstSheetName];

console.log(`Using sheet: "${firstSheetName}"\n`);

// Convert to array of arrays
const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

// Headers: row 1, from column B onward (skip first empty cell)
const headers = rawData[0].slice(1).map(h => String(h).trim());

// Generate SQL
let sql = '';
let count = 0;

for (const row of rawData.slice(1)) {
    if (!row || !row[0]) continue;
    const uoFrom = String(row[0]).trim();
    if (!uoFrom) continue;

    for (let i = 0; i < headers.length; i++) {
        const cell = row[i + 1];
        const uoTo = headers[i];

        let access = '0';
        const val = String(cell || '').trim().toLowerCase();
        if (val === 'm') access = 'm';
        else if (val === 'l') access = 'l';
        else if (val === '0' || cell === 0) access = '0';

        sql += `INSERT INTO entity_right_matrix (id, "access", uo_from, uo_to) VALUES (nextval('seqerm'), '${access}', '${ENTITY_CODE}:${uoFrom}', '${ENTITY_CODE}:${uoTo}');\n`;
        count++;
    }
    sql += '\n';
}

// Save file
const outputFile = `insert_matrix_${ENTITY_CODE}.sql`;
fs.writeFileSync(outputFile, sql.trim() + '\n');

console.log('✅ SUCCESS!');
console.log(`Generated: ${outputFile}`);
console.log(`Total INSERTs: ${count}`);