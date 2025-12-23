import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Config
const EXCEL_FILE = path.join(__dirname, 'sheets', 'NAME OF YOUR FILE'); // <-- CHANGE THIS TO YOUR EXCEL FILE.xlsx
const EXCEL_PATH = path.resolve(__dirname, EXCEL_FILE);

// Get entity code from command line
const args = process.argv.slice(2);
if (args.length === 0) {
    console.log(`
Usage:
  node generate-business-unit.js <ENTITY_CODE> <NAME>
File: ${EXCEL_FILE}
Generates INSERT ... ON CONFLICT for business_unit table
`);
    process.exit(1);
}

const ENTITY_CODE = args[0].trim();
const ENTITY_NAME = args[1] + ' ' + args[2].trim() + ' ' + args[3].trim();

// Check file exists
if (!fs.existsSync(EXCEL_PATH)) {
    console.error(`ERROR: File not found: ${EXCEL_PATH}`);
    process.exit(1);
}

console.log(`Generating business_unit INSERTs for entity: ${ENTITY_CODE}`);
console.log(`Reading: ${EXCEL_FILE}\n`);

// Read workbook
const workbook = XLSX.readFile(EXCEL_PATH);
const firstSheet = workbook.SheetNames[0];
const secondSheet = workbook.SheetNames[1];
const sheetMatrix = workbook.Sheets[firstSheet]; //Sheet 1 must be for Matrix
const sheetUos = workbook.Sheets[secondSheet]; //Sheet 2 must be for Matrix

// Read as array of arrays
const rawDataMat = XLSX.utils.sheet_to_json(sheetMatrix, { header: 1, defval: '' });
const rawDataUos = XLSX.utils.sheet_to_json(sheetUos, { header: 1, defval: '' });

// Skip header row
const dataRowsMat = rawDataMat.slice(1);
const dataRowsUos = rawDataUos.slice(1);

let sql = '';
let count = 0;

sql += `begin;\n \n`;
sql += `----- delete old UO matrix -----\n`;
sql += `delete from entity_right_matrix where uo_from like '${ENTITY_CODE}:%' or uo_to  like '${ENTITY_CODE}:%';\n \n`;

sql += `----- entity -----\n`;
sql += `INSERT INTO entity (code, descr)\n`;
sql += `VALUES('${ENTITY_CODE}', '${ENTITY_NAME}') ON CONFLICT (code) DO UPDATE SET descr = EXCLUDED.descr; \n`;
sql += `UPDATE business_unit SET "type" = 6 where entity_code = '${ENTITY_CODE}';\n\n`;
sql += `----- UO list -----\n`;

for (const row of dataRowsUos) {
    if (!row || row.length < 3) continue;

    const codeUO = String(row[0]).trim();      // Column A: CODE UO
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

const headers = rawDataMat[0].slice(1).map(h => String(h).trim());
sql += `\n`;
sql += `----- UO matrix -----\n`;

for (const row of rawDataMat.slice(1)) {
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

sql += 'commit; \n \n';
sql += '\n';
const outputFile = `${ENTITY_CODE}_UO_MATRIX.sql`;
fs.writeFileSync(outputFile, sql.trim() + '\n');

console.log('✅ SUCCESS!');
console.log(`Generated: ${outputFile}`);
console.log(`Total INSERTs: ${count}`);