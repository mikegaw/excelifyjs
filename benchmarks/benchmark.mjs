import { Workbook } from '../index.js';
import { statSync, existsSync, unlinkSync } from 'fs';

const ROWS = 100_000;
const COLS = 10;

console.log(`Benchmark: Writing ${ROWS.toLocaleString()} rows x ${COLS} columns (${(ROWS * COLS).toLocaleString()} cells)`);
console.log('='.repeat(60));

// Cleanup previous benchmark file
if (existsSync('benchmark.xlsx')) {
  unlinkSync('benchmark.xlsx');
}

// Benchmark: Create workbook and write data
const startWrite = performance.now();

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Data');

for (let row = 0; row < ROWS; row++) {
  for (let col = 0; col < COLS; col++) {
    const type = col % 3;
    if (type === 0) {
      sheet.write(row, col, `Cell ${row}-${col}`);
    } else if (type === 1) {
      sheet.write(row, col, row * col + 0.5);
    } else {
      sheet.write(row, col, row % 2 === 0);
    }
  }
}

const endWrite = performance.now();
const writeTime = endWrite - startWrite;

console.log(`Write time:  ${writeTime.toFixed(2)} ms`);
console.log(`Write rate:  ${Math.round((ROWS * COLS) / (writeTime / 1000)).toLocaleString()} cells/sec`);

// Benchmark: Save to file
const startSave = performance.now();
workbook.save('benchmark.xlsx');
const endSave = performance.now();
const saveTime = endSave - startSave;

console.log(`Save time:   ${saveTime.toFixed(2)} ms`);
console.log(`Total time:  ${(writeTime + saveTime).toFixed(2)} ms`);

// File size
if (existsSync('benchmark.xlsx')) {
  const stats = statSync('benchmark.xlsx');
  const sizeMB = stats.size / 1024 / 1024;
  console.log(`File size:   ${sizeMB < 1 ? (stats.size / 1024).toFixed(2) + ' KB' : sizeMB.toFixed(2) + ' MB'}`);
}

console.log('='.repeat(60));
