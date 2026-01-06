import { Workbook } from './index.js';

const workbook = new Workbook();

const sheet = workbook.addWorksheet('Sheet1');

sheet.write(0, 0, 'Name');
sheet.write(0, 1, 'Age');
sheet.write(0, 2, 'Active');

sheet.write(1, 0, 'Alice');
sheet.write(1, 1, 30);
sheet.write(1, 2, true);

sheet.write(2, 0, 'Bob');
sheet.write(2, 1, 25);
sheet.write(2, 2, false);

sheet.write(3, 0, 'Charlie');
sheet.write(3, 1, 35);
sheet.write(3, 2, true);

console.log(`Created worksheet: ${sheet.name}`);
console.log(`Total worksheets: ${workbook.worksheetCount}`);

workbook.save('example.xlsx');
console.log('Saved to example.xlsx');
