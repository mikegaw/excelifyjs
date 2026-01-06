## Complete Examples

### Basic Usage

```javascript
import { Workbook } from 'excelifyjs';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Data');

// Write headers
sheet.write(0, 0, 'ID');
sheet.write(0, 1, 'Name');
sheet.write(0, 2, 'Score');

// Write data
sheet.write(1, 0, 1);
sheet.write(1, 1, 'Alice');
sheet.write(1, 2, 95.5);

sheet.write(2, 0, 2);
sheet.write(2, 1, 'Bob');
sheet.write(2, 2, 87.3);

workbook.save('students.xlsx');
```

### Multiple Worksheets

```javascript
import { Workbook } from 'excelifyjs';

const workbook = new Workbook();

// Create multiple sheets
const sales = workbook.addWorksheet('Sales');
const expenses = workbook.addWorksheet('Expenses');
const summary = workbook.addWorksheet('Summary');

// Write to different sheets
sales.write(0, 0, 'Q1 Sales');
sales.write(1, 0, 100000);

expenses.write(0, 0, 'Q1 Expenses');
expenses.write(1, 0, 45000);

summary.write(0, 0, 'Net Profit');
summary.write(1, 0, 55000);

console.log(`Created ${workbook.worksheetCount} worksheets`);
workbook.save('financial-report.xlsx');
```

### Large Dataset

```javascript
import { Workbook } from 'excelifyjs';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Large Dataset');

// Write headers
const headers = ['ID', 'Name', 'Email', 'Age', 'Active'];
headers.forEach((header, col) => {
  sheet.write(0, col, header);
});

// Write 10,000 rows of data
for (let row = 1; row <= 10000; row++) {
  sheet.write(row, 0, row);
  sheet.write(row, 1, `User ${row}`);
  sheet.write(row, 2, `user${row}@example.com`);
  sheet.write(row, 3, 20 + (row % 50));
  sheet.write(row, 4, row % 2 === 0);
}

workbook.save('large-dataset.xlsx');
console.log('Created Excel file with 10,000 rows');
```

### Dynamic Data from Array

```javascript
import { Workbook } from 'excelifyjs';

const data = [
  { name: 'Alice', age: 30, city: 'New York' },
  { name: 'Bob', age: 25, city: 'San Francisco' },
  { name: 'Charlie', age: 35, city: 'Chicago' }
];

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Users');

// Write headers
const headers = Object.keys(data[0]);
headers.forEach((header, col) => {
  sheet.write(0, col, header);
});

// Write data
data.forEach((row, rowIndex) => {
  headers.forEach((header, colIndex) => {
    sheet.write(rowIndex + 1, colIndex, row[header]);
  });
});

workbook.save('users.xlsx');
```

---

## Type Reference

### CellInput

The `write()` method accepts the following value types:

| Type | Description | Example |
|------|-------------|---------||
| `string` | Text values | `'Hello World'` |
| `number` | Numeric values (integers or decimals) | `42`, `3.14`, `-100.5` |
| `boolean` | Boolean values | `true`, `false` |

---

## Performance Tips

1. **Batch writes**: Group write operations together rather than saving after each write
2. **Use appropriate data types**: The library automatically handles type conversion
3. **Pre-allocate data**: If you know the data size, structure your writes sequentially
4. **Memory management**: For very large datasets (millions of rows), consider processing in chunks

**Example - Efficient writing:**
```javascript
const workbook = new Workbook();
const sheet = workbook.addWorksheet('Data');

// Efficient: Write all data, then save once
for (let row = 0; row < 100000; row++) {
  for (let col = 0; col < 10; col++) {
    sheet.write(row, col, `Cell ${row}-${col}`);
  }
}

workbook.save('output.xlsx'); // Single save operation
```
