
## Worksheet

The `Worksheet` class represents a single worksheet within a workbook.

### Methods

#### `write(row: number, col: number, value: CellInput): void`

Writes a value to a cell at the specified row and column. Both row and column indices are **zero-based** (start from 0).

**Parameters:**
- `row` (number): Zero-based row index (0 = first row)
- `col` (number): Zero-based column index (0 = column A)
- `value` (CellInput): The value to write. Can be:
  - `string` - Text values
  - `number` - Numeric values (integers or decimals)
  - `boolean` - Boolean values (true/false)

**Examples:**

```javascript
// Write header row (row 0)
sheet.write(0, 0, 'Product');  // A1
sheet.write(0, 1, 'Price');    // B1
sheet.write(0, 2, 'In Stock'); // C1

// Write data rows
sheet.write(1, 0, 'Laptop');   // A2 - String
sheet.write(1, 1, 999.99);     // B2 - Number
sheet.write(1, 2, true);       // C2 - Boolean

sheet.write(2, 0, 'Mouse');    // A3
sheet.write(2, 1, 25.50);      // B3
sheet.write(2, 2, false);      // C3
```

### Properties

#### `name: string`

Gets the name of the worksheet (read-only).

**Example:**
```javascript
const sheetName = sheet.name;
console.log(`Worksheet name: ${sheetName}`);
```