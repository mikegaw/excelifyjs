## Workbook

The `Workbook` class represents an Excel workbook (.xlsx file).

### Constructor

```javascript
new Workbook()
```

Creates a new Excel workbook.

**Example:**
```javascript
import { Workbook } from 'excelifyjs';

const workbook = new Workbook();
```

### Methods

#### `addWorksheet(name: string): Worksheet`

Adds a new worksheet to the workbook with the specified name.

**Parameters:**
- `name` (string): The name of the worksheet

**Returns:** A `Worksheet` instance

**Example:**
```javascript
const sheet = workbook.addWorksheet('Sales Data');
```

#### `save(path: string): void`

Saves the workbook to a file at the specified path.

**Parameters:**
- `path` (string): The file path where the workbook will be saved

**Example:**
```javascript
workbook.save('output.xlsx');
workbook.save('/path/to/report.xlsx');
```

### Properties

#### `worksheetCount: number`

Gets the total number of worksheets in the workbook.

**Example:**
```javascript
const count = workbook.worksheetCount;
console.log(`Total worksheets: ${count}`);
```