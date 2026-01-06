# Getting Started

Learn how to install and use excelifyjs to create Excel files in Node.js.

## Installation

Install excelifyjs using your preferred package manager:

::: code-group

```bash [npm]
npm install excelifyjs
```

```bash [yarn]
yarn add excelifyjs
```

```bash [pnpm]
pnpm add excelifyjs
```

:::

## Requirements

- Node.js >= 16
- Supported platforms:
  - Windows (x64, ARM64)
  - macOS (x64, ARM64)
  - Linux (x64, ARM64) - both GNU and musl libc

## Quick Start

Create your first Excel file in just a few lines:

```javascript
import { Workbook } from 'excelifyjs';

// Create a new workbook
const workbook = new Workbook();

// Add a worksheet
const sheet = workbook.addWorksheet('Sheet1');

// Write data to cells (row, column, value)
sheet.write(0, 0, 'Name');
sheet.write(0, 1, 'Age');
sheet.write(0, 2, 'Active');

sheet.write(1, 0, 'Alice');
sheet.write(1, 1, 30);
sheet.write(1, 2, true);

sheet.write(2, 0, 'Bob');
sheet.write(2, 1, 25);
sheet.write(2, 2, false);

// Save the workbook
workbook.save('output.xlsx');
```

::: tip
Row and column indices are zero-based. Row 0 is the first row, and column 0 is column A.
:::

## Core Concepts

### Workbook

A workbook is the main container for your Excel file. It can contain multiple worksheets.

```javascript
const workbook = new Workbook();
```

### Worksheet

A worksheet is a single sheet within a workbook where you write data.

```javascript
const sheet = workbook.addWorksheet('MySheet');
```

### Cell Coordinates

Cells are addressed using zero-based row and column indices:

- `write(0, 0, value)` → Cell A1
- `write(0, 1, value)` → Cell B1
- `write(1, 0, value)` → Cell A2

### Supported Data Types

excelifyjs supports three data types:

- **String**: Text values
- **Number**: Integers and decimals
- **Boolean**: `true` or `false`

```javascript
sheet.write(0, 0, 'Text');     // String
sheet.write(0, 1, 42.5);       // Number
sheet.write(0, 2, true);       // Boolean
```

## Next Steps

- Check out the [API Reference](/api-examples) for detailed documentation
- View more [examples in the repository](https://github.com/user/excelifyjs/tree/main/examples)
- Run the [benchmark](https://github.com/user/excelifyjs/tree/main/benchmarks) to see performance metrics
