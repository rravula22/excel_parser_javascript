# Excel Parser in JavaScript

This is a simple Excel parser in JavaScript that allows you to read and parse data from Excel files. The parser uses a third-party library called `xlsx` to handle the Excel file processing.

## Installation

To use the Excel parser, you'll need to install the required dependencies using npm or yarn.

```bash
npm install xlsx
```

## Usage

1. Import the `xlsx` library in your JavaScript file:

```javascript
const XLSX = require('xlsx');
```

2. Load the Excel file:

```javascript
const workbook = XLSX.readFile('path/to/your/excel-file.xlsx');
```

3. Get the sheet name:

```javascript
const sheetName = workbook.SheetNames[0]; // Assuming the first sheet in the Excel file
```

4. Access the sheet data:

```javascript
const worksheet = workbook.Sheets[sheetName];
```

5. Parse the data:

```javascript
const parsedData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
```

Now, `parsedData` will contain the Excel data in a two-dimensional array. Each row of the Excel file will be an array element, and each cell's value will be an element within the row array.

Example:

```javascript
const data = [
  ['Name', 'Age', 'Email'],
  ['John', 30, 'john@example.com'],
  ['Alice', 25, 'alice@example.com'],
  ['Bob', 28, 'bob@example.com'],
];
```

You can access the data and perform any further processing as needed.

## Excel File Format

The parser supports various Excel file formats, including `.xlsx`, `.xlsb`, `.xlsm`, `.xls`, `.ods`, `.fods`, and `.csv`.

## License

This project is licensed under the [MIT License](LICENSE). You are free to use, modify, and distribute it as per the terms of the license.

## Acknowledgments

The Excel parser utilizes the `xlsx` library by SheetJS. Thanks to the SheetJS team for their great work!

For more detailed information and options regarding the `xlsx` library, please refer to the official [SheetJS documentation](https://sheetjs.com/).
