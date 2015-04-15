# PySpreadsheet
A high-performance spreadsheet library for Node, powered by Python open source libraries. PySpreadsheet can be used
to read and write Excel files in both XLS and XLSX formats.

IMPORTANT: PySpreadsheet is a work in progress. This is not a stable release.

## Features

+ Faster and more memory efficient than most JS-only alternatives
+ Uses child processes for isolation and parallelization (will not leak the Node process)
+ Support for both XLS and XLSX formats
+ Can stream large files with a familiar API
+ Native integration with Javascript objects

## Limitations

+ Reading does not parse formats, only data
+ Cannot edit existing files
+ Basic XLS writing capabilities
+ Incomplete API

## Installation
```bash
npm install pyspreadsheet
```

Python dependencies are installed automatically by downloading the latest version from their repositories. These dependencies are:

+ [xlrd](http://github.com/python-excel/xlrd) and [xlwt](http://github.com/python-excel/xlwt), by Python Excel
+ [XlsxWriter](http://github.com/jmcnamara/XlsxWriter), by John McNamara

## Documentation

### Reading a file with the `SpreadsheetReader` class

Use the `SpreadsheetReader` class to read spreadsheet files. It can be used for reading an entire file into memory or as a stream-like object.

#### Reading a file into memory

Reading a file into memory will output the entire contents of the file in an easy-to-use structure.

```js
var SpreadsheetReader = require('pyspreadsheet').SpreadsheetReader;

SpreadsheetReader.read('input.xlsx', function (err, workbook) {
  // Iterate on sheets
  workbook.sheets.forEach(function (sheet) {
    console.log('sheet: %s', sheet.name);
    // Iterate on rows
    sheet.rows.forEach(function (row) {
      // Iterate on cells
      row.forEach(function (cell) {
        console.log('%s: %s', cell.address, cell.value);
      });
    });
  });
});
```

#### Streaming a large file

You can use `SpreadsheetReader` just like a Node readable stream. This is the preferred method for reading larger files.

```javascript
var SpreadsheetReader = require('pyspreadsheet').SpreadsheetReader;
var reader = new SpreadsheetReader('examples/sample.xlsx');

reader.on('open', function (workbook) {
  // file is open
  console.log('opened ' + workbook.file);
}).on('data', function (data) {
  // data is being received
  console.log('buffer contains %d rows from sheet "%s"', data.rows.length, data.sheet.name);
}).on('close', function () {
  // file is now closed
  console.log('file closed');
}).on('error', function (err) {
  // got an error
  throw err;
});
```

### Writing a file with the `SpreadsheetWriter` class

Use the `SpreadsheetWriter` class to write files. It can only write new files, it cannot edit existing files.

```js
var SpreadsheetWriter = require('pyspreadsheet').SpreadsheetWriter;
var writer = new SpreadsheetWriter('examples/output.xlsx');

// write a string at cell A1
writer.write(0, 0, 'hello world!');

writer.save(function (err) {
  if (err) throw err;
  console.log('file saved!');
});
```

#### Adding sheets

Use the `addSheet` method to add a new sheet. If data is written without adding a sheet first, a default "Sheet1" is automatically added.

```js
var SpreadsheetWriter = require('pyspreadsheet').SpreadsheetWriter;
var writer = new SpreadsheetWriter('examples/output.xlsx');
writer.addSheet('my sheet').write(0, 0, 'hello');
```

#### Writing data

Use the `write` method to write data to the designated cell. All Javascript built-in types are supported.

```js
writer.write(0, 0, 'hello world!');
```

##### Formulas

Formulas are parsed from strings. To write formulas, just prepend your string value with "=".

```js
writer.write(2, 0, '=A1+A2');
```

Note that values calculated from formulas cannot be obtained until the file has been opened once with a spreadsheet client (like Microsoft Excel).

##### Writing multiple cells

Use arrays to write multiple cells at once horizontally.

```js
writer.write(0, 0, ['a', 'b', 'c', 'd', 'e']);
```

Use two-dimensional arrays to write multiple rows at once.

```js
writer.write(0, 0, [
  ['a', 'b', 'c', 'd', 'e'],
  ['1', '2', '3', '4', '5'],
]);
```

Cells can be merged by using the `merge_range` method.

```js
writer.merge_range('B2:E5', "merge_range Test");
```

### Formatting

Cells can be formatted by specifying format properties.

```js
writer.write(0, 0, 'hello', {
  font: {
    name: 'Calibri',
    size: 12
  }
});
```

Formats can also be reused by using the `addFormat` method.

```js
writer.addFormat('title', {
  font: { bold: true, color: '#ffffff' },
  fill: '#000000'
});

writer.write(0, 0, ['heading 1', 'heading 2', 'heading 3'], 'title');
```

Set a row height

```js
writer.set_row(0,{height:100});
```

## API Reference

### SpreadsheetReader

The `SpreadsheetReader` is used to read spreadsheet files from various formats.

#### #ctor(path, options)

Creates a new `SpreadsheetReader` instance.

* `path` - the path of the file to read, also accepting arrays for reading multiple files at once
* `options` - the reading options (optional)
  * `meta` - load only workbook metadata, without iterating on rows
  * `sheet` || `sheets` - load sheet(s) selectively, either by name or by index
  * `maxRows` - the maximum number of rows to load per sheet
  * `bufferSize` - the maximum number of rows to accumulate in the buffer (default: 20)

#### #read(path, options, callback)

Reads an entire file into memory.

* `path` - the path of the file to read, also accepting arrays for reading multiple files at once
* `options` - the reading options (optional)
  * `meta` - load only workbook metadata, without iterating on rows
  * `sheet` || `sheets` - load sheet(s) selectively, either by name or by index
  * `maxRows` - the maximum number of rows to load per sheet
* `callback(err, workbook)` - the callback function to invoke when the operation has completed
  * `err` - the error, if any
  * `workbook` - the parsed workbook instance, will be an array if `path` was also an array
    * `file` - the file used to open the workbook
    * `meta` - the metadata for this workbook
      * `user` - the owner of the file
      * `sheets` - an array of strings containing the name of sheets (available without iteration)
    * `sheets` - the array of Sheet objects that were loaded
      * `index` - the ordinal position of the sheet within the workbook
      * `name` - the name of the sheet
      * `bounds` - the data range for the sheet
        * `rows` - the largest number of rows in the sheet
        * `columns` - the largest number of columns in the sheet
      * `visibility` - the sheet visibility, possible values are `visible`, `hidden` and `very hidden`
      * `rows` - the array of rows that were loaded - rows are arrays of cells
        * `row` - the ordinal row number
        * `column` - the ordinal column number
        * `address` - the cell address ("A1", "B12", etc.)
        * `value` - the cell value, which can be of the following types:
          * `Number` - for numeric values
          * `Date` - for cells formatted as dates
          * `Error` - for cells with errors (such as #NAME?)
          * `Boolean` - for cells formatted as booleans
          * `String` - for anything else
      * `cell(address)` - a function returning the cell at a specific location (ex: B12), same as accessing the `rows` array

#### Event: 'open'

Emitted when a workbook file is open. The data included with this event includes:

* `file` - the file used to open the workbook
* `meta` - the metadata for this workbook
  * `user` - the owner of the file
  * `sheets` - an array of strings containing the name of sheets (available without iteration)

This event can be emitted more than once if multiple files are being read.

#### Event: 'data'

Emitted as rows are being read from the file. The data for this event consists of:

* `sheet` - the currently open sheet
  * `index` - the ordinal position of the sheet within the workbook
  * `name` - the sheet name
  * `bounds` - the data range for the sheet
    * `rows` - the largest number of rows in the sheet
    * `columns` - the largest number of columns in the sheet
  * `visibility` - the sheet visibility, possible values are `visible`, `hidden` and `very hidden`
* `rows` - the array of rows that were loaded (number of rows returned depend on the buffer size)
  * `row` - the ordinal row number
  * `column` - the ordinal column number
  * `address` - the cell address ("A1", "B12", etc.)
  * `value` - the cell value, which can be of the following types:
    * `Number` - for numeric values
    * `Date` - for cells formatted as dates
    * `Error` - for cells with errors, such as #NAME?
    * `Boolean` - for cells formatted as booleans
    * `String` - for anything else

#### Event: 'close'

Emitted when a workbook file is closed. This event can be emitted more than once if multiple files are being read.

#### Event: 'error'

Emitted when an error is encountered.

### SpreadsheetWriter

The `SpreadsheetWriter` is used to write spreadsheet files into various formats. All writer methods return the same instance, so feel free to chain your calls.

#### #ctor(path, options)

Creates a new `SpreadsheetWriter` instance.

* `path` - the path of the file to write
* `options` - the writer options (optional)
  * `format` - the file format
  * `defaultDateFormat` - the default date format (only for XLSX files)
  * `properties` - the workbook properties
    * `title`
    * `subject`
    * `author`
    * `manager`
    * `company`
    * `category`
    * `keywords`
    * `comments`
    * `status`

#### .addSheet(name, options)

Adds a new sheet to the workbook.

* `name` - the sheet name (must be unique within the workbook)
* `options` - the sheet options
  * `hidden`
  * `activated`
  * `selected`
  * `rightToLeft`
  * `hideZeroValues`
  * `selection`

#### .activateSheet(sheet)

Activates a previously added sheet.

* `sheet` - the sheet name or index

#### .addFormat(name, properties)

Registers a reusable format.

* `name` - the format name
* `properties` - the formatting properties
  * `font`
    * `name`
    * `size`
    * `color`
    * `bold`
    * `italic`
    * `underline`
    * `strikeout`
    * `superscript`
    * `subscript`
  * `numberFormat`
  * `locked`
  * `hidden`
  * `alignment`
  * `rotation`
  * `indent`
  * `shrinkToFit`
  * `justifyLastText`
  * `fill`
    * `pattern`
    * `backgroundColor`
    * `foregroundColor`
  * `borders`
    * `top` | `left` | `right` | `bottom`
      * `style`
      * `color`

#### .set_row(row, options)

Set a row properties.

* `row` - Row index
* `options` - the row properties [more info](http://xlsxwriter.readthedocs.org/worksheet.html#set_row)
  * `height`
  * `format`
  * `options`

#### .merge_range(range, data, format)

Merge cells.

* `range` - merged range
* `data` - value of merged cell
* `format` - the format name

#### .write(row, column, data, format)

Writes data to the specified cell with an optional format.

* `row` - the row index
* `column` - the column index
* `data` - the value to write, supported types are: String, Number, Date, Boolean and Array
* `format` - the format name or properties to use (optional)

#### .write(cell, data, format)

Same as previous, except cell is a string such as "A1", "B2" or "1,1"

#### .append(data, format)

Appends data at the first column of the next row. The next row is determined by the last call to `write`.

* `data` - the value to write (use arrays to write multiple cells at once)
* `format` - the format name or properties to use (optional)

#### .save(callback)

Save and close the resulting workbook file.

* `callback(err)` - the callback function to invoke when the file is saved
  * `err` - the error, if any

#### Event: open

Emitted when the file is open.

#### Event: close

Emitted when the file is closed.

#### Event: error

Emitted when an error occurs.

## Compatibility

+ Tested with Node 0.10.x
+ Tested on Mac OS X 10.8
+ Tested on Ubuntu Linux 12.04
+ Tested on Heroku

## Dependencies

+ Python version 2.7+
+ [xlrd](http://www.python-excel.org/) version 0.7.4+
+ [xlwt](http://www.python-excel.org/) version 0.7.5+
+ [XlsxWriter](http://xlsxwriter.readthedocs.org/en/latest/index.html) version 0.3.6+
+ bash (installation script)
+ git (installation script)

## License

The MIT License (MIT)

Copyright (c) 2013 Nicolas Mercier

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
