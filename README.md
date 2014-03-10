# PySpreadsheet

A high-performance spreadsheet library for Node, powered by Python open source libraries. PySpreadsheet can be used
to read and write Excel files in both XLS and XLSX formats.

IMPORTANT: PySpreadsheet is a work in progress. This is not a stable release.

## Features

+ Faster and more memory efficient than most JS-only alternatives
+ Uses child processes for isolation and parallelization (will not leak the Node process)
+ Can stream large files with a Node stream-like API
+ Can write large files with a simple forward-only API
+ Support for both XLS and XLSX formats
+ Native integration with Javascript objects

## Limitations

+ Reading does not parse formats, only data
+ Basic XLS writing capabilities
+ Only the forward-only write API is available

## Installation
```bash
npm install pyspreadsheet
```

Python dependencies are installed automatically by downloading the latest version from their repositories. These
dependencies are:

+ [xlrd](http://github.com/python-excel/xlrd), by Python Excel
+ [xlwt](http://github.com/python-excel/xlwt), also by Python Excel
+ [XlsxWriter](http://github.com/jmcnamara/XlsxWriter), by John McNamara

## Documentation

### Reading a file with the `SpreadsheetReader` class

The `SpreadsheetReader` class allows you to read spreadsheet files. It can be used as a method for reading an entire
file into memory, or as a stream-like object for streaming the data.

#### Reading a file into memory

If you want to read a file into memory and have the entire contents of the file in an object structure, simply use the
`read` method. There is no need to create an instance for this scenario.

* `#read(path, options, callback)` - reads an entire spreadsheet file into memory
  * `path` - the path of the file to read, accepting arrays for reading multiple files at once
  * `options` - the reading options (optional)
    * `meta` - load only workbook metadata, without iterating on rows
    * `sheet` || `sheets` - load sheet(s) selectively, either by name or by index, also accepting arrays
    * `maxRows` - the maximum number of rows to load per sheet
  * `callback(err, workbook)` - the callback function to invoke when the operation is completed
    * `err` - the error, if any
    * `workbook` - the parsed workbook instance, will be an array if `path` was also an array
      * `file` - the file used to open the workbook
      * `meta` - the metadata for this workbook
        * `user` - the owner of the file
        * `sheets` - an array of strings containing the name of sheets (available without any iteration)
      * `sheets[]` - the array of Sheet objects that were loaded
        * `index` - the ordinal position of the sheet within the workbook
        * `name` - the name of the sheet
        * `bounds` - an object specifying the data range for the sheet
          * `rows` - the total number of rows in the sheet
          * `columns` - the total number of columns in the sheet
        * `visibility` - the sheet visibility - possible values are `visible`, `hidden` and `very hidden`
        * `rows[]` - the array of rows that were loaded - rows are arrays of cells
          * `row` - the ordinal row number
          * `column` - the ordinal column number
          * `address` - the cell address ("A1", "B12", etc.)
          * `value` - the cell value, which can be of the following types:
            * `Number` - for numeric values
            * `Date` - for cells formatted as dates
            * `Error` - for cells with errors, such as #NAME?
            * `Boolean` - for cells formatted as booleans
            * `String` - for anything else
        * `cell(address)` - a function returning the cell at a specific location (ex: B12), same as accessing the `rows` array

Example of reading an entire file with `read`:

```javascript
var SpreadsheetReader = require('pyspreadsheet').SpreadsheetReader;

SpreadsheetReader.read('input.xlsx', function (err, workbook) {
  // Iterate on sheets
  workbook.sheets.forEach(function (sheet) {
    console.log('sheet: ' + sheet.name);
    // Iterate on rows
    sheet.rows.forEach(function (row) {
      // Iterate on cells
      row.forEach(function (cell) {
        console.log(cell.address + ': ' + cell.value);
      });
    });
  });
});
```

#### Reading a file from the stream interface

If you need to open a large file, you may want to stream the data instead to avoid loading the entire file contents into
memory. The `SpreadsheetReader` class exposes a stream-like interface that will allow you to read the data
progressively.

Here is an example:

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
  throw err;
});
```

First, create an instance of `SpreadsheetReader` using the constructor.

* `#ctor(path, options)` - creates a new instance of SpreadsheetReader
  * `path` - the path of the file(s) to read, accepting arrays for reading multiple files
  * `options` - the reading options (optional)
    * `meta` - load only workbook metadata, without iterating on rows
    * `sheet` || `sheets` - load sheet(s) selectively, either by name or by index, also accepting arrays
    * `maxRows` - the maximum number of rows to load per sheet

The constructor will open the file immediately for reading, so make sure you listen to the appropriate events. The
events emitted by `SpreadsheetReader` are:

* `open(workbook)` - fires when a workbook has been opened, before any iteration on sheet data
  * `workbook` - the workbook object
    * `file` - the file used to open the workbook
    * `meta` - the metadata for this workbook
      * `user` - the owner of the file
      * `sheets` - an array of strings containing the name of sheets (available without any iteration)

* `data(workbook, sheet, rows)` - fires repeatedly as data is being read from the file
  * `workbook` - the current workbook object
  * `sheet` - the current sheet object
  * `rows` - the current batch of rows

* `error(err)` - fires every time an error is encountered while parsing the file, the process is stopped only if a fatal
error is encountered
  * `err` - the error object

* `close()` - fires only once, after all files and data have been read

### Writing a file

Use the `SpreadsheetWriter` class to write a new file. It can only write new files, it cannot change an existing file.

Example:

```javascript
var SpreadsheetWriter = require('pyspreadsheet').SpreadsheetWriter;
var writer = new SpreadsheetWriter('examples/output.xlsx');

// write a string at cell A1
writer.write(0, 0, 'hello world!');

writer.save(function (err) {
  if (err) throw err;
  console.log('file saved!');
});
```

#### The `SpreadsheetWriter` class

* `#ctor(path, options)` - creates a new instance of `SpreadsheetWriter`
  * `path` - the path where to write the spreadsheet file
  * `options` - the workbook options (optional)
    * `format` - the workbook format, "xlsx" for OpenOffice file or "xls" for legacy binary format
    * `defaultDateFormat` - the default number format to apply when writing a date - default : "yyyy-mm-dd"
    * `properties` - workbook properties
      * `title`
      * `subject`
      * `author`
      * `manager`
      * `company`
      * `category`
      * `keywords`
      * `comments`
      * `status`

* `addSheet(name, options)` - adds a new sheet to the workbook
  * `name` - the name of the sheet (optional)
  * `options` - the worksheet options (optional)
    * `hidden`
    * `activated`
    * `selected`
    * `rightToLeft`
    * `hideZeroValues`
    * `selection`

* `addFormat(id, format)` - defines a reusable format that can be used with `write`
  * `id` - the format ID to use with `write`
  * `format` - the format object
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
      * `top`|`left`|`right`|`bottom`
        * `style`
        * `color`

* `write(row|cell, col, data, format)` - writes data into the current sheet
  * `row`|`cell` - the ordinal row index, or the cell location (ex: A1, B12, etc.)
  * `col` - the ordinal column index - ignore if using a cell location
  * `data` - the data to write - can be a primitive, an array for writing a row, or a 2-D array for writing multiple rows
  * `format` - the format ID or format object to use (optional)

* `append(data, format)` - appends data to the current sheet
  * `data` - the data to write - can be a primitive, an array for writing a row, or a 2-D array for writing multiple rows
  * `format` - the format ID or format object to use (optional)

* `save(callback)` - saves the workbook file
  * `callback(err)` - the callback function to invoke when the save operation is completed
    * `err` - the error, if any

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
+ wget or curl (installation script)

## License

MIT license
