# PySpreadsheet

A high-performance spreadsheet library for Node, powered by Python open source libraries. PySpreadsheet can be used
to read and write Excel files in both XLS and XLSX formats.

This module interfaces with a Python shell process to stream JSON fragments through stdio using a child process. The
Python shell uses multiple open source libraries to perform spreadsheet operations.

## Features

+ Faster and more memory efficient than most alternatives
+ Uses child processes for isolation and parallelization (will not leak the Node process)
+ Can stream large files (tested with 250K+ rows)
+ Can write large files with a simple forward-only API
+ Support for both XLS and XLSX formats
+ Native integration with Javascript objects

## Limitations

+ Requires Python 2.7+
+ Reading does not parse formats, only data
+ XLS writing is not yet implemented

## Installation
```bash
npm install pyspreadsheet
```

Python dependencies are installed automatically by downloading the latest version from their repositories. These
dependencies are:

+ [xlrd](http://github.com/python-excel/xlrd), by Python Excel
+ [xlwt](http://github.com/python-excel/xlwt), also by Python Excel
+ [XlsxWriter](http://github.com/jmcnamara/XlsxWriter), by John McNamara
+ [openpyxl](http://bitbucket.org/ericgazoni/openpyxl), by Eric Gazoni

## Documentation

### Reading a file

Files can be read entirely into a workbook object, containing sheets, rows and cells.

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

#### Return value

The returned object is either a workbook object, or an array of workbook objects if multiple files were specified as the
source. The workbook object contains a structure of sheets, rows and cells to represent the data.

The Workbook object contains the following members:

* `file` - the file used to open the workbook
* `meta` - metadata for this workbook
  * `user` - the owner of the file
  * `sheets` - an array of strings containing the name of sheets (available without any iteration)
* `sheets` - the array of Sheet objects that were loaded

Sheets can also be accessed by name. For example:
```javascript
var sheet1 = workbook.sheets['Sheet1'];
```

The Sheet object contains the following members:

* `index` - the ordinal position of the sheet within the workbook
* `name` - the name of the sheet
* `bounds` - an object specifying the data range for the sheet
  * `rows` - the total number of rows in the sheet
  * `columns` - the total number of columns in the sheet
* `visibility` - the sheet visibility - possible values are `visible`, `hidden` and `very hidden`
* `rows` - the array of rows that were loaded - rows are arrays of Cell objects
* `cell(address)` - a function returning the cell at a specific location (ex: B12)

The Cell object contains the following members:

* `row` - the ordinal row number
* `column` - the ordinal column number
* `address` - the cell address ("A1", "B12", etc.)
* `value` - the cell value

Cell values can be of the following types:

* `Number` - for numeric values
* `Date` - for cells formatted as dates
* `Error` - for cells with errors, such as #NAME?
* `Boolean` - for cells formatted as booleans
* `String` - for anything else

#### Streaming a large file

For large files, you may want to stream the data. The SpreadsheetReader is a stream-like object with events firing as
data is being read.

```javascript
var SpreadsheetReader = require('pyspreadsheet').SpreadsheetReader;

var reader = new SpreadsheetReader('input.xlsx');

reader.on('open', function (workbook) {
	console.log('successfully opened ' + workbook.file);
}).on('data', function (data) {

	var currentWorkbook = data.workbook,
		currentSheet = data.sheet,
		batchOfRows = data.rows;

	// TODO: handle streaming logic here

}).on('error', function (err) {
	// TODO: handle error here
}).on('close', function () {
	// TODO: finishing logic here
});
```

#### Events

The following events are fired from a SpreadsheetReader instance.

* `open` - fires when a workbook is opened (sheets are not available at this point)

  Arguments: the workbook object

* `data` - fires repeatedly as data is being read from the file

  Arguments: a data object containing the following:

  * `workbook`: the current workbook object
  * `sheet`: the current sheet object
  * `rows`: the current batch of rows

* `error` - fires every time an error is encountered while parsing the file, the process is stopped only if a fatal
error is encountered

  Arguments: the error object

* `close` - fires only once, after all files and data have been read

  Arguments: none

#### Options

An object can be passed along with the file path for additional options.

* `meta` - load only workbook metadata, without iterating on rows - `Boolean`
* `sheet` || `sheets` - load sheet(s) selectively, either by name or by index - `String`, `Number` or `Array`
* `maxRows` - the maximum number of rows to load per sheet - `Number`
* `debug` - log output from the xlrd-parser child process - `Boolean`

##### Examples:

Output sheet names without loading any data:

```javascript
xlrd.parse('myfile.xlsx', { meta: true }, function (err, workbook) {
  console.log(workbook.meta.sheets);
});
```

Load only the first 10 rows from the first sheet:

```javascript
xlrd.parse('myfile.xlsx', { sheet: 0, maxRows: 10 }, function (err, workbook) {
  // workbook will contain only the first sheet
});
```

Load only a sheet named "products":

```javascript
var stream = xlrd.stream('myfile.xlsx', { sheet: 'products' });
```

### Writing a file

Use the SpreadsheetWriter class to write a new file.

```javascript
var SpreadsheetWriter = require('../lib').SpreadsheetWriter;
var writer = new SpreadsheetWriter();

// write a string at cell A1
writer.write(0, 0, 'hello world!');

writer.save('examples/output.xlsx', function (err) {
	if (err) throw err;
	console.log('file saved!');
});
```

More to come...

## Compatibility

+ Tested with Node 0.10.x
+ Tested on Mac OS X 10.8
+ Tested on Ubuntu Linux 12.04 (requires prior installation of curl: apt-get install curl)

## Dependencies

+ Python version 2.7+
+ [xlrd](http://www.python-excel.org/) version 0.7.4+
+ [xlwt](http://www.python-excel.org/) version 0.7.5+
+ [XlsxWriter](http://xlsxwriter.readthedocs.org/en/latest/index.html) version 0.3.6+
+ [openpyxl](http://openpyxl.readthedocs.org/en/latest/) version 1.6.2+
+ underscore
+ bash (installation script)
+ curl (installation script)

## License

MIT license
