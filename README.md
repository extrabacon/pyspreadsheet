# xlrd-parser

High performance XLS/XLSX parser based on the xlrd library from [www.python-excel.org](www.python-excel.org) for
efficiently reading Excel files from all versions and all sizes.

This module interfaces with a Python shell to stream JSON fragments from stdout. It is not a port of xlrd from Python to
Javascript (which is surely possible) and it does not use native bindings.

## Features

+ Much faster and more memory efficient than most alternatives
+ Support for both XLS and XLSX formats
+ Can read multiple sheets
+ Can efficiently stream large files (tested with 250K+ rows)

## Documentation

### Installation
```bash
npm install xlrd-parser
```

### Parsing a file

Parsing a file loads the entire file into an object structure composed of a workbook, sheets, rows and cells.

```javascript
var xlrd = require('xlrd');

xlrd.parse('myfile.xlsx', function (err, workbook) {
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

Cell values are accurately parsed as native strings, numbers and dates.

For more details on the API, see the included unit tests.

### Streaming a large file

For large files, you may want to stream the data. The stream method returns a familiar EventEmitter instance.

```javascript
var xlrd = require('xlrd');

xlrd.stream('myfile.xlsx').on('open', function (workbook) {
	console.log('successfully opened ' + workbook.file);
}).on('data', function (data) {

	var currentWorkbook = data.workbook,
		currentSheet = data.sheet,
		rows = data.rows;

	// TODO: handle streaming logic here

}).on('error', function (err) {
	// TODO: handle error here
}).on('close', function () {
	// TODO: finishing logic here
});
```

## Compatibility

+ Tested with Node 0.8
+ Tested on Mac OS X 10.8
+ Tested on Ubuntu Linux 12.04 (requires prior installation of curl: apt-get install curl)

## Dependencies

+ Python version 2.6+
+ xlrd version 0.7.4+
+ underscore.js
+ bash (installation script)
+ curl (installation script)

Windows platform is not yet supported. I will accept contributions for an alternate install script that will also work
on Windows.

## Limitations

+ Cannot parse file selectively (will be addressed in a future release)
+ Does not parse formatting info (might be addressed in a future release)

## Thanks

Many thanks to the authors of the xlrd library ([here](http://github.com/python-excel/xlrd)). It is the best and most
efficient open-source library I could find.

## License

	Portions copyright © 2005-2009, Stephen John Machin, Lingfo Pty Ltd
	All rights reserved.

	Redistribution and use in source and binary forms, with or without
	modification, are permitted provided that the following conditions are met:

	1. Redistributions of source code must retain the above copyright notice,
	this list of conditions and the following disclaimer.

	2. Redistributions in binary form must reproduce the above copyright notice,
	this list of conditions and the following disclaimer in the documentation
	and/or other materials provided with the distribution.

	3. None of the names of Stephen John Machin, Lingfo Pty Ltd and any
	contributors may be used to endorse or promote products derived from this
	software without specific prior written permission.

	THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
	AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
	THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
	PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS
	BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
	CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
	SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
	INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
	CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
	ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
	THE POSSIBILITY OF SUCH DAMAGE.

	/*-
	 * Copyright (c) 2001 David Giffin.
	 * All rights reserved.
	 *
	 * Based on the the Java version: Andrew Khan Copyright (c) 2000.
	 *
	 *
	 * Redistribution and use in source and binary forms, with or without
	 * modification, are permitted provided that the following conditions
	 * are met:
	 *
	 * 1. Redistributions of source code must retain the above copyright
	 *    notice, this list of conditions and the following disclaimer.
	 *
	 * 2. Redistributions in binary form must reproduce the above copyright
	 *    notice, this list of conditions and the following disclaimer in
	 *    the documentation and/or other materials provided with the
	 *    distribution.
	 *
	 * 3. All advertising materials mentioning features or use of this
	 *    software must display the following acknowledgment:
	 *    "This product includes software developed by
	 *     David Giffin <david@giffin.org>."
	 *
	 * 4. Redistributions of any form whatsoever must retain the following
	 *    acknowledgment:
	 *    "This product includes software developed by
	 *     David Giffin <david@giffin.org>."
	 *
	 * THIS SOFTWARE IS PROVIDED BY DAVID GIFFIN ``AS IS'' AND ANY
	 * EXPRESSED OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
	 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
	 * PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL DAVID GIFFIN OR
	 * ITS CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
	 * SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
	 * NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
	 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
	 * HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
	 * STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
	 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
	 * OF THE POSSIBILITY OF SUCH DAMAGE.
	 */
