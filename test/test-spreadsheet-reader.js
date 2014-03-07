var assert = require('assert');
var async = require('async');
var SpreadsheetReader = require('../lib/spreadsheet-reader');

var sampleFiles = [
    // OO XML workbook (Excel 2003 and later)
    'test/input/sample.xlsx',
    // Legacy binary format (before Excel 2003)
    'test/input/sample.xls'
];

describe('SpreadsheetReader', function () {
    describe('#ctor(file)', function () {

        it('should emit a single "open" event when file is ready', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                new SpreadsheetReader(file).on('open', function (workbook) {
                    workbook.should.be.an.Object.and.have.property('file', file);
                    return next();
                }).on('error', function (err) {
                    throw err;
                });
            }, done);
        });

        it('should emit "data" events as data is being received', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                var events = [];
                var total = {};

                new SpreadsheetReader(file).on('open', function (workbook) {
                    workbook.should.have.property('file', file);
                    workbook.should.have.property('meta');
                }).on('data', function (data) {
                    events.push(data);
                }).on('error', function (err) {
                    throw err;
                }).on('close', function () {
                    events.length.should.be.above(0);
                    events.forEach(function (e) {
                        e.should.have.property('workbook');
                        e.workbook.should.have.property('file', file);
                        e.should.have.property('sheet');
                        e.should.have.property('rows').and.be.an.Array;
                        e.rows.length.should.be.above(0);
                        total[e.sheet.name] = (total[e.sheet.name] || 0) + e.rows.length;
                    });

                    total.should.have.property('Sheet1', 51);
                    total.should.have.property('second sheet', 1384);

                    return next();
                });

            }, done);
        });

    });

    describe('#read(file, callback)', function () {

        it('should load a workbook/sheet/row/cell structure', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, function (err, workbook) {
                    if (err) throw err;

                    workbook.should.have.property('file', file);

                    workbook.should.have.property('meta');
                    workbook.meta.should.have.property('user', 'Nicolas Mercier-Gaboury');
                    workbook.meta.should.have.property('sheets');
                    workbook.meta.sheets.should.be.an.Array.and.have.lengthOf(2);
                    workbook.meta.sheets.should.containDeep(['Sheet1', 'second sheet']);

                    workbook.should.have.property('sheets');
                    workbook.sheets.should.have.lengthOf(workbook.meta.sheets.length);

                    workbook.sheets.forEach(function (sheet, index) {
                        sheet.should.have.property('index', index);
                        sheet.should.have.property('name', index === 0 ? 'Sheet1' : 'second sheet');
                        sheet.should.have.property('bounds');
                        sheet.bounds.should.be.an.Object;
                        sheet.bounds.should.have.properties('rows', 'columns');
                        sheet.should.have.property('visibility', 'visible');
                        sheet.should.have.property('rows');
                        sheet.rows.should.have.lengthOf(sheet.bounds.rows);

                        if (index === 0) {
                            sheet.bounds.should.have.property('columns', 26);
                            sheet.bounds.should.have.property('rows', 51);
                        } else if (index === 1) {
                            sheet.bounds.should.have.property('columns', 3);
                            sheet.bounds.should.have.property('rows', 1384);
                        }

                        sheet.rows.forEach(function (row) {
                            row.should.be.an.Array.and.have.lengthOf(sheet.bounds.columns);
                            row.forEach(function (cell) {
                                cell.should.have.properties('row', 'column', 'address', 'value');
                            });
                        });
                    });

                    return next();
                });
            }, done);
        });

        it('should parse cell addresses', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, function (err, workbook) {
                    if (err) throw err;

                    var rows = workbook.sheets[0].rows;
                    rows[0][0].should.have.property('address', 'A1');
                    rows[0][4].should.have.property('address', 'E1');
                    rows[5][0].should.have.property('address', 'A6');
                    rows[5][4].should.have.property('address', 'E6');
                    rows[50][0].should.have.property('address', 'A51');
                    rows[50][4].should.have.property('address', 'E51');

                    return next();
                });
            }, done);
        });

        it('should parse numeric values', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, function (err, workbook) {

                    if (err) throw err;

                    var row1 = workbook.sheets[0].rows[1];
                    var row2 = workbook.sheets[0].rows[2];
                    var row3 = workbook.sheets[0].rows[3];
                    var row4 = workbook.sheets[0].rows[4];
                    var row5 = workbook.sheets[0].rows[5];

                    [row1, row2, row3, row4, row5].forEach(function (row) {
                        row[0].value.should.be.a.Number;
                        row[1].value.should.be.a.Number;
                    });

                    row1[0].value.should.be.exactly(1);
                    row1[1].value.should.be.exactly(1.0001);
                    row2[0].value.should.be.exactly(2);
                    row2[1].value.should.be.exactly(2);
                    row3[0].value.should.be.exactly(3);
                    row3[1].value.should.be.exactly(3.00967676764465);
                    row4[0].value.should.be.exactly(4);
                    row4[1].value.should.be.exactly(0);
                    row5[0].value.should.be.exactly(5);
                    row5[1].value.should.be.exactly(5.00005);

                    return next();
                });
            }, done);
        });

        it('should parse string values', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, function (err, workbook) {
                    if (err) throw err;

                    var cell1 = workbook.sheets[0].rows[1][2];
                    var cell2 = workbook.sheets[0].rows[2][2];
                    var cell3 = workbook.sheets[0].rows[3][2];
                    var cell4 = workbook.sheets[0].rows[4][2];
                    var cell5 = workbook.sheets[0].rows[5][2];

                    [cell1, cell2, cell3, cell4, cell5].forEach(function (cell) {
                        cell.value.should.be.a.String;
                    });

                    cell1.value.should.be.exactly('Some text');
                    cell2.value.should.be.exactly('{ "property": "value" }');
                    cell3.value.should.be.exactly('ÉéÀàçÇùÙ');
                    cell4.value.should.be.exactly('some "quoted" text');
                    cell5.value.should.be.exactly('more \'quoted\' "text"');

                    return next();
                });
            }, done);
        });

        it('should parse date values', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, function (err, workbook) {
                    if (err) throw err;

                    var row1 = workbook.sheets[0].rows[1];
                    var row2 = workbook.sheets[0].rows[2];
                    var row3 = workbook.sheets[0].rows[3];
                    var row4 = workbook.sheets[0].rows[4];
                    var row5 = workbook.sheets[0].rows[5];

                    [row1, row2, row3, row4, row5].forEach(function (row) {
                        row[3].value.should.be.a.Date;
                        row[4].value.should.be.a.Date;
                    });

                    row1[3].value.should.eql(new Date(2013, 0, 1, 0, 0, 0));
                    row1[4].value.should.eql(new Date(2013, 0, 1, 12, 54, 21));
                    row2[3].value.should.eql(new Date(2013, 0, 2, 0, 0, 0));
                    //row2[4].value.should.eql(new Date(0, 0, 0, 0, 0, 34));
                    row3[4].value.should.eql(new Date(2013, 0, 3, 3, 45, 20));
                    row4[4].value.should.eql(new Date(2013, 0, 4, 0, 0, 0));
                    row5[4].value.should.eql(new Date(2013, 0, 5, 16, 0, 0));

                    return next();
                });
            }, done);
        });

        it('should parse empty cells as nulls', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, function (err, workbook) {
                    if (err) throw err;

                    workbook.sheets[0].rows.splice(1).forEach(function (row) {
                        assert.ok(row[5].value === null);
                    });

                    return next();
                });
            }, done);
        });

        it('should parse cells with errors', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, function (err, workbook) {
                    if (err) throw err;

                    var row1 = workbook.sheets[0].rows[1];
                    var row2 = workbook.sheets[0].rows[2];
                    var row3 = workbook.sheets[0].rows[3];

                    row1[6].value.should.be.an.Error.and.have.property('errorCode', '#DIV/0!');
                    row2[6].value.should.be.an.Error.and.have.property('errorCode', '#NAME?');
                    row3[6].value.should.be.an.Error.and.have.property('errorCode', '#VALUE!');

                    return next();
                });
            }, done);
        });

        it('should parse cells with booleans', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, function (err, workbook) {
                    if (err) throw err;

                    var row1 = workbook.sheets[0].rows[1];
                    var row2 = workbook.sheets[0].rows[2];

                    row1[7].value.should.be.true;
                    row2[7].value.should.be.false;

                    return next();
                });
            }, done);
        });

        it('should parse only the selected sheet by index', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, { sheet: 0 }, function (err, workbook) {
                    if (err) throw err;
                    workbook.sheets.should.be.an.Array.and.have.lengthOf(1);
                    workbook.sheets[0].should.have.properties({ index: 0, name: 'Sheet1' });
                    return next();
                });
            }, done);
        });

        it('should parse only the selected sheet by name', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, { sheet: 'Sheet1' }, function (err, workbook) {
                    if (err) throw err;
                    workbook.sheets.should.be.an.Array.and.have.lengthOf(1);
                    workbook.sheets[0].should.have.properties({ index: 0, name: 'Sheet1' });
                    return next();
                });
            }, done);
        });

        it('should parse only metadata', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, { meta: true }, function (err, workbook) {
                    if (err) throw err;
                    if (workbook.sheets) throw new Error('workbook.sheets should be undefined');
                    return next();
                });
            }, done);
        });

        it('should parse up to 10 rows on the first sheet', function (done) {
            async.eachSeries(sampleFiles, function (file, next) {
                SpreadsheetReader.read(file, { sheet: 0, maxRows: 10 }, function (err, workbook) {
                    if (err) throw err;
                    workbook.sheets[0].should.have.property('name', 'Sheet1');
                    workbook.sheets[0].rows.should.have.lengthOf(10);
                    return next();
                });
            }, done);
        });

        it('should fail if the file does not exist', function (done) {
            SpreadsheetReader.read('unknown.xlsx', function (err, workbook) {
                if (workbook) throw new Error('workbook should be undefined');
                err.should.be.an.Error.and.have.property('id', 'file_not_found');
                return done();
            });
        });

        it('should fail if the file is not a valid workbook', function (done) {
            SpreadsheetReader.read('package.json', function (err, workbook) {
                if (workbook) throw new Error('workbook should be undefined');
                err.should.be.an.Error.and.have.property('id', 'open_workbook_failed');
                return done();
            });
        });
    });

});
