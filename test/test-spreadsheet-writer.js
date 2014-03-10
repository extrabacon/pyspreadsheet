var SpreadsheetReader = require('../lib/spreadsheet-reader');
var SpreadsheetWriter = require('../lib/spreadsheet-writer');

SpreadsheetWriter.prototype.saveAndRead = function (callback) {
    var filePath = this.filePath;
    this.save(function (err) {
        if (err) return callback(err);
        SpreadsheetReader.read(filePath, callback);
    });
};

describe.only('SpreadsheetWriter', function () {

    describe('#ctor(options)', function () {
        it('should only support xls and xlsx formats', function () {
            new SpreadsheetWriter({ format: 'xls' }).save();
            new SpreadsheetWriter({ format: 'xlsx' }).save();
            (function () {
                new SpreadsheetWriter({ format: 'something else' });
            }).should.throw(Error);
        });
        it('should emit "open" event', function (done) {
            var writer = new SpreadsheetWriter({ format: 'xls' });
            writer.on('open', function () {
                done();
            });
        });
    });

    describe('.addSheet(name)', function () {
        it('should write a simple file with multiple sheets', function (done) {
            var writer = new SpreadsheetWriter('test/output/multiple_sheets.xlsx');
            writer.addSheet('first').addSheet('second').addSheet('third');
            writer.saveAndRead(function (err, workbook) {
                if (err) return done(err);
                workbook.meta.sheets.should.have.lengthOf(3);
                workbook.meta.sheets.should.containDeep(['first', 'second', 'third']);
                done();
            });
        });
    });

    describe('.addFormat(id, properties)', function (done) {
        it('should write with a simple format', function (done) {
            var writer = new SpreadsheetWriter('test/output/simple_format.xlsx');
            writer.addFormat('my format', { font: { bold: true, color: 'red' } });
            writer.write(0, 0, 'Hello World!', 'my format');
            writer.saveAndRead(function (err, workbook) {
                if (err) return done(err);
                workbook.sheets[0].rows[0][0].value.should.be.exactly('Hello World!');
                done();
            });
        });
    });

    describe('.write(row, col, value)', function () {
        it('should write to the correct cells from indexed positions', function (done) {
            var writer = new SpreadsheetWriter('test/output/write2.xlsx');
            writer.write(0, 0, 1);
            writer.write(0, 2, 2);
            writer.write(1, 1, 3);
            writer.write(3, 3, 4);
            writer.saveAndRead(function (err, workbook) {
                if (err) return done(err);
                workbook.sheets[0].cell('A1').value.should.be.exactly(1);
                workbook.sheets[0].cell('C1').value.should.be.exactly(2);
                workbook.sheets[0].cell('B2').value.should.be.exactly(3);
                workbook.sheets[0].cell('D4').value.should.be.exactly(4);
                done();
            });
        });
        it('should write supported data types', function (done) {
            var writer = new SpreadsheetWriter('test/output/write3.xlsx');
            var number = 1934587.9812858;
            var string = '<html>some html <body>tags</body></html>';
            var date = new Date(2014, 0, 1, 0, 0, 0, 0);

            writer.write(0, 0, number);
            writer.write(0, 1, string);
            writer.write(0, 2, date);

            writer.saveAndRead(function (err, workbook) {
                if (err) return done(err);
                var row = workbook.sheets[0].rows[0];
                row[0].value.should.be.exactly(number);
                row[1].value.should.be.exactly(string);
                row[2].value.should.eql(date);
                done();
            });
        });
    });

    describe('.write(cell, value)', function () {
        it('should write to the correct cell from its address', function (done) {
            var writer = new SpreadsheetWriter('test/output/write1.xlsx');
            writer.write('A1', 1);
            writer.write('C1', 2);
            writer.write('B2', 3);
            writer.write('D4', 4);
            writer.saveAndRead(function (err, workbook) {
                if (err) return done(err);
                workbook.sheets[0].cell('A1').value.should.be.exactly(1);
                workbook.sheets[0].cell('C1').value.should.be.exactly(2);
                workbook.sheets[0].cell('B2').value.should.be.exactly(3);
                workbook.sheets[0].cell('D4').value.should.be.exactly(4);
                done();
            });
        });
    });

    describe('.save(callback)', function () {
        it('should save a hello world file', function (done) {
            var writer = new SpreadsheetWriter('test/output/helloworld.xlsx');
            writer.write(0, 0, 'Hello World!');
            writer.saveAndRead(function (err, workbook) {
                if (err) return done(err);
                workbook.sheets[0].rows[0][0].value.should.be.exactly('Hello World!');
                done();
            });
        });
        it('should emit a "close" event', function (done) {
            var writer = new SpreadsheetWriter('test/output/helloworld.xlsx');
            writer.write(0, 0, 'Hello World!');
            writer.on('close', function () {
                done();
            }).save();
        });
    });

});
