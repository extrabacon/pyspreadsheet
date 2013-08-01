var async = require('async'),
	expect = require('chai').expect,
	pyspreadsheet = require('../lib'),
	sampleFiles = [
		// OO XML workbook (Excel 2003 and later)
		'tests/sample.xlsx',
		// Legacy binary format (before Excel 2003)
		'tests/sample.xls'
	];

describe('pyspreadsheet', function () {

	describe('read()', function () {

		it('should emit a single "open" event when file is ready', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file).on('open', function (workbook) {
					expect(workbook).to.be.an('object').and.have.property('file', file);
					return next();
				}).on('error', function (err) {
					throw err;
				});
			}, done);
		});

		it('should emit "data" events as data is being received', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {

				var events = [],
					total = {};

				pyspreadsheet.read(file).on('open', function (workbook) {
					expect(workbook).not.be.null;
					expect(workbook.file).to.equal(file);
					expect(workbook.meta).not.be.null;
					expect(workbook.sheets).to.be.undefined;
				}).on('data',function (data) {
					events.push(data);
				}).on('error',function (err) {
					throw err;
				}).on('close', function () {

					expect(events).to.have.length.above(0);
					expect(events).to.have.deep.property('[0].workbook').and.have.property('file', file);
					expect(events).to.have.deep.property('[0].sheet');
					expect(events).to.have.deep.property('[0].rows').and.have.length.above(0);

					events.forEach(function (data) {
						total[data.sheet.name] = (total[data.sheet.name] || 0) + data.rows.length;
					});

					expect(total['Sheet1']).to.equal(51);
					expect(total['second sheet']).to.equal(1384);

					return next();
				});

			}, done);
		});

	});

	describe('read(callback)', function () {

		it('should load a workbook/sheet/row/cell structure', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, function (err, workbook) {

					if (err) throw err;

					expect(workbook).not.be.null;
					expect(workbook.file).to.equal(file);
					expect(workbook.meta.user).to.equal('Nicolas Mercier-Gaboury');
					expect(workbook.meta.sheets.length).to.equal(2);
					expect(workbook.meta.sheets).to.contain('Sheet1').and.to.contain('second sheet');
					expect(workbook.sheets).to.be.an('array').and.have.length(workbook.meta.sheets.length);

					workbook.sheets.forEach(function (sheet, index) {
						expect(sheet).to.have.property('index', index);
						expect(sheet).to.have.property('name', index === 0 ? 'Sheet1' : 'second sheet');
						expect(sheet).to.have.property('bounds').and.to.be.an('object');
						expect(sheet).to.have.property('visibility', 'visible');
						expect(sheet).to.have.property('rows').and.have.length(sheet.bounds.rows);
						//expect(sheet).to.equal(workbook.sheets[index]).and.to.equal(workbook.sheets[sheet.name]);

						if (index === 0) {
							expect(sheet.bounds).to.have.property('columns', 8);
							expect(sheet.bounds).to.have.property('rows', 51);
						} else if (index === 1) {
							expect(sheet.bounds).to.have.property('columns', 3);
							expect(sheet.bounds).to.have.property('rows', 1384);
						}

						sheet.rows.forEach(function (row) {
							expect(row).to.be.an('array').and.have.length(sheet.bounds.columns);
							expect(row).to.have.deep.property('[0].row');
							expect(row).to.have.deep.property('[0].column');
							expect(row).to.have.deep.property('[0].address');
							expect(row).to.have.deep.property('[0].value');
						});
					});

					return next();
				});
			}, done);
		});

		it('should parse cell addresses', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, function (err, workbook) {

					if (err) throw err;

					var rows = workbook.sheets[0].rows;
					expect(rows[0][0]).to.have.property('address', 'A1');
					expect(rows[0][4]).to.have.property('address', 'E1');
					expect(rows[5][0]).to.have.property('address', 'A6');
					expect(rows[5][4]).to.have.property('address', 'E6');
					expect(rows[50][0]).to.have.property('address', 'A51');
					expect(rows[50][4]).to.have.property('address', 'E51');

					return next();
				});
			}, done);
		});

		it('should parse numeric values', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, function (err, workbook) {

					if (err) throw err;

					var row1 = workbook.sheets[0].rows[1],
						row2 = workbook.sheets[0].rows[2],
						row3 = workbook.sheets[0].rows[3],
						row4 = workbook.sheets[0].rows[4],
						row5 = workbook.sheets[0].rows[5];

					[row1, row2, row3, row4, row5].forEach(function (row) {
						expect(row[0].value).to.be.a('number');
						expect(row[1].value).to.be.a('number');
					})

					expect(row1[0].value).to.equal(1);
					expect(row1[1].value).to.equal(1.0001);
					expect(row2[0].value).to.equal(2);
					expect(row2[1].value).to.equal(2);
					expect(row3[0].value).to.equal(3);
					expect(row3[1].value).to.equal(3.00967676764465);
					expect(row4[0].value).to.equal(4);
					expect(row4[1].value).to.equal(0);
					expect(row5[0].value).to.equal(5);
					expect(row5[1].value).to.equal(5.00005);

					return next();
				});
			}, done);
		});

		it('should parse string values', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, function (err, workbook) {

					if (err) throw err;

					var cell1 = workbook.sheets[0].rows[1][2],
						cell2 = workbook.sheets[0].rows[2][2],
						cell3 = workbook.sheets[0].rows[3][2],
						cell4 = workbook.sheets[0].rows[4][2],
						cell5 = workbook.sheets[0].rows[5][2];

					[cell1, cell2, cell3, cell4, cell5].forEach(function (cell) {
						expect(cell.value).to.be.a('string');
					});

					expect(cell1.value).to.equal('Some text');
					expect(cell2.value).to.equal('{ "property": "value" }');
					expect(cell3.value).to.equal('ÉéÀàçÇùÙ');
					expect(cell4.value).to.equal('some "quoted" text');
					expect(cell5.value).to.equal('more \'quoted\' "text"');

					return next();
				});
			}, done);
		});

		it('should parse date values', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, function (err, workbook) {

					if (err) throw err;

					var row1 = workbook.sheets[0].rows[1],
						row2 = workbook.sheets[0].rows[2],
						row3 = workbook.sheets[0].rows[3],
						row4 = workbook.sheets[0].rows[4],
						row5 = workbook.sheets[0].rows[5];

					[row1, row2, row3, row4, row5].forEach(function (row) {
						expect(row[3].value).to.be.a('date');
						expect(row[4].value).to.be.a('date');
					});

					expect(row1[3].value).to.eql(new Date(2013, 0, 1, 0, 0, 0));
					expect(row1[4].value).to.eql(new Date(2013, 0, 1, 12, 54, 21));
					expect(row2[3].value).to.eql(new Date(2013, 0, 2, 0, 0, 0));
					//expect(row2[4].value).to.eql(new Date(0, 0, 0, 0, 0, 34));
					expect(row3[4].value).to.eql(new Date(2013, 0, 3, 3, 45, 20));
					expect(row4[4].value).to.eql(new Date(2013, 0, 4, 0, 0, 0));
					expect(row5[4].value).to.eql(new Date(2013, 0, 5, 16, 0, 0));

					return next();
				});
			}, done);
		});

		it('should parse empty cells as nulls', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, function (err, workbook) {

					if (err) throw err;

					workbook.sheets[0].rows.splice(1).forEach(function (row) {
						expect(row[5].value).to.be.null;
					});

					return next();
				});
			}, done);
		});

		it('should parse cells with errors', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, function (err, workbook) {

					if (err) throw err;

					var row1 = workbook.sheets[0].rows[1];
					var row2 = workbook.sheets[0].rows[2];
					var row3 = workbook.sheets[0].rows[3];

					expect(row1[6].value).to.be.an.instanceof(Error).and.have.property('errorCode', '#DIV/0!');
					expect(row2[6].value).to.be.an.instanceof(Error).and.have.property('errorCode', '#NAME?');
					expect(row3[6].value).to.be.an.instanceof(Error).and.have.property('errorCode', '#VALUE!');

					return next();
				});
			}, done);
		});

		it('should parse cells with booleans', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, function (err, workbook) {

					if (err) throw err;

					var row1 = workbook.sheets[0].rows[1];
					var row2 = workbook.sheets[0].rows[2];

					expect(row1[7].value).to.be.true;
					expect(row2[7].value).to.be.false;

					return next();
				});
			}, done);
		});

		it('should parse only the selected sheet by index', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, { sheet: 0 }, function (err, workbook) {
					if (err) throw err;
					expect(workbook.sheets).to.have.length(1);
					expect(workbook.sheets[0]).to.have.property('index', 0);
					expect(workbook.sheets[0]).to.have.property('name', 'Sheet1');
					return next();
				});
			}, done);
		});

		it('should parse only the selected sheet by name', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, { sheet: 'Sheet1' }, function (err, workbook) {
					if (err) throw err;
					expect(workbook.sheets).to.have.length(1);
					expect(workbook.sheets[0]).to.have.property('index', 0);
					expect(workbook.sheets[0]).to.have.property('name', 'Sheet1');
					return next();
				});
			}, done);
		});

		it('should parse only metadata', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, { meta: true }, function (err, workbook) {
					if (err) throw err;
					expect(workbook.sheets).to.be.undefined;
					return next();
				});
			}, done);
		});

		it('should parse up to 10 rows on the first sheet', function (done) {
			async.eachSeries(sampleFiles, function (file, next) {
				pyspreadsheet.read(file, { sheet: 0, maxRows: 10 }, function (err, workbook) {
					if (err) throw err;
					expect(workbook.sheets[0]).to.have.property('name', 'Sheet1');
					expect(workbook.sheets[0].rows).to.have.length(10);
					return next();
				});
			}, done);
		});

		it('should fail if the file does not exist', function (done) {
			pyspreadsheet.read('unknown.xlsx', function (err, workbook) {
				expect(workbook).to.not.be.ok;
				expect(err).to.be.an.instanceof(Error).and.have.property('id', 'file_not_found');
				return done();
			});
		});

		it('should fail if the file is not a valid workbook', function (done) {
			pyspreadsheet.read('package.json', function (err, workbook) {
				expect(workbook).to.not.be.ok;
				expect(err).to.be.an.instanceof(Error).and.have.property('id', 'open_workbook_failed');
				return done();
			});
		});
	});


});
