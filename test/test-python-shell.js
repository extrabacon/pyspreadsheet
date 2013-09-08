var async = require('async'),
	expect = require('chai').expect,
	PythonShell = require('../lib/python-shell');

describe('PythonShell', function () {

	describe('ctor', function () {

		it('should spawn a Python process', function (done) {

			var pyshell = new PythonShell('-');
			expect(pyshell.terminated).to.be.false;

			pyshell.on('close', function () {
				expect(pyshell.exitCode).to.equal(0);
				expect(pyshell.terminated).to.be.true;
				done();
			}).end();
		});

		it('should emit error on failure', function (done) {

			var pyshell = new PythonShell('tests/error.py');

			pyshell.on('error', function (err) {
				expect(err).to.not.be.null;
				expect(err.toString()).to.contain(' >> Traceback ');
				expect(err).to.have.property('data');
				expect(err.data.toString()).to.contain('ZeroDivisionError: integer division or modulo by zero');
			}).on('close', function () {
				expect(pyshell.exitCode).to.equal(1);
				expect(pyshell.terminated).to.be.true;
				done();
			});
		});

	});

	describe('send', function () {
		it('should send and receive JSON messages in the same order', function (done) {

			var pyshell = new PythonShell('tests/echo.py');
			var count = 0;

			pyshell.send('command1');
			pyshell.send('command2', 'string');
			pyshell.send('command3', 1, 2, 3);

			pyshell.on('message', function (command, args) {

				switch (count) {
					case 0:
						expect(command).to.equal('command1');
						expect(args).to.be.undefined;
						break;
					case 1:
						expect(command).to.equal('command2');
						expect(args).to.equal('string');
						break;
					case 2:
						expect(command).to.equal('command3');
						expect(args).to.have.length(3);
						expect(args[0]).to.equal(1);
						expect(args[1]).to.equal(2);
						expect(args[2]).to.equal(3);
						break;
				}
				count++;

			}).on('close', function () {
				expect(count).to.equal(3);
				done();
			}).end();
		});
	});

});
