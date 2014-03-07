var PythonShell = require('../lib/python-shell');

describe('PythonShell', function () {

    var rootPath;
    before(function () {
        rootPath = PythonShell.rootPath;
        PythonShell.rootPath = '.';
    });
    after(function () {
        PythonShell.rootPath = rootPath;
    });

    describe('#ctor', function () {

        it('should spawn a Python process', function (done) {

            var pyshell = new PythonShell('-');
            pyshell.terminated.should.be.false;

            pyshell.on('close', function () {
                pyshell.exitCode.should.be.exactly(0);
                pyshell.terminated.should.be.true;
                done();
            }).end();
        });

        it('should emit error on failure', function (done) {

            var pyshell = new PythonShell('test/error.py');

            pyshell.on('error', function (err) {
                err.message.should.match(/ >> Traceback /);
                err.should.have.property('data');
                err.data.should.containEql('ZeroDivisionError: integer division or modulo by zero');
            }).on('close', function () {
                pyshell.exitCode.should.be.exactly(1);
                pyshell.terminated.should.be.true;
                done();
            });
        });

    });

    describe('.send(command, args)', function () {
        it('should send and receive JSON messages in the same order', function (done) {

            var pyshell = new PythonShell('test/echo.py');
            var count = 0;

            pyshell.send('command1');
            pyshell.send('command2', 'string');
            pyshell.send('command3', 1, 2, 3);

            pyshell.on('message', function (command, args) {
                switch (count) {
                    case 0:
                        command.should.be.exactly('command1');
                        break;
                    case 1:
                        command.should.be.exactly('command2');
                        args.should.be.exactly('string');
                        break;
                    case 2:
                        command.should.be.exactly('command3');
                        args.should.be.an.Array.and.have.lengthOf(3);
                        args[0].should.be.exactly(1);
                        args[1].should.be.exactly(2);
                        args[2].should.be.exactly(3);
                        break;
                }
                count++;
            }).on('close', function () {
                count.should.be.exactly(3);
                done();
            }).end();
        });
    });
});
