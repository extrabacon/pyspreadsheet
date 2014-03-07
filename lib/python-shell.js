var _ = require('./utilities');
var EventEmitter = require('events').EventEmitter;
var path = require('path');
var util = require('util');
var spawn = require('child_process').spawn;
var debug = require('debug')('pyspreadsheet:shell');

/**
 * An interactive Python shell streaming JSON fragments through stdio
 * @this PythonShell
 * @param {string} script The python script to execute
 * @param {Array} args The arguments to pass to the script
 * @param {object} [options] The process spawn options (passed to child_process.spawn)
 * @constructor
 */
var PythonShell = function (script, args, options) {
    var remaining = '';
    var errorData = '';
    var self = this;

    EventEmitter.call(this);

    script = path.join(PythonShell.rootPath, script);
    if (args) {
        args = args.slice(0);
        args.unshift(script);
    } else {
        args = [script];
    }

    this.terminated = false;
    this.childProcess = spawn('python', args, options);
    this.childProcess.stdout.setEncoding('utf8');
    debug('start: python ' + args.join(' '));

    // listen for incoming data on stdout, parsing into messages
    this.childProcess.stdout.on('data', function (data) {

        var lines = data.split(/\n/g);
        var lastLine = lines.pop();

        // fix the first line with the remaining from the previous iteration of 'data'
        lines[0] = remaining + lines[0];
        // keep the remaining for the next iteration of 'data'
        remaining = lastLine;

        lines.forEach(function (line) {
            debug('receive: ' + line);
            try {
                var record = JSON.parse(line);
                if (record.length === 1) {
                    self.emit('message', record[0]);
                } else if (record.length === 2) {
                    self.emit('message', record[0], record[1]);
                } else if (record.length > 2) {
                    self.emit('message', record[0], record.slice(1));
                }
            } catch (err) {
                self.emit('error', _.extend(
                    new Error('invalid or malformed message: ' + line),
                    { data: line, args: args }
                ));
            }
        });
    });

    // listen to stderr and emit errors for incoming data
    this.childProcess.stderr.on('data', function (data) {
        errorData += ''+data;
    });

    this.childProcess.on('exit', function (code) {

        debug('end: exit code %d', code);

        errorData && self.emit('error', _.extend(
            new Error('an unexpected error occurred while executing "' + script + '" >> ' + errorData),
            { data: errorData, args: args }
        ));

        self.exitCode = code;
        self.terminated = true;
        self.emit('close');
    });
};
util.inherits(PythonShell, EventEmitter);

PythonShell.rootPath = path.join(__dirname, '..', 'python');

/**
 * Sends a command to the Python shell through stdin
 * @param {string} command The command to execute
 * @param {...} args The command arguments
 */
PythonShell.prototype.send = function (command, args) {
    var commandToSend = Array.prototype.slice.call(arguments).filter(function (val) {
        return typeof val !== 'undefined' && val !== null;
    });
    debug('send: ' + JSON.stringify(commandToSend));
    this.childProcess.stdin.write(JSON.stringify(commandToSend) + '\n');
    return this;
};

/**
 * Closes the stdin stream, which should cause to the process to finish its work and close
 */
PythonShell.prototype.end = function () {
    // this will cause the process to end its listening loop
    this.childProcess.stdin.end();
    return this;
};

module.exports = PythonShell;
