var _ = require('underscore'),
	EventEmitter = require('events').EventEmitter,
	util = require('util'),
	spawn = require('child_process').spawn;

/**
 * An interactive Python shell which can stream JSON through stdio
 * @this PythonShell
 * @param {string} script The python script to execute
 * @param {Array} args The arguments to pass to the script
 * @param {object} [options] The process spawn options (passed to child_process.spawn)
 * @constructor
 */
var PythonShell = function (script, args, options) {

	var remaining = '',
		self = this;

	EventEmitter.call(this);

	/**
	 * Sends a command to the Python shell through stdin
	 * @param {string} command The command to execute
	 * @param {...} args The command arguments
	 */
	PythonShell.prototype.send = function (command, args) {

		var commandToSend = _.chain(arguments).rest().reject(function (val) {
			return val === null || _.isUndefined(val);
		}).unshift(command).value();

		self.childProcess.stdin.write(JSON.stringify(commandToSend) + '\n');

	};

	/**
	 * Ends the stdin stream, which should cause to the process to finish its work
	 */
	PythonShell.prototype.end = function () {
		// this will cause the process to close
		self.childProcess.stdin.end();
	};

	args = args || [];
	this.terminated = false;
	this.childProcess = spawn('python', args.slice(0).unshift(script), options);
	this.childProcess.stdout.setEncoding('utf8');

	// listen for incoming data on stdout
	this.childProcess.stdout.on('data', function (data) {

		var lines = data.split(/\n/g),
			lastLine = _.last(lines);

		// Fix the first line with the remaining from the previous iteration of 'data'
		lines[0] = remaining + lines[0];
		// Keep the remaining for the next iteration of 'data'
		remaining = lastLine;

		_.initial(lines).forEach(function (line) {
			var record = JSON.parse(line);
			if (record.length == 2) {
				self.emit('message', record[0], record[1]);
			} else if (record.length > 2) {
				self.emit('message', record[0], _.rest(record));
			} else {
				self.emit('error', _.extend(
					new Error('invalid message received from ' + script + ' >> ' + line),
					{ script: script, args: args, data: line }
				));
			}
		});

		self.emit('batchCompleted');
	});

	// listen to stderr and emit errors for incoming data
	this.childProcess.stderr.on('data', function (data) {
		self.emit('error', _.extend(
			new Error('an unexpected error occurred while executing ' + script + ' >> ' + data),
			{ script: script, args: args, data: data }
		));
	});

	this.childProcess.on('exit', function (code) {
		// look for an error from the exit code
		if (code != 0) {
			self.emit('error', _.extend(
				new Error('exit code ' + code + ' returned from ' + script),
				{ script: script, args: args, code: code })
			);
		}

		self.terminated = true;
		self.emit('close', code);
	});

};

util.inherits(PythonShell, EventEmitter);
module.exports = PythonShell;
