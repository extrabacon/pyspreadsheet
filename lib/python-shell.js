var _ = require('underscore'),
	EventEmitter = require('events').EventEmitter,
	util = require('util'),
	spawn = require('child_process').spawn,
	debug = false;

/**
 * An interactive Python shell which streams JSON fragments through stdio
 * @this PythonShell
 * @param {string} script The python script to execute
 * @param {Array} args The arguments to pass to the script
 * @param {object} [options] The process spawn options (passed to child_process.spawn)
 * @constructor
 */
var PythonShell = function (script, args, options) {

	var remaining = '',
		errorData = '',
		self = this;

	options = options || { stdio: 'pipe', debug: debug };
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

		options.debug && console.log(' >> ' + JSON.stringify(commandToSend));
		self.childProcess.stdin.write(JSON.stringify(commandToSend) + '\n');
		return self;

	};

	/**
	 * Closes the stdin stream, which should cause to the process to finish its work and close
	 */
	PythonShell.prototype.end = function () {
		// this will cause the process to end its listening loop
		self.childProcess.stdin.end();
		return self;
	};

	if (args) {
		args = args.slice(0);
		args.unshift(script);
	} else {
		args = [script];
	}

	this.terminated = false;
	this.childProcess = spawn('python', args, options);
	this.childProcess.stdout.setEncoding('utf8');

	// listen for incoming data on stdout
	this.childProcess.stdout.on('data', function (data) {

		var lines = data.split(/\n/g),
			lastLine = _.last(lines);

		// fix the first line with the remaining from the previous iteration of 'data'
		lines[0] = remaining + lines[0];
		// keep the remaining for the next iteration of 'data'
		remaining = lastLine;

		_.initial(lines).forEach(function (line) {
			options.debug && console.log(' << ' + line);
			try {
				var record = JSON.parse(line);
				if (record.length == 1) {
					self.emit('message', record[0]);
				} else if (record.length == 2) {
					self.emit('message', record[0], record[1]);
				} else if (record.length > 2) {
					self.emit('message', record[0], _.rest(record));
				}
			} catch(err) {
				self.emit('error', _.extend(
					new Error('invalid or malformed message: ' + line),
					{ args: args, data: line }
				));
			}
		});
	});

	// listen to stderr and emit errors for incoming data
	this.childProcess.stderr.on('data', function (data) {
		errorData += ''+data;
	});

	this.childProcess.on('exit', function (code) {

		errorData && self.emit('error', _.extend(
			new Error('an unexpected error occurred while executing "' + script + '" >> ' + errorData),
			{ args: args, data: errorData }
		));

		self.exitCode = code;
		self.terminated = true;
		self.emit('close');
	});

};

util.inherits(PythonShell, EventEmitter);
module.exports = PythonShell;
