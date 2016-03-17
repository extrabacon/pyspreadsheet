# 0.1.6
=======
* fix that avoids skipping end of file when python console finishes before parsing console.

# 0.1.4
=======
* installing with git instead of wget and curl

# 0.1.3
=======
* maintenance: simplified Python scripts, improved performance
* vastly improved documentation

# 0.1.2
=======
* fixed resolving of scripts with `python-shell`
* extended errors with traceback support for exceptions thrown from Python code

# 0.1.1
=======
* moved `python-shell` into a separate and more robust module

# 0.1.0
=======
* writing spreadsheet files no longer rely on a temporary file
* `SpreadsheetWriter` is now an `EventEmitter` and fires `open`, `close` and `error` events
* `SpreadsheetWriter.save` no longer accept a path argument, path is now assigned via the constructor instead
* `SpreadsheetWriter.save` no longer return a `stream.Readable` when callback is omitted
* `SpreadsheetWriter.destroy` has been removed, use `save` instead to close the underlying process
* Incomplete integration of xlwt for writing native XLS files

# 0.0.4
=======
* JSHint linting and fixes
* minor changes to coding style
* refactoring for removal of dependencies (underscore and async)
* improved compatibility of installation script

# 0.0.3
=======
* bug fixes
* improved debugging with [debug](visionmedia/debug)

# 0.0.2
=======
* documentation

# 0.0.1
=======
* initial version
* moved from [extrabacon/xlrd-parser](https://github.com/extrabacon/xlrd-parser)
* added support for writing XLSX
