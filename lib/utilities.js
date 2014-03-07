var letters = ['',
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
    'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'
];

exports.extend = function (obj) {
    Array.prototype.slice.call(arguments, 1).forEach(function (source) {
        if (source) {
            for (var key in source) {
                obj[key] = source[key];
            }
        }
    });
    return obj;
};

exports.flatten = function (array) {
    var results = [];
    var self = arguments.callee;
    array.forEach(function(item) {
        Array.prototype.push.apply(results, Array.isArray(item) ? self(item) : [item]);
    });
    return results;
};

exports.find = function (array, test) {
    for (var i = 0, len = array.length; i < len; i++) {
        if (test(array[i])) {
            return array[i];
        }
    }
    return null;
};

exports.cellToRowCol = function (cell) {
    var match;
    if (!cell) {
        return [0, 0];
    } else if (cell.indexOf(',') > 0) {
        // 'row,col' notation
        match = /(\d+),(\d+)/.exec(cell);
        if (match) {
            return [parseInt(match[1], 10), parseInt(match[2], 10)];
        } else {
            throw new Error('invalid cell reference: ' + cell);
        }
    } else {
        // 'A1' notation
        match = /(\$?)([A-Z]{1,3})(\$?)(\d+)/.exec(cell);
        if (match) {
            return [parseInt(match[4], 10) - 1, this.colToInt(match[2]) - 1];
        } else {
            throw new Error('invalid cell reference: ' + cell);
        }
    }
};

/**
 * Converts a column letter into a column number (A -> 1, B -> 2, etc.)
 * @param col The column letter
 * @returns {number} The column number
 */
exports.colToInt = function (col) {
    var n = 0;
    col = col.trim().split('');
    for (var i = 0; i < col.length; i++) {
        n *= 26;
        n += letters.indexOf(col[i]);
    }
    return n;
};

