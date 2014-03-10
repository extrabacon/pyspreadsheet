var _ = require('../lib/utilities');

describe('utilities', function () {
    describe('extend', function () {
        it('should extend an object with a source', function () {
            _.extend({}, { a: 'b' }).should.have.property('a', 'b');
        });
        it('should extend an object with multiple sources', function () {
            _.extend({}, { a: 'b' }, { c: 'd' }).should.have.properties({
                a: 'b',
                c: 'd'
            });
        });
        it('should extend an error with properties', function () {
            _.extend(new Error('error'), { a: 'b' }, { c: 'd' }).should.be.an.Error.and.have.properties({
                a: 'b',
                c: 'd'
            });
        });
    });

    describe('flatten', function () {
        it('should flatten an array', function () {
            _.flatten([1, 2, 3, [4]]).should.containDeep([1, 2, 3, 4]);
        });
        it('should flatten an array recursively', function () {
            _.flatten([[1, [2]], [[[3]]], 4]).should.containDeep([1, 2, 3, 4]);
        });
    });

    describe('find', function () {
        it('should return the item matching the truth test', function () {
            _.find([1, 2, 3], function (el) {
                return el === 2;
            }).should.be.exactly(2);
        });
    });

    describe('randomString', function () {
        it('should return a random string', function () {
            _.randomString(8).should.have.lengthOf(8).and.match(/\w+/);
        });
    });

    describe('cellToRowCol', function () {
        it('should return row and column indexes', function () {
            _.cellToRowCol('A1').should.containDeep([0, 0]);
            _.cellToRowCol('C3').should.containDeep([2, 2]);
            _.cellToRowCol('AA12').should.containDeep([11, 26]);
        });
    });

    describe('colToInt', function () {
        it('should convert a column letter to a zero-based index', function () {
            _.colToInt('A').should.be.exactly(1);
            _.colToInt('C').should.be.exactly(3);
            _.colToInt('AA').should.be.exactly(27);
        });
    });
});
