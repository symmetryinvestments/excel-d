/**
	MemoryManager.D

	Ported from MemoryManager.cpp by Laeeth Isharc
//
// Platform:    Microsoft Windows
//
///***************************************************************************
*/
module xlld.memorymanager;

import xlld.xlcall: XLOPER12, LPXLOPER12;
import xlld.any: Any;
import std.experimental.allocator.mallocator: Mallocator;
import std.traits: isArray;
import std.meta: allSatisfy;

version(unittest) {
    import unit_threaded;
    import xlld.test_util;
    import std.experimental.allocator.gc_allocator: GCAllocator;
    alias theMallocator = Mallocator.instance;
    alias theGC = GCAllocator.instance;
}


alias allocator = Mallocator.instance;
alias autoFreeAllocator = Mallocator.instance;

enum StartingMemorySize = 10240;
enum MaxMemorySize=100*1024*1024;

alias MemoryPool = MemoryPoolImpl!Mallocator;

/**
   This allocator isn't public since it's best not to use it from outside.
   It can have surprising consequences on code depending on it.
 */
package MemoryPool gTempAllocator;


struct MemoryPoolImpl(T) {

    alias _allocator = T.instance;

    package ubyte[] data;
    package size_t curPos;

    version(unittest) {
        private int _numReallocations;
        private size_t _largestReservation;
    }

    this(size_t startingMemorySize) {
        import std.experimental.allocator: makeArray;
        data = _allocator.makeArray!(ubyte)(startingMemorySize);
    }

    ~this() nothrow {
        import std.experimental.allocator: dispose;
        _allocator.dispose(data);
    }

    ubyte[] allocate(size_t numBytes) {

        import std.algorithm: min, max;
        import std.experimental.allocator: expandArray;

        if (numBytes == 0)
            return null;

        if (curPos + numBytes > data.length)
        {
            const newAllocationSize = min(MaxMemorySize, max(data.length * 2, curPos + numBytes));

            if (newAllocationSize <= data.length)
                return null;

            const delta = newAllocationSize - data.length;
            _allocator.expandArray(data, delta, 0);
            version(unittest) ++_numReallocations;
        }

        auto ret = data[curPos .. curPos + numBytes];
        curPos += numBytes;

        return ret;
    }

    @("ensure size of data after forced reallocation")
    unittest {
        auto pool = MemoryPool(4);
        pool.allocate(10); // must reallocate to service request
        pool.data.length.shouldEqual(10);
    }

    @("issue 22 - expansion must cover the allocation request itself")
    unittest {
        auto pool = memoryPool;
        pool.allocate(32_000);
    }

    @("issue 22 - expansion must take into account number of bytes already allocated")
    unittest {
        auto pool = memoryPool;
        pool.allocate(10);
        pool.allocate(32_000);
    }

    // Frees all the temporary memory by setting the index for available memory back to the beginning
    bool deallocateAll() {
        curPos = 0;
        return true;
    }

    // pre-allocate memory to serve a large request
    bool reserve(in size_t numBytes) {
        import std.experimental.allocator: expandArray;
        version(unittest) import std.algorithm: max;

        version(unittest) _largestReservation = max(_largestReservation, numBytes);

        if(numBytes < data.length - curPos) return true;
        if(numBytes + curPos > MaxMemorySize) return false;

        const delta = numBytes - (data.length - curPos);
        return _allocator.expandArray(data, delta, 0);
    }

    @("reserve")
    unittest {
        auto pool = MemoryPool(10);

        pool.allocate(5);
        pool.reserve(3).shouldBeTrue;
        pool.curPos.shouldEqual(5);
        pool.data.length.shouldEqual(10);

        pool.reserve(MaxMemorySize).shouldBeFalse;

        pool.reserve(7).shouldBeTrue;

        const ptr = pool.data.ptr;
        const length = pool.data.length;
        pool.curPos.shouldEqual(5);
        length.shouldEqual(12);

        pool.allocate(7);

        pool.data.ptr.shouldEqual(ptr);
        pool.data.length.shouldEqual(length);
        pool.curPos.shouldEqual(12);
    }

    version(unittest) {
        int numReallocations() @safe @nogc pure nothrow const {
            return _numReallocations;
        }

        size_t largestReservation() @safe @nogc pure nothrow const {
            return _largestReservation;
        }
    }
}

auto memoryPool() {
    return MemoryPool(StartingMemorySize);
}

enum isOperPtr(T) = is(T == XLOPER12*);

// returns the total number of bytes that will have to be allocated for the D
// arguments to be passed to wrappedFunc base on the opers passed in.
// This allows to reserve enough space for all the D arguments
size_t numBytesForDArgs(alias wrappedFunc, T...)(auto ref T opers) if(allSatisfy!(isOperPtr, T)) {
    import std.traits: Parameters;

    size_t numBytes;

    foreach(i, T; Parameters!wrappedFunc) {
        numBytes += numBytesFor!T(*opers[i]);
    }

    return numBytes;
}

@("numBytesForDArgs (double, double)")
unittest {
    import xlld.wrap: toXlOper;
    double func(double a, double b) { return a + b; }
    auto oper1 = (5.0).toXlOper(theGC);
    auto oper2 = (7.0).toXlOper(theGC);
    numBytesForDArgs!func(&oper1, &oper2).shouldEqual(0);
}

@("numBytesForDArgs (double, double)")
unittest {
    import xlld.wrap: toXlOper;
    double func(double a, string b) { return a + b.length; }
    auto oper1 = (5.0).toXlOper(theGC);
    auto oper2 = "foo".toXlOper(theGC);
    // "foo" has length 3, which in wide chars is 6
    // since XL uses 16 bits for the length, that equals 8 bytes
    numBytesForDArgs!func(&oper1, &oper2).shouldEqual(8);
}


size_t numBytesFor(T)(ref const(XLOPER12) oper) if(is(T == double) || is(T == Any)) {
    return 0;
}

@("numBytesFor!double")
unittest {
    import xlld.wrap: toXlOper, fromXlOper;

    auto arg = 2.0;
    auto oper = arg.toXlOper(theGC);
    auto pool = MemoryPool(1);
    pool.reserve(numBytesFor!(double)(oper));
    auto back = oper.fromXlOper!(double)(pool);

    pool.numReallocations.shouldEqual(0);
    back.shouldEqual(arg);
}

size_t numBytesFor(T)(ref XLOPER12 oper) if(is(T == double[]) || is(T == Any[]) || is(T == string[])) {
    import xlld.wrap: isMulti;

    if(!isMulti(oper))
        return 0;

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;

    return typeof(T.init[0]).sizeof * (rows * cols) + countStringBytes!(typeof(T.init[0]))(oper);
}

private size_t countStringBytes(T)(ref XLOPER12 oper) {

    size_t elemAllocBytes;

    static if(is(T == string)) {
        import xlld.wrap: apply, numOperStringBytes;
        try
            oper.apply!(T, (shouldConvert, row, col, cellVal) {
                if(shouldConvert)
                    elemAllocBytes += numOperStringBytes(cellVal);
            });
        catch(Exception ex) {
            return 0;
        }
    }

    return elemAllocBytes;
}

@("numBytesFor!double[]")
unittest {
    import xlld.wrap: toXlOper, fromXlOper;

    auto arg = [1.0, 2.0];
    auto oper = arg.toXlOper(theGC);

    auto pool = MemoryPool(1);
    pool.reserve(numBytesFor!(double[])(oper));
    auto back = oper.fromXlOper!(double[])(pool);

    pool.numReallocations.shouldEqual(0);
    back.shouldEqual(arg);
}

size_t numBytesFor(T)(ref XLOPER12 oper) if(is(T == double[][]) || is(T == string[][]) || is(T == Any[][])) {
    return numBytesForArray2D!(typeof(T.init[0][0]))(oper);
}


@("numBytesFor!double[][]")
unittest {
    import xlld.wrap: toXlOper, fromXlOper;

    auto arg = [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]];
    auto oper = arg.toXlOper(theGC);

    auto pool = MemoryPool(1);
    pool.reserve(numBytesFor!(double[][])(oper));
    auto back = oper.fromXlOper!(double[][])(pool);

    pool.numReallocations.shouldEqual(0);
    back.shouldEqual(arg);
}


size_t numBytesFor(T)(ref const(XLOPER12) oper) if(is(T == string)) {
    import xlld.wrap: numOperStringBytes;
    return numOperStringBytes(oper);
}

@("numBytesFor!string")
unittest {
    import xlld.wrap: toXlOper, fromXlOper;

    auto arg = "foo";
    auto oper = arg.toXlOper(theGC);

    // 2 bytes for the length + 3 chars * wchar.sizeof = 8
    numBytesFor!string(oper).shouldEqual(8);

    auto pool = MemoryPool(1);
    pool.reserve(numBytesFor!(string)(oper));
    auto back = oper.fromXlOper!(string)(pool);

    pool.numReallocations.shouldEqual(0);
    back.shouldEqual(arg);
}


@("numBytesFor!string[]")
unittest {
    import xlld.wrap: toXlOper, fromXlOper;

    auto arg = [
        "the quick brown fox jumps over something something darkside",
        "some nonsensical string for size purposes",
        "quux",
    ];
    auto oper = arg.toXlOper(theGC);

    auto pool = MemoryPool(1);
    pool.reserve(numBytesFor!(string[])(oper));
    auto back = oper.fromXlOper!(string[])(pool);

    pool.numReallocations.shouldEqual(0);
    back.shouldEqual(arg);
}


@("numBytesFor!string[][]")
unittest {
    import xlld.wrap: toXlOper, fromXlOper;

    auto arg = [
        [
            "the quick brown fox jumps over something something darkside",
            "some nonsensical string for size purposes",
            "quux",
        ],
        [
            "it's hard to come up with different strings",
            "lorem ipsum something something darkside",
            "foobar",
        ],
    ];
    auto oper = arg.toXlOper(theGC);

    auto pool = MemoryPool(1);
    pool.reserve(numBytesFor!(string[][])(oper) - 10);
    auto back = oper.fromXlOper!(string[][])(pool);

    pool.numReallocations.shouldEqual(0);
    back.shouldEqual(arg);
}


@("numBytesFor!any double")
unittest {
    import xlld.wrap: toXlOper, fromXlOper;
    import xlld.any: any;

    auto arg = any(2.0, theGC);
    auto oper = arg.toXlOper(theGC);
    auto pool = MemoryPool(1);
    pool.reserve(numBytesFor!Any(oper));
    auto back = oper.fromXlOper!Any(pool);

    pool.numReallocations.shouldEqual(0);
    back.shouldEqual(arg);
}

@("numBytesFor!any string")
unittest {
    import xlld.wrap: toXlOper, fromXlOper;
    import xlld.any: any;

    auto arg = any("The quick brown fox jumps over the lazy dog",
                   theGC);
    auto oper = arg.toXlOper(theGC);
    auto pool = MemoryPool(1);
    pool.reserve(numBytesFor!Any(oper));
    auto back = oper.fromXlOper!Any(pool);

    pool.numReallocations.shouldEqual(0);
    back.shouldEqual(arg);
}


@("numBytesFor!any[]")
unittest {
    import xlld.wrap: toXlOper, fromXlOper;

    Any[] arg;
    with(allocatorContext(theGC)) {
        arg =
            [
                any(2.0),
                any(3.0),
                any("the quick brown fox jumps over the lazy dog"),
            ];
    }
    auto oper = arg.toXlOper(theGC);
    auto pool = MemoryPool(1);
    pool.reserve(numBytesFor!(Any[])(oper));
    auto back = oper.fromXlOper!(Any[])(pool);

    pool.numReallocations.shouldEqual(0);
    back.shouldEqual(arg);
}


@("numBytesFor!any[][]")
unittest {
    import xlld.wrap: toXlOper, fromXlOper;

    Any[][] arg;
    with(allocatorContext(theGC)) {
        arg =
            [
                [any(1.0), any(2.0), any("the quick brown fox jumps over the lazy dog"),],
                [any(4.0), any(5.0), any("the quick brown fox jumps over the lazy dog"),],
            ];
    }
    auto oper = arg.toXlOper(theGC);
    auto pool = MemoryPool(1);
    pool.reserve(numBytesFor!(Any[][])(oper));
    auto back = oper.fromXlOper!(Any[][])(pool);

    pool.numReallocations.shouldEqual(0);
    back.shouldEqual(arg);
}



T[][] makeArray2D(T, A)(ref A allocator, ref XLOPER12 oper) {
    import xlld.xlcall: XlType;
    import xlld.wrap: isMulti;
    import std.experimental.allocator: makeArray;

    if(!isMulti(oper))
        return T[][].init;

    static if(__traits(compiles, allocator.reserve(1)))
        allocator.reserve(numBytesForArray2D!T(oper));

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;

    auto ret = allocator.makeArray!(T[])(rows);
    foreach(ref row; ret)
        row = allocator.makeArray!T(cols);

    return ret;
}

@("issue 22 - makeArray with 2D array causing relocations")
unittest {
    import xlld.wrap: toXlOper;

    auto pool = memoryPool;
    string[][] strings;
    strings.length = 4000;
    foreach(ref s; strings) s.length = 37;
    auto oper = strings.toXlOper(Mallocator.instance);
    auto arr2d = pool.makeArray2D!string(oper);
}


// the number of bytes that need to be allocated for a 2D array of type T[][]
private size_t numBytesForArray2D(T)(size_t rows, size_t cols) {
    return rows * (T[].sizeof + T.sizeof * cols);
}

@("numBytesForArray2D!string rows cols")
@safe pure unittest {
    version(X86)         const expected = 1_216_000;
    else version(X86_64) const expected = 2_432_000;
    else static assert(0);

    numBytesForArray2D!string(4000, 37).shouldEqual(expected);
}

@("numBytesForArray2D!int rows cols")
@safe pure unittest {
    version(X86)         const expected = 624_000;
    else version(X86_64) const expected = 656_000;
    else static assert(0);
    numBytesForArray2D!int(4000, 37).shouldEqual(expected);
}


// the number of bytes that need to be allocated to convert oper to T[][]
private size_t numBytesForArray2D(T)(ref XLOPER12 oper) {
    import xlld.wrap: isMulti;

    if(!isMulti(oper))
        return 0;

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;

    return numBytesForArray2D!T(rows, cols) + countStringBytes!T(oper);
}


@("numBytesForArray2D!double oper")
unittest {
    import xlld.wrap: toXlOper;

    auto doubles = [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]];
    const rows = doubles.length; const cols = doubles[0].length;

    auto oper = doubles.toXlOper(theGC);
    // no allocation for doubles so the memory requirements are just the array itself
    numBytesForArray2D!double(oper).shouldEqual(numBytesForArray2D!double(rows, cols));
    numBytesForArray2D!double(oper).shouldEqual(numBytesFor!(double[][])(oper));
}

@("numBytesForArray2D!string oper")
unittest {
    import xlld.wrap: toXlOper, numOperStringBytes;
    import std.array: join;
    import std.algorithm: fold;

    auto strings = [["foo", "bar"], ["quux", "toto"], ["a", "b"]];
    const rows = strings.length; const cols = strings[0].length;

    auto oper = strings.toXlOper(Mallocator.instance);

    size_t seed;
    const bytesForStrings = strings.join.fold!((a, b) => a + numOperStringBytes(b))(seed);
    bytesForStrings.shouldEqual(8 + 8 + 10 + 10 + 4 + 4);

    // no allocation for doubles so the memory requirements are just the array itself
    numBytesForArray2D!string(oper).shouldEqual(numBytesForArray2D!string(rows, cols) + bytesForStrings);
    numBytesForArray2D!string(oper).shouldEqual(numBytesFor!(string[][])(oper));
}


// the function called by the Excel callback
void autoFree(LPXLOPER12 arg) nothrow {
    import xlld.framework: freeXLOper;
    freeXLOper(arg, autoFreeAllocator);
}

struct AllocatorContext(A) {

    A* _allocator_;

    this(ref A allocator) {
        _allocator_ = &allocator;
    }

    auto any(T)(auto ref T value) {
        import xlld.any: _any = any;
        return _any(value, *_allocator_);
    }

    auto fromXlOper(T, U)(U oper) {
        import xlld.wrap: wrapFromXlOper = fromXlOper;
        return wrapFromXlOper!T(oper, _allocator_);
    }

    auto toXlOper(T)(T val) {
        import xlld.wrap: wrapToXlOper = toXlOper;
        return wrapToXlOper(val, _allocator_);
    }

    version(unittest) {
        auto toSRef(T)(T val) {
            import xlld.test_util: toSRef_ = toSRef;
            return toSRef_(val, _allocator_);
        }
    }
}

auto allocatorContext(A)(ref A allocator) {
    return AllocatorContext!A(allocator);
}


// this shouldn't be needed IMHO and is a bug in std.experimental.allocator that dispose
// doesn't handle 2D arrays correctly
void dispose(A, T)(auto ref A allocator, T[] array) {
    static import std.experimental.allocator;
    import std.traits: Unqual;

    static if(isArray!T) {
        foreach(ref e; array) {
            dispose(allocator, e);
        }
    }

    alias U = Unqual!T;
    std.experimental.allocator.dispose(allocator, cast(U[])array);
}

void dispose(A, T)(auto ref A allocator, T value) if(!isArray!T) {
    static import std.experimental.allocator;
    std.experimental.allocator.dispose(allocator, value);
}
