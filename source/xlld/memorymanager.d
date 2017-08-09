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
import std.experimental.allocator.mallocator: Mallocator;
import std.traits: isArray;

version(unittest) import unit_threaded;

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

    ubyte[] data;
    size_t curPos;

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
}

auto memoryPool() {
    return MemoryPool(StartingMemorySize);
}

T[][] makeArray2D(T, A)(ref A allocator, ref XLOPER12 oper) {
    import xlld.xlcall: XlType;
    import xlld.wrap: isMulti;
    import std.experimental.allocator: makeArray;

    static if(__traits(compiles, allocator.reserve(5))) {
        allocator.reserve(numBytesForArray2D!T(oper));
    }

    if(!isMulti(oper))
        return T[][].init;

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
    import xlld.wrap: dlangToXlOperType, isMulti, numOperStringBytes, apply;

    if(!isMulti(oper))
        return 0;

    size_t elemAllocBytes;

    try
        oper.apply!(T, (shouldConvert, row, col, cellVal) {
            if(shouldConvert && is(T == string))
                elemAllocBytes += numOperStringBytes(cellVal);
        });
    catch(Exception ex) {
        return 0;
    }

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;

    return numBytesForArray2D!T(rows, cols) + elemAllocBytes;
}


@("numBytesForArray2D!double oper")
unittest {
    import xlld.wrap: toXlOper;

    auto doubles = [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]];
    const rows = doubles.length; const cols = doubles[0].length;

    auto oper = doubles.toXlOper(Mallocator.instance);
    // no allocation for doubles so the memory requirements are just the array itself
    numBytesForArray2D!double(oper).shouldEqual(numBytesForArray2D!double(rows, cols));
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
}


// the function called by the Excel callback
void autoFree(LPXLOPER12 arg) {
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
