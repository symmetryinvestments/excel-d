/**
	MemoryManager.D

	Ported from MemoryManager.cpp by Laeeth Isharc
//
// Platform:    Microsoft Windows
//
///***************************************************************************
*/
module xlld.memorymanager;

import std.experimental.allocator.mallocator: Mallocator;
import xlld.xlcall: XLOPER12, LPXLOPER12;

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
    import xlld.xlcall: xlbitDLLFree, XlType;
    import std.experimental.allocator: makeArray;

    static if(__traits(compiles, allocator.reserve(5))) {
        allocator.reserve(numBytesForArray2D!T(oper));
    }

    const realType = oper.xltype & ~xlbitDLLFree;
    if(realType != XlType.xltypeMulti)
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
    numBytesForArray2D!string(4000, 37).shouldEqual(1_216_000);
}

@("numBytesForArray2D!int rows cols")
@safe pure unittest {
    numBytesForArray2D!int(4000, 37).shouldEqual(624_000);
}


// the number of bytes that need to be allocated to convert val to T[][]
private size_t numBytesForArray2D(T)(ref XLOPER12 val) {
    import xlld.xlcall: xlbitDLLFree, XlType;
    import xlld.xl: coerce, free;
    import xlld.wrap: dlangToXlOperType;
    import xlld.any: Any;
    version(unittest) import xlld.test_util: gNumXlCoerce, gNumXlFree;

    const realType = val.xltype & ~xlbitDLLFree;
    if(realType != XlType.xltypeMulti)
        return 0;

    const rows = val.val.array.rows;
    const cols = val.val.array.columns;
    auto values = val.val.array.lparray[0 .. (rows * cols)];
    size_t elemAllocBytes;

    foreach(const row; 0 .. rows) {
        foreach(const col; 0 .. cols) {
            auto cellVal = coerce(&values[row * cols + col]);
            version(unittest) --gNumXlCoerce; // ignore this for testing
            scope(exit) {
                free(&cellVal);
                version(unittest) --gNumXlFree; // ignore this for testing
            }

            // try to convert doubles to string if trying to convert everything to an
            // array of strings
            const shouldConvert = (cellVal.xltype == dlangToXlOperType!T.Type) ||
                (cellVal.xltype == XlType.xltypeNum && dlangToXlOperType!T.Type == XlType.xltypeStr)
                || is(T == Any);
            if(shouldConvert && is(T == string))
                elemAllocBytes += (cellVal.val.str[0] + 1) * wchar.sizeof;
        }
    }

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
    import xlld.wrap: toXlOper;
    import std.array: join;
    import std.algorithm: fold;

    auto strings = [["foo", "bar"], ["quux", "toto"], ["a", "b"]];
    const rows = strings.length; const cols = strings[0].length;

    auto oper = strings.toXlOper(Mallocator.instance);

    size_t bytesPerString(in string str) {
        // XLOPER12 strings are wide strings where the first "character"
        // is the length, the real string is in [1.. $]
        return (str.length + 1) * wchar.sizeof;
    }

    const bytesForStrings = strings.join.fold!((a, b) => a + bytesPerString(b))(0);
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

    A* allocator;

    this(ref A allocator) {
        this.allocator = &allocator;
    }

    auto any(T)(auto ref T value) {
        import xlld.any: _any = any;
        return _any(value, *allocator);
    }

    auto fromXlOper(T, U)(U oper) {
        import xlld.wrap: wrapFromXlOper = fromXlOper;
        return wrapFromXlOper!T(oper, allocator);
    }

    auto toXlOper(T)(T val) {
        import xlld.wrap: wrapToXlOper = toXlOper;
        return wrapToXlOper(val, allocator);
    }

}

auto allocatorContext(A)(ref A allocator) {
    return AllocatorContext!A(allocator);
}
