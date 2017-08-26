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
import std.experimental.allocator.building_blocks.allocator_list: AllocatorList;
import std.experimental.allocator.mallocator: Mallocator;
import std.experimental.allocator.building_blocks.region: Region;
import std.algorithm.comparison: max;
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

package alias MemoryPool = AllocatorList!((size_t n) => Region!Mallocator(max(n, size_t(1024 * 1024))), Mallocator);

package MemoryPool gTempAllocator;

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

    auto pool = MemoryPool();
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
