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
import xlld.xlcall: LPXLOPER12;

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
}

auto memoryPool() {
    return MemoryPool(StartingMemorySize);
}

auto makeArray2D(T, A)(ref A allocator, int rows, int cols) {
    import std.experimental.allocator: makeArray;

    auto ret = allocator.makeArray!(T[])(rows);
    foreach(ref row; ret)
        row = allocator.makeArray!T(cols);

    return ret;
}

@("issue 22 - makeArray with 2D array causing relocations")
unittest {
    auto pool = memoryPool;
    auto arr2d = pool.makeArray2D!string(4000, 37);
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
