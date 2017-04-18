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
private MemoryPool gMemoryPool;

static this() {
    gMemoryPool = MemoryPool(StartingMemorySize);
}

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
        import std.algorithm: min;
        import std.experimental.allocator: expandArray;

        if (numBytes <= 0)
            return null;

        if (curPos + numBytes > data.length)
        {
            auto newAllocationSize = min(MaxMemorySize, data.length * 2);
            if (newAllocationSize <= data.length)
                return null;
            _allocator.expandArray(data, newAllocationSize, 0);
        }

        auto ret = data[curPos .. curPos + numBytes];
        curPos += numBytes;

        return ret;
    }

    // Frees all the temporary memory by setting the index for available memory back to the beginning
    bool deallocateAll() {
        curPos = 0;
        return true;
    }
}

// the function called by the Excel callback
void autoFree(LPXLOPER12 arg) {
    import xlld.framework: freeXLOper;
    freeXLOper(arg, autoFreeAllocator);
}
