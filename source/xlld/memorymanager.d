/**
	MemoryManager.D

	Ported from MemoryManager.cpp by Laeeth Isharc
//
// Platform:    Microsoft Windows
//
///***************************************************************************
*/
module xlld.memorymanager;

import std.typecons: Flag, Yes;
import std.experimental.allocator.mallocator: Mallocator;

alias Allocator = Mallocator;
alias allocator = Allocator.instance;

enum StartingMemorySize = 10240;
enum MaxMemorySize=100*1024*1024;

private MemoryPool excelCallPool;

static this() {
    excelCallPool = MemoryPool(StartingMemorySize);
}


struct MemoryPool {

    ubyte[] data;
    size_t curPos=0;

    this(size_t startingMemorySize) nothrow @nogc {
        import std.experimental.allocator: makeArray;

        if (data.length==0)
            data=allocator.makeArray!(ubyte)(startingMemorySize);
        curPos=0;
    }

    ~this() nothrow @nogc {
        dispose;
    }

    void dispose() nothrow @nogc {
        import std.experimental.allocator: dispose;

        if(data.length>0)
            allocator.dispose(data);
        data = [];
        curPos=0;
    }

    void dispose(string) nothrow @nogc {
        dispose;
    }

    ubyte[] allocate(size_t numBytes) nothrow @nogc {
        import std.algorithm: min;
        import std.experimental.allocator: expandArray;

        if (numBytes<=0)
            return null;

        if (curPos + numBytes > data.length)
        {
            auto newAllocationSize = min(MaxMemorySize, data.length * 2);
            if (newAllocationSize <= data.length)
                return null;
            allocator.expandArray(data, newAllocationSize, 0);
        }

        auto lpMemory = data[curPos .. curPos+numBytes];
        curPos += numBytes;

        return lpMemory;
    }

    // Frees all the temporary memory by setting the index for available memory back to the beginning
    void freeAll() nothrow @nogc {
        curPos = 0;
    }
}
