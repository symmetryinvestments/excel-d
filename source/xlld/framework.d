/**
   Utility functions for writing XLLs
 */
module xlld.framework;


import xlld.xlcall;


/**
        Will free any malloc'd memory associated with the given
        LPXLOPER, assuming it has any memory associated with it

   Parameters:

        LPXLOPER pxloper    Pointer to the XLOPER whose associated
*/
void freeXLOper(T, A)(T pxloper, ref A allocator)
    if(is(T == LPXLOPER) || is(T == LPXLOPER12))
{
    import std.experimental.allocator: dispose;

    switch (pxloper.xltype & ~XlType.xlbitDLLFree) with(XlType) {
        case xltypeStr:
            if (pxloper.val.str !is null) {
                void* bytesPtr = pxloper.val.str;
                const numBytes = (pxloper.val.str[0] + 1) * wchar.sizeof;
                allocator.dispose(bytesPtr[0 .. numBytes]);
            }
            break;
        case xltypeRef:
            if (pxloper.val.mref.lpmref !is null)
                allocator.dispose(pxloper.val.mref.lpmref);
            break;
        case xltypeMulti:
            auto cxloper = pxloper.val.array.rows * pxloper.val.array.columns;
            const numOpers = cxloper;
            if (pxloper.val.array.lparray !is null)
            {
                auto pxloperFree = pxloper.val.array.lparray;
                while (cxloper > 0)
                {
                    freeXLOper(pxloperFree, allocator);
                    pxloperFree++;
                    cxloper--;
                }
                allocator.dispose(pxloper.val.array.lparray[0 .. numOpers]);
            }
            break;
        case xltypeBigData:
            if (pxloper.val.bigdata.h.lpbData !is null)
                allocator.dispose(pxloper.val.bigdata.h.lpbData);
            break;
        default: // todo: add error handling
            break;
    }
}

@("Free regular XLOPER")
unittest {
    import xlld.memorymanager: allocator;
    XLOPER oper;
    freeXLOper(&oper, allocator);
}

@("Free XLOPER12")
unittest {
    import xlld.memorymanager: allocator;
    XLOPER12 oper;
    freeXLOper(&oper, allocator);
}

/**
   Wrapper for the Excel12 function that allows passing D arrays

   Purpose:
        A fancy wrapper for the Excel12() function. It also
        does the following:

        (1) Checks that none of the LPXLOPER12 arguments are 0,
            which would indicate that creating a temporary XLOPER12
            has failed. In this case, it doesn't call Excel12
            but it does print a debug message.
        (2) If an error occurs while calling Excel12,
            print a useful debug message.
        (3) When done, free all temporary memory.

        #1 and #2 require _DEBUG to be defined.

   Parameters:

        int xlfn            Function number (xl...) to call
        LPXLOPER12 pxResult Pointer to a place to stuff the result,
                            or 0 if you don't care about the result.
        int count           Number of arguments
        ...                 (all LPXLOPER12s) - the arguments.

   Returns:

        A return code (Some of the xlret... values, as defined
        in XLCALL.H, OR'ed together).

   Comments:
*/

int Excel12f(int xlfn, LPXLOPER12 pxResult, LPXLOPER12[] args) nothrow @nogc
{
    import xlld.xlcallcpp: Excel12v;
    import std.algorithm: all;

    assert(args.all!(a => a !is null));
    return Excel12v(xlfn, pxResult, cast(int)args.length, cast(LPXLOPER12*)args.ptr);
}

int Excel12f(int xlfn, LPXLOPER12 pxResult, LPXLOPER12[] args...) nothrow @nogc {
    return Excel12f(xlfn, pxResult, args);
}

/**
   D version of Excel12f. "D version" in the sense that the types
   are all D types. Avoids having to manually convert to XLOPER12.
   e.g. excel12(xlfFoo, 1.0, 2.0);

   Returns a value of type T to specified at compile time which
   must be freed with xlld.memorymanager: autoFreeAllocator
 */
T excel12(T, A...)(int xlfn, auto ref A args) @trusted {
    import xlld.memorymanager: gTempAllocator, autoFreeAllocator;
    import xlld.wrap: toXlOper, fromXlOper;
    import xlld.xlcall: xlretSuccess;

    static immutable exception = new Exception("Error calling Excel12f");
    XLOPER12[A.length] operArgs;
    LPXLOPER12[A.length] operArgPtrs;

    scope(exit) gTempAllocator.deallocateAll;

    foreach(i, _; A) {
        operArgs[i] = args[i].toXlOper(gTempAllocator);
        operArgPtrs[i] = &operArgs[i];
    }

    XLOPER12 result;
    if(Excel12f(xlfn, &result, operArgPtrs) != xlretSuccess)
        throw exception;

    return result.fromXlOper!T(autoFreeAllocator);
}

/**
   Version of excel12 that avoids automatic memory management
   by asking the caller to supply a compile-time function
   to call on the result.
 */
auto excel12Then(alias F, A...)(int xlfn, auto ref A args) {
    import std.traits: Parameters, Unqual, isArray, ReturnType;
    import std.experimental.allocator: dispose;
    import xlld.memorymanager: autoFreeAllocator;

    static assert(Parameters!F.length == 1, "Must pass function of one argument");
    alias T = Parameters!F[0];
    auto excelRet = excel12!T(xlfn, args);
    static if(is(ReturnType!F == void))
        F(excelRet);
    else
        auto ret = F(excelRet);

    static if(isArray!T) {
        import std.range: ElementType;
        alias RealType = Unqual!(ElementType!T)[];
    } else
        alias RealType = Unqual!T;

    void freeRet(U)() @trusted {
        autoFreeAllocator.dispose(cast()excelRet);
    }

    static if(__traits(compiles, freeRet!RealType()))
        freeRet!RealType();

    static if(!is(ReturnType!F == void))
        return ret;
}
