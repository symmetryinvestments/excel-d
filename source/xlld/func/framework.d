/**
   Easier calling of Excel12f
 */
module xlld.func.framework;

///
__gshared immutable excel12Exception = new Exception("Error calling Excel12f");

/**
   D version of Excel12f. "D version" in the sense that the types
   are all D types. Avoids having to manually convert to XLOPER12.
   e.g. excel12(xlfFoo, 1.0, 2.0);

   Returns a value of type T to specified at compile time which
   must be freed with xlld.memorymanager: autoFreeAllocator
 */
T excel12(T, A...)(int xlfn, auto ref A args) @trusted {
    import xlld.memorymanager: gTempAllocator, autoFreeAllocator;
    import xlld.conv.from: fromXlOper;
    import xlld.conv: toXlOper;
    import xlld.sdk.xlcall: XLOPER12, LPXLOPER12, xlretSuccess;
    import xlld.sdk.framework: Excel12f;
    import std.meta: allSatisfy;
    import std.traits: Unqual;
    import std.experimental.allocator.gc_allocator: GCAllocator;
    import std.experimental.allocator.mallocator: Mallocator;

    // doubles never need the allocator anyway, so we use Mallocator
    // to guarantee @nogc when needed
    enum nogc(T) = is(Unqual!T == double);
    static if(allSatisfy!(nogc, A))
        alias allocator = Mallocator.instance;
    else
        alias allocator = GCAllocator.instance;

    XLOPER12[A.length] operArgs;
    LPXLOPER12[A.length] operArgPtrs;

    foreach(i, _; A) {
        operArgs[i] = args[i].toXlOper(allocator);
        operArgPtrs[i] = &operArgs[i];
    }

    XLOPER12 result;
    if(Excel12f(xlfn, &result, operArgPtrs) != xlretSuccess)
        throw excel12Exception;

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
        autoFreeAllocator.dispose(cast(U)excelRet);
    }

    static if(__traits(compiles, freeRet!RealType()))
        freeRet!RealType();

    static if(!is(ReturnType!F == void))
        return ret;
}
