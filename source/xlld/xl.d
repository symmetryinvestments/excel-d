/**
 Wraps calls to xlXXX "functions" via the Excel4/Excel12 functions
 */
module xlld.xl;

import xlld.xlcall: XLOPER12, LPXLOPER12;

version(unittest) {


}

XLOPER12 coerce(LPXLOPER12 oper) nothrow @nogc @trusted {
    import xlld.framework: Excel12f;
    import xlld.xlcall: xlCoerce;

    XLOPER12 coerced;
    LPXLOPER12[1] arg = [oper];
    Excel12f(xlCoerce, &coerced, arg);
    return coerced;
}

void free(ref XLOPER12 oper) nothrow @nogc @trusted {
    free(&oper);
}

void free(LPXLOPER12 oper) nothrow @nogc @trusted {
    import xlld.framework: Excel12f;
    import xlld.xlcall: xlFree;

    LPXLOPER12[1] arg = [oper];
    Excel12f(xlFree, null, arg);
}


struct Coerced {
    XLOPER12 oper;
    private bool coerced;

    alias oper this;


    this(LPXLOPER12 oper) @safe @nogc nothrow {
        this.oper = coerce(oper);
        this.coerced = true;
    }

    ~this() @safe @nogc nothrow {
        if(coerced) free(oper);
    }
}

/**
   Coerces an oper and returns an RAII struct that automatically frees memory
 */
auto scopedCoerce(LPXLOPER12 oper) @safe @nogc nothrow {
    return Coerced(oper);
}
