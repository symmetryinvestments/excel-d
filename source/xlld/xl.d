/**
 Wraps calls to xlXXX "functions" via the Excel4/Excel12 functions
 */
module xlld.xl;

import xlld.xlcall: XLOPER12, LPXLOPER12;

version(unittest) {


}

///
XLOPER12 coerce(in LPXLOPER12 oper) nothrow @nogc @trusted {
    import xlld.framework: Excel12f;
    import xlld.xlcall: xlCoerce;

    XLOPER12 coerced;
    Excel12f(xlCoerce, &coerced, oper);
    return coerced;
}

///
void free(ref const(XLOPER12) oper) nothrow @nogc @trusted {
    free(&oper);
}

///
void free(in LPXLOPER12 oper) nothrow @nogc @trusted {
    import xlld.framework: Excel12f;
    import xlld.xlcall: xlFree;

    const(XLOPER12)*[1] arg = [oper];
    Excel12f(xlFree, null, arg);
}


///
struct Coerced {
    XLOPER12 oper;
    private bool coerced;

    alias oper this;


    this(inout(XLOPER12)* oper) @safe @nogc nothrow inout {
        this.oper = coerce(oper);
        this.coerced = true;
    }

    this(inout(XLOPER12) oper) @safe @nogc nothrow inout {
        this.oper = () @trusted { return coerce(&oper); }();
        this.coerced = true;
    }

    ~this() @safe @nogc nothrow {
        if(coerced) free(oper);
    }
}

/**
   Coerces an oper and returns an RAII struct that automatically frees memory
 */
auto scopedCoerce(in LPXLOPER12 oper) @safe @nogc nothrow {
    return Coerced(oper);
}

auto scopedCoerce(in XLOPER12 oper) @trusted @nogc nothrow {
    return Coerced(oper);
}


alias coerced = scopedCoerce;
