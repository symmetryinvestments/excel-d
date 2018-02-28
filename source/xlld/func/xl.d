/**
 Wraps calls to xlXXX "functions" via the Excel4/Excel12 functions
 */
module xlld.func.xl;

import xlld.sdk.xlcall: XLOPER12, LPXLOPER12;


///
XLOPER12 coerce(in LPXLOPER12 oper) nothrow @nogc @trusted {
    import xlld.sdk.framework: Excel12f;
    import xlld.sdk.xlcall: xlCoerce;

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
    import xlld.sdk.framework: Excel12f;
    import xlld.sdk.xlcall: xlFree;

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
   Automatically calls xlfFree when destroyed
 */
struct ScopedOper {
    XLOPER12 oper;
    private bool _free;

    alias oper this;

    this(inout(XLOPER12) oper) @safe @nogc nothrow inout {
        this.oper = oper;
        _free = true;
    }

    this(inout(XLOPER12*) oper) @safe @nogc nothrow inout {
        this.oper = *oper;
        _free = true;
    }

    ~this() @safe @nogc nothrow {
        if(_free) free(oper);
    }
}
