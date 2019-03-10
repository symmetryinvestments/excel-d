/**
   Miscelleanous functions that assist in type conversions.
 */
module xlld.conv.misc;

import xlld.from;

///
template isUserStruct(T) {
    import xlld.any: Any;
    import xlld.sdk.xlcall: XLOPER12;
    import std.datetime: DateTime;
    import std.typecons: Tuple;
    import std.traits: Unqual;
    import std.range.primitives: isInputRange;

    enum isUserStruct =
        is(T == struct)
        && !is(Unqual!T == Any)
        && !is(Unqual!T == DateTime)
        && !is(Unqual!T: Tuple!A, A...)
        && !is(Unqual!T == XLOPER12)
        && !isVector!T
        && !isInputRange!T
        ;
}

template isVector(T) {
    version(Have_automem) {
        import automem.vector: Vector;
        import std.traits: Unqual;
        enum isVector = is(Unqual!T == Vector!(E,  A), E, A);
    } else
        enum isVector = false;
}


///
__gshared immutable gDupMemoryException = new Exception("Failed to allocate memory in dup");


/**
   Deep copy of an oper
 */
from!"xlld.sdk.xlcall".XLOPER12 dup(A)(from!"xlld.sdk.xlcall".XLOPER12 oper, ref A allocator) @safe {

    import xlld.sdk.xlcall: XLOPER12, XlType;
    import std.experimental.allocator: makeArray;

    XLOPER12 ret;

    ret.xltype = oper.xltype;

    switch(stripMemoryBitmask(oper.xltype)) with(XlType) {

        default:
            ret = oper;
            return ret;

        case xltypeStr:
            const length = operStringLength(oper) + 1;

            () @trusted {
                ret.val.str = allocator.makeArray!wchar(length).ptr;
                if(ret.val.str is null)
                    throw gDupMemoryException;
            }();

            () @trusted { ret.val.str[0 .. length] = oper.val.str[0 .. length]; }();
            return ret;

        case xltypeMulti:
            () @trusted {
                ret.val.array.rows = oper.val.array.rows;
                ret.val.array.columns = oper.val.array.columns;
                const length = oper.val.array.rows * oper.val.array.columns;
                ret.val.array.lparray = allocator.makeArray!XLOPER12(length).ptr;

                if(ret.val.array.lparray is null)
                    throw gDupMemoryException;

                foreach(i; 0 .. length) {
                    ret.val.array.lparray[i] = oper.val.array.lparray[i].dup(allocator);
                }
            }();

            return ret;
    }

    assert(0);
}



from!"xlld.sdk.xlcall".XlType stripMemoryBitmask(in from!"xlld.sdk.xlcall".XlType type) @safe @nogc pure nothrow {
    import xlld.sdk.xlcall: XlType, xlbitXLFree, xlbitDLLFree;
    return cast(XlType)(type & ~(xlbitXLFree | xlbitDLLFree));
}

///
ushort operStringLength(T)(in T value) {
    import xlld.sdk.xlcall: XlType;
    import nogc.exception: enforce;

    enforce(value.xltype == XlType.xltypeStr,
            "Cannot calculate string length for oper of type ", cast(int) value.xltype);

    return cast(ushort)value.val.str[0];
}


// can't be pure because to!double isn't pure
string toString(in from!"xlld.sdk.xlcall".XLOPER12 oper) @safe {
    import xlld.sdk.xlcall: XlType;
    import xlld.conv.misc: stripMemoryBitmask;
    import std.conv: text;
    import std.format: format;

    string ret;

    ret ~= "XLOPER12(";
    switch(stripMemoryBitmask(oper.xltype)) {
    default:
        ret ~= oper.xltype.stripMemoryBitmask.text;
        break;

    case XlType.xltypeSRef:
        import xlld.func.xl: Coerced;
        auto coerced = () @trusted { return Coerced(&oper); }();
        return "SRef[ " ~ coerced.toString ~ " ]";

    case XlType.xltypeNum:
        ret ~= format!"%.6f"(oper.val.num);
        break;

    case XlType.xltypeStr:
        ret ~= `"`;
        () @trusted {
            const ulong length = oper.val.str[0];
            ret ~= text(oper.val.str[1 .. 1 + length]);
        }();
        ret ~= `"`;
        break;

    case XlType.xltypeInt:
        ret ~= text(oper.val.w);
        break;

    case XlType.xltypeBool:
        ret ~= text(oper.val.bool_);
        break;

    case XlType.xltypeErr:
        ret ~= "ERROR";
        break;

    case XlType.xltypeBigData:
        () @trusted {
            ret ~= "BIG(";
            ret ~= text(oper.val.bigdata.h.hdata);
            ret ~= ", ";
            ret ~= text(oper.val.bigdata.cbData);
            ret ~= ")";
        }();
    }
    ret ~= ")";
    return ret;
}

///
__gshared immutable multiMemoryException = new Exception("Failed to allocate memory for multi oper");

/// Returns an array XLOPER12
from!"xlld.sdk.xlcall".XLOPER12 multi(A)(int rows, int cols, ref A allocator) @trusted {
    import xlld.sdk.xlcall: XLOPER12, XlType;

    auto ret = XLOPER12();

    ret.xltype = XlType.xltypeMulti;
    () @trusted {
        ret.val.array.rows = rows;
        ret.val.array.columns = cols;
    }();

    auto slice = allocator.allocate(rows * cols * ret.sizeof);
    ret.val.array.lparray = () @trusted { return cast(XLOPER12*) slice.ptr; }();

    if(ret.val.array.lparray is null)
        throw multiMemoryException;

    return ret;
}
