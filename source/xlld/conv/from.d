/**
   Conversions from XLOPER12 to D types
 */
module xlld.conv.from;

import xlld.from;
import xlld.sdk.xlcall: XLOPER12;
import xlld.any: Any;
import std.traits: Unqual;
import std.datetime: DateTime;


alias ToEnumConversionFunction = int delegate(string);
package __gshared ToEnumConversionFunction[string] gToEnumConversions;
shared from!"core.sync.mutex".Mutex gToEnumMutex;


///
auto fromXlOper(T, A)(ref XLOPER12 val, ref A allocator) {
    return (&val).fromXlOper!T(allocator);
}

/// RValue overload
auto fromXlOper(T, A)(XLOPER12 val, ref A allocator) {
    return fromXlOper!T(val, allocator);
}

__gshared immutable fromXlOperDoubleWrongTypeException = new Exception("Wrong type for fromXlOper!double");
///
auto fromXlOper(T, A)(XLOPER12* val, ref A allocator) if(is(Unqual!T == double)) {
    import xlld.sdk.xlcall: XlType;
    import xlld.conv.misc: stripMemoryBitmask;

    if(val.xltype.stripMemoryBitmask == XlType.xltypeMissing)
        return double.init;

    if(val.xltype.stripMemoryBitmask == XlType.xltypeInt)
        return cast(T)val.val.w;

    if(val.xltype.stripMemoryBitmask != XlType.xltypeNum)
        throw fromXlOperDoubleWrongTypeException;

    return val.val.num;
}

__gshared immutable fromXlOperIntWrongTypeException = new Exception("Wrong type for fromXlOper!int");

///
auto fromXlOper(T, A)(XLOPER12* val, ref A allocator) if(is(Unqual!T == int)) {
    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.sdk.xlcall: XlType;

    if(val.xltype.stripMemoryBitmask == XlType.xltypeMissing)
        return int.init;

    if(val.xltype.stripMemoryBitmask == XlType.xltypeNum)
        return cast(typeof(return))val.val.num;

    if(val.xltype.stripMemoryBitmask != XlType.xltypeInt)
        throw fromXlOperIntWrongTypeException;

    return val.val.w;
}


///
__gshared immutable fromXlOperMemoryException = new Exception("Could not allocate memory for array of char");
///
__gshared immutable fromXlOperConvException = new Exception("Could not convert double to string");

__gshared immutable fromXlOperStringTypeException = new Exception("Wrong type for fromXlOper!string");

///
auto fromXlOper(T, A)(XLOPER12* val, ref A allocator) if(is(Unqual!T == string)) {

    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.sdk.xlcall: XlType;
    import std.experimental.allocator: makeArray;
    import std.utf: byChar;
    import std.range: walkLength;

    const stripType = stripMemoryBitmask(val.xltype);

    if(stripType == XlType.xltypeMissing)
        return null;

    if(stripType != XlType.xltypeStr && stripType != XlType.xltypeNum)
        throw fromXlOperStringTypeException;


    if(stripType == XlType.xltypeStr) {

        auto chars = () @trusted { return val.val.str[1 .. val.val.str[0] + 1].byChar; }();
        const length = chars.save.walkLength;
        auto ret = () @trusted { return allocator.makeArray!char(length); }();

        if(ret is null && length > 0)
            throw fromXlOperMemoryException;

        int i;
        foreach(ch; () @trusted { return val.val.str[1 .. val.val.str[0] + 1].byChar; }())
            ret[i++] = ch;

        return () @trusted {  return cast(string)ret; }();
    } else {

        // if a double, try to convert it to a string
        import std.math: isNaN;
        import core.stdc.stdio: snprintf;

        char[1024] buffer;
        const numChars = () @trusted {
            if(val.val.num.isNaN)
                return snprintf(&buffer[0], buffer.length, "#NaN");
            else
                return snprintf(&buffer[0], buffer.length, "%lf", val.val.num);
        }();
        if(numChars > buffer.length - 1)
            throw fromXlOperConvException;
        auto ret = () @trusted { return allocator.makeArray!char(numChars); }();

        if(ret is null && numChars > 0)
            throw fromXlOperMemoryException;

        ret[] = buffer[0 .. numChars];
        return () @trusted { return cast(string)ret; }();
    }
}


///
T fromXlOper(T, A)(XLOPER12* oper, ref A allocator) if(is(Unqual!T == Any)) {
    import xlld.conv.misc: dup;
    return Any((*oper).dup(allocator));
}


///
auto fromXlOper(T, A)(XLOPER12* val, ref A allocator)
    if(is(T: E[][], E) &&
       (is(Unqual!E == string) || is(Unqual!E == double) || is(Unqual!E == int)
        || is(Unqual!E == Any) || is(Unqual!E == DateTime)))
{
    return val.fromXlOperMulti!(Dimensions.Two, typeof(T.init[0][0]))(allocator);
}



private enum Dimensions {
    One,
    Two,
}


/// 1D slices
auto fromXlOper(T, A)(XLOPER12* val, ref A allocator)
    if(is(T: E[], E) &&
       (is(Unqual!E == string) || is(Unqual!E == double) || is(Unqual!E == int)
        || is(Unqual!E == Any) || is(Unqual!E == DateTime)))
{
    return val.fromXlOperMulti!(Dimensions.One, typeof(T.init[0]))(allocator);
}



///
__gshared immutable fromXlOperMultiOperException = new Exception("fromXlOper: oper not of multi type");
///
__gshared immutable fromXlOperMultiMemoryException = new Exception("fromXlOper: Could not allocate memory in fromXlOperMulti");

private auto fromXlOperMulti(Dimensions dim, T, A)(XLOPER12* val, ref A allocator) {
    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.func.xl: coerce, free;
    import xlld.memorymanager: makeArray2D;
    import xlld.sdk.xlcall: XlType;
    import std.experimental.allocator: makeArray;

    if(stripMemoryBitmask(val.xltype) == XlType.xltypeNil) {
        static if(dim == Dimensions.Two)
            return T[][].init;
        else static if(dim == Dimensions.One)
            return T[].init;
        else
            static assert(0, "Unknown number of dimensions in fromXlOperMulti");
    }

    if(stripMemoryBitmask(val.xltype) == XlType.xltypeNum) {
        static if(dim == Dimensions.Two) {
            import std.experimental.allocator: makeMultidimensionalArray;
            auto ret = allocator.makeMultidimensionalArray!T(1, 1);
            ret[0][0] = val.fromXlOper!T(allocator);
            return ret;
        } else static if(dim == Dimensions.One) {
            auto ret = allocator.makeArray!T(1);
            ret[0] = val.fromXlOper!T(allocator);
            return ret;
        } else
            static assert(0, "Unknown number of dimensions in fromXlOperMulti");
    }

    if(!isMulti(*val)) {
        throw fromXlOperMultiOperException;
    }

    const rows = val.val.array.rows;
    const cols = val.val.array.columns;

    assert(rows > 0 && cols > 0, "Multi opers may not have 0 rows or columns");

    static if(dim == Dimensions.Two) {
        auto ret = allocator.makeArray2D!T(*val);
    } else static if(dim == Dimensions.One) {
        auto ret = allocator.makeArray!T(rows * cols);
    } else
        static assert(0, "Unknown number of dimensions in fromXlOperMulti");

    if(&ret[0] is null)
        throw fromXlOperMultiMemoryException;

    (*val).apply!(T, (shouldConvert, row, col, cellVal) {

        auto value = shouldConvert ? cellVal.fromXlOper!T(allocator) : T.init;

        static if(dim == Dimensions.Two)
            ret[row][col] = value;
        else
            ret[row * cols + col] = value;
    });

    return ret;
}


// apply a function to an oper of type xltypeMulti
// the function must take a boolean value indicating if the cell value
// is to be converted or not, the row index, the column index,
// and a reference to the cell value itself
private void apply(T, alias F)(ref XLOPER12 oper) {
    import xlld.sdk.xlcall: XlType;
    import xlld.func.xl: coerce, free;
    import xlld.any: Any;
    version(unittest) import xlld.test.util: gNumXlAllocated, gNumXlFree;

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;
    auto values = oper.val.array.lparray[0 .. (rows * cols)];

    foreach(const row; 0 .. rows) {
        foreach(const col; 0 .. cols) {

            auto cellVal = coerce(&values[row * cols + col]);

            // Issue 22's unittest ends up coercing more than xlld.test.util can handle
            // so we undo the side-effect here
            version(unittest) --gNumXlAllocated; // ignore this for testing

            scope(exit) {
                free(&cellVal);
                // see comment above about gNumXlCoerce
                version(unittest) --gNumXlFree;
            }

            // try to convert doubles to string if trying to convert everything to an
            // array of strings
            const shouldConvert =
                (cellVal.xltype == dlangToXlOperType!T.Type) ||
                (cellVal.xltype == XlType.xltypeNum && dlangToXlOperType!T.Type == XlType.xltypeStr) ||
                is(Unqual!T == Any);


            F(shouldConvert, row, col, cellVal);
        }
    }
}


__gshared immutable fromXlOperDateTimeTypeException = new Exception("Wrong type for fromXlOper!DateTime");

///
T fromXlOper(T, A)(XLOPER12* oper, ref A allocator)
    if(is(Unqual!T == DateTime))
{
    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.sdk.framework: Excel12f;
    import xlld.sdk.xlcall: XlType, xlretSuccess, xlfYear, xlfMonth, xlfDay, xlfHour, xlfMinute, xlfSecond;

    if(oper.xltype.stripMemoryBitmask != XlType.xltypeNum)
        throw fromXlOperDateTimeTypeException;

    XLOPER12 ret;

    auto get(int fn) @trusted {
        const code = Excel12f(fn, &ret, oper);
        if(code != xlretSuccess)
            throw new Exception("Error calling xlf datetime part function");

        // for some reason the Excel API returns doubles
        assert(ret.xltype == XlType.xltypeNum, "xlf datetime part return not xltypeNum");
        return cast(int)ret.val.num;
    }

    return T(get(xlfYear), get(xlfMonth), get(xlfDay),
             get(xlfHour), get(xlfMinute), get(xlfSecond));
}

T fromXlOper(T, A)(XLOPER12* oper, ref A allocator) if(is(Unqual!T == bool)) {

    import xlld.sdk.xlcall: XlType;
    import std.uni: toLower;

    if(oper.xltype == XlType.xltypeStr) {
        return oper.fromXlOper!string(allocator).toLower == "true";
    }

    return cast(T)oper.val.bool_;
}

T fromXlOper(T, A)(XLOPER12* oper, ref A allocator) if(is(T == enum)) {
    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.sdk.xlcall: XlType;
    import std.conv: to;
    import std.traits: fullyQualifiedName;

    static immutable typeException = new Exception("Wrong type for fromXlOper!" ~ T.stringof);
    if(oper.xltype.stripMemoryBitmask != XlType.xltypeStr)
        throw typeException;

    enum name = fullyQualifiedName!T;
    auto str = oper.fromXlOper!string(allocator);

    return () @trusted {
        assert(gToEnumMutex !is null, "gToEnumMutex is null");

        gToEnumMutex.lock_nothrow;
        scope(exit) gToEnumMutex.unlock_nothrow;

        return name in gToEnumConversions
                           ? cast(T) gToEnumConversions[name](str)
                           : str.to!T;
    }();
}


T fromXlOper(T, A)(XLOPER12* oper, ref A allocator)
    if(is(T == struct) && !is(Unqual!T == Any) && !is(Unqual!T == DateTime))
{
    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.sdk.xlcall: XlType;
    import std.conv: text;
    import std.exception: enforce;

    static immutable multiException = new Exception("Can only convert arrays to structs. Must be either 1xN, Nx1, 2xN or Nx2");
    if(oper.xltype.stripMemoryBitmask != XlType.xltypeMulti)
        throw multiException;

    const length =  oper.val.array.rows * oper.val.array.columns;

    if(oper.val.array.rows == 1 || oper.val.array.columns == 1)
        enforce(length == T.tupleof.length,
               text("1D array length must match number of members in ", T.stringof,
                    ". Expected ", T.tupleof.length, ", got ", length));
    else
        enforce((oper.val.array.rows == 2 && oper.val.array.columns == T.tupleof.length) ||
               (oper.val.array.rows == T.tupleof.length && oper.val.array.columns == 2),
               text("2D array must be 2x", T.tupleof.length, " or ", T.tupleof.length, "x2 for ", T.stringof));

    T ret;

    size_t ptrIndex(size_t i) {

        if(oper.val.array.rows == 1 || oper.val.array.columns == 1)
            return i;

        if(oper.val.array.rows == 2)
            return i + oper.val.array.columns;

        if(oper.val.array.columns == 2)
            return i * 2 + 1;

        assert(0);
    }

    static immutable wrongTypeException = new Exception("Wrong type converting oper to " ~ T.stringof);

    foreach(i, ref member; ret.tupleof) {
        try
            member = oper.val.array.lparray[ptrIndex(i)].fromXlOper!(typeof(member))(allocator);
        catch(Exception _)
            throw wrongTypeException;
    }

    return ret;
}


///
auto fromXlOperCoerce(T)(XLOPER12* val) {
    return fromXlOperCoerce(*val);
}

///
auto fromXlOperCoerce(T, A)(XLOPER12* val, auto ref A allocator) {
    return fromXlOperCoerce!T(*val, allocator);
}


///
auto fromXlOperCoerce(T)(ref XLOPER12 val) {
    import xlld.memorymanager: allocator;
    return fromXlOperCoerce!T(val, allocator);
}


///
auto fromXlOperCoerce(T, A)(ref XLOPER12 val, auto ref A allocator) {
    import xlld.func.xl: coerce, free;

    auto coerced = coerce(&val);
    scope(exit) free(&coerced);

    return coerced.fromXlOper!T(allocator);
}


private enum invalidXlOperType = 0xdeadbeef;

/**
 Maps a D type to two integer xltypes from XLOPER12.
 InputType is the type actually passed in by the spreadsheet,
 whilst Type is the Type that it gets coerced to.
 */
template dlangToXlOperType(T) {
    import xlld.sdk.xlcall: XlType;
    static if(is(Unqual!T == string[])   || is(Unqual!T == string[][]) ||
              is(Unqual!T == double[])   || is(Unqual!T == double[][]) ||
              is(Unqual!T == int[])      || is(Unqual!T == int[][]) ||
              is(Unqual!T == DateTime[]) || is(Unqual!T == DateTime[][]))
    {
        enum InputType = XlType.xltypeSRef;
        enum Type = XlType.xltypeMulti;
    } else static if(is(Unqual!T == double)) {
        enum InputType = XlType.xltypeNum;
        enum Type = XlType.xltypeNum;
    } else static if(is(Unqual!T == string)) {
        enum InputType = XlType.xltypeStr;
        enum Type = XlType.xltypeStr;
    } else static if(is(Unqual!T == DateTime)) {
        enum InputType = XlType.xltypeNum;
        enum Type = XlType.xltypeNum;
    } else {
        enum InputType = invalidXlOperType;
        enum Type = invalidXlOperType;
    }
}

/**
   If an oper is of multi type
 */
bool isMulti(ref const(XLOPER12) oper) @safe @nogc pure nothrow {
    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.sdk.xlcall: XlType;

    return stripMemoryBitmask(oper.xltype) == XlType.xltypeMulti;
}

/**
   Register a custom conversion from string to an enum type. This function will
   be called before converting any enum arguments to be passed to a wrapped
   D function.
 */
void registerConversionTo(T)(ToEnumConversionFunction func) @trusted {
    import std.traits: fullyQualifiedName;

    assert(gToEnumMutex !is null, "gToEnumMutex is null");

    gToEnumMutex.lock_nothrow;
    scope(exit)gToEnumMutex.unlock_nothrow;

    gToEnumConversions[fullyQualifiedName!T] = func;
}

void unregisterConversionTo(T)() @trusted {
    import std.traits: fullyQualifiedName;

    assert(gToEnumMutex !is null, "gToEnumMutex is null");

    gToEnumMutex.lock_nothrow;
    scope(exit)gToEnumMutex.unlock_nothrow;

    gToEnumConversions.remove(fullyQualifiedName!T);
}
