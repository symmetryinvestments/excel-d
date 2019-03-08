/**
   Conversions from D types to XLOPER12
 */
module xlld.conv.to;

import xlld.from;
import xlld.conv.misc: isUserStruct, isVector;
import xlld.sdk.xlcall: XLOPER12, XlType;
import xlld.any: Any;
import xlld.wrap.wrap: isWantedType;
import std.traits: isIntegral, Unqual;
import std.datetime: DateTime;
import std.typecons: Tuple;


alias FromEnumConversionFunction = string delegate(int) @safe;
package __gshared FromEnumConversionFunction[string] gFromEnumConversions;
shared from!"core.sync.mutex".Mutex gFromEnumMutex;


///
XLOPER12 toXlOper(T)(T val) {
    import std.experimental.allocator: theAllocator;
    return toXlOper(val, theAllocator);
}


///
XLOPER12 toXlOper(T, A)(in T val, ref A allocator) if(isIntegral!T) {
    import xlld.sdk.xlcall: XlType;

    auto ret = XLOPER12();
    ret.xltype = XlType.xltypeInt;
    ret.val.w = val;

    return ret;
}


///
XLOPER12 toXlOper(T, A)(in T val, ref A allocator) if(is(Unqual!T == double)) {
    import xlld.sdk.xlcall: XlType;

    auto ret = XLOPER12();
    ret.xltype = XlType.xltypeNum;
    ret.val.num = val;

    return ret;
}


///
__gshared immutable toXlOperMemoryException = new Exception("Failed to allocate memory for string oper");


///
XLOPER12 toXlOper(T, A)(in T val, ref A allocator)
    if(is(Unqual!T == string) || is(Unqual!T == wstring))
{
    import xlld.sdk.xlcall: XCHAR;
    import std.utf: byWchar;

    const numBytes = numOperStringBytes(val);
    auto wval = () @trusted { return cast(wchar[]) allocator.allocate(numBytes); }();
    if(&wval[0] is null)
        throw toXlOperMemoryException;

    int i = 1;
    foreach(ch; val.byWchar) {
        wval[i++] = ch;
    }

    wval[0] = cast(ushort)(i - 1);

    auto ret = XLOPER12();
    ret.xltype = XlType.xltypeStr;
    () @trusted { ret.val.str = cast(XCHAR*)&wval[0]; }();

    return ret;
}



/// the number of bytes required to store `str` as an XLOPER12 string
package size_t numOperStringBytes(T)(in T str) if(is(Unqual!T == string) || is(Unqual!T == wstring)) {
    // XLOPER12 strings are wide strings where index 0 is the length
    // and [1 .. $] is the actual string
    return (str.length + 1) * wchar.sizeof;
}


private template isRange1D(T) {
    import std.traits: isSomeString;
    import std.range.primitives: isForwardRange, ElementType;

    enum isRange1D =
        isForwardRange!T
        && (!isForwardRange!(ElementType!T) || isSomeString!(ElementType!T))
        && !isVector!T
        && !isSomeString!T
        ;
}


private template isRange2D(T) {
    import std.traits: isSomeString;
    import std.range.primitives: isForwardRange, ElementType;

    enum isRange2D =
        isForwardRange!T
        && isForwardRange!(ElementType!T)
        && !isVector!T
        && !isSomeString!(ElementType!T)
        ;
}


XLOPER12 toXlOper(T, A)(T range, ref A allocator)
    if(isRange1D!T)
{
    import std.range: only;
    return only(range).toXlOper(allocator);
}

XLOPER12 toXlOper(T, A)(T range, ref A allocator)
    if(isRange2D!T)
{
    import xlld.conv.misc: multi;
    import std.range: walkLength;
    import std.array: save, front;
    import std.algorithm: any;

    static __gshared immutable shapeException = new Exception("# of columns must all be the same and aren't");
    const rows = cast(int) range.save.walkLength;
    const frontLength = range.front.save.walkLength;

    if(range.save.any!(r => r.save.walkLength != frontLength))
        throw shapeException;

    const cols = cast(int) range.front.save.walkLength;
    auto ret = multi(rows, cols, allocator);
    auto opers = () @trusted { return ret.val.array.lparray[0 .. rows*cols]; }();

    int i = 0;
    foreach(ref subRange; range) {
        foreach(ref elt; subRange) {
            opers[i++] = elt.toXlOper(allocator);
        }
    }

    return ret;
}


XLOPER12 toXlOper(T, A)(T value, ref A allocator)
    if(isVector!T)
{
    import std.experimental.allocator: makeArray, dispose;

    enum is2D = isVector!(typeof(value[0]));

    static if(is2D) {
        alias E = typeof(value[0][0]);
        assert(value.length <= size_t.sizeof);

        auto arr = allocator.makeArray!(E[])(cast(size_t) value.length);
        scope(exit) allocator.dispose(arr);

        foreach(i; 0 .. value.length) {
            arr[cast(size_t) i] = value[i][];
        }

        return arr.toXlOper(allocator);

    } else {
        return value[].toXlOper(allocator);
    }
}

///
XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(Unqual!T == Any)) {
    import xlld.conv.misc: dup;
    return value.dup(allocator);
}


XLOPER12 toXlOper(T, A)(T value, ref A allocator, in string file = __FILE__, in size_t line = __LINE__)
    @safe if(is(Unqual!T == DateTime))
{
    import xlld.sdk.framework: Excel12f;
    import xlld.sdk.xlcall: xlfDate, xlfTime, xlretSuccess;

    XLOPER12 ret, date, time;

    auto year = value.year.toXlOper(allocator);
    auto month = () @trusted { return toXlOper(cast(int)value.month, allocator); }();
    auto day = value.day.toXlOper(allocator);

    () @trusted {
        assert(year.xltype == XlType.xltypeInt);
        assert(month.xltype == XlType.xltypeInt);
        assert(day.xltype == XlType.xltypeInt);
    }();

    const dateCode = () @trusted { return Excel12f(xlfDate, &date, &year, &month, &day); }();
    if(dateCode != xlretSuccess)
        throw new Exception("Error calling xlfDate", file, line);
    () @trusted { assert(date.xltype == XlType.xltypeNum); }();

    auto hour = value.hour.toXlOper(allocator);
    auto minute = value.minute.toXlOper(allocator);
    auto second = value.second.toXlOper(allocator);

    const timeCode = () @trusted { return Excel12f(xlfTime, &time, &hour, &minute, &second); }();
    if(timeCode != xlretSuccess)
        throw new Exception("Error calling xlfTime", file, line);

    () @trusted { assert(time.xltype == XlType.xltypeNum); }();

    ret.xltype = XlType.xltypeNum;
    ret.val.num = date.val.num + time.val.num;
    return ret;
}


XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(Unqual!T == bool)) {
    import xlld.sdk.xlcall: XlType;
    XLOPER12 ret;
    ret.xltype = XlType.xltypeBool;
    ret.val.bool_ = cast(typeof(ret.val.bool_)) value;
    return ret;
}

/**
   Register a custom conversion from enum (going through integer) to a string.
   This function will be called to convert enum return values from wrapped
   D functions into strings in Excel.
 */
void registerConversionFrom(T)(FromEnumConversionFunction func) @trusted {
    import std.traits: fullyQualifiedName;

    assert(gFromEnumMutex !is null, "gFromEnumMutex is null");

    gFromEnumMutex.lock_nothrow;
    scope(exit)gFromEnumMutex.unlock_nothrow;

    gFromEnumConversions[fullyQualifiedName!T] = func;
}


void unregisterConversionFrom(T)() @trusted {
    import std.traits: fullyQualifiedName;

    assert(gFromEnumMutex !is null, "gFromEnumMutex is null");

    gFromEnumMutex.lock_nothrow;
    scope(exit)gFromEnumMutex.unlock_nothrow;

    gFromEnumConversions.remove(fullyQualifiedName!T);
}




XLOPER12 toXlOper(T, A)(T value, ref A allocator) @trusted if(is(T == enum)) {

    import std.conv: text;
    import std.traits: fullyQualifiedName;
    import core.memory: GC;

    enum name = fullyQualifiedName!T;

    {
        assert(gFromEnumMutex !is null, "gFromEnumMutex is null");

        gFromEnumMutex.lock_nothrow;
        scope(exit) gFromEnumMutex.unlock_nothrow;

        if(name in gFromEnumConversions)
            return gFromEnumConversions[name](value).toXlOper(allocator);
    }

    scope str = text(value);
    auto ret = str.toXlOper(allocator);
    () @trusted { GC.free(cast(void*)str.ptr); }();

    return ret;
}

XLOPER12 toXlOper(T, A)(T value, ref A allocator)
    if(isUserStruct!T)
{
    import std.conv: text;
    import core.memory: GC;

    scope str = text(value);

    auto ret = str.toXlOper(allocator);
    () @trusted { GC.free(cast(void*)str.ptr); }();

    return ret;
}

XLOPER12 toXlOper(T, A)(T value, ref A allocator)
    @trusted
    if(is(T: Tuple!A, A...))
{
    import std.experimental.allocator: makeArray;

    XLOPER12 oper;

    oper.xltype = XlType.xltypeMulti;
    oper.val.array.rows = 1;
    oper.val.array.columns = value.length;
    oper.val.array.lparray = allocator.makeArray!XLOPER12(T.length).ptr;

    static foreach(i; 0 .. T.length) {
        oper.val.array.lparray[i] = value[i].toXlOper(allocator);
    }

    return oper;
}


XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(Unqual!T == XLOPER12))
{
    return value;
}


/**
  creates an XLOPER12 that can be returned to Excel which
  will be freed by Excel itself
 */
XLOPER12 toAutoFreeOper(T)(T value) {
    import xlld.memorymanager: autoFreeAllocator;
    import xlld.sdk.xlcall: XlType;

    auto result = value.toXlOper(autoFreeAllocator);
    result.xltype |= XlType.xlbitDLLFree;
    return result;
}
