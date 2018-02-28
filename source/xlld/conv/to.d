/**
   Conversions from D types to XLOPER12
 */
module xlld.conv.to;

import xlld.sdk.xlcall: XLOPER12, XlType;
import xlld.any: Any;
import std.traits: isIntegral, Unqual;
import std.datetime: DateTime;
import core.sync.mutex: Mutex;

version(unittest) {
    import xlld.any: any;
    import xlld.sdk.framework: freeXLOper;
    import xlld.test_util: TestAllocator, shouldEqualDlang,
        MockXlFunction, FailingAllocator;
    import unit_threaded;
    import std.experimental.allocator.gc_allocator: GCAllocator;
    alias theGC = GCAllocator.instance;
}
alias FromEnumConversionFunction = string delegate(int) @safe;
package __gshared FromEnumConversionFunction[string] gFromEnumConversions;
package shared Mutex gFromEnumMutex;


///
XLOPER12 toXlOper(T, A)(in T val, ref A allocator) if(isIntegral!T) {
    auto ret = XLOPER12();
    ret.xltype = XlType.xltypeInt;
    ret.val.w = val;
    return ret;
}

///
@("toExcelOper!int")
unittest {
    auto oper = 42.toXlOper(theGC);
    oper.xltype.shouldEqual(XlType.xltypeInt);
    oper.val.w.shouldEqual(42);
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
@("toExcelOper!double")
unittest {
    auto oper = (42.0).toXlOper(theGC);
    oper.xltype.shouldEqual(XlType.xltypeNum);
    oper.val.num.shouldEqual(42.0);
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
    auto wval = () @trusted { return cast(wchar[])allocator.allocate(numBytes); }();
    if(wval is null)
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

///
@("toXlOper!string utf8")
@system unittest {
    import std.conv: to;
    import xlld.memorymanager: allocator;

    const str = "foo";
    auto oper = str.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);

    oper.xltype.shouldEqual(XlType.xltypeStr);
    (cast(int)oper.val.str[0]).shouldEqual(str.length);
    (cast(wchar*)oper.val.str)[1 .. str.length + 1].to!string.shouldEqual(str);
}

///
@("toXlOper!string utf16")
@system unittest {
    import xlld.memorymanager: allocator;

    const str = "foo"w;
    auto oper = str.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);

    oper.xltype.shouldEqual(XlType.xltypeStr);
    (cast(int)oper.val.str[0]).shouldEqual(str.length);
    (cast(wchar*)oper.val.str)[1 .. str.length + 1].shouldEqual(str);
}

///
@("toXlOper!string TestAllocator")
@system unittest {
    auto allocator = TestAllocator();
    auto oper = "foo".toXlOper(allocator);
    allocator.numAllocations.shouldEqual(1);
    freeXLOper(&oper, allocator);
}

///
@("toXlOper!string unicode")
@system unittest {
    import std.utf: byWchar;
    import std.array: array;

    "é".byWchar.array.length.shouldEqual(1);
    "é"w.byWchar.array.length.shouldEqual(1);

    auto oper = "é".toXlOper(theGC);
    const ushort length = oper.val.str[0];
    length.shouldEqual("é"w.length);
}

@("toXlOper!string failing allocator")
@safe unittest {
    import xlld.conv.misc: dup;
    auto allocator = FailingAllocator();
    "foo".toXlOper(theGC).dup(allocator).shouldThrowWithMessage("Failed to allocate memory for string oper");
}

/// the number of bytes required to store `str` as an XLOPER12 string
package size_t numOperStringBytes(T)(in T str) if(is(Unqual!T == string) || is(Unqual!T == wstring)) {
    // XLOPER12 strings are wide strings where index 0 is the length
    // and [1 .. $] is the actual string
    return (str.length + 1) * wchar.sizeof;
}


///
__gshared immutable toXlOperShapeException = new Exception("# of columns must all be the same and aren't");


///
XLOPER12 toXlOper(T, A)(T[][] values, ref A allocator)
    if(is(Unqual!T == double) || is(Unqual!T == string) || is(Unqual!T == Any)
       || is(Unqual!T == int) || is(Unqual!T == DateTime))
{
    import xlld.conv.misc: multi;
    import std.algorithm: map, all;
    import std.array: array;

    if(!values.all!(a => a.length == values[0].length))
       throw toXlOperShapeException;

    const rows = cast(int)values.length;
    const cols = values.length ? cast(int)values[0].length : 0;
    auto ret = multi(rows, cols, allocator);
    auto opers = () @trusted { return ret.val.array.lparray[0 .. rows*cols]; }();

    int i;
    foreach(ref row; values) {
        foreach(ref val; row) {
            opers[i++] = val.toXlOper(allocator);
        }
    }

    return ret;
}


///
@("toXlOper string[][]")
@system unittest {
    import xlld.memorymanager: allocator;

    auto oper = [["foo", "bar", "baz"], ["toto", "titi", "quux"]].toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);

    oper.xltype.shouldEqual(XlType.xltypeMulti);
    oper.val.array.rows.shouldEqual(2);
    oper.val.array.columns.shouldEqual(3);
    auto opers = oper.val.array.lparray[0 .. oper.val.array.rows * oper.val.array.columns];

    opers[0].shouldEqualDlang("foo");
    opers[3].shouldEqualDlang("toto");
    opers[5].shouldEqualDlang("quux");
}

///
@("toXlOper string[][] TestAllocator")
@system unittest {
    TestAllocator allocator;
    auto oper = [["foo", "bar", "baz"], ["toto", "titi", "quux"]].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(7);
    freeXLOper(&oper, allocator);
}

///
@("toXlOper double[][]")
@system unittest {
    TestAllocator allocator;
    auto oper = [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(1);
    freeXLOper(&oper, allocator);
}

@("toXlOper!double[][] failing allocation")
@safe unittest {
    auto allocator = FailingAllocator();
    [33.3].toXlOper(allocator).shouldThrowWithMessage("Failed to allocate memory for multi oper");
}

@("toXlOper!double[][] wrong shape")
@safe unittest {
    [[33.3], [1.0, 2.0]].toXlOper(theGC).shouldThrowWithMessage("# of columns must all be the same and aren't");
}


///
XLOPER12 toXlOper(T, A)(T values, ref A allocator) if(is(Unqual!T == string[]) || is(Unqual!T == double[]) ||
                                                      is(Unqual!T == int[]) || is(Unqual!T == DateTime[])) {
    T[1] realValues = [values];
    return realValues[].toXlOper(allocator);
}


///
@("toXlOper string[]")
@system unittest {
    TestAllocator allocator;
    auto oper = ["foo", "bar", "baz", "toto", "titi", "quux"].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(7);
    freeXLOper(&oper, allocator);
}

///
XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(Unqual!T == Any)) {
    import xlld.conv.misc: dup;
    return value.dup(allocator);
}

///
@("toXlOper any double")
unittest {
    any(5.0, theGC).toXlOper(theGC).shouldEqualDlang(5.0);
}

///
@("toXlOper any string")
unittest {
    any("foo", theGC).toXlOper(theGC).shouldEqualDlang("foo");
}

///
@("toXlOper any double[][]")
unittest {
    any([[1.0, 2.0], [3.0, 4.0]], theGC)
        .toXlOper(theGC).shouldEqualDlang([[1.0, 2.0], [3.0, 4.0]]);
}

///
@("toXlOper any string[][]")
unittest {
    any([["foo", "bar"], ["quux", "toto"]], theGC)
        .toXlOper(theGC).shouldEqualDlang([["foo", "bar"], ["quux", "toto"]]);
}


///
XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(Unqual!T == Any[])) {
    return [value].toXlOper(allocator);
}

///
@("toXlOper any[]")
unittest {
    import xlld.memorymanager: allocatorContext;

    with(allocatorContext(theGC)) {
        auto oper = toXlOper([any(42.0), any("foo")]);
        oper.xltype.shouldEqual(XlType.xltypeMulti);
        oper.val.array.lparray[0].shouldEqualDlang(42.0);
        oper.val.array.lparray[1].shouldEqualDlang("foo");
    }
}


///
@("toXlOper mixed 1D array of any")
unittest {
    const a = any([any(1.0, theGC), any("foo", theGC)],
                  theGC);
    auto oper = a.toXlOper(theGC);
    oper.xltype.shouldEqual(XlType.xltypeMulti);

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;
    auto opers = oper.val.array.lparray[0 .. rows * cols];
    opers[0].shouldEqualDlang(1.0);
    opers[1].shouldEqualDlang("foo");
}

///
@("toXlOper any[][]")
unittest {
    import xlld.memorymanager: allocatorContext;

    with(allocatorContext(theGC)) {
        auto oper = toXlOper([[any(42.0), any("foo"), any("quux")], [any("bar"), any(7.0), any("toto")]]);
        oper.xltype.shouldEqual(XlType.xltypeMulti);
        oper.val.array.rows.shouldEqual(2);
        oper.val.array.columns.shouldEqual(3);
        oper.val.array.lparray[0].shouldEqualDlang(42.0);
        oper.val.array.lparray[1].shouldEqualDlang("foo");
        oper.val.array.lparray[2].shouldEqualDlang("quux");
        oper.val.array.lparray[3].shouldEqualDlang("bar");
        oper.val.array.lparray[4].shouldEqualDlang(7.0);
        oper.val.array.lparray[5].shouldEqualDlang("toto");
    }
}


///
@("toXlOper mixed 2D array of any")
unittest {
    const a = any([
                     [any(1.0, theGC), any(2.0, theGC)],
                     [any("foo", theGC), any("bar", theGC)]
                 ],
                 theGC);
    auto oper = a.toXlOper(theGC);
    oper.xltype.shouldEqual(XlType.xltypeMulti);

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;
    auto opers = oper.val.array.lparray[0 .. rows * cols];
    opers[0].shouldEqualDlang(1.0);
    opers[1].shouldEqualDlang(2.0);
    opers[2].shouldEqualDlang("foo");
    opers[3].shouldEqualDlang("bar");
}



XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(Unqual!T == DateTime)) {
    import xlld.sdk.framework: Excel12f;
    import xlld.sdk.xlcall: xlfDate, xlfTime, xlretSuccess;
    import nogc.conv: text;

    XLOPER12 ret, date, time;

    auto year = value.year.toXlOper(allocator);
    auto month = value.month.toXlOper(allocator);
    auto day = value.day.toXlOper(allocator);

    const dateCode = () @trusted { return Excel12f(xlfDate, &date, &year, &month, &day); }();
    assert(dateCode == xlretSuccess, "Error calling xlfDate");
    () @trusted { assert(date.xltype == XlType.xltypeNum, text("date is not xltypeNum but ", date.xltype)); }();

    auto hour = value.hour.toXlOper(allocator);
    auto minute = value.minute.toXlOper(allocator);
    auto second = value.second.toXlOper(allocator);

    const timeCode = () @trusted { return Excel12f(xlfTime, &time, &hour, &minute, &second); }();
    assert(timeCode == xlretSuccess, "Error calling xlfTime");
    () @trusted { assert(time.xltype == XlType.xltypeNum, text("time is not xltypeNum but ", time.xltype)); }();

    ret.xltype = XlType.xltypeNum;
    ret.val.num = date.val.num + time.val.num;
    return ret;
}

@("toXlOper!DateTime")
@safe unittest {

    import xlld.sdk.xlcall: xlfDate, xlfTime;

    const dateTime = DateTime(2017, 12, 31, 1, 2, 3);
    {
        auto mockDate = MockXlFunction(xlfDate, 0.1.toXlOper(theGC));
        auto mockTime = MockXlFunction(xlfTime, 0.2.toXlOper(theGC));

        auto oper = dateTime.toXlOper(theGC);

        oper.xltype.shouldEqual(XlType.xltypeNum);
        oper.val.num.shouldApproxEqual(0.3);
    }

    {
        auto mockDate = MockXlFunction(xlfDate, 1.1.toXlOper(theGC));
        auto mockTime = MockXlFunction(xlfTime, 1.2.toXlOper(theGC));

        auto oper = dateTime.toXlOper(theGC);

        oper.xltype.shouldEqual(XlType.xltypeNum);
        oper.val.num.shouldApproxEqual(2.3);
    }
}


XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(Unqual!T == bool)) {
    import xlld.sdk.xlcall: XlType;
    XLOPER12 ret;
    ret.xltype = XlType.xltypeBool;
    ret.val.bool_ = cast(typeof(ret.val.bool_)) value;
    return ret;
}

///
@("toXlOper!bool when bool")
@system unittest {
    import xlld.sdk.xlcall: XlType;
    {
        const oper = true.toXlOper(theGC);
        oper.xltype.shouldEqual(XlType.xltypeBool);
        oper.val.bool_.shouldEqual(1);
    }

    {
        const oper = false.toXlOper(theGC);
        oper.xltype.shouldEqual(XlType.xltypeBool);
        oper.val.bool_.shouldEqual(0);
    }
}

/**
   Register a custom conversion from enum (going through integer) to a string.
   This function will be called to convert enum return values from wrapped
   D functions into strings in Excel.
 */
void registerConversionFrom(T)(FromEnumConversionFunction func) @trusted {
    import std.traits: fullyQualifiedName;

    gFromEnumMutex.lock_nothrow;
    scope(exit)gFromEnumMutex.unlock_nothrow;

    gFromEnumConversions[fullyQualifiedName!T] = func;
}


void unregisterConversionFrom(T)() @trusted {
    import std.traits: fullyQualifiedName;

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

@("toXlOper!enum")
@safe unittest {

    enum Enum {
        foo,
        bar,
        baz,
    }

    Enum.bar.toXlOper(theGC).shouldEqualDlang("bar");
}

XLOPER12 toXlOper(T, A)(T value, ref A allocator)
    if(is(T == struct) && !is(Unqual!T == Any) && !is(Unqual!T == DateTime))
{
    import std.conv: text;
    import core.memory: GC;

    scope str = text(value);

    auto ret = str.toXlOper(allocator);
    () @trusted { GC.free(cast(void*)str.ptr); }();

    return ret;
}

@("toXlOper!struct")
@safe unittest {
    static struct Foo { int x, y; }
    Foo(2, 3).toXlOper(theGC).shouldEqualDlang("Foo(2, 3)");
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
