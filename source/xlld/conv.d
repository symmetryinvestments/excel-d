module xlld.conv;

import xlld.xlcall: XLOPER12, XlType;
import xlld.any: Any;
import std.traits: isIntegral, Unqual;
import std.datetime: DateTime;

version(unittest) {
    import xlld.any: any;
    import xlld.framework: freeXLOper;
    import xlld.memorymanager: autoFree;
    import xlld.test_util: TestAllocator, shouldEqualDlang, toSRef,
        MockXlFunction, MockDateTime, FailingAllocator;
    import unit_threaded;
    import std.experimental.allocator.gc_allocator: GCAllocator;
    alias theGC = GCAllocator.instance;
}


/**
   Deep copy of an oper
 */
XLOPER12 dup(A)(XLOPER12 oper, ref A allocator) @safe {

    import xlld.xlcall: XlType;
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
                    throw toXlOperMemoryException;
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
                    throw toXlOperMemoryException;

                foreach(i; 0 .. length) {
                    ret.val.array.lparray[i] = oper.val.array.lparray[i].dup(allocator);
                }
            }();

            return ret;
    }

    assert(0);
}


///
@("dup")
@safe unittest {
    auto int_ = 42.toXlOper(theGC);
    int_.dup(theGC).shouldEqualDlang(42);

    auto double_ = (33.3).toXlOper(theGC);
    double_.dup(theGC).shouldEqualDlang(33.3);

    auto string_ = "foobar".toXlOper(theGC);
    string_.dup(theGC).shouldEqualDlang("foobar");

    auto array = () @trusted {
        return [
            ["foo", "bar", "baz"],
            ["quux", "toto", "brzz"]
        ]
        .toXlOper(theGC);
    }();

    array.dup(theGC).shouldEqualDlang(
        [
            ["foo", "bar", "baz"],
            ["quux", "toto", "brzz"],
        ]
    );
}

@("dup string allocator fails")
@safe unittest {
    auto allocator = FailingAllocator();
    "foo".toXlOper(theGC).dup(allocator).shouldThrowWithMessage("Failed to allocate memory for string oper");
}

@("dup multi allocator fails")
@safe unittest {
    auto allocator = FailingAllocator();
    auto oper = () @trusted { return [33.3].toXlOper(theGC); }();
    oper.dup(allocator).shouldThrowWithMessage("Failed to allocate memory for string oper");
}


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
    import xlld.xlcall: XlType;

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
    import xlld.xlcall: XCHAR;
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
__gshared immutable multiMemoryException = new Exception("Failed to allocate memory for multi oper");

private XLOPER12 multi(A)(int rows, int cols, ref A allocator) @trusted {
    auto ret = XLOPER12();

    ret.xltype = XlType.xltypeMulti;
    ret.val.array.rows = rows;
    ret.val.array.columns = cols;

    ret.val.array.lparray = cast(XLOPER12*)allocator.allocate(rows * cols * ret.sizeof).ptr;
    if(ret.val.array.lparray is null)
        throw multiMemoryException;

    return ret;
}


@("multi")
@safe unittest {
    auto allocator = FailingAllocator();
    multi(2, 3, allocator).shouldThrowWithMessage("Failed to allocate memory for multi oper");
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
    import xlld.framework: Excel12f;
    import xlld.xlcall: xlfDate, xlfTime, xlretSuccess;
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

    import xlld.xlcall: xlfDate, xlfTime;

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
    import xlld.xlcall: XlType;
    XLOPER12 ret;
    ret.xltype = XlType.xltypeBool;
    ret.val.bool_ = cast(typeof(ret.val.bool_)) value;
    return ret;
}

///
@("toXlOper!bool when bool")
@system unittest {
    import xlld.xlcall: XlType;
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

XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(T == enum)) {

    import std.conv: text;
    import core.memory: GC;

    auto str = text(value);
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

///
auto fromXlOper(T, A)(ref XLOPER12 val, ref A allocator) {
    return (&val).fromXlOper!T(allocator);
}

/// RValue overload
auto fromXlOper(T, A)(XLOPER12 val, ref A allocator) {
    return fromXlOper!T(val, allocator);
}

///
auto fromXlOper(T, A)(XLOPER12* val, ref A allocator) if(is(Unqual!T == double)) {
    if(val.xltype == XlType.xltypeMissing)
        return double.init;

    return val.val.num;
}

@("fromXlOper!double")
@system unittest {

    TestAllocator allocator;
    auto num = 4.0;
    auto oper = num.toXlOper(allocator);
    auto back = oper.fromXlOper!double(allocator);
    back.shouldEqual(num);

    freeXLOper(&oper, allocator);
}


///
@("isNan for fromXlOper!double")
@system unittest {
    import std.math: isNaN;
    import xlld.memorymanager: allocator;
    XLOPER12 oper;
    oper.xltype = XlType.xltypeMissing;
    fromXlOper!double(&oper, allocator).isNaN.shouldBeTrue;
}

///
auto fromXlOper(T, A)(XLOPER12* val, ref A allocator) if(is(Unqual!T == int)) {
    import xlld.xlcall: XlType;

    if(val.xltype == XlType.xltypeMissing)
        return int.init;

    if(val.xltype == XlType.xltypeNum)
        return cast(typeof(return))val.val.num;

    return val.val.w;
}

///
@("fromXlOper!int")
@system unittest {
    42.toXlOper(theGC).fromXlOper!int(theGC).shouldEqual(42);
}

///
@("fromXlOper!int when given xltypeNum")
@system unittest {
    42.0.toXlOper(theGC).fromXlOper!int(theGC).shouldEqual(42);
}

///
@("0 for fromXlOper!int missing oper")
@system unittest {
    XLOPER12 oper;
    oper.xltype = XlType.xltypeMissing;
    oper.fromXlOper!int(theGC).shouldEqual(0);
}


///
__gshared immutable fromXlOperMemoryException = new Exception("Could not allocate memory for array of char");
///
__gshared immutable fromXlOperConvException = new Exception("Could not convert double to string");

///
auto fromXlOper(T, A)(XLOPER12* val, ref A allocator) if(is(Unqual!T == string)) {

    import std.experimental.allocator: makeArray;
    import std.utf: byChar;
    import std.range: walkLength;

    const stripType = stripMemoryBitmask(val.xltype);
    if(stripType != XlType.xltypeStr && stripType != XlType.xltypeNum)
        return null;


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
@("fromXlOper!string missing")
@system unittest {
    import xlld.memorymanager: allocator;
    XLOPER12 oper;
    oper.xltype = XlType.xltypeMissing;
    fromXlOper!string(&oper, allocator).shouldBeNull;
}

///
@("fromXlOper!string")
@system unittest {
    import std.experimental.allocator: dispose;
    TestAllocator allocator;
    auto oper = "foo".toXlOper(allocator);
    auto str = fromXlOper!string(&oper, allocator);
    allocator.numAllocations.shouldEqual(2);

    freeXLOper(&oper, allocator);
    str.shouldEqual("foo");
    allocator.dispose(cast(void[])str);
}

///
@("fromXlOper!string unicode")
@system unittest {
    auto oper = "é".toXlOper(theGC);
    auto str = fromXlOper!string(&oper, theGC);
    str.shouldEqual("é");
}

@("fromXlOper!string allocation failure")
@system unittest {
    auto allocator = FailingAllocator();
    "foo".toXlOper(theGC).fromXlOper!string(allocator).shouldThrowWithMessage("Could not allocate memory for array of char");
}

@("fromXlOper!string conversion failure")
@system unittest {
    auto allocator = FailingAllocator();
    33.3.toXlOper(theGC).fromXlOper!string(allocator).shouldThrowWithMessage("Could not allocate memory for array of char");
}


package XlType stripMemoryBitmask(in XlType type) @safe @nogc pure nothrow {
    import xlld.xlcall: xlbitXLFree, xlbitDLLFree;
    return cast(XlType)(type & ~(xlbitXLFree | xlbitDLLFree));
}

///
T fromXlOper(T, A)(XLOPER12* oper, ref A allocator) if(is(Unqual!T == Any)) {
    return Any((*oper).dup(allocator));
}

///
@("fromXlOper any double")
@system unittest {
    any(5.0, theGC).fromXlOper!Any(theGC).shouldEqual(any(5.0, theGC));
}

///
@("fromXlOper any string")
@system unittest {
    any("foo", theGC).fromXlOper!Any(theGC)._impl
        .fromXlOper!string(theGC).shouldEqual("foo");
}

///
auto fromXlOper(T, A)(XLOPER12* val, ref A allocator)
    if(is(T: E[][], E) &&
       (is(Unqual!E == string) || is(Unqual!E == double) || is(Unqual!E == int)
        || is(Unqual!E == Any) || is(Unqual!E == DateTime)))
{
    return val.fromXlOperMulti!(Dimensions.Two, typeof(T.init[0][0]))(allocator);
}

///
@("fromXlOper!string[][]")
unittest {
    import xlld.memorymanager: allocator;

    auto strings = [["foo", "bar", "baz"], ["toto", "titi", "quux"]];
    auto oper = strings.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);
    oper.fromXlOper!(string[][])(allocator).shouldEqual(strings);
}

///
@("fromXlOper!double[][]")
unittest {
    import xlld.memorymanager: allocator;

    auto doubles = [[1.0, 2.0], [3.0, 4.0]];
    auto oper = doubles.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);
    oper.fromXlOper!(double[][])(allocator).shouldEqual(doubles);
}

///
@("fromXlOper!string[][] TestAllocator")
unittest {
    import std.experimental.allocator: disposeMultidimensionalArray;
    TestAllocator allocator;
    auto strings = [["foo", "bar", "baz"], ["toto", "titi", "quux"]];
    auto oper = strings.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(string[][])(allocator);

    allocator.numAllocations.shouldEqual(16);

    freeXLOper(&oper, allocator);
    backAgain.shouldEqual(strings);
    allocator.disposeMultidimensionalArray(cast(void[][][])backAgain);
}

///
@("fromXlOper!string[][] when not all opers are strings")
unittest {
    import std.experimental.allocator.mallocator: Mallocator;
    alias allocator = theGC;

    const rows = 2;
    const cols = 3;
    auto array = multi(rows, cols, allocator);
    auto opers = array.val.array.lparray[0 .. rows*cols];
    const strings = ["foo", "bar", "baz"];
    const numbers = [1.0, 2.0, 3.0];

    int i;
    foreach(r; 0 .. rows) {
        foreach(c; 0 .. cols) {
            if(r == 0)
                opers[i++] = strings[c].toXlOper(allocator);
            else
                opers[i++] = numbers[c].toXlOper(allocator);
        }
    }

    opers[3].fromXlOper!string(allocator).shouldEqual("1.000000");
    // sanity checks
    opers[0].fromXlOper!string(allocator).shouldEqual("foo");
    opers[3].fromXlOper!double(allocator).shouldEqual(1.0);
    // the actual assertion
    array.fromXlOper!(string[][])(allocator).shouldEqual([["foo", "bar", "baz"],
                                                          ["1.000000", "2.000000", "3.000000"]]);
}


///
@("fromXlOper!double[][] TestAllocator")
unittest {
    import std.experimental.allocator: disposeMultidimensionalArray;
    TestAllocator allocator;
    auto doubles = [[1.0, 2.0], [3.0, 4.0]];
    auto oper = doubles.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(double[][])(allocator);

    allocator.numAllocations.shouldEqual(4);

    freeXLOper(&oper, allocator);
    backAgain.shouldEqual(doubles);
    allocator.disposeMultidimensionalArray(backAgain);
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
@("fromXlOper!string[]")
unittest {
    import xlld.memorymanager: allocator;

    auto strings = ["foo", "bar", "baz", "toto", "titi", "quux"];
    auto oper = strings.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);
    oper.fromXlOper!(string[])(allocator).shouldEqual(strings);
}

///
@("fromXlOper!double[] from row")
unittest {
    import xlld.xlcall: xlfCaller;

    XLOPER12 caller;
    caller.xltype = XlType.xltypeSRef;
    caller.val.sref.ref_.rwFirst = 1;
    caller.val.sref.ref_.rwLast = 1;
    caller.val.sref.ref_.colFirst = 2;
    caller.val.sref.ref_.colLast = 4;

    with(MockXlFunction(xlfCaller, caller)) {
        auto doubles = [1.0, 2.0, 3.0, 4.0];
        auto oper = doubles.toXlOper(theGC);
        oper.shouldEqualDlang(doubles);
    }
}

///
@("fromXlOper!double[]")
unittest {
    auto doubles = [1.0, 2.0, 3.0, 4.0];
    doubles.toXlOper(theGC).fromXlOper!(double[])(theGC).shouldEqual(doubles);
}


///
@("fromXlOper!string[] TestAllocator")
unittest {
    import std.experimental.allocator: disposeMultidimensionalArray;
    TestAllocator allocator;
    auto strings = ["foo", "bar", "baz", "toto", "titi", "quux"];
    auto oper = strings.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(string[])(allocator);

    allocator.numAllocations.shouldEqual(14);

    backAgain.shouldEqual(strings);
    freeXLOper(&oper, allocator);
    allocator.disposeMultidimensionalArray(cast(void[][])backAgain);
}

///
@("fromXlOper!double[] TestAllocator")
unittest {
    import std.experimental.allocator: dispose;
    TestAllocator allocator;
    auto doubles = [1.0, 2.0, 3.0, 4.0];
    auto oper = doubles.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(double[])(allocator);

    allocator.numAllocations.shouldEqual(2);

    backAgain.shouldEqual(doubles);
    freeXLOper(&oper, allocator);
    allocator.dispose(backAgain);
}

@("fromXlOper!double[][] nil")
@system unittest {
    XLOPER12 oper;
    oper.xltype = XlType.xltypeNil;
    double[][] empty;
    oper.fromXlOper!(double[][])(theGC).shouldEqual(empty);
}

///
__gshared immutable fromXlOperMultiOperException = new Exception("fromXlOper: oper not of multi type");
///
__gshared immutable fromXlOperMultiMemoryException = new Exception("fromXlOper: Could not allocate memory in fromXlOperMulti");

private auto fromXlOperMulti(Dimensions dim, T, A)(XLOPER12* val, ref A allocator) {
    import xlld.xl: coerce, free;
    import xlld.memorymanager: makeArray2D;
    import xlld.xlcall: XlType;
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
    import xlld.xlcall: XlType;
    import xlld.xl: coerce, free;
    import xlld.any: Any;
    version(unittest) import xlld.test_util: gNumXlAllocated, gNumXlFree;

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;
    auto values = oper.val.array.lparray[0 .. (rows * cols)];

    foreach(const row; 0 .. rows) {
        foreach(const col; 0 .. cols) {

            auto cellVal = coerce(&values[row * cols + col]);

            // Issue 22's unittest ends up coercing more than test_util can handle
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


package bool isMulti(ref const(XLOPER12) oper) @safe @nogc pure nothrow {
    const realType = stripMemoryBitmask(oper.xltype);
    return realType == XlType.xltypeMulti;
}


@("fromXlOper any 1D array")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    with(allocatorContext(theGC)) {
        auto array = [any(1.0), any("foo")];
        auto oper = toXlOper(array);
        auto back = fromXlOper!(Any[])(oper);
        back.shouldEqual(array);
    }
}


///
@("fromXlOper Any 2D array")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    with(allocatorContext(theGC)) {
        auto array = [[any(1.0), any(2.0)], [any("foo"), any("bar")]];
        auto oper = toXlOper(array);
        auto back = fromXlOper!(Any[][])(oper);
        back.shouldEqual(array);
    }
}

///
T fromXlOper(T, A)(XLOPER12* oper, ref A allocator) if(is(Unqual!T == DateTime)) {
    import xlld.framework: Excel12f;
    import xlld.xlcall: XlType, xlretSuccess, xlfYear, xlfMonth, xlfDay, xlfHour, xlfMinute, xlfSecond;

    XLOPER12 ret;

    auto get(int fn) @trusted {
        const code = Excel12f(fn, &ret, oper);
        assert(code == xlretSuccess, "Error calling xlf datetime part function");
        // for some reason the Excel API returns doubles
        assert(ret.xltype == XlType.xltypeNum, "xlf datetime part return not xltypeNum");
        return cast(int)ret.val.num;
    }

    return T(get(xlfYear), get(xlfMonth), get(xlfDay),
             get(xlfHour), get(xlfMinute), get(xlfSecond));
}

///
@("fromXlOper!DateTime")
@system unittest {
    XLOPER12 oper;
    auto mockDateTime = MockDateTime(2017, 12, 31, 1, 2, 3);

    const dateTime = oper.fromXlOper!DateTime(theGC);

    dateTime.year.shouldEqual(2017);
    dateTime.month.shouldEqual(12);
    dateTime.day.shouldEqual(31);
    dateTime.hour.shouldEqual(1);
    dateTime.minute.shouldEqual(2);
    dateTime.second.shouldEqual(3);
}

T fromXlOper(T, A)(XLOPER12* oper, ref A allocator) if(is(Unqual!T == bool)) {

    import xlld.xlcall: XlType;
    import std.uni: toLower;

    if(oper.xltype == XlType.xltypeStr) {
        return oper.fromXlOper!string(allocator).toLower == "true";
    }

    return cast(T)oper.val.bool_;
}

@("fromXlOper!bool when bool")
@system unittest {
    import xlld.xlcall: XLOPER12, XlType;
    XLOPER12 oper;
    oper.xltype = XlType.xltypeBool;
    oper.val.bool_ = 1;
    oper.fromXlOper!bool(theGC).shouldEqual(true);

    oper.val.bool_ = 0;
    oper.fromXlOper!bool(theGC).shouldEqual(false);

    oper.val.bool_ = 2;
    oper.fromXlOper!bool(theGC).shouldEqual(true);
}

@("fromXlOper!bool when int")
@system unittest {
    42.toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(true);
    0.toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(false);
}

@("fromXlOper!bool when double")
@system unittest {
    33.3.toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(true);
    0.0.toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(false);
}


@("fromXlOper!bool when string")
@system unittest {
    "true".toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(true);
    "True".toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(true);
    "TRUE".toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(true);
    "false".toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(false);
}

T fromXlOper(T, A)(XLOPER12* oper, ref A allocator) if(is(T == enum)) {
    import std.conv: to;
    return oper.fromXlOper!string(allocator).to!T;
}

@system unittest {
    enum Enum {
        foo, bar, baz,
    }

    "bar".toXlOper(theGC).fromXlOper!Enum(theGC).shouldEqual(Enum.bar);
    "quux".toXlOper(theGC).fromXlOper!Enum(theGC).shouldThrowWithMessage("Enum does not have a member named 'quux'");
}

/**
  creates an XLOPER12 that can be returned to Excel which
  will be freed by Excel itself
 */
XLOPER12 toAutoFreeOper(T)(T value) {
    import xlld.memorymanager: autoFreeAllocator;
    import xlld.xlcall: XlType;

    auto result = value.toXlOper(autoFreeAllocator);
    result.xltype |= XlType.xlbitDLLFree;
    return result;
}

///
ushort operStringLength(T)(in T value) {
    import nogc.exception: enforce;

    enforce(value.xltype == XlType.xltypeStr,
            "Cannot calculate string length for oper of type ", value.xltype);

    return cast(ushort)value.val.str[0];
}

///
@("operStringLength")
unittest {
    import std.experimental.allocator.mallocator: Mallocator;
    auto oper = "foobar".toXlOper(theGC);
    const length = () @nogc { return operStringLength(oper); }();
    length.shouldEqual(6);
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
    import xlld.xl: coerce, free;

    auto coerced = coerce(&val);
    scope(exit) free(&coerced);

    return coerced.fromXlOper!T(allocator);
}


///
@("fromXlOperCoerce")
unittest {
    double[][] doubles = [[1, 2, 3, 4], [11, 12, 13, 14]];
    auto doublesOper = toSRef(doubles, theGC);
    doublesOper.fromXlOper!(double[][])(theGC).shouldThrowWithMessage(
        "fromXlOper: oper not of multi type");
    doublesOper.fromXlOperCoerce!(double[][]).shouldEqual(doubles);
}

private enum invalidXlOperType = 0xdeadbeef;

/**
 Maps a D type to two integer xltypes from XLOPER12.
 InputType is the type actually passed in by the spreadsheet,
 whilst Type is the Type that it gets coerced to.
 */
template dlangToXlOperType(T) {
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
