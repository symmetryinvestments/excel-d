/**
   Conversions from XLOPER12 to D types
 */
module xlld.conv.from;

import xlld.from;
import xlld.sdk.xlcall: XLOPER12;
import xlld.any: Any;
import std.traits: Unqual;
import std.datetime: DateTime;
import core.sync.mutex: Mutex;

version(testingExcelD) {
    import xlld.conv.to: toXlOper;
    import xlld.sdk.framework: freeXLOper;
    import xlld.test.util: TestAllocator, FailingAllocator, toSRef, MockDateTime, MockXlFunction, shouldEqualDlang;
    import xlld.sdk.xlcall: XlType;
    import xlld.any: any;
    import unit_threaded;
    import std.experimental.allocator.gc_allocator: GCAllocator;
    alias theGC = GCAllocator.instance;
}

alias ToEnumConversionFunction = int delegate(string);
package __gshared ToEnumConversionFunction[string] gToEnumConversions;
package shared Mutex gToEnumMutex;


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

@("fromXlOper!double wrong oper type")
@system unittest {
    "foo".toXlOper(theGC).fromXlOper!double(theGC).shouldThrowWithMessage("Wrong type for fromXlOper!double");
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

@("fromXlOper!int wrong oper type")
@system unittest {
    "foo".toXlOper(theGC).fromXlOper!int(theGC).shouldThrowWithMessage("Wrong type for fromXlOper!int");
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


@("fromXlOper!string wrong oper type")
@system unittest {
    42.toXlOper(theGC).fromXlOper!string(theGC).shouldThrowWithMessage("Wrong type for fromXlOper!string");
}

///
T fromXlOper(T, A)(XLOPER12* oper, ref A allocator) if(is(Unqual!T == Any)) {
    import xlld.conv.misc: dup;
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
    import xlld.conv.misc: multi;
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
    import xlld.sdk.xlcall: xlfCaller;

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

__gshared immutable fromXlOperDateTimeTypeException = new Exception("Wrong type for fromXlOper!DateTime");

///
T fromXlOper(T, A)(XLOPER12* oper, ref A allocator) if(is(Unqual!T == DateTime)) {
    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.sdk.framework: Excel12f;
    import xlld.sdk.xlcall: XlType, xlretSuccess, xlfYear, xlfMonth, xlfDay, xlfHour, xlfMinute, xlfSecond;

    if(oper.xltype.stripMemoryBitmask != XlType.xltypeNum)
        throw fromXlOperDateTimeTypeException;

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

@("fromXlOper!DateTime wrong oper type")
@system unittest {
    42.toXlOper(theGC).fromXlOper!DateTime(theGC).shouldThrowWithMessage(
        "Wrong type for fromXlOper!DateTime");
}

T fromXlOper(T, A)(XLOPER12* oper, ref A allocator) if(is(Unqual!T == bool)) {

    import xlld.sdk.xlcall: XlType;
    import std.uni: toLower;

    if(oper.xltype == XlType.xltypeStr) {
        return oper.fromXlOper!string(allocator).toLower == "true";
    }

    return cast(T)oper.val.bool_;
}

@("fromXlOper!bool when bool")
@system unittest {
    import xlld.sdk.xlcall: XLOPER12, XlType;
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
        gToEnumMutex.lock_nothrow;
        scope(exit) gToEnumMutex.unlock_nothrow;

        return name in gToEnumConversions
                           ? cast(T) gToEnumConversions[name](str)
                           : str.to!T;
    }();
}

@("fromXlOper!enum")
@system unittest {
    enum Enum {
        foo, bar, baz,
    }

    "bar".toXlOper(theGC).fromXlOper!Enum(theGC).shouldEqual(Enum.bar);
    "quux".toXlOper(theGC).fromXlOper!Enum(theGC).shouldThrowWithMessage("Enum does not have a member named 'quux'");
}

@("fromXlOper!enum wrong type")
@system unittest {
    enum Enum { foo, bar, baz, }

    42.toXlOper(theGC).fromXlOper!Enum(theGC).shouldThrowWithMessage(
        "Wrong type for fromXlOper!Enum");
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

    ulong ptrIndex(ulong i) {

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

@("1D array to struct")
@system unittest {
    static struct Foo { int x, y; }
    [2, 3].toXlOper(theGC).fromXlOper!Foo(theGC).shouldEqual(Foo(2, 3));
}

@("wrong oper type to struct")
@system unittest {
    static struct Foo { int x, y; }

    2.toXlOper(theGC).fromXlOper!Foo(theGC).shouldThrowWithMessage(
        "Can only convert arrays to structs. Must be either 1xN, Nx1, 2xN or Nx2");
}

@("1D array to struct with wrong length")
@system unittest {

    static struct Foo { int x, y; }

    [2, 3, 4].toXlOper(theGC).fromXlOper!Foo(theGC).shouldThrowWithMessage(
        "1D array length must match number of members in Foo. Expected 2, got 3");

    [2].toXlOper(theGC).fromXlOper!Foo(theGC).shouldThrowWithMessage(
        "1D array length must match number of members in Foo. Expected 2, got 1");
}

@("1D array to struct with wrong type")
@system unittest {
    static struct Foo { int x, y; }

    ["foo", "bar"].toXlOper(theGC).fromXlOper!Foo(theGC).shouldThrowWithMessage(
        "Wrong type converting oper to Foo");
}

@("2D horizontal array to struct")
unittest {
    import xlld.memorymanager: allocatorContext;

    static struct Foo { int x, y, z; }

    with(allocatorContext(theGC)) {
        [[any("x"), any("y"), any("z")], [any(2), any(3), any(4)]].toFrom!Foo.shouldEqual(Foo(2, 3, 4));
    }
}

@("2D vertical array to struct")
unittest {
    import xlld.memorymanager: allocatorContext;

    static struct Foo { int x, y, z; }

    with(allocatorContext(theGC)) {
        [[any("x"), any(2)], [any("y"), any(3)], [any("z"), any(4)]].toFrom!Foo.shouldEqual(Foo(2, 3, 4));
    }
}

@("2D array wrong size")
unittest {
    import xlld.memorymanager: allocatorContext;

    static struct Foo { int x, y, z; }

    with(allocatorContext(theGC)) {
        [[any("x"), any(2)], [any("y"), any(3)], [any("z"), any(4)], [any("w"), any(5)]].toFrom!Foo.
            shouldThrowWithMessage("2D array must be 2x3 or 3x2 for Foo");
    }

}

private auto toFrom(R, T)(T val) {
    import std.experimental.allocator.gc_allocator: GCAllocator;
    return val.toXlOper(GCAllocator.instance).fromXlOper!R(GCAllocator.instance);
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

    gToEnumMutex.lock_nothrow;
    scope(exit)gToEnumMutex.unlock_nothrow;

    gToEnumConversions[fullyQualifiedName!T] = func;
}

void unregisterConversionTo(T)() @trusted {
    import std.traits: fullyQualifiedName;

    gToEnumMutex.lock_nothrow;
    scope(exit)gToEnumMutex.unlock_nothrow;

    gToEnumConversions.remove(fullyQualifiedName!T);
}
