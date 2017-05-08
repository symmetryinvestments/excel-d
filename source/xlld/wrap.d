module xlld.wrap;

import xlld.xlcall;
import xlld.traits: isSupportedFunction;
import xlld.memorymanager: autoFree;
import xlld.framework: freeXLOper;
import xlld.worksheet;
import xlld.any: Any;
import std.traits: isArray, Unqual;

// here to prevent cyclic dependency
static this() {
    import xlld.memorymanager: gTempAllocator, MemoryPool, StartingMemorySize;
    gTempAllocator = MemoryPool(StartingMemorySize);
}


version(unittest) {
    import unit_threaded;
    import xlld.test_util: TestAllocator, shouldEqualDlang, toSRef;
    import std.experimental.allocator.mallocator: Mallocator;
    import xlld.any: any;
    alias theMallocator = Mallocator.instance;

}

// this shouldn't be needed IMHO and is a bug in std.experimental.allocator that dispose
// doesn't handle 2D arrays correctly
void dispose(A, T)(auto ref A allocator, T[] array) {
    static import std.experimental.allocator;
    import std.traits: Unqual;

    static if(isArray!T) {
        foreach(ref e; array) {
            dispose(allocator, e);
        }
    }

    alias U = Unqual!T;
    std.experimental.allocator.dispose(allocator, cast(U[])array);
}

void dispose(A, T)(auto ref A allocator, T value) if(!isArray!T) {
    static import std.experimental.allocator;
    std.experimental.allocator.dispose(allocator, value);
}


XLOPER12 toXlOper(T, A)(in T val, ref A allocator) if(is(T == double)) {
    auto ret = XLOPER12();
    ret.xltype = XlType.xltypeNum;
    ret.val.num = val;
    return ret;
}


XLOPER12 toXlOper(T, A)(in T val, ref A allocator)
    if(is(T == string) || is(T == wstring))
{
    import std.utf: byWchar;
    import std.stdio;

    // extra space for the length
    auto wval = cast(wchar*)allocator.allocate((val.length + 1) * wchar.sizeof).ptr;
    wval[0] = cast(wchar)val.length;

    int i = 1;
    foreach(ch; val.byWchar) {
        wval[i++] = ch;
    }

    auto ret = XLOPER12();
    ret.xltype = XlType.xltypeStr;
    ret.val.str = cast(XCHAR*)wval;

    return ret;
}


@("toXlOper!string utf8")
@system unittest {
    import std.conv: to;
    import xlld.memorymanager: allocator;

    const str = "foo";
    auto oper = str.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);

    oper.xltype.shouldEqual(xltypeStr);
    (cast(int)oper.val.str[0]).shouldEqual(str.length);
    (cast(wchar*)oper.val.str)[1 .. str.length + 1].to!string.shouldEqual(str);
}

@("toXlOper!string utf16")
@system unittest {
    import xlld.memorymanager: allocator;

    const str = "foo"w;
    auto oper = str.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);

    oper.xltype.shouldEqual(xltypeStr);
    (cast(int)oper.val.str[0]).shouldEqual(str.length);
    (cast(wchar*)oper.val.str)[1 .. str.length + 1].shouldEqual(str);
}

@("toXlOper!string allocator")
@system unittest {
    auto allocator = TestAllocator();
    auto oper = "foo".toXlOper(allocator);
    allocator.numAllocations.shouldEqual(1);
    freeXLOper(&oper, allocator);
}

XLOPER12 toXlOper(T, A)(T[][] values, ref A allocator)
    if(is(T == double) || is(T == string))
{
    import std.algorithm: map, all;
    import std.array: array;

    static const exception = new Exception("# of columns must all be the same and aren't");
    if(!values.all!(a => a.length == values[0].length))
       throw exception;

    const rows = cast(int)values.length;
    const cols = values.length ? cast(int)values[0].length : 0;
    auto ret = multi(cast(int)values.length, cols, allocator);
    auto opers = ret.val.array.lparray[0 .. rows*cols];

    int i;
    foreach(ref row; values) {
        foreach(ref val; row) {
            opers[i++] = val.toXlOper(allocator);
        }
    }

    return ret;
}

XLOPER12 multi(A)(int rows, int cols, ref A allocator) {
    auto ret = XLOPER12();
    ret.xltype = XlType.xltypeMulti;
    ret.val.array.rows = rows;
    ret.val.array.columns = cols;
    ret.val.array.lparray = cast(XLOPER12*)allocator.allocate(rows * cols * ret.sizeof).ptr;
    return ret;
}


@("toXlOper string[][]")
@system unittest {
    import xlld.memorymanager: allocator;

    auto oper = [["foo", "bar", "baz"], ["toto", "titi", "quux"]].toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);

    oper.xltype.shouldEqual(xltypeMulti);
    oper.val.array.rows.shouldEqual(2);
    oper.val.array.columns.shouldEqual(3);
    auto opers = oper.val.array.lparray[0 .. oper.val.array.rows * oper.val.array.columns];

    opers[0].shouldEqualDlang("foo");
    opers[3].shouldEqualDlang("toto");
    opers[5].shouldEqualDlang("quux");
}

@("toXlOper string[][] allocator")
@system unittest {
    TestAllocator allocator;
    auto oper = [["foo", "bar", "baz"], ["toto", "titi", "quux"]].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(7);
    freeXLOper(&oper, allocator);
}

@("toXlOper double[][] allocator")
@system unittest {
    TestAllocator allocator;
    auto oper = [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(1);
    freeXLOper(&oper, allocator);
}


XLOPER12 toXlOper(T, A)(T values, ref A allocator) if(is(T == string[]) || is(T == double[])) {
    T[1] realValues = [values];
    return realValues.toXlOper(allocator);
}


@("toXlOper string[] allocator")
@system unittest {
    TestAllocator allocator;
    auto oper = ["foo", "bar", "baz", "toto", "titi", "quux"].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(7);
    freeXLOper(&oper, allocator);
}

XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(Unqual!T == Any)) {
    return value._impl;
}

@("toXlOper any double")
unittest {
    any(5.0, Mallocator.instance).toXlOper(theMallocator).shouldEqualDlang(5.0);
}

@("toXlOper any string")
unittest {
    any("foo", Mallocator.instance).toXlOper(theMallocator).shouldEqualDlang("foo");
}

@("toXlOper any double[][]")
unittest {
    any([[1.0, 2.0], [3.0, 4.0]], Mallocator.instance)
        .toXlOper(theMallocator).shouldEqualDlang([[1.0, 2.0], [3.0, 4.0]]);
}

@("toXlOper any string[][]")
unittest {
    any([["foo", "bar"], ["quux", "toto"]], Mallocator.instance)
        .toXlOper(theMallocator).shouldEqualDlang([["foo", "bar"], ["quux", "toto"]]);
}


XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(Unqual!T == Any[])) {
    return [value].toXlOper(allocator);
}

@("toXlOper mixed 1D array of any")
unittest {
    const a = any([any(1.0, theMallocator), any("foo", theMallocator)],
                  theMallocator);
    auto oper = a.toXlOper(theMallocator);
    oper.xltype.shouldEqual(XlType.xltypeMulti);

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;
    auto opers = oper.val.array.lparray[0 .. rows * cols];
    opers[0].shouldEqualDlang(1.0);
    opers[1].shouldEqualDlang("foo");
    autoFree(&oper); // normally this is done by Excel
}

XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(Unqual!T == Any[][])) {

    import std.experimental.allocator: makeArray;

    XLOPER12 ret;
    ret.xltype = XlType.xltypeMulti;
    ret.val.array.rows = cast(typeof(ret.val.array.rows)) value.length;
    ret.val.array.columns = cast(typeof(ret.val.array.columns)) value[0].length;
    const length = ret.val.array.rows * ret.val.array.columns;
    ret.val.array.lparray = &allocator.makeArray!XLOPER12(length)[0];

    int i;
    foreach(ref row; value) {
        foreach(ref cell; row) {
            ret.val.array.lparray[i++] = cell;
        }
    }

    return ret;
}

@("toXlOper mixed 2D array of any")
unittest {
    const a = any([
                     [any(1.0, theMallocator), any(2.0, theMallocator)],
                     [any("foo", theMallocator), any("bar", theMallocator)]
                 ],
                 theMallocator);
    auto oper = a.toXlOper(theMallocator);
    oper.xltype.shouldEqual(XlType.xltypeMulti);

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;
    auto opers = oper.val.array.lparray[0 .. rows * cols];
    opers[0].shouldEqualDlang(1.0);
    opers[1].shouldEqualDlang(2.0);
    opers[2].shouldEqualDlang("foo");
    opers[3].shouldEqualDlang("bar");
    autoFree(&oper); // normally this is done by Excel
}

XLOPER12 toXlOper(T, A)(T value, ref A allocator) if(is(T == int)) {
    XLOPER12 ret;
    ret.xltype = XlType.xltypeInt;
    ret.val.w = value;
    return ret;
}

@("toExcelOper!int")
unittest {
    auto oper = 42.toXlOper(theMallocator);
    oper.xltype.shouldEqual(XlType.xltypeInt);
    oper.val.w.shouldEqual(42);
}

auto fromXlOper(T, A)(ref XLOPER12 val, ref A allocator) {
    return (&val).fromXlOper!T(allocator);
}

// RValue overload
auto fromXlOper(T, A)(XLOPER12 val, ref A allocator) {
    return fromXlOper!T(val, allocator);
}

auto fromXlOper(T, A)(LPXLOPER12 val, ref A allocator) if(is(T == double)) {
    if(val.xltype == xltypeMissing)
        return double.init;

    return val.val.num;
}

@("fromXlOper double allocator")
@system unittest {

    TestAllocator allocator;
    auto num = 4.0;
    auto oper = num.toXlOper(allocator);
    auto back = oper.fromXlOper!double(allocator);
    back.shouldEqual(num);

    freeXLOper(&oper, allocator);
}


@("isNan for fromXlOper!double")
@system unittest {
    import std.math: isNaN;
    import xlld.memorymanager: allocator;
    XLOPER12 oper;
    oper.xltype = XlType.xltypeMissing;
    fromXlOper!double(&oper, allocator).isNaN.shouldBeTrue;
}


auto fromXlOper(T, A)(LPXLOPER12 val, ref A allocator)
    if(is(T: E[][], E) && (is(E == string) || is(E == double)))
{
    return val.fromXlOperMulti!(Dimensions.Two, typeof(T.init[0][0]))(allocator);
}

@("fromXlOper!string[][]")
unittest {
    import xlld.memorymanager: allocator;

    auto strings = [["foo", "bar", "baz"], ["toto", "titi", "quux"]];
    auto oper = strings.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);
    oper.fromXlOper!(string[][])(allocator).shouldEqual(strings);
}

@("fromXlOper!double[][]")
unittest {
    import xlld.memorymanager: allocator;

    auto doubles = [[1.0, 2.0], [3.0, 4.0]];
    auto oper = doubles.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);
    oper.fromXlOper!(double[][])(allocator).shouldEqual(doubles);
}

@("fromXlOper!string[][] allocator")
unittest {
    TestAllocator allocator;
    auto strings = [["foo", "bar", "baz"], ["toto", "titi", "quux"]];
    auto oper = strings.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(string[][])(allocator);

    allocator.numAllocations.shouldEqual(16);

    freeXLOper(&oper, allocator);
    backAgain.shouldEqual(strings);
    allocator.dispose(backAgain);
}

@("fromXlOper!string[][] when not all opers are strings")
unittest {
    import std.experimental.allocator.mallocator: Mallocator;
    alias allocator = Mallocator.instance;

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


@("fromXlOper!double[][] allocator")
unittest {
    TestAllocator allocator;
    auto doubles = [[1.0, 2.0], [3.0, 4.0]];
    auto oper = doubles.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(double[][])(allocator);

    allocator.numAllocations.shouldEqual(4);

    freeXLOper(&oper, allocator);
    backAgain.shouldEqual(doubles);
    allocator.dispose(backAgain);
}


private enum Dimensions {
    One,
    Two,
}


// 1D slices
auto fromXlOper(T, A)(LPXLOPER12 val, ref A allocator)
    if(is(T: E[], E) && (is(E == string) || is(E == double)))
{
    return val.fromXlOperMulti!(Dimensions.One, typeof(T.init[0]))(allocator);
}


@("fromXlOper!string[]")
unittest {
    import xlld.memorymanager: allocator;

    auto strings = ["foo", "bar", "baz", "toto", "titi", "quux"];
    auto oper = strings.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);
    oper.fromXlOper!(string[])(allocator).shouldEqual(strings);
}

@("fromXlOper!double[]")
unittest {
    import xlld.memorymanager: allocator;

    auto doubles = [1.0, 2.0, 3.0, 4.0];
    auto oper = doubles.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);
    oper.fromXlOper!(double[])(allocator).shouldEqual(doubles);
}

@("fromXlOper!string[] allocator")
unittest {
    TestAllocator allocator;
    auto strings = ["foo", "bar", "baz", "toto", "titi", "quux"];
    auto oper = strings.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(string[])(allocator);

    allocator.numAllocations.shouldEqual(14);

    backAgain.shouldEqual(strings);
    freeXLOper(&oper, allocator);
    allocator.dispose(backAgain);
}

@("fromXlOper!double[] allocator")
unittest {
    TestAllocator allocator;
    auto doubles = [1.0, 2.0, 3.0, 4.0];
    auto oper = doubles.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(double[])(allocator);

    allocator.numAllocations.shouldEqual(2);

    backAgain.shouldEqual(doubles);
    freeXLOper(&oper, allocator);
    allocator.dispose(backAgain);
}


private auto fromXlOperMulti(Dimensions dim, T, A)(LPXLOPER12 val, ref A allocator) {
    import xlld.xl: coerce, free;
    import std.experimental.allocator: makeArray;

    static const exception = new Exception("fromXlOperMulti failed - oper not of multi type");

    const realType = val.xltype & ~xlbitDLLFree;
    if(realType != xltypeMulti)
        throw exception;

    const rows = val.val.array.rows;
    const cols = val.val.array.columns;

    static if(dim == Dimensions.Two) {
        auto ret = allocator.makeArray!(T[])(rows);
        foreach(ref row; ret)
            row = allocator.makeArray!T(cols);
    } else static if(dim == Dimensions.One) {
        auto ret = allocator.makeArray!T(rows * cols);
    } else
        static assert(0);

    auto values = val.val.array.lparray[0 .. (rows * cols)];

    foreach(const row; 0 .. rows) {
        foreach(const col; 0 .. cols) {
            auto cellVal = coerce(&values[row * cols + col]);
            scope(exit) free(&cellVal);

            // try to convert doubles to string if trying to convert everything to an
            // array of strings
            const shouldConvert = (cellVal.xltype == dlangToXlOperType!T.Type) ||
                (cellVal.xltype == XlType.xltypeNum && dlangToXlOperType!T.Type == XlType.xltypeStr)
                || is(T == Any);

            auto value = shouldConvert ? cellVal.fromXlOper!T(allocator) : T.init;

            static if(dim == Dimensions.Two)
                ret[row][col] = value;
            else
                ret[row * cols + col] = value;
        }
    }

    return ret;
}


auto fromXlOper(T, A)(LPXLOPER12 val, ref A allocator) if(is(T == string)) {

    import std.experimental.allocator: makeArray;
    import std.utf;

    const stripType = val.xltype & ~(xlbitXLFree | xlbitDLLFree);
    if(stripType != XlType.xltypeStr && stripType != XlType.xltypeNum)
        return null;

    if(stripType == XlType.xltypeStr) {

        auto ret = allocator.makeArray!char(val.val.str[0]);
        int i;
        foreach(ch; val.val.str[1 .. ret.length + 1].byChar)
            ret[i++] = ch;

        return cast(string)ret;
    } else {
        // if a double, try to convert it to a string
        import core.stdc.stdio: snprintf;
        char[1024] buffer;
        static const exception = new Exception("Could not convert double to string");
        const numChars = snprintf(&buffer[0], buffer.length, "%lf", val.val.num);
        if(numChars > buffer.length - 1)
            throw exception;
        auto ret = allocator.makeArray!char(numChars);
        ret[] = buffer[0 .. numChars];
        return cast(string)ret;
    }
}

@("fromXlOper missing")
@system unittest {
    import xlld.memorymanager: allocator;
    XLOPER12 oper;
    oper.xltype = XlType.xltypeMissing;
    fromXlOper!string(&oper, allocator).shouldBeNull;
}

@("fromXlOper string allocator")
@system unittest {
    TestAllocator allocator;
    auto oper = "foo".toXlOper(allocator);
    auto str = fromXlOper!string(&oper, allocator);
    allocator.numAllocations.shouldEqual(2);

    freeXLOper(&oper, allocator);
    str.shouldEqual("foo");
    allocator.dispose(cast(void[])str);
}

T fromXlOper(T, A)(LPXLOPER12 oper, ref A allocator) if(is(T == Any)) {
    // FIXME: deep copy
    return Any(*oper);
}

@("fromXlOper any double")
@system unittest {
    any(5.0, theMallocator).fromXlOper!Any(theMallocator).shouldEqual(any(5.0, theMallocator));
}

@("fromXlOper any string")
@system unittest {
    any("foo", theMallocator).fromXlOper!Any(theMallocator)._impl
        .fromXlOper!string(theMallocator).shouldEqual("foo");
}

T fromXlOper(T, A)(LPXLOPER12 oper, ref A allocator) if(is(T == Any[])) {
    return oper.fromXlOperMulti!(Dimensions.One, Any)(allocator);
}


@("fromXlOper any 1D array")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    with(allocatorContext(theMallocator)) {
        auto array = [any(1.0), any("foo")];
        auto oper = toXlOper(array);
        auto back = fromXlOper!(Any[])(oper);
        back.shouldEqual(array);
    }
}


T fromXlOper(T, A)(LPXLOPER12 oper, ref A allocator) if(is(T == Any[][])) {
    return oper.fromXlOperMulti!(Dimensions.Two, typeof(T.init[0][0]))(allocator);
}


@("fromXlOper Any 2D array")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    with(allocatorContext(theMallocator)) {
        auto array = [[any(1.0), any(2.0)], [any("foo"), any("bar")]];
        auto oper = toXlOper(array);
        auto back = fromXlOper!(Any[][])(oper);
        back.shouldEqual(array);
    }
}


private enum isWorksheetFunction(alias F) =
    isSupportedFunction!(F, double, double[][], string[][], string[], double[], string, Any, Any[], Any[][]);

@safe pure unittest {
    import xlld.test_d_funcs;
    // the line below checks that the code still compiles even with a private function
    // it might stop compiling in a future version when the deprecation rules for
    // visibility kick in
    static assert(!isWorksheetFunction!shouldNotBeAProblem);
    static assert(!isWorksheetFunction!FuncThrows);
    static assert(isWorksheetFunction!DoubleArrayToAnyArray);
}

/**
   A string to mixin that wraps all eligible functions in the
   given module.
 */
string wrapModuleWorksheetFunctionsString(string moduleName)() {
    if(!__ctfe) {
        return "";
    }

    import xlld.traits: Identity;
    import std.array: join;
    import std.traits: ReturnType, Parameters;

    mixin(`import ` ~ moduleName ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    string ret = `static import ` ~ moduleName ~ ";\n\n";

    foreach(moduleMemberStr; __traits(allMembers, module_)) {
        alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));

        static if(isWorksheetFunction!moduleMember) {
            ret ~= wrapModuleFunctionStr!(moduleName, moduleMemberStr);
        }
    }

    return ret;
}


@("Wrap double[][] -> double")
@system unittest {
    import xlld.memorymanager: allocator;

    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]], allocator);
    FuncAddEverything(&arg).shouldEqualDlang(60.0);

    arg = toSRef(cast(double[][])[[0, 1, 2, 3], [4, 5, 6, 7]], allocator);
    FuncAddEverything(&arg).shouldEqualDlang(28.0);
}

@("Wrap double[][] -> double[][]")
@system unittest {
    import xlld.memorymanager: allocator;

    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]], allocator);
    FuncTripleEverything(&arg).shouldEqualDlang(cast(double[][])[[3, 6, 9, 12], [33, 36, 39, 42]]);

    arg = toSRef(cast(double[][])[[0, 1, 2, 3], [4, 5, 6, 7]], allocator);
    FuncTripleEverything(&arg).shouldEqualDlang(cast(double[][])[[0, 3, 6, 9], [12, 15, 18, 21]]);
}


@("Wrap string[][] -> double")
@system unittest {

    import xlld.memorymanager: allocator;

    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]], allocator);
    FuncAllLengths(&arg).shouldEqualDlang(29.0);

    arg = toSRef([["", "", "", ""], ["", "", "", ""]], allocator);
    FuncAllLengths(&arg).shouldEqualDlang(0.0);
}

@("Wrap string[][] -> double[][]")
@system unittest {

    import xlld.memorymanager: allocator;

    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]], allocator);
    FuncLengths(&arg).shouldEqualDlang(cast(double[][])[[3, 3, 3, 4], [4, 4, 4, 4]]);

    arg = toSRef([["", "", ""], ["", "", "huh"]], allocator);
    FuncLengths(&arg).shouldEqualDlang(cast(double[][])[[0, 0, 0], [0, 0, 3]]);
}

@("Wrap string[][] -> string[][]")
@system unittest {

    import xlld.memorymanager: allocator;

    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]], allocator);
    FuncBob(&arg).shouldEqualDlang([["foobob", "barbob", "bazbob", "quuxbob"],
                                    ["totobob", "titibob", "tutubob", "tetebob"]]);
}

@("Wrap string[] -> double")
@system unittest {
    import xlld.memorymanager: allocator;

    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toSRef([["foo", "bar"], ["baz", "quux"]], allocator);
    FuncStringSlice(&arg).shouldEqualDlang(4.0);
}

@("Wrap double[] -> double")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toSRef([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], allocator);
    FuncDoubleSlice(&arg).shouldEqualDlang(6.0);
}

@("Wrap double[] -> double[]")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toSRef([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], allocator);
    FuncSliceTimes3(&arg).shouldEqualDlang([3.0, 6.0, 9.0, 12.0, 15.0, 18.0]);
}

@("Wrap string[] -> string[]")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toSRef(["quux", "toto"], allocator);
    StringsToStrings(&arg).shouldEqualDlang(["quuxfoo", "totofoo"]);
}

@("Wrap string[] -> string")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toSRef(["quux", "toto"], allocator);
    StringsToString(&arg).shouldEqualDlang("quux, toto");
}

@("Wrap string -> string")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toXlOper("foo", allocator);
    StringToString(&arg).shouldEqualDlang("foobar");
}

@("Wrap string, string, string -> string")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg0 = toXlOper("foo", allocator);
    auto arg1 = toXlOper("bar", allocator);
    auto arg2 = toXlOper("baz", allocator);
    ManyToString(&arg0, &arg1, &arg2).shouldEqualDlang("foobarbaz");
}

@("Only look at nothrow functions")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toXlOper(2.0, allocator);
    static assert(!__traits(compiles, FuncThrows(&arg)));
}

@("FuncAddEverything wrapper is @nogc")
@system @nogc unittest {
    import std.experimental.allocator.mallocator: Mallocator;
    import xlld.framework: freeXLOper;

    mixin(wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toXlOper(2.0, Mallocator.instance);
    scope(exit) freeXLOper(&arg, Mallocator.instance);
    FuncAddEverything(&arg);
}

private enum invalidXlOperType = 0xdeadbeef;

/**
 Maps a D type to two integer xltypes from XLOPER12.
 InputType is the type actually passed in by the spreadsheet,
 whilst Type is the Type that it gets coerced to.
 */
template dlangToXlOperType(T) {
    static if(is(T == double[][]) || is(T == string[][]) || is(T == double[]) || is(T == string[])) {
        enum InputType = XlType.xltypeSRef;
        enum Type = XlType.xltypeMulti;
    } else static if(is(T == double)) {
        enum InputType = XlType.xltypeNum;
        enum Type = XlType.xltypeNum;
    } else static if(is(T == string)) {
        enum InputType = XlType.xltypeStr;
        enum Type = XlType.xltypeStr;
    } else {
        enum InputType = invalidXlOperType;
        enum Type = invalidXlOperType;
    }
}

/**
 A string to use with `mixin` that wraps a D function
 */
string wrapModuleFunctionStr(string moduleName, string funcName)() {
    if(!__ctfe) {
        return "";
    }

    import std.array: join;
    import std.traits: Parameters, functionAttributes, FunctionAttribute, getUDAs;
    import std.conv: to;
    import std.algorithm: map;
    import std.range: iota;
    mixin("import " ~ moduleName ~ ": " ~ funcName ~ ";");

    const argsLength = Parameters!(mixin(funcName)).length;
    // e.g. LPXLOPER12 arg0, LPXLOPER12 arg1, ...
    const argsDecl = argsLength.iota.map!(a => `LPXLOPER12 arg` ~ a.to!string).join(", ");
    // e.g. arg0, arg1, ...
    const argsCall = argsLength.iota.map!(a => `arg` ~ a.to!string).join(", ");
    const nogc = functionAttributes!(mixin(funcName)) & FunctionAttribute.nogc
        ? "@nogc "
        : "";
    const safe = functionAttributes!(mixin(funcName)) & FunctionAttribute.safe
        ? "@trusted "
        : "";

    alias registerAttrs = getUDAs!(mixin(funcName), Register);
    static assert(registerAttrs.length == 0 || registerAttrs.length == 1,
                  "Invalid number of @Register on " ~ funcName);

    string register;
    static if(registerAttrs.length)
        register = `@` ~ registerAttrs[0].to!string;

    return [
        register,
        `extern(Windows) LPXLOPER12 ` ~ funcName ~ `(` ~ argsDecl ~ `) nothrow ` ~ nogc ~ safe ~ `{`,
        `    static import ` ~ moduleName ~ `;`,
        `    import xlld.memorymanager: gTempAllocator;`,
        `    alias wrappedFunc = ` ~ moduleName ~ `.` ~ funcName ~ `;`,
        `    return wrapModuleFunctionImpl!wrappedFunc(gTempAllocator, ` ~ argsCall ~  `);`,
        `}`,
    ].join("\n");
}

@system unittest {
    import xlld.worksheet;
    import std.traits: getUDAs;

    mixin(wrapModuleFunctionStr!("xlld.test_d_funcs", "FuncAddEverything"));
    alias registerAttrs = getUDAs!(FuncAddEverything, Register);
    static assert(registerAttrs[0].argumentText.value == "Array to add");
}

/**
 Implement a wrapper for a regular D function
 */
LPXLOPER12 wrapModuleFunctionImpl(alias wrappedFunc, A, T...)
                                  (ref A tempAllocator, auto ref T args) {
    import xlld.xl: coerce, free;
    import xlld.worksheet: Dispose;
    import std.traits: Parameters;
    import std.typecons: Tuple;
    import std.traits: hasUDA, getUDAs;

    static XLOPER12 ret;

    XLOPER12[T.length] realArgs;
    // must 1st convert each argument to the "real" type.
    // 2D arrays are passed in as SRefs, for instance
    foreach(i, InputType; Parameters!wrappedFunc) {
        if(args[i].xltype == xltypeMissing) {
             realArgs[i] = *args[i];
             continue;
        }
        realArgs[i] = coerce(args[i]);
    }

    // free any coerced memory
    scope(exit)
        foreach(ref arg; realArgs)
            () @trusted { free(&arg); }();

    Tuple!(Parameters!wrappedFunc) dArgs; // the D types to pass to the wrapped function

    void freeAll() {
        static if(__traits(compiles, tempAllocator.deallocateAll))
            tempAllocator.deallocateAll;
        else {
            foreach(ref dArg; dArgs) {
                import std.traits: isPointer;
                static if(isArray!(typeof(dArg)) || isPointer!(typeof(dArg)))
                    tempAllocator.dispose(dArg);
            }
        }
    }

    // get rid of the temporary memory allocations for the conversions
    scope(exit) freeAll;

    // convert all Excel types to D types
    foreach(i, InputType; Parameters!wrappedFunc) {
        try {
            dArgs[i] = () @trusted { return fromXlOper!InputType(&realArgs[i], tempAllocator); }();
        } catch(Exception ex) {
            ret.xltype = XlType.xltypeErr;
            ret.val.err = -1;
            return &ret;
        }
    }

    try {

        // call the wrapped function with D types
        auto wrappedRet = wrappedFunc(dArgs.expand);
        // convert the return value to an Excel type, tell Excel to call
        // us back to free it afterwards
        ret = toAutoFreeOper(wrappedRet);

        // dispose of the memory allocated in the wrapped function
        static if(hasUDA!(wrappedFunc, Dispose)) {
            alias disposes = getUDAs!(wrappedFunc, Dispose);
            static assert(disposes.length == 1, "Too many @Dispose for " ~ wrappedFunc.stringof);
            disposes[0].dispose(wrappedRet);
        }

    } catch(Exception ex) {

        version(unittest) {
            import core.stdc.stdio: printf;
            static char[1024] buffer;
            buffer[0 .. ex.msg.length] = ex.msg[];
            buffer[ex.msg.length + 1] = 0;
            () @trusted { printf("Could not call wrapped function: %s\n", &buffer[0]); }();
        }

        return null;
    }

    return &ret;
}

@("No memory allocation bugs in wrapModuleFunctionImpl for double return Mallocator")
@system unittest {
    import std.experimental.allocator.mallocator: Mallocator;
    import xlld.test_d_funcs: FuncAddEverything;

    TestAllocator allocator;
    auto arg = toSRef([1.0, 2.0], Mallocator.instance);
    auto oper = wrapModuleFunctionImpl!FuncAddEverything(allocator, &arg);
    (oper.xltype & xlbitDLLFree).shouldBeTrue;
    allocator.numAllocations.shouldEqual(2);
    oper.shouldEqualDlang(3.0);
    autoFree(oper); // normally this is done by Excel
}

@("No memory allocation bugs in wrapModuleFunctionImpl for double[][] return Mallocator")
@system unittest {
    import std.experimental.allocator.mallocator: Mallocator;
    import xlld.test_d_funcs: FuncTripleEverything;

    TestAllocator allocator;
    auto arg = toSRef([1.0, 2.0, 3.0], Mallocator.instance);
    auto oper = wrapModuleFunctionImpl!FuncTripleEverything(allocator, &arg);
    (oper.xltype & xlbitDLLFree).shouldBeTrue;
    (oper.xltype & ~xlbitDLLFree).shouldEqual(xltypeMulti);
    allocator.numAllocations.shouldEqual(2);
    oper.shouldEqualDlang([[3.0, 6.0, 9.0]]);
    autoFree(oper); // normally this is done by Excel
}

@("No memory allocation bugs in wrapModuleFunctionImpl for double[][] return pool")
@system unittest {
    import xlld.memorymanager: gTempAllocator;
    import xlld.test_d_funcs: FuncTripleEverything;

    auto arg = toSRef([1.0, 2.0, 3.0], gTempAllocator);
    auto oper = wrapModuleFunctionImpl!FuncTripleEverything(gTempAllocator, &arg);
    gTempAllocator.curPos.shouldEqual(0);
    oper.shouldEqualDlang([[3.0, 6.0, 9.0]]);
    autoFree(oper); // normally this is done by Excel
}

@("No memory allocation bugs in wrapModuleFunctionImpl for string")
@system unittest {
    import xlld.memorymanager: gTempAllocator;
    import xlld.test_d_funcs: StringToString;

    auto arg = "foo".toSRef(gTempAllocator);
    auto oper = wrapModuleFunctionImpl!StringToString(gTempAllocator, &arg);
    gTempAllocator.curPos.shouldEqual(0);
    oper.shouldEqualDlang("foobar");
}


string wrapWorksheetFunctionsString(Modules...)() {

    if(!__ctfe) {
        return "";
    }

    string ret;
    foreach(module_; Modules) {
        ret ~= wrapModuleWorksheetFunctionsString!module_;
    }

    return ret;
}


string wrapAll(Modules...)(in string mainModule = __MODULE__) {

    if(!__ctfe) {
        return "";
    }

    import xlld.traits: implGetWorksheetFunctionsString;
    return
        wrapWorksheetFunctionsString!Modules ~
        "\n" ~
        implGetWorksheetFunctionsString!(mainModule) ~
        "\n" ~
        `mixin GenerateDllDef!"` ~ mainModule ~ `";` ~
        "\n";
}

@("wrapAll")
unittest  {
    import xlld.memorymanager: allocator;
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll

    mixin(wrapAll!("xlld.test_d_funcs"));
    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]], allocator);
    FuncAddEverything(&arg).shouldEqualDlang(60.0);
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

ushort operStringLength(T)(in T value) {
    import nogc.exception: enforce;

    enforce(value.xltype == XlType.xltypeStr,
            "Cannot calculate string length for oper of type ", value.xltype);

    return cast(ushort)value.val.str[0];
}

@("operStringLength")
unittest {
    import std.experimental.allocator.mallocator: Mallocator;
    auto oper = "foobar".toXlOper(Mallocator.instance);
    const length = () @nogc { return operStringLength(oper); }();
    length.shouldEqual(6);
}

auto fromXlOperCoerce(T)(LPXLOPER12 val) {
    return fromXlOperCoerce(*val);
}

auto fromXlOperCoerce(T, A)(LPXLOPER12 val, auto ref A allocator) {
    return fromXlOperCoerce!T(*val, allocator);
}


auto fromXlOperCoerce(T)(ref XLOPER12 val) {
    import xlld.memorymanager: allocator;
    return fromXlOperCoerce!T(val, allocator);
}


auto fromXlOperCoerce(T, A)(ref XLOPER12 val, auto ref A allocator) {
    import std.experimental.allocator: dispose;
    import xlld.xl: coerce, free;

    auto coerced = coerce(&val);
    scope(exit) free(&coerced);

    return coerced.fromXlOper!T(allocator);
}


@("fromXlOperCoerce")
unittest {
    import xlld.memorymanager: allocator;

    double[][] doubles = [[1, 2, 3, 4], [11, 12, 13, 14]];
    auto doublesOper = toSRef(doubles, allocator);
    doublesOper.fromXlOper!(double[][])(allocator).shouldThrowWithMessage("fromXlOperMulti failed - oper not of multi type");
    doublesOper.fromXlOperCoerce!(double[][]).shouldEqual(doubles);
}

struct TempMemoryPool {

    import xlld.memorymanager: gTempAllocator;
    alias _allocator = gTempAllocator;

    static auto fromXlOper(T, U)(U oper) {
        import xlld.wrap: wrapFromXlOper = fromXlOper;
        return wrapFromXlOper!T(oper, _allocator);
    }

    static auto toXlOper(T)(T val) {
        import xlld.wrap: wrapToXlOper = toXlOper;
        return wrapToXlOper(val, _allocator);
    }

    ~this() @safe {
        _allocator.deallocateAll;
    }
}


@("TempMemoryPool")
unittest {
    import xlld.memorymanager: pool = gTempAllocator;

    with(TempMemoryPool()) {
        auto strOper = toXlOper("foo");
        auto str = fromXlOper!string(strOper);
        pool.curPos.shouldNotEqual(0);
        str.shouldEqual("foo");
    }

    pool.curPos.shouldEqual(0);
}

@("wrap function with @Dispose")
@safe unittest {
    import xlld.test_util: gTestAllocator;
    import xlld.memorymanager: gTempAllocator;
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll

    // this is needed since gTestAllocator is global, so we can't rely
    // on its destructor
    scope(exit) gTestAllocator.verify;

    mixin(wrapAll!("xlld.test_d_funcs"));
    double[4] args = [1.0, 2.0, 3.0, 4.0];
    auto oper = args[].toSRef(gTempAllocator); // don't use TestAllocator
    auto arg = () @trusted { return &oper; }();
    auto ret = () @safe @nogc { return FuncReturnArrayNoGc(arg); }();
    ret.shouldEqualDlang([2.0, 4.0, 6.0, 8.0]);
}

@("wrapModuleFunctionStr function that returns Any[][]")
@safe unittest {
    mixin(wrapModuleFunctionStr!("xlld.test_d_funcs", "DoubleArrayToAnyArray"));

    auto oper = [[1.0, 2.0], [3.0, 4.0]].toSRef(theMallocator);
    auto arg = () @trusted { return &oper; }();
    auto ret = DoubleArrayToAnyArray(arg);

    auto opers = () @trusted { return ret.val.array.lparray[0 .. 4]; }();
    opers[0].shouldEqualDlang(2.0);
    opers[1].shouldEqualDlang(6.0);
    opers[2].shouldEqualDlang("3quux");
    opers[3].shouldEqualDlang("4toto");
}


@("wrapAll function that returns Any[][]")
@safe unittest {
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll

    mixin(wrapAll!("xlld.test_d_funcs"));

    auto oper = [[1.0, 2.0], [3.0, 4.0]].toSRef(theMallocator);
    auto arg = () @trusted { return &oper; }();
    auto ret = DoubleArrayToAnyArray(arg);
    scope(exit) () @trusted { autoFree(ret); }(); // usually done by Excel

    auto opers = () @trusted { return ret.val.array.lparray[0 .. 4]; }();
    opers[0].shouldEqualDlang(2.0);
    opers[1].shouldEqualDlang(6.0);
    opers[2].shouldEqualDlang("3quux");
    opers[3].shouldEqualDlang("4toto");
}

@("wrapAll function that takes Any[][]")
unittest {
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll
    import xlld.memorymanager: allocatorContext;

    mixin(wrapAll!("xlld.test_d_funcs"));

    LPXLOPER12 ret;
    with(allocatorContext(theMallocator)) {
        auto oper = [[any(1.0), any(2.0)], [any(3.0), any(4.0)], [any("foo"), any("bar")]].toXlOper(theMallocator);
        auto arg = () @trusted { return &oper; }();
        ret = AnyArrayToDoubleArray(arg);
    }

    auto opers = () @trusted { return ret.val.array.lparray[0 .. 2]; }();
    opers[0].shouldEqualDlang(3.0); // number of rows
    opers[1].shouldEqualDlang(2.0); // number of columns
}
