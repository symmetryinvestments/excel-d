module xlld.wrap;

import xlld.xlcall;
import xlld.traits: isSupportedFunction;
import xlld.memorymanager: autoFree;
import xlld.framework: freeXLOper;
import xlld.worksheet;
import std.traits: isArray;

version(unittest) {
    import unit_threaded;

    /// emulates SRef types by storing what the referenced type actually is
    XlType gReferencedType;

    // tracks calls to `coerce` and `free` to make sure memory allocations/deallocations match
    int gNumXlCoerce;
    int gNumXlFree;
    enum maxCoerce = 1000;
    void*[maxCoerce] gCoerced;
    void*[maxCoerce] gFreed;

    // automatically converts from oper to compare with a D type
    void shouldEqualDlang(U)(LPXLOPER12 actual, U expected, string file = __FILE__, size_t line = __LINE__) {
        import xlld.memorymanager: allocator;
        if(actual.xltype == xltypeErr)
            fail("XLOPER is of error type", file, line);
        actual.fromXlOper!U(allocator).shouldEqual(expected, file, line);
    }

    // automatically converts from oper to compare with a D type
    void shouldEqualDlang(U)(ref XLOPER12 actual, U expected, string file = __FILE__, size_t line = __LINE__) {
        shouldEqualDlang(&actual, expected, file, line);
    }

    XLOPER12 toSRef(T, A)(T val, ref A allocator) {
        auto ret = toXlOper(val, allocator);
        //hide real type somewhere to retrieve it
        gReferencedType = ret.xltype;
        ret.xltype = XlType.xltypeSRef;
        return ret;
    }

    // tracks allocations and throws in the destructor if there is a memory leak
    // it also throws when there is an attempt to deallocate memory that wasn't
    // allocated
    struct TestAllocator {
        import std.experimental.allocator.common: platformAlignment;
        import std.experimental.allocator.mallocator: Mallocator;

        alias allocator = Mallocator.instance;

        private static struct ByteRange {
            void* ptr;
            size_t length;
        }
        private ByteRange[] _allocations;
        private int _numAllocations;

        enum uint alignment = platformAlignment;

        void[] allocate(size_t numBytes) {
            ++_numAllocations;
            auto ret = allocator.allocate(numBytes);
            writelnUt("+ Allocated  ptr ", ret.ptr, " of ", ret.length, " bytes length");
            _allocations ~= ByteRange(ret.ptr, ret.length);
            return ret;
        }

        bool deallocate(void[] bytes) {
            import std.algorithm: remove, canFind;
            import std.exception: enforce;
            import std.conv: text;

            writelnUt("- Deallocate ptr ", bytes.ptr, " of ", bytes.length, " bytes length");

            bool pred(ByteRange other) { return other.ptr == bytes.ptr && other.length == bytes.length; }

            enforce(_allocations.canFind!pred,
                    text("Unknown deallocate byte range. Ptr: ", bytes.ptr, " length: ", bytes.length,
                         " allocations: ", _allocations));
            _allocations = _allocations.remove!pred;
            return allocator.deallocate(bytes);
        }

        auto numAllocations() @safe pure nothrow const {
            return _numAllocations;
        }

        ~this() {
            import std.exception: enforce;
            import std.conv: text;
            enforce(!_allocations.length, text("Memory leak in TestAllocator. Allocations: ", _allocations));
        }
    }
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
    // should throw unless allocations match deallocations
    TestAllocator allocator;
    auto oper = "foo".toXlOper(allocator);
    allocator.numAllocations.shouldEqual(1);
    freeXLOper(&oper, allocator);
}


XLOPER12 toXlOper(T, A)(T[][] values, ref A allocator)
    if(is(T == double) || is(T == string))
{
    import std.algorithm: map, all;
    import std.array: array;
    import std.exception: enforce;
    import std.conv: text;

    static const exception = new Exception("# of columns must all be the same and aren't");
    if(!values.all!(a => a.length == values[0].length))
       throw exception;

    auto ret = XLOPER12();
    ret.xltype = XlType.xltypeMulti;
    const rows = cast(int)values.length;
    ret.val.array.rows = rows;
    const cols = cast(int)values[0].length;
    ret.val.array.columns = cols;

    ret.val.array.lparray = cast(XLOPER12*)allocator.allocate(rows * cols * ret.sizeof).ptr;
    auto opers = ret.val.array.lparray[0 .. rows*cols];

    int i;
    foreach(ref row; values)
        foreach(ref val; row) {
            opers[i++] = val.toXlOper(allocator);
        }

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

auto fromXlOper(T, A)(ref XLOPER12 val, ref A allocator) {
    return (&val).fromXlOper!T(allocator);
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
    import std.exception: enforce;
    import std.experimental.allocator: makeArray;

    static const exception = new Exception("XL oper not of multi type");

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

            auto value = cellVal.xltype == dlangToXlOperType!T.Type ? cellVal.fromXlOper!T(allocator) : T.init;
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
    if(stripType != xltypeStr)
        return null;

    auto ret = allocator.makeArray!char(val.val.str[0]);
    int i;
    foreach(ch; val.val.str[1 .. ret.length + 1].byChar)
        ret[i++] = ch;

    return cast(string)ret;
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

private enum isWorksheetFunction(alias F) =
    isSupportedFunction!(F, double, double[][], string[][], string[], double[], string);

@safe pure unittest {
    import xlld.test_d_funcs;
    // the line below checks that the code still compiles even with a private function
    // it might stop compiling in a future version when the deprecation rules for
    // visibility kick in
    static assert(!isWorksheetFunction!shouldNotBeAProblem);
    static assert(!isWorksheetFunction!FuncThrows);
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
        enum InputType = xltypeSRef;
        enum Type= xltypeMulti;
    } else static if(is(T == double)) {
        enum InputType = xltypeNum;
        enum Type = xltypeNum;
    } else static if(is(T == string)) {
        enum InputType = xltypeStr;
        enum Type = xltypeStr;
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

    alias registerAttrs = getUDAs!(mixin(funcName), Register);
    static assert(registerAttrs.length == 0 || registerAttrs.length == 1,
                  "Invalid number of @Register on " ~ funcName);

    string register;
    static if(registerAttrs.length)
        register = `@` ~ registerAttrs[0].to!string;

    return [
        register,
        `extern(Windows) LPXLOPER12 ` ~ funcName ~ `(` ~ argsDecl ~ `) nothrow ` ~ nogc ~ `{`,
        `    static import ` ~ moduleName ~ `;`,
        `    import xlld.memorymanager: gMemoryPool;`,
        `    alias wrappedFunc = ` ~ moduleName ~ `.` ~ funcName ~ `;`,
        `    return wrapModuleFunctionImpl!wrappedFunc(gMemoryPool, ` ~ argsCall ~  `);`,
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
                                  (ref A allocator, auto ref T args) {
    import xlld.xl: free;
    import std.traits: Parameters;
    import std.typecons: Tuple;

    static XLOPER12 ret;

    XLOPER12[T.length] realArgs;
    // must 1st convert each argument to the "real" type.
    // 2D arrays are passed in as SRefs, for instance
    foreach(i, InputType; Parameters!wrappedFunc) {
        if(args[i].xltype == xltypeMissing) {
             realArgs[i] = *args[i];
             continue;
        }
        try
            realArgs[i] = convertInput!InputType(args[i]);
        catch(Exception ex) {
            ret.xltype = XlType.xltypeErr;
            ret.val.err = -1;
            return &ret;
        }
    }

    scope(exit)
        foreach(ref arg; realArgs)
            free(&arg);

    Tuple!(Parameters!wrappedFunc) dArgs; // the D types to pass to the wrapped function

    void freeAll() {
        static if(__traits(compiles, allocator.deallocateAll))
            allocator.deallocateAll;
        else {
            foreach(ref dArg; dArgs) {
                import std.traits: isPointer;
                static if(isArray!(typeof(dArg)) || isPointer!(typeof(dArg)))
                    allocator.dispose(dArg);
            }
        }
    }

    scope(exit) freeAll;

    // next call the wrapped function with D types
    foreach(i, InputType; Parameters!wrappedFunc) {
        try {
            dArgs[i] = fromXlOper!InputType(&realArgs[i], allocator);
        } catch(Exception ex) {
            ret.xltype = XlType.xltypeErr;
            ret.val.err = -1;
            return &ret;
        }
    }

    try
        ret = toAutoFreeOper(wrappedFunc(dArgs.expand));
    catch(Exception ex)
        return null;

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
    import xlld.memorymanager: gMemoryPool;
    import xlld.test_d_funcs: FuncTripleEverything;

    auto arg = toSRef([1.0, 2.0, 3.0], gMemoryPool);
    auto oper = wrapModuleFunctionImpl!FuncTripleEverything(gMemoryPool, &arg);
    gMemoryPool.curPos.shouldEqual(0);
    oper.shouldEqualDlang([[3.0, 6.0, 9.0]]);
    autoFree(oper); // normally this is done by Excel
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

    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef;
    mixin(wrapAll!("xlld.test_d_funcs"));
    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]], allocator);
    FuncAddEverything(&arg).shouldEqualDlang(60.0);
}


XLOPER12 convertInput(T)(LPXLOPER12 arg) {
    import xlld.xl: coerce, free;

    static exception = new const Exception("Error converting input");

    if(arg.xltype != dlangToXlOperType!T.InputType)
        throw exception;

    auto realArg = coerce(arg);

    if(realArg.xltype != dlangToXlOperType!T.Type) {
        free(&realArg);
        throw exception;
    }

    return realArg;
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

ushort operStringLength(T)(T value) {
    import std.exception: enforce;
    import std.conv: text;

    enforce(value.xltype == xltypeStr,
            text("Cannot calculate string length for oper of type ", value.xltype));

    return cast(ushort)value.val.str[0];
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
    doublesOper.fromXlOper!(double[][])(allocator).shouldThrowWithMessage("XL oper not of multi type");
    doublesOper.fromXlOperCoerce!(double[][]).shouldEqual(doubles);
}

struct TempMemoryPool {

    import xlld.memorymanager: gMemoryPool;
    alias _allocator = gMemoryPool;

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
    import xlld.memorymanager: pool = gMemoryPool;

    with(TempMemoryPool()) {
        auto strOper = toXlOper("foo");
        auto str = fromXlOper!string(strOper);
        pool.curPos.shouldNotEqual(0);
        str.shouldEqual("foo");
    }

    pool.curPos.shouldEqual(0);
}
