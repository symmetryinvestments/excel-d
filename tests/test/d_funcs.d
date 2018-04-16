/**
 Only exists to test the wrapping functionality Contains functions
 with regular D types that will get wrapped so they can be called by
 the spreadsheet.
 */

module test.d_funcs;

version(unittest):

import xlld;
import std.datetime: DateTime;


///
@Register(ArgumentText("Array to add"),
          HelpTopic("Adds all cells in an array"),
          FunctionHelp("Adds all cells in an array"),
          ArgumentHelp(["The array to add"]))
double FuncAddEverything(double[][] args) nothrow @nogc {
    import std.algorithm: fold;
    import std.math: isNaN;

    double ret = 0;
    foreach(row; args)
        ret += row.fold!((a, b) => b.isNaN ? 0.0 : a + b)(0.0);
    return ret;
}

///
double[][] FuncTripleEverything(double[][] args) nothrow {
    double[][] ret;
    ret.length = args.length;
    foreach(i; 0 .. args.length) {
        ret[i].length = args[i].length;
        foreach(j; 0 .. args[i].length)
            ret[i][j] = args[i][j] * 3;
    }

    return ret;
}

///
double FuncAllLengths(string[][] args) nothrow @nogc {
    import std.algorithm: fold;

    double ret = 0;
    foreach(row; args)
        ret += row.fold!((a, b) => a + b.length)(0.0);
    return ret;
}

///
double[][] FuncLengths(string[][] args) nothrow {
    double[][] ret;

    ret.length = args.length;
    foreach(i; 0 .. args.length) {
        ret[i].length = args[i].length;
        foreach(j; 0 .. args[i].length)
            ret[i][j] = args[i][j].length;
    }

    return ret;
}


///
string[][] FuncBob(string[][] args) nothrow {
    string[][] ret;

    ret.length = args.length;
    foreach(i; 0 .. args.length) {
        ret[i].length = args[i].length;
        foreach(j; 0 .. args[i].length)
            ret[i][j] = args[i][j] ~ "bob";
    }

    return ret;
}


///
double FuncDoubleSlice(double[] arg) nothrow @nogc {
    return arg.length;
}

///
double FuncStringSlice(string[] arg) nothrow @nogc {
    return arg.length;
}

///
double[] FuncSliceTimes3(double[] arg) nothrow {
    import std.algorithm;
    import std.array;
    return arg.map!(a => a * 3).array;
}

///
string[] StringsToStrings(string[] args) nothrow {
    import std.algorithm;
    import std.array;
    return args.map!(a => a ~ "foo").array;
}

///
string StringsToString(string[] args) nothrow {
    import std.string;
    return args.join(", ");
}

///
string StringToString(string arg) nothrow {
    return arg ~ "bar";
}

private string shouldNotBeAProblem(string, string[]) nothrow {
    return "";
}

///
string ManyToString(string arg0, string arg1, string arg2) nothrow {
    return arg0 ~ arg1 ~ arg2;
}

// shouldn't get wrapped
double FuncThrows(double) {
    throw new Exception("oops");
}

///
double FuncAsserts(double) {
    assert(false);
}


/**
    @Dispose is used to tell the framework how to free memory that is dynamically
    allocated by the D function. After returning, the value is converted to an
    Excel type sand the D value is freed using the lambda defined here.
    In this example we're using TestAllocator to make sure that there are no
    memory leaks.
*/
@Dispose!((ret) {
    import xlld.test.util: gTestAllocator;
    import std.experimental.allocator: dispose;
    gTestAllocator.dispose(ret);
})

///
double[] FuncReturnArrayNoGc(double[] numbers) @nogc @safe nothrow {
    import xlld.test.util: gTestAllocator;
    import std.experimental.allocator: makeArray;
    import std.algorithm: map;

    try {
        return () @trusted { return gTestAllocator.makeArray(numbers.map!(a => a * 2)); }();
    } catch(Exception _) {
        return [];
    }
}


///
Any[][] DoubleArrayToAnyArray(double[][] values) @safe nothrow {
    import std.experimental.allocator.mallocator: Mallocator;
    import std.conv: to;

    alias allocator = Mallocator.instance;

    string third, fourth;
    try {
        third = values[1][0].to!string;
        fourth = values[1][1].to!string;
    } catch(Exception ex) {
        third = "oops";
        fourth = "oops";
    }

    return () @trusted {
        with(allocatorContext(allocator)) {
            try
                return [
                    [any(values[0][0] * 2), any(values[0][1] * 3)],
                    [any(third ~ "quux"),   any(fourth ~ "toto")],
                ];
            catch(Exception _) {
                Any[][] empty;
                return empty;
            }
        }
    }();
}


///
double[] AnyArrayToDoubleArray(Any[][] values) nothrow {
    return [values.length, values.length ? values[0].length : 0];
}


///
Any[][] AnyArrayToAnyArray(Any[][] values) nothrow {
    return values;
}

///
Any[][] FirstOfTwoAnyArrays(Any[][] a, Any[][]) nothrow {
    return a;
}

///
string[] EmptyStrings1D(Any) nothrow {
    string[] empty;
    return empty;
}


///
string[][] EmptyStrings2D(Any) nothrow {
    string[][] empty;
    return empty;
}

///
string[][] EmptyStringsHalfEmpty2D(Any) nothrow {
    string[][] empty;
    empty.length = 1;
    assert(empty[0].length == 0);
    return empty;
}

///
int Twice(int i) @safe nothrow {
    return i * 2;
}

///
double FuncConstDouble(const double a) @safe nothrow {
    return a;
}

double DateTimeToDouble(DateTime d) @safe nothrow {
    return d.year * 2;
}

string DateTimesToString(DateTime[] ds) @safe nothrow {
    import std.algorithm: map;
    import std.string: join;
    import std.conv: text;
    return ds.map!(d => d.day.text).join(", ");
}


@Async
double AsyncDoubleToDouble(double d) @safe nothrow {
    return d * 2;
}

import core.stdc.stdio;
double Overloaded(double d) @safe @nogc nothrow {
    () @trusted { printf("double\n"); }();
    return d * 2;
}

double Overloaded(string s) @safe @nogc nothrow {
    () @trusted { printf("string\n"); }();
    return s.length;
}

double NaN() {
    return double.init;
}

int BoolToInt(bool b) {
    return cast(int)b;
}

enum MyEnum {
    foo, bar, baz,
}

string FuncMyEnumArg(MyEnum val) @safe {
    import std.conv: text;
    return "prefix_" ~ val.text;
}

MyEnum FuncMyEnumRet(int i) @safe {
    return cast(MyEnum)i;
}

struct Point {
    int x, y;
}

int FuncPointArg(Point p) @safe {
    return p.x + p.y;
}

Point FuncPointRet(int x, int y) @safe {
    return Point(x, y);
}

auto FuncSimpleTupleRet(int i, string s) @safe {
    import std.typecons: tuple;
    return tuple(i, s);
}

auto FuncComplexTupleRet(int d1, int d2) @safe {
    import std.typecons: tuple;
    return tuple([DateTime(2017, 1, d1), DateTime(2017, 2, d1)],
                 [DateTime(2018, 1, d1), DateTime(2018, 2, d2)]);
}

auto FuncTupleArrayRet() @safe {
    import std.typecons: tuple;
    return [
        tuple(DateTime(2017, 1, 1), 11.1),
        tuple(DateTime(2018, 1, 1), 22.2),
        tuple(DateTime(2019, 1, 1), 33.3),
    ];
}

struct DateAndString {
    DateTime dateTime;
    string string_;
}

DateAndString[] FuncDateAndStringRet() {
    return [DateAndString(DateTime(2017, 1, 2), "foobar")];
}


void FuncEnumArray(MyEnum[]) {

}
