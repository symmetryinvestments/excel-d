/**
 Only exists to test the wrapping functionality Contains functions
 with regular D types that will get wrapped so they can be called by
 the spreadsheet.
 */

import xlld;
import std.datetime: DateTime;

double[][] dlookup(string[][] haystack, string[] needles, double columnNumberD) nothrow
{
    import std.exception;
    import std.algorithm:map,countUntil;
    import std.range:repeat,transposed;
    import std.array:array;
    import std.conv:to;

    double toDouble(long pos)
    {
        return (pos==-1) ? double.nan : pos.to!double + 1.0;
    }
    try
    {
        auto columnNumber = columnNumberD.to!int;
        auto haystackColumn = haystack.map!(row => row[columnNumber-1].to!string).array;
        return needles.map!( needle => [toDouble(haystackColumn.countUntil(needle).to!long)]).array;
    }
    catch(Exception e)
    {
        return ([double.nan].repeat(needles.length)).array;
    }
}
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

// @Dispose is used to tell the framework how to free memory that is dynamically
// allocated by the D function. After returning, the value is converted to an
// Excel type sand the D value is freed using the lambda defined here.
@Dispose!((ret) {
    import std.experimental.allocator.mallocator: Mallocator;
    import std.experimental.allocator: dispose;
    Mallocator.instance.dispose(ret);
})
double[] FuncReturnArrayNoGc(double[] numbers) @nogc @safe nothrow {
    import std.experimental.allocator.mallocator: Mallocator;
    import std.experimental.allocator: makeArray;
    import std.algorithm: map;

    try {
        // Allocate memory here in order to return an array of doubles.
        // The memory will be freed after the call by calling the
        // function in `@Dispose` above
        return Mallocator.instance.makeArray(numbers.map!(a => a * 2));
    } catch(Exception _) {
        return [];
    }
}

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

double FuncAllLengths(string[][] args) nothrow @nogc {
    import std.algorithm: fold;

    double ret = 0;
    foreach(row; args)
        ret += row.fold!((a, b) => a + b.length)(0.0);
    return ret;
}

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


double FuncDoubleSlice(double[] arg) nothrow @nogc {
    return arg.length;
}

double FuncStringSlice(string[] arg) nothrow @nogc {
    return arg.length;
}

double[] FuncSliceTimes3(double[] arg) nothrow {
    import std.algorithm;
    import std.array;
    return arg.map!(a => a * 3).array;
}

string[] StringsToStrings(string[] args) nothrow {
    import std.algorithm;
    import std.array;
    return args.map!(a => a ~ "foo").array;
}

string StringsToString(string[] args) nothrow {
    import std.string;
    return args.join(", ");
}

string StringToString(string arg) nothrow {
    return arg ~ "bar";
}

private string shouldNotBeAProblem(string, string[]) nothrow {
    return "";
}

string ManyToString(string arg0, string arg1, string arg2) nothrow {
    return arg0 ~ arg1 ~ arg2;
}

double FuncThrows(double) {
    throw new Exception("oops");
}

Any[][] FirstOfTwoAnyArrays(Any[][] lhs, Any[][] rhs) nothrow {
    return lhs;
}

string[][] FirstOfTwoAnyArraysToString(Any[][] testarg, Any[][] rhs) nothrow {
    import std.array, std.algorithm, std.conv;
    try {
        return testarg.map!(map!(to!string)).map!array.array;
    } catch(Exception e) {
        return [[e.msg]];
    }
}

string DateTimeToString(DateTime dt) @safe {
    import std.conv: text;
    return text("year: ", dt.year, ", month: ", dt.month, ", day: ", dt.day,
                ", hour: ", dt.hour, ", minute: ", dt.minute, ", second: ", dt.second);
}

string DateTimesToString(DateTime[] dts) @safe {
    import std.conv: text;
    import std.algorithm: map;
    import std.string: join;
    return dts.map!(dt => text("year: ", dt.year, ", month: ", dt.month, ", day: ", dt.day,
                               ", hour: ", dt.hour, ", minute: ", dt.minute, ", second: ", dt.second)).join("\n");
}

double FuncTwice(double d) @safe nothrow @nogc {
    return d * 2;
}

@Async
double FuncTwiceAsync(double d) {
    import core.thread;
    Thread.sleep(5.seconds);
    return d * 2;
}

double IntToDouble(int i) {
    import xlld.xll: log;
    log("IntToDouble(", i, ")");
    return i * 2;
}

int DoubleToInt(double d) {
    return cast(int)(d * 2);
}

int IntToInt(int i) @safe nothrow @nogc {
    return i * 2;
}

DateTime[] DateTimes(int year, int month, int day) {
    return [
        DateTime(year, month, day),
        DateTime(year + 1, month + 1, day + 1),
        DateTime(year + 2, month + 2, day + 2),
    ];
}

string FuncCaller() @safe {
    import xlld.xlf: caller;
    import xlld.xlcall: XlType;

    auto res = caller;

    switch(res.xltype) with(XlType) {
    default:
        return "Unknown caller type";
    case xltypeSRef:
        return "Called from a cell";
    case xltypeRef:
        return "Called from a multi-cell array formula";
    }
}

string FuncCallerAdjacent() @safe {
    import xlld.xl: Coerced;
    import xlld.xlf: caller;
    import xlld.xlcall: XlType;
    import std.exception: enforce;

    auto res = caller;

    enforce(res.xltype == XlType.xltypeSRef);

    ++res.val.sref.ref_.colFirst;
    ++res.val.sref.ref_.colLast;

    auto coerced = Coerced(res);
    return "Guy next to me: " ~ coerced.toString;
}

string FuncCallerCell() @safe {
    import xlld.xlf: callerCell;
    import xlld.xl: free;

    auto caller = callerCell;

    return "Was called from a cell";
}
