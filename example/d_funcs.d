/**
 Only exists to test the wrapping functionality Contains functions
 with regular D types that will get wrapped so they can be called by
 the spreadsheet.
 */

import xlld;

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
