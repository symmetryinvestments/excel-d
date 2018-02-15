/**
    Wrapper module for Excel. This module contains the functionality that autowraps
    D code for use within Excel.
*/
module xlld.wrap;

import xlld.worksheet;
import xlld.xlcall: XLOPER12;
import xlld.traits: isSupportedFunction;
import xlld.memorymanager: autoFree;
import xlld.framework: freeXLOper;
import xlld.any: Any;
import std.datetime: DateTime;
import std.experimental.allocator: theAllocator;

version(unittest) {
    import xlld.conv: toXlOper;
    import xlld.any: any;
    import xlld.test_util: TestAllocator, shouldEqualDlang, toSRef, gDates, gTimes,
        gYears, gMonths, gDays, gHours, gMinutes, gSeconds;

    import unit_threaded;
    import std.experimental.allocator.mallocator: Mallocator;
    import std.experimental.allocator.gc_allocator: GCAllocator;
    alias theMallocator = Mallocator.instance;
    alias theGC = GCAllocator.instance;
}

static if(!is(Flaky))
    enum Flaky;




private enum isWorksheetFunction(alias F) =
    isSupportedFunction!(F,
                         bool,
                         int,
                         double, double[], double[][],
                         string, string[], string[][],
                         Any, Any[], Any[][],
                         DateTime, DateTime[], DateTime[][],
    );

@safe pure unittest {
    import xlld.test_d_funcs;
    // the line below checks that the code still compiles even with a private function
    // it might stop compiling in a future version when the deprecation rules for
    // visibility kick in
    static assert(!isWorksheetFunction!shouldNotBeAProblem);
    static assert(isWorksheetFunction!FuncThrows);
    static assert(isWorksheetFunction!DoubleArrayToAnyArray);
    static assert(isWorksheetFunction!Twice);
    static assert(isWorksheetFunction!DateTimeToDouble);
    static assert(isWorksheetFunction!BoolToInt);
}

/**
   A string to mixin that wraps all eligible functions in the
   given module.
 */
string wrapModuleWorksheetFunctionsString(string moduleName)(string callingModule = __MODULE__) {
    if(!__ctfe) {
        return "";
    }

    import xlld.traits: Identity;
    import std.array: join;
    import std.traits: ReturnType, Parameters;

    mixin(`import ` ~ moduleName ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    string ret;

    foreach(moduleMemberStr; __traits(allMembers, module_)) {

        alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));

        static if(isWorksheetFunction!moduleMember) {
            enum numOverloads = __traits(getOverloads, mixin(moduleName), moduleMemberStr).length;
            static if(numOverloads == 1)
                ret ~= wrapModuleFunctionStr!(moduleName, moduleMemberStr)(callingModule);
            else
                pragma(msg, "excel-d WARNING: Not wrapping ", moduleMemberStr, " due to it having ",
                       cast(int)numOverloads, " overloads");
        } else {
            /// trying to get a pointer to something is a good way of making sure we can
            /// attempt to evaluate `isSomeFunction` - it's not always possible
            enum canGetPointerToIt = __traits(compiles, &moduleMember);
            static if(canGetPointerToIt) {
                import xlld.worksheet: Register;
                import std.traits: getUDAs;
                alias registerAttrs = getUDAs!(moduleMember, Register);
                static assert(registerAttrs.length == 0,
                              "excel-d ERROR: Function `" ~ moduleMemberStr ~ "` not eligible for wrapping");
            }
        }
    }

    return ret;
}


///
@("Wrap double[][] -> double")
@system unittest {
    import xlld.memorymanager: allocator;

    mixin(wrapTestFuncsString);

    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]], allocator);
    FuncAddEverything(&arg).shouldEqualDlang(60.0);

    arg = toSRef(cast(double[][])[[0, 1, 2, 3], [4, 5, 6, 7]], allocator);
    FuncAddEverything(&arg).shouldEqualDlang(28.0);
}

///
@("Wrap double[][] -> double[][]")
@system unittest {
    import xlld.memorymanager: allocator;

    mixin(wrapTestFuncsString);

    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]], allocator);
    FuncTripleEverything(&arg).shouldEqualDlang(cast(double[][])[[3, 6, 9, 12], [33, 36, 39, 42]]);

    arg = toSRef(cast(double[][])[[0, 1, 2, 3], [4, 5, 6, 7]], allocator);
    FuncTripleEverything(&arg).shouldEqualDlang(cast(double[][])[[0, 3, 6, 9], [12, 15, 18, 21]]);
}


///
@("Wrap string[][] -> double")
@system unittest {

    import xlld.memorymanager: allocator;

    mixin(wrapTestFuncsString);

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]], allocator);
    FuncAllLengths(&arg).shouldEqualDlang(29.0);

    arg = toSRef([["", "", "", ""], ["", "", "", ""]], allocator);
    FuncAllLengths(&arg).shouldEqualDlang(0.0);
}

///
@("Wrap string[][] -> double[][]")
@system unittest {

    import xlld.memorymanager: allocator;

    mixin(wrapTestFuncsString);

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]], allocator);
    FuncLengths(&arg).shouldEqualDlang(cast(double[][])[[3, 3, 3, 4], [4, 4, 4, 4]]);

    arg = toSRef([["", "", ""], ["", "", "huh"]], allocator);
    FuncLengths(&arg).shouldEqualDlang(cast(double[][])[[0, 0, 0], [0, 0, 3]]);
}

///
@("Wrap string[][] -> string[][]")
@system unittest {

    import xlld.memorymanager: allocator;

    mixin(wrapTestFuncsString);

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]], allocator);
    FuncBob(&arg).shouldEqualDlang([["foobob", "barbob", "bazbob", "quuxbob"],
                                    ["totobob", "titibob", "tutubob", "tetebob"]]);
}

///
@("Wrap string[] -> double")
@system unittest {
    import xlld.memorymanager: allocator;

    mixin(wrapTestFuncsString);
    auto arg = toSRef([["foo", "bar"], ["baz", "quux"]], allocator);
    FuncStringSlice(&arg).shouldEqualDlang(4.0);
}

///
@("Wrap double[] -> double")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapTestFuncsString);
    auto arg = toSRef([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], allocator);
    FuncDoubleSlice(&arg).shouldEqualDlang(6.0);
}

///
@("Wrap double[] -> double[]")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapTestFuncsString);
    auto arg = toSRef([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], allocator);
    FuncSliceTimes3(&arg).shouldEqualDlang([3.0, 6.0, 9.0, 12.0, 15.0, 18.0]);
}

///
@("Wrap string[] -> string[]")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapTestFuncsString);
    auto arg = toSRef(["quux", "toto"], allocator);
    StringsToStrings(&arg).shouldEqualDlang(["quuxfoo", "totofoo"]);
}

///
@("Wrap string[] -> string")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapTestFuncsString);
    auto arg = toSRef(["quux", "toto"], allocator);
    StringsToString(&arg).shouldEqualDlang("quux, toto");
}

///
@("Wrap string -> string")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapTestFuncsString);
    auto arg = toXlOper("foo", allocator);
    StringToString(&arg).shouldEqualDlang("foobar");
}

///
@("Wrap string, string, string -> string")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapTestFuncsString);
    auto arg0 = toXlOper("foo", allocator);
    auto arg1 = toXlOper("bar", allocator);
    auto arg2 = toXlOper("baz", allocator);
    ManyToString(&arg0, &arg1, &arg2).shouldEqualDlang("foobarbaz");
}

///
@("nothrow functions")
@system unittest {
    import xlld.memorymanager: allocator;
    mixin(wrapTestFuncsString);
    auto arg = toXlOper(2.0, allocator);
    static assert(__traits(compiles, FuncThrows(&arg)));
}

///
@("FuncAddEverything wrapper is @nogc")
@system @nogc unittest {
    import std.experimental.allocator.mallocator: Mallocator;
    import xlld.framework: freeXLOper;

    mixin(wrapTestFuncsString);
    auto arg = toXlOper(2.0, Mallocator.instance);
    scope(exit) freeXLOper(&arg, Mallocator.instance);
    FuncAddEverything(&arg);
}

///
@("Wrap a function that throws")
@system unittest {
    mixin(wrapTestFuncsString);
    auto arg = toSRef(33.3, theGC);
    FuncThrows(&arg); // should not actually throw
}

///
@("Wrap a function that asserts")
@system unittest {
    mixin(wrapTestFuncsString);
    auto arg = toSRef(33.3, theGC);
    FuncAsserts(&arg); // should not actually throw
}

///
@("Wrap a function that accepts DateTime")
@system unittest {
    import xlld.xlcall: XlType;
    import xlld.conv: stripMemoryBitmask;

    mixin(wrapTestFuncsString);

    const dateTime = DateTime(2017, 12, 31, 1, 2, 3);
    gDates = [42.0];
    gTimes = [33.0];
    gYears = [dateTime.year];
    gMonths = [dateTime.month];
    gDays = [dateTime.day];
    gHours = [dateTime.hour];
    gMinutes = [dateTime.minute];
    gSeconds = [dateTime.second];

    auto arg = dateTime.toXlOper(theGC);
    const ret = DateTimeToDouble(&arg);

    ret.xltype.stripMemoryBitmask.shouldEqual(XlType.xltypeNum);
    ret.val.num.shouldEqual(2017 * 2);
}

///
@("Wrap a function that accepts DateTime[]")
@system unittest {
    mixin(wrapTestFuncsString);

    const dateTime = DateTime(2017, 12, 31, 1, 2, 3);
    gYears = [2017, 2017];
    gMonths = [12, 12];
    gDays = [31, 30];
    gHours = [1, 1];
    gMinutes = [2, 2];
    gSeconds = [3, 3];
    gDates = [10.0, 20.0];
    gTimes = [1.0, 2.0];

    auto arg = [DateTime(2017, 12, 31, 1, 2, 3),
                DateTime(2017, 12, 30, 1, 2, 3)].toXlOper(theGC);
    auto ret = DateTimesToString(&arg);

    ret.shouldEqualDlang("31, 30");
}


/**
 A string to use with `mixin` that wraps a D function
 */
string wrapModuleFunctionStr(string moduleName, string funcName)(in string callingModule = __MODULE__) {

    if(!__ctfe) {
        return "";
    }

    assert(callingModule != moduleName,
           "Cannot use `wrapAll` with __MODULE__");

    import xlld.traits: Async, Identity;
    import xlld.worksheet: Register;
    import std.array: join;
    import std.traits: Parameters, functionAttributes, FunctionAttribute, getUDAs, hasUDA;
    import std.conv: to;
    import std.algorithm: map;
    import std.range: iota;
    import std.format: format;

    mixin("import " ~ moduleName ~ ": " ~ funcName ~ ";");

    alias func = Identity!(mixin(funcName));

    const argsLength = Parameters!(mixin(funcName)).length;
    // e.g. XLOPER12* arg0, XLOPER12* arg1, ...
    auto argsDecl = argsLength.iota.map!(a => `XLOPER12* arg` ~ a.to!string).join(", ");
    // e.g. arg0, arg1, ...
    static if(!hasUDA!(func, Async))
        const argsCall = argsLength.iota.map!(a => `arg` ~ a.to!string).join(", ");
    else {
        import std.range: only, chain, iota;
        import std.conv: text;
        argsDecl ~= ", XLOPER12* asyncHandle";
        const argsCall = chain(only("*asyncHandle"), argsLength.iota.map!(a => `*arg` ~ a.text)).
            map!(a => `cast(immutable)` ~ a)
            .join(", ");
    }
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
    string async;
    static if(hasUDA!(func, Async))
        async = "@Async";

    const returnType = hasUDA!(func, Async) ? "void" : "XLOPER12*";
    const return_ = hasUDA!(func, Async) ? "" : "return ";
    const wrap = hasUDA!(func, Async)
        ? q{wrapAsync!wrappedFunc(Mallocator.instance, %s)}.format(argsCall)
        : q{wrapModuleFunctionImpl!wrappedFunc(gTempAllocator, %s)}.format(argsCall);

    return [
        register,
        async,
        q{
            extern(Windows) %s %s(%s) nothrow %s %s {
                static import %s;
                import xlld.memorymanager: gTempAllocator;
                import std.experimental.allocator.mallocator: Mallocator;
                alias wrappedFunc = %s.%s;
                %s%s;
            }
        }.format(returnType, funcName, argsDecl, nogc, safe,
                 moduleName,
                 moduleName, funcName,
                 return_, wrap),
    ].join("\n");
}

///
@("wrapModuleFunctionStr")
@system unittest {
    import xlld.worksheet;
    import std.traits: getUDAs;

    mixin(wrapModuleFunctionStr!("xlld.test_d_funcs", "FuncAddEverything"));
    alias registerAttrs = getUDAs!(FuncAddEverything, Register);
    static assert(registerAttrs[0].argumentText.value == "Array to add");
}

void wrapAsync(alias F, A, T...)(ref A allocator, immutable XLOPER12 asyncHandle, T args) {

    import xlld.xlcall: XlType, xlAsyncReturn;
    import xlld.framework: Excel12f;
    import std.concurrency: spawn;
    import std.format: format;

    string toDArgsStr() {
        import std.string: join;
        import std.conv: text;

        string[] pointers;
        foreach(i; 0 .. T.length) {
            pointers ~= text("&args[", i, "]");
        }

        return q{toDArgs!F(allocator, %s)}.format(pointers.join(", "));
    }

    mixin(q{alias DArgs = typeof(%s);}.format(toDArgsStr));
    DArgs dArgs;

    // Convert all Excel types to D types. This needs to be done here because the
    // asynchronous part of the computation can't call back into Excel, and converting
    // to D types requires calling coerce.
    try {
        mixin(q{dArgs = %s;}.format(toDArgsStr));
    } catch(Throwable t) {
        static if(isGC!F) {
            import xlld.xll: log;
            log("ERROR: Could not convert to D args for asynchronous function " ~
                __traits(identifier, F));
        }
    }

    try
        spawn(&wrapAsyncImpl!(F, A, DArgs), allocator, asyncHandle, cast(immutable)dArgs);
    catch(Exception ex) {
        XLOPER12 functionRet, ret;
        functionRet.xltype = XlType.xltypeErr;
        Excel12f(xlAsyncReturn, &ret, cast(XLOPER12*)&asyncHandle, &functionRet);
    }
}

void wrapAsyncImpl(alias F, A, T)(ref A allocator, XLOPER12 asyncHandle, T dArgs) {
    import xlld.framework: Excel12f;
    import xlld.xlcall: xlAsyncReturn;
    import std.traits: Unqual;

    // get rid of the temporary memory allocations for the conversions
    scope(exit) freeDArgs(allocator, cast(Unqual!T)dArgs);

    auto functionRet = callWrapped!F(dArgs);
    XLOPER12 xl12ret;
    const errorCode = () @trusted {
        return Excel12f(xlAsyncReturn, &xl12ret, &asyncHandle, functionRet);
    }();
}

// if a function is not @nogc, i.e. it uses the GC
private bool isGC(alias F)() {
    import std.traits: functionAttributes, FunctionAttribute;
    enum nogc = functionAttributes!F & FunctionAttribute.nogc;
    return !nogc;
}

// this has to be a top-level function and can't be declared in the unittest
version(unittest) private double twice(double d) { return d * 2; }

@Flaky
@("wrapAsync")
@system unittest {
    import xlld.test_util: asyncReturn, newAsyncHandle;
    import xlld.conv: fromXlOper;
    import core.time: MonoTime;
    import core.thread;

    const start = MonoTime.currTime;
    auto asyncHandle = newAsyncHandle;
    auto oper = (3.2).toXlOper(theGC);
    wrapAsync!twice(theGC, cast(immutable)asyncHandle, oper);
    const expected = 6.4;
    while(asyncReturn(asyncHandle).fromXlOper!double(theGC) != expected &&
          MonoTime.currTime - start < 1.seconds)
    {
        Thread.sleep(10.msecs);
    }
    asyncReturn(asyncHandle).shouldEqualDlang(expected);
}

/**
 Implement a wrapper for a regular D function
 */
XLOPER12* wrapModuleFunctionImpl(alias wrappedFunc, A, T...)
                                  (ref A allocator, T args) {
    static XLOPER12 ret;

    alias DArgs = typeof(toDArgs!wrappedFunc(allocator, args));
    DArgs dArgs;
    // convert all Excel types to D types
    try {
        dArgs = toDArgs!wrappedFunc(allocator, args);
    } catch(Exception ex) {
        ret = stringOper("#ERROR converting argument to call " ~ __traits(identifier, wrappedFunc));
        return &ret;
    } catch(Throwable t) {
        ret = stringOper("#FATAL ERROR converting argument to call " ~ __traits(identifier, wrappedFunc));
        return &ret;
    }

    // get rid of the temporary memory allocations for the conversions
    scope(exit) freeDArgs(allocator, dArgs);

    return callWrapped!wrappedFunc(dArgs);
}

/**
   Converts a variadic number of XLOPER12* to their equivalent D types
   and returns a tuple
 */
private auto toDArgs(alias wrappedFunc, A, T...)
                    (ref A allocator, T args)
{
    import xlld.xl: coerce, free;
    import xlld.xlcall: XlType;
    import xlld.conv: fromXlOper;
    import std.traits: Parameters, Unqual;
    import std.typecons: Tuple;
    import std.meta: staticMap;

    static XLOPER12 ret;

    XLOPER12[T.length] realArgs;
    // must 1st convert each argument to the "real" type.
    // 2D arrays are passed in as SRefs, for instance
    foreach(i, InputType; Parameters!wrappedFunc) {
        if(args[i].xltype == XlType.xltypeMissing) {
            realArgs[i] = *args[i];
            continue;
        }
        realArgs[i] = coerce(args[i]);
    }

    // scopedCoerce doesn't work with actual Excel
    scope(exit) {
        foreach(ref arg; realArgs)
            free(&arg);
    }

    // the D types to pass to the wrapped function
    Tuple!(staticMap!(Unqual, Parameters!wrappedFunc)) dArgs;

    // convert all Excel types to D types
    foreach(i, InputType; Parameters!wrappedFunc) {
        dArgs[i] = () @trusted { return fromXlOper!InputType(&realArgs[i], allocator); }();
    }

    return dArgs;
}


@("xltypeNum can convert to array")
unittest {
    import std.typecons: tuple;

    void fun(double[] arg) {}
    auto arg = 33.3.toSRef(theGC);
    toDArgs!fun(theGC, &arg).shouldEqual(tuple([33.3]));
}

@("xltypeNil can convert to array")
unittest {
    import xlld.xlcall: XlType;
    import std.typecons: tuple;

    void fun(double[] arg) {}
    XLOPER12 arg;
    arg.xltype = XlType.xltypeNil;
    double[] empty;
    toDArgs!fun(theGC, &arg).shouldEqual(tuple(empty));
}

// Takes a tuple returned by `toDArgs`, calls the wrapped function and returns
// the XLOPER12 result
private XLOPER12* callWrapped(alias wrappedFunc, T)(T dArgs) {

    import xlld.worksheet: Dispose;
    import std.traits: hasUDA, getUDAs;

    static XLOPER12 ret;

     try {
        // call the wrapped function with D types
        auto wrappedRet = wrappedFunc(dArgs.expand);
        ret = excelRet(wrappedRet);

        // dispose of the memory allocated in the wrapped function
        static if(hasUDA!(wrappedFunc, Dispose)) {
            alias disposes = getUDAs!(wrappedFunc, Dispose);
            static assert(disposes.length == 1, "Too many @Dispose for " ~ wrappedFunc.stringof);
            disposes[0].dispose(wrappedRet);
        }

        return &ret;

    } catch(Exception ex) {
         ret = stringOper("#ERROR calling " ~ __traits(identifier, wrappedFunc));
         return &ret;
    } catch(Throwable t) {
         ret = stringOper("#FATAL ERROR calling " ~ __traits(identifier, wrappedFunc));
         return &ret;
    }
}


private XLOPER12 stringOper(in string msg) @safe @nogc nothrow {
    import xlld.conv: toAutoFreeOper;
    import xlld.xlcall: XlType;

    try
        return () @trusted { return msg.toAutoFreeOper; }();
    catch(Exception _) {
        XLOPER12 ret;
        ret.xltype = XlType.xltypeErr;
        return ret;
    }
}



// get excel return value from D return value of wrapped function
private XLOPER12 excelRet(T)(T wrappedRet) {

    import xlld.conv: toAutoFreeOper;
    import std.traits: isArray;

    // Excel crashes if it's returned an empty array, so stop that from happening
    static if(isArray!(typeof(wrappedRet))) {
        if(wrappedRet.length == 0) {
            return "#ERROR: empty result".toAutoFreeOper;
        }

        static if(isArray!(typeof(wrappedRet[0]))) {
            if(wrappedRet[0].length == 0) {
                return "#ERROR: empty result".toAutoFreeOper;
            }
        }
    }

    // convert the return value to an Excel type, tell Excel to call
    // us back to free it afterwards
    return toAutoFreeOper(wrappedRet);
}


private void freeDArgs(A, T)(ref A allocator, ref T dArgs) {
    static if(__traits(compiles, allocator.deallocateAll))
        allocator.deallocateAll;
    else {
        foreach(ref dArg; dArgs) {
            import std.traits: isPointer, isArray;
            static if(isArray!(typeof(dArg)))
            {
                import std.experimental.allocator: disposeMultidimensionalArray;
                allocator.disposeMultidimensionalArray(dArg[]);
            }
            else
                static if(isPointer!(typeof(dArg)))
                {
                    import std.experimental.allocator: dispose;
                    allocator.dispose(dArg);
                }
        }
    }
}

///
@("No memory allocation bugs in wrapModuleFunctionImpl for double return Mallocator")
@system unittest {
    import xlld.test_d_funcs: FuncAddEverything;
    import xlld.xlcall: xlbitDLLFree;

    TestAllocator allocator;
    auto arg = toSRef([1.0, 2.0], theMallocator);
    auto oper = wrapModuleFunctionImpl!FuncAddEverything(allocator, &arg);
    (oper.xltype & xlbitDLLFree).shouldBeTrue;
    allocator.numAllocations.shouldEqual(2);
    oper.shouldEqualDlang(3.0);
    autoFree(oper); // normally this is done by Excel
}

///
@("No memory allocation bugs in wrapModuleFunctionImpl for double[][] return Mallocator")
@system unittest {
    import xlld.test_d_funcs: FuncTripleEverything;
    import xlld.xlcall: xlbitDLLFree, XlType;

    TestAllocator allocator;
    auto arg = toSRef([1.0, 2.0, 3.0], theMallocator);
    auto oper = wrapModuleFunctionImpl!FuncTripleEverything(allocator, &arg);
    (oper.xltype & xlbitDLLFree).shouldBeTrue;
    (oper.xltype & ~xlbitDLLFree).shouldEqual(XlType.xltypeMulti);
    allocator.numAllocations.shouldEqual(2);
    oper.shouldEqualDlang([[3.0, 6.0, 9.0]]);
    autoFree(oper); // normally this is done by Excel
}

///
@("No memory allocation bugs in wrapModuleFunctionImpl for double[][] return pool")
@system unittest {
    import std.typecons: Ternary;
    import xlld.memorymanager: gTempAllocator;
    import xlld.test_d_funcs: FuncTripleEverything;

    auto arg = toSRef([1.0, 2.0, 3.0], gTempAllocator);
    auto oper = wrapModuleFunctionImpl!FuncTripleEverything(gTempAllocator, &arg);
    gTempAllocator.empty.shouldEqual(Ternary.yes);
    oper.shouldEqualDlang([[3.0, 6.0, 9.0]]);
    autoFree(oper); // normally this is done by Excel
}

///
@("No memory allocation bugs in wrapModuleFunctionImpl for string")
@system unittest {
    import std.typecons: Ternary;
    import xlld.memorymanager: gTempAllocator;
    import xlld.test_d_funcs: StringToString;

    auto arg = "foo".toSRef(gTempAllocator);
    auto oper = wrapModuleFunctionImpl!StringToString(gTempAllocator, &arg);
    gTempAllocator.empty.shouldEqual(Ternary.yes);
    oper.shouldEqualDlang("foobar");
}

///
@("No memory allocation bugs in wrapModuleFunctionImpl for Any[][] -> Any[][] -> Any[][] mallocator")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    import xlld.test_d_funcs: FirstOfTwoAnyArrays;

    with(allocatorContext(theGC)) {
        auto dArg = [[any(1.0), any("foo"), any(3.0)], [any(4.0), any(5.0), any(6.0)]];
        auto arg = toXlOper(dArg);
        auto oper = wrapModuleFunctionImpl!FirstOfTwoAnyArrays(theMallocator, &arg, &arg);
        oper.shouldEqualDlang(dArg);
    }
}

///
@("No memory allocation bugs in wrapModuleFunctionImpl for Any[][] -> Any[][] -> Any[][] TestAllocator")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    import xlld.test_d_funcs: FirstOfTwoAnyArrays;

    auto testAllocator = TestAllocator();

    with(allocatorContext(theGC)) {
        auto dArg = [
            [ any(1.0), any("foo"), any(3.0) ],
            [ any(4.0), any(5.0),   any(6.0) ],
        ];
        auto arg = toXlOper(dArg);
        auto oper = wrapModuleFunctionImpl!FirstOfTwoAnyArrays(testAllocator, &arg, &arg);
        oper.shouldEqualDlang(dArg);
    }
}

///
@("Correct number of coercions and frees in wrapModuleFunctionImpl")
@system unittest {
    import xlld.test_d_funcs: FuncAddEverything;
    import xlld.test_util: gNumXlAllocated, gNumXlFree;

    const oldNumAllocated = gNumXlAllocated;
    const oldNumFree = gNumXlFree;

    auto arg = toSRef([1.0, 2.0], theGC);
    auto oper = wrapModuleFunctionImpl!FuncAddEverything(theGC, &arg);

    (gNumXlAllocated - oldNumAllocated).shouldEqual(1);
    (gNumXlFree   - oldNumFree).shouldEqual(1);
}


///
@("Can't return empty 1D array to Excel")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    import xlld.test_d_funcs: EmptyStrings1D;

    with(allocatorContext(theGC)) {
        auto dArg = any(1.0);
        auto arg = toXlOper(dArg);
        auto oper = wrapModuleFunctionImpl!EmptyStrings1D(theGC, &arg);
        oper.shouldEqualDlang("#ERROR: empty result");
    }
}


///
@("Can't return empty 2D array to Excel")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    import xlld.test_d_funcs: EmptyStrings2D;

    with(allocatorContext(theGC)) {
        auto dArg = any(1.0);
        auto arg = toXlOper(dArg);
        auto oper = wrapModuleFunctionImpl!EmptyStrings2D(theGC, &arg);
        oper.shouldEqualDlang("#ERROR: empty result");
    }
}

///
@("Can't return half empty 2D array to Excel")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    import xlld.test_d_funcs: EmptyStringsHalfEmpty2D;

    with(allocatorContext(theGC)) {
        auto dArg = any(1.0);
        auto arg = toXlOper(dArg);
        auto oper = wrapModuleFunctionImpl!EmptyStringsHalfEmpty2D(theGC, &arg);
        oper.shouldEqualDlang("#ERROR: empty result");
    }
}

///
@("issue 25 - make sure to reserve memory for all dArgs")
@system unittest {
    import std.typecons: Ternary;
    import xlld.memorymanager: allocatorContext, MemoryPool;
    import xlld.test_d_funcs: FirstOfTwoAnyArrays;

    auto pool = MemoryPool();

    with(allocatorContext(theGC)) {
        auto dArg = [[any(1.0), any("foo"), any(3.0)], [any(4.0), any(5.0), any(6.0)]];
        auto arg = toSRef(dArg);
        auto oper = wrapModuleFunctionImpl!FirstOfTwoAnyArrays(pool, &arg, &arg);
    }

    pool.empty.shouldEqual(Ternary.yes); // deallocateAll in wrapImpl
}

///
string wrapWorksheetFunctionsString(Modules...)(string callingModule = __MODULE__) {

    if(!__ctfe) {
        return "";
    }

    string ret;
    foreach(module_; Modules) {
        ret ~= wrapModuleWorksheetFunctionsString!module_(callingModule);
    }

    return ret;
}


///
string wrapAll(Modules...)(in string mainModule = __MODULE__) {

    if(!__ctfe) {
        return "";
    }

    import xlld.traits: implGetWorksheetFunctionsString;
    return
        wrapWorksheetFunctionsString!Modules(mainModule) ~
        "\n" ~
        implGetWorksheetFunctionsString!(mainModule) ~
        "\n" ~
        `mixin GenerateDllDef!"` ~ mainModule ~ `";` ~
        "\n";
}

///
@("wrapAll")
unittest  {
    import xlld.memorymanager: allocator;
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll

    mixin(wrapAllTestFuncsString);
    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]], allocator);
    FuncAddEverything(&arg).shouldEqualDlang(60.0);
}



///
@("wrap function with @Dispose")
@safe unittest {
    import xlld.test_util: gTestAllocator;
    import xlld.memorymanager: gTempAllocator;
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll

    // this is needed since gTestAllocator is global, so we can't rely
    // on its destructor
    scope(exit) gTestAllocator.verify;

    mixin(wrapAllTestFuncsString);
    double[4] args = [1.0, 2.0, 3.0, 4.0];
    auto oper = args[].toSRef(gTempAllocator); // don't use TestAllocator
    auto arg = () @trusted { return &oper; }();
    auto ret = () @safe @nogc { return FuncReturnArrayNoGc(arg); }();
    ret.shouldEqualDlang([2.0, 4.0, 6.0, 8.0]);
}

///
@("wrapModuleFunctionStr function that returns Any[][]")
@safe unittest {
    mixin(wrapModuleFunctionStr!("xlld.test_d_funcs", "DoubleArrayToAnyArray"));

    auto oper = [[1.0, 2.0], [3.0, 4.0]].toSRef(theGC);
    auto arg = () @trusted { return &oper; }();
    auto ret = DoubleArrayToAnyArray(arg);

    auto opers = () @trusted { return ret.val.array.lparray[0 .. 4]; }();
    opers[0].shouldEqualDlang(2.0);
    opers[1].shouldEqualDlang(6.0);
    opers[2].shouldEqualDlang("3quux");
    opers[3].shouldEqualDlang("4toto");
}

///
@("wrapModuleFunctionStr int -> int")
@safe unittest {
    mixin(wrapModuleFunctionStr!("xlld.test_d_funcs", "Twice"));

    auto oper = 3.toSRef(theGC);
    auto arg = () @trusted { return &oper; }();
    Twice(arg).shouldEqualDlang(6);
}

///
@("issue 31 - D functions can have const arguments")
@safe unittest {
    mixin(wrapModuleFunctionStr!("xlld.test_d_funcs", "FuncConstDouble"));

    auto oper = (3.0).toSRef(theGC);
    auto arg = () @trusted { return &oper; }();
    FuncConstDouble(arg).shouldEqualDlang(3.0);
}


@("wrapModuleFunctionStr async double -> double")
unittest {
    import xlld.conv: fromXlOper;
    import xlld.traits: Async;
    import xlld.test_util: asyncReturn, newAsyncHandle;
    import core.time: MonoTime;
    import core.thread;

    mixin(wrapModuleFunctionStr!("xlld.test_d_funcs", "AsyncDoubleToDouble"));

    auto oper = (3.0).toXlOper(theGC);
    auto arg = () @trusted { return &oper; }();
    auto asyncHandle = newAsyncHandle;
    () @trusted { AsyncDoubleToDouble(arg, &asyncHandle); }();

    const start = MonoTime.currTime;
    const expected = 6.0;
    while(asyncReturn(asyncHandle).fromXlOper!double(theGC) != expected &&
          MonoTime.currTime - start < 1.seconds)
    {
        Thread.sleep(10.msecs);
    }
    asyncReturn(asyncHandle).shouldEqualDlang(expected);
}

@("wrapModuleFunctionStr () -> NaN")
unittest {
    mixin(wrapModuleFunctionStr!("xlld.test_d_funcs", "NaN"));
    NaN().shouldEqualDlang("#NaN");
}



///
@("wrapAll function that returns Any[][]")
@safe unittest {
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll

    mixin(wrapAllTestFuncsString);

    auto oper = [[1.0, 2.0], [3.0, 4.0]].toSRef(theGC);
    auto arg = () @trusted { return &oper; }();
    auto ret = DoubleArrayToAnyArray(arg);
    scope(exit) () @trusted { autoFree(ret); }(); // usually done by Excel

    auto opers = () @trusted { return ret.val.array.lparray[0 .. 4]; }();
    opers[0].shouldEqualDlang(2.0);
    opers[1].shouldEqualDlang(6.0);
    opers[2].shouldEqualDlang("3quux");
    opers[3].shouldEqualDlang("4toto");
}

///
@("wrapAll function that takes Any[][]")
unittest {
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll
    import xlld.memorymanager: allocatorContext;

    mixin(wrapAllTestFuncsString);

    XLOPER12* ret;
    with(allocatorContext(theGC)) {
        auto oper = [[any(1.0), any(2.0)], [any(3.0), any(4.0)], [any("foo"), any("bar")]].toXlOper(theGC);
        auto arg = () @trusted { return &oper; }();
        ret = AnyArrayToDoubleArray(arg);
    }

    auto opers = () @trusted { return ret.val.array.lparray[0 .. 2]; }();
    opers[0].shouldEqualDlang(3.0); // number of rows
    opers[1].shouldEqualDlang(2.0); // number of columns
}


///
@("wrapAll Any[][] -> Any[][]")
unittest {
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll
    import xlld.memorymanager: allocatorContext;
    import xlld.any: Any;

    mixin(wrapAllTestFuncsString);

    XLOPER12* ret;
    with(allocatorContext(theGC)) {
        auto oper = [[any(1.0), any(2.0)], [any(3.0), any(4.0)], [any("foo"), any("bar")]].toXlOper(theGC);
        auto arg = () @trusted { return &oper; }();
        ret = AnyArrayToAnyArray(arg);
    }

    auto opers = () @trusted { return ret.val.array.lparray[0 .. 6]; }();
    ret.val.array.rows.shouldEqual(3);
    ret.val.array.columns.shouldEqual(2);
    opers[0].shouldEqualDlang(1.0);
    opers[1].shouldEqualDlang(2.0);
    opers[2].shouldEqualDlang(3.0);
    opers[3].shouldEqualDlang(4.0);
    opers[4].shouldEqualDlang("foo");
    opers[5].shouldEqualDlang("bar");
}

///
@("wrapAll Any[][] -> Any[][] -> Any[][]")
unittest {
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll
    import xlld.memorymanager: allocatorContext;
    import xlld.any: Any;

    mixin(wrapAllTestFuncsString);

    XLOPER12* ret;
    with(allocatorContext(theGC)) {
        auto oper = [[any(1.0), any("foo"), any(3.0)], [any(4.0), any(5.0), any(6.0)]].toXlOper(theGC);
        auto arg = () @trusted { return &oper; }();
        ret = FirstOfTwoAnyArrays(arg, arg);
    }

    auto opers = () @trusted { return ret.val.array.lparray[0 .. 6]; }();
    ret.val.array.rows.shouldEqual(2);
    ret.val.array.columns.shouldEqual(3);
    opers[0].shouldEqualDlang(1.0);
    opers[1].shouldEqualDlang("foo");
    opers[2].shouldEqualDlang(3.0);
    opers[3].shouldEqualDlang(4.0);
    opers[4].shouldEqualDlang(5.0);
    opers[5].shouldEqualDlang(6.0);
}

///
@("wrapAll overloaded functions are not wrapped")
unittest {
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll

    mixin(wrapAllTestFuncsString);

    auto double_ = (42.0).toXlOper(theGC);
    auto string_ = "foobar".toXlOper(theGC);
    static assert(!__traits(compiles, Overloaded(&double_).shouldEqualDlang(84.0)));
    static assert(!__traits(compiles, Overloaded(&string_).shouldEqualDlang(84.0)));
}

///
@("wrapAll bool -> int")
@safe unittest {
    import xlld.traits: getAllWorksheetFunctions, GenerateDllDef; // for wrapAll

    mixin(wrapAllTestFuncsString);
    auto string_ = "true".toXlOper(theGC);
    () @trusted { BoolToInt(&string_).shouldEqualDlang(1); }();
}


version(unittest):

string wrapTestFuncsString() {
    return "import xlld.traits: Async;\n" ~
        wrapModuleWorksheetFunctionsString!"xlld.test_d_funcs";
}

string wrapAllTestFuncsString() {
    return "import xlld.traits: Async;\n" ~ wrapAll!"xlld.test_d_funcs";
}
