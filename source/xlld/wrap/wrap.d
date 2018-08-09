/**
    Wrapper module for Excel. This module contains the functionality that autowraps
    D code for use within Excel.
*/
module xlld.wrap.wrap;

import xlld.wrap.worksheet;
import xlld.sdk.xlcall: XLOPER12;
import std.typecons: Flag, No;




/**
   Wrap all modules given as strings.
   Also deals with some necessary boilerplate.
 */
string wrapAll(Modules...)
              (Flag!"onlyExports" onlyExports = No.onlyExports,
               in string mainModule = __MODULE__)
{

    if(!__ctfe) {
        return "";
    }

    import xlld.wrap.traits: implGetWorksheetFunctionsString;
    return
        q{import xlld;} ~
        "\n" ~
        wrapWorksheetFunctionsString!Modules(onlyExports, mainModule) ~
        "\n" ~
        implGetWorksheetFunctionsString!(mainModule) ~
        "\n" ~
        `mixin GenerateDllDef!"` ~ mainModule ~ `";` ~
        "\n";
}


/**
   Wrap all modules given as strings.
 */
string wrapWorksheetFunctionsString(Modules...)
                                   (Flag!"onlyExports" onlyExports = No.onlyExports, string callingModule = __MODULE__)
{
    if(!__ctfe) {
        return "";
    }

    string ret;
    foreach(module_; Modules) {
        ret ~= wrapModuleWorksheetFunctionsString!module_(onlyExports, callingModule);
    }

    return ret;
}


/**
   A string to mixin that wraps all eligible functions in the
   given module.
 */
string wrapModuleWorksheetFunctionsString(string moduleName)
                                         (Flag!"onlyExports" onlyExports = No.onlyExports, string callingModule = __MODULE__)
{
    if(!__ctfe) {
        return "";
    }

    import xlld.wrap.traits: Identity;

    mixin(`import ` ~ moduleName ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    string ret;

    foreach(moduleMemberStr; __traits(allMembers, module_))
        ret ~= wrapModuleMember!(moduleName, moduleMemberStr)(onlyExports, callingModule);

    return ret;
}

string wrapModuleMember(string moduleName, string moduleMemberStr)
                       (Flag!"onlyExports" onlyExports = No.onlyExports, string callingModule = __MODULE__)
{
    if(!__ctfe) return "";

    import xlld.wrap.traits: Identity;
    import std.traits: functionAttributes, FunctionAttribute;

    mixin(`import ` ~ moduleName ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    string ret;

    alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));

    static if(isWorksheetFunction!moduleMember) {
        enum numOverloads = __traits(getOverloads, mixin(moduleName), moduleMemberStr).length;
        static if(numOverloads == 1) {
            // if onlyExports is true, then only functions that are "export" are allowed
            // Otherwise, any function will do as long as they're visible (i.e. public)
            const shouldWrap = onlyExports ? __traits(getProtection, moduleMember) == "export" : true;
            if(shouldWrap)
                ret ~= wrapModuleFunctionStr!(moduleName, moduleMemberStr)(callingModule);
        } else
            pragma(msg, "excel-d WARNING: Not wrapping ", moduleMemberStr, " due to it having ",
                   cast(int)numOverloads, " overloads");
    } else {
        /// trying to get a pointer to something is a good way of making sure we can
        /// attempt to evaluate `isSomeFunction` - it's not always possible
        enum canGetPointerToIt = __traits(compiles, &moduleMember);
        static if(canGetPointerToIt) {
            import xlld.wrap.worksheet: Register;
            import std.traits: getUDAs;
            alias registerAttrs = getUDAs!(moduleMember, Register);
            static assert(registerAttrs.length == 0,
                          "excel-d ERROR: Function `" ~ moduleMemberStr ~ "` not eligible for wrapping");
        }
    }

    return ret;
}

/**
 A string to use with `mixin` that wraps a D function
 */
string wrapModuleFunctionStr(string moduleName, string funcName)
                            (in string callingModule = __MODULE__)
{
    if(!__ctfe) {
        return "";
    }

    assert(callingModule != moduleName,
           "Cannot use `wrapAll` with __MODULE__");

    import xlld.wrap.traits: Async, Identity;
    import xlld.wrap.worksheet: Register;
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
    // The function name that Excel actually calls in the binary
    const xlFuncName = pascalCase(funcName);
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
        }.format(returnType, xlFuncName, argsDecl, nogc, safe,
                 moduleName,
                 moduleName, funcName,
                 return_, wrap),
    ].join("\n");
}

string pascalCase(in string func) @safe pure {
    import std.uni: toUpper;
    import std.conv: to;
    return (func[0].toUpper ~ func[1..$].to!dstring).to!string;
}

void wrapAsync(alias F, A, T...)(ref A allocator, immutable XLOPER12 asyncHandle, T args) {

    import xlld.sdk.xlcall: XlType, xlAsyncReturn;
    import xlld.sdk.framework: Excel12f;
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
            import xlld.sdk.xll: log;
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
    import xlld.sdk.framework: Excel12f;
    import xlld.sdk.xlcall: xlAsyncReturn;
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


/**
 Implement a wrapper for a regular D function
 */
XLOPER12* wrapModuleFunctionImpl(alias wrappedFunc, A, T...)
                                (ref A allocator, T args)
{
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
        ret = stringOper("#FATAL ERROR converting argument to call " ~
                         __traits(identifier, wrappedFunc));
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
auto toDArgs(alias wrappedFunc, A, T...)
            (ref A allocator, T args)
{
    import xlld.func.xl: coerce, free;
    import xlld.sdk.xlcall: XlType;
    import xlld.conv.from: fromXlOper;
    import std.traits: Parameters, ParameterDefaults, Unqual;
    import std.typecons: Tuple;
    import std.meta: staticMap;

    static XLOPER12 ret;

    XLOPER12[T.length] coercedOperArgs;
    // must 1st convert each argument to the "real" type.
    // 2D arrays are passed in as SRefs, for instance
    foreach(i, InputType; Parameters!wrappedFunc) {
        if(args[i].xltype == XlType.xltypeMissing) {
            coercedOperArgs[i] = *args[i];
            continue;
        }
        coercedOperArgs[i] = coerce(args[i]);
    }

    // scopedCoerce doesn't work with actual Excel
    scope(exit) {
        static foreach(i; 0 .. args.length) {
            if(args[i].xltype != XlType.xltypeMissing)
                free(&coercedOperArgs[i]);
        }
    }

    // the D types to pass to the wrapped function
    Tuple!(staticMap!(Unqual, Parameters!wrappedFunc)) dArgs;

    // convert all Excel types to D types
    static foreach(i, InputType; Parameters!wrappedFunc) {

        // here we must be careful to use a default value if it exists _and_
        // the oper that was passed in was xlTypeMissing
        static if(is(ParameterDefaults!wrappedFunc[i] == void))
            dArgs[i] = () @trusted { return fromXlOper!InputType(&coercedOperArgs[i], allocator); }();
        else
            dArgs[i] = args[i].xltype == XlType.xltypeMissing
                ? ParameterDefaults!wrappedFunc[i]
                : () @trusted { return fromXlOper!InputType(&coercedOperArgs[i], allocator); }();
    }

    return dArgs;
}


// Takes a tuple returned by `toDArgs`, calls the wrapped function and returns
// the XLOPER12 result
private XLOPER12* callWrapped(alias wrappedFunc, T)(T dArgs) {

    import xlld.wrap.worksheet: Dispose;
    import xlld.sdk.xlcall: XlType;
    import nogc.conv: text;
    import std.traits: hasUDA, getUDAs, ReturnType;

    static XLOPER12 ret;

     try {
        // call the wrapped function with D types
         static if(is(ReturnType!wrappedFunc == void)) {
             wrappedFunc(dArgs.expand);
             ret.xltype = XlType.xltypeNil;
             return &ret;
         } else {
             auto wrappedRet = wrappedFunc(dArgs.expand);
             ret = excelRet(wrappedRet);

             // dispose of the memory allocated in the wrapped function
             static if(hasUDA!(wrappedFunc, Dispose)) {
                 alias disposes = getUDAs!(wrappedFunc, Dispose);
                 static assert(disposes.length == 1, "Too many @Dispose for " ~ wrappedFunc.stringof);
                 disposes[0].dispose(wrappedRet);
             }

             return &ret;
         }

    } catch(Exception ex) {
         ret = stringOper(text("#ERROR calling ", __traits(identifier, wrappedFunc), ": ", ex.msg));
         return &ret;
    } catch(Throwable t) {
         ret = stringOper(text("#FATAL ERROR calling ", __traits(identifier, wrappedFunc), ": ", t.msg));
         return &ret;
    }
}


private XLOPER12 stringOper(in string msg) @safe @nogc nothrow {
    import xlld.conv: toAutoFreeOper;
    import xlld.sdk.xlcall: XlType;

    try
        return () @trusted { return msg.toAutoFreeOper; }();
    catch(Exception _) {
        XLOPER12 ret;
        ret.xltype = XlType.xltypeErr;
        return ret;
    }
}



// get excel return value from D return value of wrapped function
XLOPER12 excelRet(T)(T wrappedRet) {

    import xlld.conv: toAutoFreeOper;
    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.sdk.xlcall: XlType;
    import std.traits: isArray;

    static if(isArray!(typeof(wrappedRet))) {

        // Excel crashes if it's returned an empty array, so stop that from happening
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
    auto ret = toAutoFreeOper(wrappedRet);

    // convert 1D arrays called from a column into a column instead of the default row
    static if(isArray!(typeof(wrappedRet))) {
        static if(!isArray!(typeof(wrappedRet[0]))) { // 1D array
            import xlld.func.xlf: xlfCaller = caller;
            import std.algorithm: swap;

            try {
                auto caller = xlfCaller;
                if(caller.xltype.stripMemoryBitmask == XlType.xltypeSRef) {
                    const isColumnCaller = caller.val.sref.ref_.colLast == caller.val.sref.ref_.colFirst;
                    if(isColumnCaller) swap(ret.val.array.rows, ret.val.array.columns);
                }
            } catch(Exception _) {}
        }
    }

    return ret;
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

// if a function can be wrapped to be baclled by Excel
private template isWorksheetFunction(alias F) {
    import xlld.wrap.traits: isSupportedFunction;
    enum isWorksheetFunction = isSupportedFunction!(F, isWantedType);
}

template isWantedType(T) {
    import xlld.wrap.traits: isOneOf;
    import xlld.any: Any;
    import std.datetime: DateTime;
    import std.traits: Unqual;

    alias U = Unqual!T;

    enum isOneOfTypes = isOneOf!(
        U,
        bool,
        int,
        double, double[], double[][],
        string, string[], string[][],
        Any, Any[], Any[][],
        DateTime, DateTime[], DateTime[][],
        );

    static if(isOneOfTypes)
        enum isWantedType = true;
    else static if(is(U == enum) || is(U == struct))
        enum isWantedType = true;
    else static if(is(U: E[], E))
        enum isWantedType = isWantedType!E;
    else
        enum isWantedType = false;
}


version(testingExcelD) {
    @("isWorksheetFunction")
        @safe pure unittest {
        static import test.d_funcs;
        // the line below checks that the code still compiles even with a private function
        // it might stop compiling in a future version when the deprecation rules for
        // visibility kick in
        static assert(!isWorksheetFunction!(test.d_funcs.shouldNotBeAProblem));
        static assert( isWorksheetFunction!(test.d_funcs.FuncThrows));
        static assert( isWorksheetFunction!(test.d_funcs.DoubleArrayToAnyArray));
        static assert( isWorksheetFunction!(test.d_funcs.Twice));
        static assert( isWorksheetFunction!(test.d_funcs.DateTimeToDouble));
        static assert( isWorksheetFunction!(test.d_funcs.BoolToInt));
        static assert( isWorksheetFunction!(test.d_funcs.FuncSimpleTupleRet));
        static assert( isWorksheetFunction!(test.d_funcs.FuncTupleArrayRet));
        static assert( isWorksheetFunction!(test.d_funcs.FuncDateAndStringRet));
    }
}
