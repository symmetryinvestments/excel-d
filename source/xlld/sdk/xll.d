/**
	Code from generic.h and generic.d
	Ported to the D Programming Language by Laeeth Isharc (2015)
	This module provides the ceremony that must be done for every XLL.
	At least one module must be linked to this one implementing the
	getWorkSheetFunctions function so that they can be registered
	with Excel.
*/

module xlld.sdk.xll;

import xlld: WorksheetFunction, LPXLOPER12;
version(testingExcelD) import unit_threaded;


alias AutoCloseFunc = void delegate() nothrow;

private AutoCloseFunc[] gAutoCloseFuncs;

/**
   Registers a delegate to be called when the XLL is unloaded
*/
void registerAutoCloseFunc(AutoCloseFunc func) nothrow {
    gAutoCloseFuncs ~= func;
}

/**
   Registers a function to be called when the XLL is unloaded
*/
void registerAutoCloseFunc(void function() nothrow func) nothrow {
    gAutoCloseFuncs ~= { func(); };
}

void callRegisteredAutoCloseFuncs() nothrow {
    foreach(func; gAutoCloseFuncs) func();
}


version(Windows) {
    version(exceldDef)
        enum dllMain = false;
    else version(unittest)
        enum dllMain = false;
    else
        enum dllMain = true;
} else
      enum dllMain = false;


static if(dllMain) {

    import core.sys.windows.windows;

    extern(Windows) BOOL DllMain( HANDLE hDLL, DWORD dwReason, LPVOID lpReserved )
    {
        import core.runtime;
        import core.sys.windows.dll;
        switch (dwReason)
        {
        case DLL_PROCESS_ATTACH:
            Runtime.initialize();
            dll_process_attach( hDLL, true );
            break;
        case DLL_PROCESS_DETACH:
            Runtime.terminate();
            dll_process_detach( hDLL, true );
            break;
        case DLL_THREAD_ATTACH:
            dll_thread_attach( true, true );
            break;
        case DLL_THREAD_DETACH:
            dll_thread_detach( true, true );
            break;
        default:
            break;
        }
        return true;
    }
}

// this function must be defined in a module compiled with
// the current module
// It's extern(C) so that it can be defined in any module
extern(C) WorksheetFunction[] getWorksheetFunctions() @safe pure nothrow;

extern(Windows) int xlAutoOpen() {
    import core.runtime: rt_init;

    rt_init(); // move to DllOpen?
    registerAllWorkSheetFunctions;

    return 1;
}

private void registerAllWorkSheetFunctions() {
    import xlld.memorymanager: allocator;
    import xlld.sdk.framework: Excel12f, freeXLOper;
    import xlld.sdk.xlcall: xlGetName, xlfRegister, XLOPER12;
    import xlld.conv: toXlOper;
    import std.algorithm: map;
    import std.array: array;

    // get name of this XLL, needed to pass to xlfRegister
    static XLOPER12 dllName;
    Excel12f(xlGetName, &dllName);

    foreach(strings; getWorksheetFunctions.map!(a => a.toStringArray)) {
        auto opers = strings.map!(a => a.toXlOper(allocator)).array;
        scope(exit) foreach(ref oper; opers) freeXLOper(&oper, allocator);

        auto args = new LPXLOPER12[opers.length + 1];
        args[0] = &dllName;
        foreach(i; 0 .. opers.length)
            args[i + 1] = &opers[i];

        Excel12f(xlfRegister, cast(LPXLOPER12)null, args);
    }
}

extern(Windows) int xlAutoClose() {
    import core.runtime: rt_term;

    callRegisteredAutoCloseFuncs;

    rt_term;
    return 1;
}

extern(Windows) int xlAutoFree12(LPXLOPER12 arg) nothrow {
    import xlld.memorymanager: autoFree;
    import xlld.sdk.xlcall: xlbitDLLFree;

    if(!(arg.xltype & xlbitDLLFree)) {
        log("[ERROR]: Trying to free XLOPER12 without xlbitDLLFree, ignoring");
        return 0;
    }

    autoFree(arg);
    return 1;
}

extern(Windows) LPXLOPER12 xlAddInManagerInfo12(LPXLOPER12 xAction) {
    import xlld.sdk.xlcall: XLOPER12, XlType, xltypeInt, xlCoerce, xlerrValue;
    import xlld.sdk.framework: Excel12f;
    import xlld.conv: toAutoFreeOper;

    static XLOPER12 xInfo, xIntAction;

    //
    // This code coerces the passed-in value to an integer. This is how the
    // code determines what is being requested. If it receives a 1,
    // it returns a string representing the long name. If it receives
    // anything else, it returns a #VALUE! error.
    //

    //we need an XLOPER12 with a _value_ of xltypeInt
    XLOPER12 arg;
    arg.xltype = XlType.xltypeInt;
    arg.val.w = xltypeInt;

    Excel12f(xlCoerce, &xIntAction, [xAction, &arg]);

    if (xIntAction.val.w == 1) {
        xInfo = "My XLL".toAutoFreeOper;
    } else {
        xInfo.xltype = XlType.xltypeErr;
        xInfo.val.err = xlerrValue;
    }

    return &xInfo;
}


version(Windows) {
    extern(Windows) void OutputDebugStringW(const wchar* fmt) @nogc nothrow;
}


/**
   Polymorphic logging function.
   Prints to the console when unit testing and on Linux,
   otherwise uses the system logger on Windows.
 */
void log(A...)(auto ref A args) @trusted {
    try {
        version(unittest) {
            version(Have_unit_threaded) {
                import unit_threaded: writelnUt;
                writelnUt(args);
            } else {
                import std.stdio: writeln;
                writeln(args);
            }
        } else version(Windows) {
            import nogc.conv: text, toWStringz;
            scope txt = text(args);
            scope wtxt = txt[].toWStringz;
            OutputDebugStringW(wtxt.ptr);
        } else {
            import std.experimental.logger: trace;
            trace(args);
        }
    } catch(Exception e) {
        import core.stdc.stdio: printf;
        printf("Error - could not log\n");
    }
}
