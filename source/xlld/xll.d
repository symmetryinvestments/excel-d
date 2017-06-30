/**
	Code from generic.h and generic.d
	Ported to the D Programming Language by Laeeth Isharc (2015)
	This module provides the ceremony that must be done for every XLL.
	At least one module must be linked to this one implementing the
	getWorkSheetFunctions function so that they can be registered
	with Excel.
*/

module xlld.xll;

version(exceldDef) {}
else {

    import xlld: WorksheetFunction, LPXLOPER12;

    version(Windows) {

        import core.sys.windows.windows;

        extern(Windows) BOOL DllMain( HANDLE hDLL, DWORD dwReason, LPVOID lpReserved )
        {
            import core.runtime;
            import std.c.windows.windows;
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

        import xlld.memorymanager: allocator;
        import xlld.framework: Excel12f, freeXLOper;
        import xlld.xlcall: xlGetName, xlfRegister, XLOPER12;
        import xlld.wrap: toXlOper;
        import std.algorithm: map;
        import std.array: array;
        import core.runtime:rt_init;

        rt_init(); // move to DllOpen?

        // get name of this XLL, needed to pass to xlfRegister
        static XLOPER12 dllName;
        Excel12f(xlGetName, &dllName, []);

        foreach(strings; getWorksheetFunctions.map!(a => a.toStringArray)) {
            auto opers = strings.map!(a => a.toXlOper(allocator)).array;
            scope(exit) foreach(ref oper; opers) freeXLOper(&oper, allocator);

            auto args = new LPXLOPER12[opers.length + 1];
            args[0] = &dllName;
            foreach(i; 0 .. opers.length)
                args[i + 1] = &opers[i];

            Excel12f(xlfRegister, cast(LPXLOPER12)null, args);
        }

        return 1;
    }

    extern(Windows) int xlAutoClose() {
        import core.runtime: rt_term;
        rt_term;
        return 1;
    }

    extern(Windows) int xlAutoFree12(LPXLOPER12 arg) {
        import xlld.memorymanager: autoFree;
        import xlld.xlcall: xlbitDLLFree;

        assert(arg.xltype & xlbitDLLFree);

        autoFree(arg);
        return 1;
    }

    extern(Windows) LPXLOPER12 xlAddInManagerInfo12(LPXLOPER12 xAction) {
        import xlld.xlcall: XLOPER12, XlType, xltypeInt, xlCoerce, xlerrValue;
        import xlld.wrap: toAutoFreeOper;
        import xlld.framework: Excel12f;

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
        extern(Windows) void OutputDebugStringW(const wchar* fmt) nothrow;

        const(wchar)* toWStringz(in wstring str) nothrow {
            return (str ~ '\0').ptr;
        }

        void log(T...)(T args) {
            import std.conv: text, to;
            try
                OutputDebugStringW(text(args).to!wstring.toWStringz);
            catch(Exception)
                OutputDebugStringW("[DataServer] outputDebug itself failed"w.toWStringz);
        }
    } else version(unittest) {
        void log(T...)(T args) {
        }
    } else {
        void log(T...)(T args) {
            import std.experimental.logger: trace;
            trace(args);
        }
    }
}

version(unittest) {
    extern(C) WorksheetFunction[] getWorksheetFunctions() @safe pure nothrow {
        return [];
    }
}
