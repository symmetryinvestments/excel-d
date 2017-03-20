/**
	Code from generic.h and generic.d
	Ported to the D Programming Language by Laeeth Isharc (2015)
	This module provides the ceremony that must be done for every XLL.
	At least one module must be linked to this one implementing the
	getWorkSheetFunctions function so that they can be registered
	with Excel.
*/

version(exceldDef) {}
else {

import xlld;
import core.sys.windows.windows;

// this function must be define in a module compiled with
// the current module
// It's extern(C) so that it can be defined in any module
extern(C) WorksheetFunction[] getWorksheetFunctions() @safe pure nothrow;


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


extern(Windows) int xlAutoOpen()
{
	import std.conv;
        import std.algorithm: map;
	import core.runtime:rt_init;

	rt_init(); // move to DllOpen?

        // get name of this XLL, needed to pass to xlfRegister
	static XLOPER12 dllName;
	Excel12f(xlGetName, &dllName, []);

	foreach(functionParams; getWorksheetFunctions.map!(a => a.toStringArray))
		Excel12f(xlfRegister, cast(LPXLOPER12)null, [cast(LPXLOPER12) &dllName] ~ TempStr12(functionParams[]));

	return 1;
}

extern(Windows) int xlAutoClose() {
    import core.runtime: rt_term;
    rt_term;
    return 1;
}

extern(Windows) int xlAutoFree12(LPXLOPER12 arg) {
    assert(arg.xltype & xlbitDLLFree);
    FreeXLOper(arg);
    return 1;
}


extern(Windows) LPXLOPER12 xlAddInManagerInfo12(LPXLOPER12 xAction)
{
	static XLOPER12 xInfo, xIntAction;

	//
	// This code coerces the passed-in value to an integer. This is how the
	// code determines what is being requested. If it receives a 1,
	// it returns a string representing the long name. If it receives
	// anything else, it returns a #VALUE! error.
	//

	Excel12f(xlCoerce, &xIntAction,[xAction, TempInt12(xltypeInt)]);

	if (xIntAction.val.w == 1)
	{
		xInfo.xltype = XlType.xltypeStr;
		xInfo.val.str = TempStr12("My XLL"w).val.str;
	}
	else
	{
		xInfo.xltype = XlType.xltypeErr;
		xInfo.val.err = xlerrValue;
	}

	//Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
	//for UDFs declared as thread safe, use alternate memory allocation mechanisms
	return cast(LPXLOPER12) &xInfo;
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
        import std.experimental.logger;
        trace(args);
    }
}


}
