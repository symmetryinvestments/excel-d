/**
  framework.d
  Translated from framework.c by Laeeth Isharc

   Purpose:	Framework library for Microsoft Excel.

       This library provides some basic functions
       that help in writing Excel DLLs. It includes
       simple functions for managing memory with XLOPER12s,
       creating temporary XLOPER12s, robustly calling
       Excel12(), and outputting debugging strings
       to the debugger for the current application.

       The main purpose of this library is to help
       you to write cleaner C code for calling Excel.
       For example, using the framework library you
       can write

           Excel12f(xlcDisplay, 0, 2, TempMissing12(), TempBool12(0));

       instead of the more verbose

           XLOPER12 xMissing, bool_;
           xMissing.xltype = xltypeMissing;
           bool_.xltype = xltypeBool;
           bool_.val.bool_ = 0;
           Excel12(xlcDisplay, 0, 2, (LPXLOPER12) &xMissing, (LPXLOPER12) &bool_);


       The library is non-reentrant.

       Define _DEBUG to use the debugging functions.

       Source code is provided so that you may
       enhance this library or optimize it for your
       own application.

   Platform:    Microsoft Windows

   Functions:
                debugPrintf
                GetTempMemory
                FreeAllTempMemory
                Excel
                Excel12f
                TempNum
                TempNum12
                TempStr
                TempStrConst
                TempStr12
                TempBool
                TempBool12
                TempInt
                TempInt12
                TempErr
                TempErr12
                TempActiveRef
                TempActiveRef12
                TempActiveCell
                TempActiveCell12
                TempActiveRow
                TempActiveRow12
                TempActiveColumn
                TempActiveColumn12
                TempMissing
                TempMissing12
                InitFramework
				QuitFramework

*/
module xlld.framework;

/**
   FreeXLOper()

   Purpose:
        Will free any malloc'd memory associated with the given
        LPXLOPER, assuming it has any memory associated with it

   Parameters:

        LPXLOPER pxloper    Pointer to the XLOPER whose associated
                            memory we want to free

   Returns:

   Comments:

*/

import xlld.xlcall;


void FreeXLOper(T)(T pxloper) if(is(T == LPXLOPER) || is(T == LPXLOPER12)) {
    import xlld.memorymanager: allocator;
    FreeXLOper(pxloper, allocator);
}

void FreeXLOper(T, A)(T pxloper, ref A allocator)
    if(is(T == LPXLOPER) || is(T == LPXLOPER12))
{
    import std.experimental.allocator: dispose;

    switch (pxloper.xltype & ~XlType.xlbitDLLFree) with(XlType)
	{
		case xltypeStr:
                    if (pxloper.val.str !is null) {
                        void* bytesPtr = pxloper.val.str;
                        const numBytes = (pxloper.val.str[0] + 1) * wchar.sizeof;
                        allocator.dispose(bytesPtr[0 .. numBytes]);
                    }
			break;
		case xltypeRef:
			if (pxloper.val.mref.lpmref !is null)
				allocator.dispose(pxloper.val.mref.lpmref);
			break;
		case xltypeMulti:
			auto cxloper = pxloper.val.array.rows * pxloper.val.array.columns;
                        const numOpers = cxloper;
			if (pxloper.val.array.lparray !is null)
			{
				auto pxloperFree = pxloper.val.array.lparray;
				while (cxloper > 0)
				{
                                    FreeXLOper(pxloperFree, allocator);
					pxloperFree++;
					cxloper--;
				}
				allocator.dispose(pxloper.val.array.lparray[0 .. numOpers]);
			}
			break;
		case xltypeBigData:
			if (pxloper.val.bigdata.h.lpbData !is null)
				allocator.dispose(pxloper.val.bigdata.h.lpbData);
			break;
		default: // todo: add error handling
			break;
	}
}

@("Free regular XLOPER")
unittest {
    XLOPER oper;
    FreeXLOper(&oper);
}

@("Free XLOPER12")
unittest {
    XLOPER12 oper;
    FreeXLOper(&oper);
}

/**
   Excel12f()

   Purpose:
        A fancy wrapper for the Excel12() function. It also
        does the following:

        (1) Checks that none of the LPXLOPER12 arguments are 0,
            which would indicate that creating a temporary XLOPER12
            has failed. In this case, it doesn't call Excel12
            but it does print a debug message.
        (2) If an error occurs while calling Excel12,
            print a useful debug message.
        (3) When done, free all temporary memory.

        #1 and #2 require _DEBUG to be defined.

   Parameters:

        int xlfn            Function number (xl...) to call
        LPXLOPER12 pxResult Pointer to a place to stuff the result,
                            or 0 if you don't care about the result.
        int count           Number of arguments
        ...                 (all LPXLOPER12s) - the arguments.

   Returns:

        A return code (Some of the xlret... values, as defined
        in XLCALL.H, OR'ed together).

   Comments:
*/

int Excel12f(int xlfn, LPXLOPER12 pxResult, LPXLOPER12[] args) nothrow @nogc
{
	import xlld.memorymanager: FreeAllTempMemory;
	import xlld.xlcallcpp: Excel12v;

	int xlret;

	xlret = Excel12v(xlfn,pxResult,cast(int)args.length, cast(LPXLOPER12 *)args.ptr);

	static if(false) //debug
	{
		if (xlret != xlretSuccess)
		{
			debugPrintf("Error! Excel12(");

			if (xlfn & xlCommand)
				debugPrintf("xlCommand | ");
			if (xlfn & xlSpecial)
				debugPrintf("xlSpecial | ");
			if (xlfn & xlIntl)
				debugPrintf("xlIntl | ");
			if (xlfn & xlPrompt)
				debugPrintf("xlPrompt | ");

			debugPrintf("%u) callback failed:",xlfn & 0x0FFF);

			/* More than one error bit may be on */

			if (xlret & xlretAbort)
			{
				debugPrintf(" Macro Halted\r");
			}

			if (xlret & xlretInvXlfn)
			{
				debugPrintf(" Invalid Function Number\r");
			}

			if (xlret & xlretInvCount)
			{
				debugPrintf(" Invalid Number of Arguments\r");
			}

			if (xlret & xlretInvXloper)
			{
				debugPrintf(" Invalid XLOPER12\r");
			}

			if (xlret & xlretStackOvfl)
			{
				debugPrintf(" Stack Overflow\r");
			}

			if (xlret & xlretFailed)
			{
				debugPrintf(" Command failed\r");
			}

			if (xlret & xlretUncalced)
			{
				debugPrintf(" Uncalced cell\r");
			}

		}
	} // debug

	FreeAllTempMemory();

	return xlret;
}


version(Windows):

debug=0;

import xlld.xlcallcpp;
import std.typecons: Flag, Yes;
import core.sys.windows.windows;

enum rwMaxO8=65536;
enum colMaxO8=256;
enum cchMaxStz=255;
enum MAXSHORTINT =0x7fff;
enum CP_ACP = 0;
enum MAXWORD = 0xFFFF;

static if(false) // debug
{

	/**
	   debugPrintf()

	   Purpose:
	        sends a string to the debugger for the current application.

	   Parameters:

	        LPSTR lpFormat  The format definition string
	        ...             The values to print

	   Returns:

	   Comments:

	*/

	void  debugPrintf(LPSTR lpFormat, ...) // cdecl
	{
		char[256] rgch;
		va_list argList;

		va_start(argList,lpFormat);
		wvsprintfA(rgch,lpFormat,argList);
		va_end(argList);
		OutputDebugStringA(rgch);
	}
}

/**
   Excel()

   Purpose:
        A fancy wrapper for the Excel4() function. It also
        does the following:

        (1) Checks that none of the LPXLOPER arguments are 0,
            which would indicate that creating a temporary XLOPER
            has failed. In this case, it doesn't call Excel
            but it does print a debug message.
        (2) If an error occurs while calling Excel,
            print a useful debug message.
        (3) When done, free all temporary memory.

        #1 and #2 require _DEBUG to be defined.

   Parameters:

        int xlfn            Function number (xl...) to call
        LPXLOPER pxResult   Pointer to a place to stuff the result,
                            or 0 if you don't care about the result.
        int count           Number of arguments
        ...                 (all LPXLOPERs) - the arguments.

   Returns:

        A return code (Some of the xlret... values, as defined
        in XLCALL.H, OR'ed together).

   Comments:
*/

int  Excel(int xlfn, LPXLOPER pxResult, LPXLOPER[] args ...) // cdecl
{
    import xlld.memorymanager: FreeAllTempMemory;
	int xlret;

	xlret = Excel4v(xlfn,pxResult,cast(int)args.length,cast(LPXLOPER *)args.ptr);

	static if(false) //debug
	{

		if (xlret != xlretSuccess)
		{
			debugPrintf("Error! Excel4(");

			if (xlfn & xlCommand)
				debugPrintf("xlCommand | ");
			if (xlfn & xlSpecial)
				debugPrintf("xlSpecial | ");
			if (xlfn & xlIntl)
				debugPrintf("xlIntl | ");
			if (xlfn & xlPrompt)
				debugPrintf("xlPrompt | ");

			debugPrintf("%u) callback failed:",xlfn & 0x0FFF);

			/* More than one error bit may be on */

			if (xlret & xlretAbort)
			{
				debugPrintf(" Macro Halted\r");
			}

			if (xlret & xlretInvXlfn)
			{
				debugPrintf(" Invalid Function Number\r");
			}

			if (xlret & xlretInvCount)
			{
				debugPrintf(" Invalid Number of Arguments\r");
			}

			if (xlret & xlretInvXloper)
			{
				debugPrintf(" Invalid XLOPER\r");
			}

			if (xlret & xlretStackOvfl)
			{
				debugPrintf(" Stack Overflow\r");
			}

			if (xlret & xlretFailed)
			{
				debugPrintf(" Command failed\r");
			}

			if (xlret & xlretUncalced)
			{
				debugPrintf(" Uncalced cell\r");
			}

		}
	} // debug

	FreeAllTempMemory();

	return xlret;
}




/**
   TempNum()

   Purpose:
        Creates a temporary numeric (IEEE floating point) XLOPER.

   Parameters:

        double d        The value

   Returns:

        LPXLOPER        The temporary XLOPER, or 0
                        if GetTempMemory() failed.

   Comments:
*/

LPXLOPER TempNum(Flag!"autoFree" autoFree = Yes.autoFree)(double d)
{
	LPXLOPER lpx;

	lpx = cast(LPXLOPER) GetTempMemory!autoFree(XLOPER.sizeof);

	if (!lpx)
	{
		return null;
	}

	lpx.xltype = xltypeNum;
	lpx.val.num = d;

	return lpx;
}

/**
   TempNum12()

   Purpose:
        Creates a temporary numeric (IEEE floating point) XLOPER12.

   Parameters:

        double d        The value

   Returns:

        LPXLOPER12      The temporary XLOPER12, or 0
                        if GetTempMemory() failed.

   Comments:


*/

LPXLOPER12 TempNum12(Flag!"autoFree" autoFree = Yes.autoFree)(double d)
{
	LPXLOPER12 lpx;

	lpx = cast(LPXLOPER12) GetTempMemory!autoFree(XLOPER12.sizeof);

	if (!lpx)
	{
		return null;
	}

	lpx.xltype = xltypeNum;
	lpx.val.num = d;

	return lpx;
}

/**
   TempStr()

   Purpose:
        Creates a temporary string XLOPER

   Parameters:

        LPSTR lpstr     The string, as a null-terminated
                        C string, with the first byte
                        undefined. This function will
                        count the bytes of the string
                        and insert that count in the
                        first byte of lpstr. Excel cannot
                        handle strings longer than 255
                        characters.

   Returns:

        LPXLOPER        The temporary XLOPER, or 0
                        if GetTempMemory() failed.

   Comments:

        (1) This function has the side effect of inserting
            the byte count as the first character of
            the created string.

        (2) For highest speed, with constant strings,
            you may want to manually count the length of
            the string before compiling, and then avoid
            using this function.

        (3) Behavior is undefined for non-null terminated
            input or strings longer than 255 characters.

   Note: If lpstr passed into TempStr is readonly, TempStr
   will crash your XLL as it will try to modify a read only
   string in place. strings declared on the stack as described below
   are read only by default in VC++

   char *str = " I am a string"

   Use extreme caution while calling TempStr on such strings. Refer to
   VC++ documentation for complier options to ensure that these strings
   are compiled as read write or use TempStrConst instead.

   TempStr is provided mainly for backwards compatability and use of
   TempStrConst is encouraged going forward.
*/

LPXLOPER TempStr(Flag!"autoFree" autoFree = Yes.autoFree)(LPSTR lpstr)
{
	LPXLOPER lpx;

	lpx = cast(LPXLOPER) GetTempMemory!autoFree(XLOPER.sizeof);

	if (!lpx)
	{
		return null;
	}

	lpstr[0] = cast(BYTE) strlen (lpstr+1);
	lpx.xltype = xltypeStr;
	lpx.val.str = lpstr;

	return lpx;
}


/**
   TempStrConst()

   Purpose:
        Creates a temporary string XLOPER from a
        const string with a local copy in temp memory

   Parameters:

        LPSTR lpstr     The string, as a null-terminated
                        C string. This function will
                        count the bytes of the string
                        and insert that count in the
                        first byte of the temp string.
                        Excel cannot handle strings
                        longer than 255 characters.

   Returns:

        LPXLOPER        The temporary XLOPER, or 0
                        if GetTempMemory() failed.

   Comments:

        Will take a string of the form "abc\0" and make a
        temp XLOPER of the form "\003abc"

*/

LPXLOPER TempStrConst(Flag!"autoFree" autoFree = Yes.autoFree)(const LPSTR lpstr)
{
	LPXLOPER lpx;
	LPSTR lps;
	size_t len;

	len = strlen(lpstr);

	lpx = cast(LPXLOPER) (GetTempMemory!autoFree(XLOPER.sizeof + len + 1));

	if (!lpx)
	{
		return null;
	}

	lps = cast(LPSTR)lpx + XLOPER.sizeof;

	lps[0] = cast(BYTE)len;
	//can't strcpy_s because of removal of null-termination
	memcpy_s( cast(ubyte*)lps+1, cast(uint)len+1, cast(ubyte*)lpstr, cast(uint)len);
	lpx.xltype = xltypeStr;
	lpx.val.str = lps;

	return lpx;
}

/**
   TempStr12()

   Purpose:
        Creates a temporary string XLOPER12 from a
        unicode const string with a local copy in
        temp memory

   Parameters:

        wchar lpstr     The string, as a null-terminated
                        unicode string. This function will
                        count the bytes of the string
                        and insert that count in the
                        first byte of the temp string.

   Returns:

        LPXLOPER12      The temporary XLOPER12, or 0
                        if GetTempMemory() failed.

   Comments:

        (1) Fix for const string pointers being passed in to TempStr.
            Note it assumes NO leading space

        (2) Also note that XLOPER12 now uses unicode for the string
            operators

        (3) Will remove the null-termination on the string



   Note: TempStr12 is different from TempStr and is more like TempStrConst
   in its behavior. We have consciously made this choice and deprecated the
   behavior of TempStr going forward. Refer to the note in comment section
   for TempStr to better understand this design decision.
*/

LPXLOPER12[] TempStr12(in wstring[] strings)
{
  LPXLOPER12[] ret;
  ret.length=strings.length;
  foreach(i,str;strings)
    ret[i]=strings[i].TempStr12;
  return ret;
}

wchar* makePascalString(Flag!"autoFree" autoFree = Yes.autoFree)(wchar* str)
{
	auto len=lstrlenW(str);
	auto lpx=GetTempMemory!autoFree((len+1)*2);
	if (lpx is null)
		return null;
	auto lps=cast(wchar*)(cast(CHAR*)lpx);
	lps[0]=cast(BYTE)len;
	wmemcpy_s( lps+1, len+1, str, len);
	return lps;
}

wchar* makePascalString(Flag!"autoFree" autoFree = Yes.autoFree)(wstring str)
{
	wchar* buf;
	auto len=lstrlenW(str);
        wchar* buf = GetTempMemory!autoFree((len + 1) * 2);
	if (buf is null)
		return null;
	buf[0]=cast(wchar)len;
	buf[1..len]=str[1..len];
	return buf;
}

LPXLOPER12 TempStr12(Flag!"autoFree" autoFree = Yes.autoFree)(wstring lpstr)
{
    import xlld.memorymanager: GetTempMemory;
	LPXLOPER12 lpx;
	wchar* lps;
	int len=cast(int)lpstr.length;

	lpx = cast(LPXLOPER12) (GetTempMemory!autoFree(XLOPER12.sizeof + (len+1)*2));

	if (lpx is null)
	{
		return null;
	}

	lps = cast(wchar*)((cast(ubyte*)lpx + XLOPER12.sizeof));

	lps[0] = cast(wchar)len;
	//can't wcscpy_s because of removal of null-termination
	wmemcpy_s( lps+1, len+1, lpstr.ptr, len);
	lpx.xltype = XlType.xltypeStr;
	lpx.val.str = lps;

	return lpx;
}

LPXLOPER12 TempStr12(Flag!"autoFree" autoFree = Yes.autoFree)(const(wchar*) lpstr)
{
	LPXLOPER12 lpx;
	wchar* lps;
	int len;

	len = lstrlenW(lpstr);

	lpx = cast(LPXLOPER12) GetTempMemory!autoFree(XLOPER12.sizeof + (len+1)*2);

	if (!lpx)
	{
		return null;
	}

	lps = cast(wchar*)((cast(CHAR*)lpx + XLOPER12.sizeof));

	lps[0] = cast(BYTE)len;
	//can't wcscpy_s because of removal of null-termination
	wmemcpy_s( lps+1, len+1, lpstr, len);
	lpx.xltype = xltypeStr;
	lpx.val.str = lps;

	return lpx;
}

/**
   TempBool()

   Purpose:
        Creates a temporary logical (true/false) XLOPER.

   Parameters:

        int b           0 - for a false XLOPER
                        Anything else - for a true XLOPER

   Returns:

        LPXLOPER        The temporary XLOPER, or 0
                        if GetTempMemory() failed.

   Comments:
*/

LPXLOPER TempBool(Flag!"autoFree" autoFree = Yes.autoFree)(int b)
{
	LPXLOPER lpx;

	lpx = cast(LPXLOPER) GetTempMemory!autoFree(XLOPER.sizeof);

	if (!lpx)
	{
		return null;
	}

	lpx.xltype = xltypeBool;
	lpx.val.bool_ = b?1:0;
	return lpx;
}

/**
   TempBool12()

   Purpose:
        Creates a temporary logical (true/false) XLOPER12.

   Parameters:

        BOOL b          0 - for a false XLOPER12
                        Anything else - for a true XLOPER12

   Returns:

        LPXLOPER12      The temporary XLOPER12, or 0
                        if GetTempMemory() failed.

   Comments:
*/

LPXLOPER12 TempBool12(Flag!"autoFree" autoFree = Yes.autoFree)(BOOL b)
{
	LPXLOPER12 lpx;

	lpx = cast(LPXLOPER12) GetTempMemory!autoFree(XLOPER12.sizeof);

	if (!lpx)
	{
		return cast(LPXLOPER12)0;
	}

	lpx.xltype = xltypeBool;
	lpx.val.bool_ = b?1:0;

	return lpx;
}

/**
   TempInt()

   Purpose:
        Creates a temporary integer XLOPER.

   Parameters:

        short int i     The integer

   Returns:

        LPXLOPER        The temporary XLOPER, or 0
                        if GetTempMemory() failed.

   Comments:
*/

LPXLOPER TempInt(Flag!"autoFree" autoFree = Yes.autoFree)(short i)
{
	LPXLOPER lpx;

	lpx = cast(LPXLOPER) GetTempMemory!autoFree(XLOPER.sizeof);

	if (!lpx)
	{
		return cast(LPXLOPER)0;
	}

	lpx.xltype = xltypeInt;
	lpx.val.w = i;

	return lpx;
}

/**
   TempInt12()

   Purpose:
            Creates a temporary integer XLOPER12.

   Parameters:

        int i           The integer

   Returns:

        LPXLOPER12      The temporary XLOPER12, or 0
                        if GetTempMemory() failed.

   Comments:

        Note that the int oper has increased in size from
        short int up to int in the 12 opers
*/

LPXLOPER12 TempInt12(Flag!"autoFree" autoFree = Yes.autoFree)(int i)
{
    import xlld.memorymanager: GetTempMemory;
	LPXLOPER12 lpx;

	lpx = cast(LPXLOPER12) GetTempMemory!autoFree(XLOPER12.sizeof);

	if (!lpx)
	{
		return cast(LPXLOPER12)0;
	}

	lpx.xltype = XlType.xltypeInt;
	lpx.val.w = i;

	return lpx;
}

/**
   TempErr()

   Purpose:
        Creates a temporary error XLOPER.

   Parameters:

        WORD err        The error code. One of the xlerr...
                        constants, as defined in XLCALL.H.
                        See the Excel user manual for
                        descriptions about the interpretation
                        of various error codes.

   Returns:

        LPXLOPER        The temporary XLOPER, or 0
                        if GetTempMemory() failed.

   Comments:
*/

LPXLOPER TempErr(Flag!"autoFree" autoFree = Yes.autoFree)(WORD err)
{
	LPXLOPER lpx;

	lpx = cast(LPXLOPER) GetTempMemory!autoFree(XLOPER.sizeof);

	if (!lpx)
	{
		return null;
	}

	lpx.xltype = xltypeErr;
	lpx.val.err = err;

	return lpx;
}

/**
   TempErr12()

   Purpose:
        Creates a temporary error XLOPER12.

   Parameters:

        int err         The error code. One of the xlerr...
                        constants, as defined in XLCALL.H.
                        See the Excel user manual for
                        descriptions about the interpretation
                        of various error codes.

   Returns:

        LPXLOPER12      The temporary XLOPER12, or 0
                        if GetTempMemory() failed.

   Comments:

        Note the paramater has changed from a WORD to an int
        in the new 12 operators
*/

LPXLOPER12 TempErr12(Flag!"autoFree" autoFree = Yes.autoFree)(int err)
{
	LPXLOPER12 lpx;

	lpx = cast(LPXLOPER12) GetTempMemory!autoFree(XLOPER12.sizeof);

	if (!lpx)
	{
		return null;
	}

	lpx.xltype = xltypeErr;
	lpx.val.err = err;

	return lpx;
}

/**
   TempActiveRef()

   Purpose:
        Creates a temporary rectangular reference to the active
        sheet. Remember that the active sheet is the sheet that
        the user sees in front, not the sheet that is currently
        being calculated.

   Parameters:

        WORD rwFirst    (0 based) The first row in the rectangle.
        WORD rwLast     (0 based) The last row in the rectangle.
        BYTE colFirst   (0 based) The first column in the rectangle.
        BYTE colLast    (0 based) The last column in the rectangle.

   Returns:

        LPXLOPER        The temporary XLOPER, or 0
                        if GetTempMemory() failed.

   Comments:
*/
LPXLOPER TempActiveRef(Flag!"autoFree" autoFree = Yes.autoFree)(WORD rwFirst, WORD rwLast, BYTE colFirst, BYTE colLast)
{
    import xlld.memorymanager: allocator;
	LPXLOPER lpx;
	LPXLMREF lpmref;
	int wRet;

	lpx = cast(LPXLOPER) GetTempMemory!autoFree(XLOPER.sizeof);
	lpmref = cast(LPXLMREF) GetTempMemory!autoFree(XLMREF.sizeof);

        if(!lpmref) {
            freePointer(lpx);
            return null;
        }

	/* calling Excel() instead of Excel4() would free all temp memory! */
	wRet = Excel4(xlSheetId, lpx, 0);

	if (wRet != xlretSuccess)
	{
            freePointer(lmref);
            freePointer(lpx);
		return null;
	}
	else
	{
		lpx.xltype = xltypeRef;
		lpx.val.mref.lpmref = lpmref;
		lpmref.count = 1;
		lpmref.reftbl[0].rwFirst = rwFirst;
		lpmref.reftbl[0].rwLast = rwLast;
		lpmref.reftbl[0].colFirst = colFirst;
		lpmref.reftbl[0].colLast = colLast;

		return lpx;
	}
}

private void freePointer(Flag!"autoFree" autoFree = Yes.autoFree)(void* ptr) {
    import xlld.memorymanager: allocator;
    static if(autoFree)
        FreeAllTempMemory;
    else
        allocator.dispose(ptr);
}


/**
   TempActiveRef12()

   Purpose:
        Creates a temporary rectangular reference to the active
        sheet. Remember that the active sheet is the sheet that
        the user sees in front, not the sheet that is currently
        being calculated.

   Parameters:

        RW rwFirst      (0 based) The first row in the rectangle.
        RW rwLast       (0 based) The last row in the rectangle.
        COL colFirst    (0 based) The first column in the rectangle.
        COL colLast     (0 based) The last column in the rectangle.

   Returns:

        LPXLOPER12      The temporary XLOPER12, or 0
                        if GetTempMemory() failed.

   Comments:

        Note that the formal parameters have changed for Excel 2007
        The valid size has increased to accomodate the increase
        in Excel 2007 workbook sizes
*/

LPXLOPER12 TempActiveRef12(Flag!"autoFree" autoFree = Yes.autoFree)(RW rwFirst,RW rwLast,COL colFirst,COL colLast)
{
	LPXLOPER12 lpx;
	LPXLMREF12 lpmref;
	int wRet;

	lpx = cast(LPXLOPER12) GetTempMemory!autoFree(XLOPER12.sizeof);
	lpmref = cast(LPXLMREF12) GetTempMemory!autoFree(XLMREF12.sizeof);

	if (!lpmref) {
            freePointer(lpx);
		return null;
	}

	/* calling Excel12f() instead of Excel12() would free all temp memory! */
	wRet = Excel12(xlSheetId, lpx, []);

	if (wRet != xlretSuccess)
	{
            freePointer(lpmref);
            freePointer(lpx);
	   return null;
	}
	else
	{
		lpx.xltype = xltypeRef;
		lpx.val.mref.lpmref = cast(LPXLMREF12)lpmref;
		lpmref.count = 1;
		lpmref.reftbl[0].rwFirst = rwFirst;
		lpmref.reftbl[0].rwLast = rwLast;
		lpmref.reftbl[0].colFirst = colFirst;
		lpmref.reftbl[0].colLast = colLast;

		return lpx;
	}
}

/**
   TempActiveCell()

   Purpose:
        Creates a temporary reference to a single cell on the active
        sheet. Remember that the active sheet is the sheet that
        the user sees in front, not the sheet that is currently
        being calculated.

   Parameters:

        WORD rw         (0 based) The row of the cell.
        BYTE col        (0 based) The column of the cell.

   Returns:

        LPXLOPER        The temporary XLOPER, or 0
                        if GetTempMemory() failed.

   Comments:
*/


LPXLOPER TempActiveCell(Flag!"autoFree" autoFree = Yes.autoFree)(WORD rw, BYTE col)
{
	return TempActiveRef!autoFree(rw, rw, col, col);
}

/**
   TempActiveCell12()

   Purpose:
        Creates a temporary reference to a single cell on the active
        sheet. Remember that the active sheet is the sheet that
        the user sees in front, not the sheet that is currently
        being calculated.

   Parameters:

        RW rw           (0 based) The row of the cell.
        COL col         (0 based) The column of the cell.

   Returns:

        LPXLOPER12      The temporary XLOPER12, or 0
                        if GetTempMemory() failed.

   Comments:

        Paramter types changed to RW and COL to accomodate the increase
        in sheet sizes introduced in Excel 2007
*/

LPXLOPER12 TempActiveCell12(Flag!"autoFree" autoFree = Yes.autoFree)(RW rw, COL col)
{
	return TempActiveRef12!autoFree(rw, rw, col, col);
}

/**
   TempActiveRow()

   Purpose:
        Creates a temporary reference to an entire row on the active
        sheet. Remember that the active sheet is the sheet that
        the user sees in front, not the sheet that is currently
        being calculated.

   Parameters:

        RW rw           (0 based) The row.

   Returns:

        LPXLOPER        The temporary XLOPER, or 0
                        if GetTempMemory() failed.

   Comments:
*/

LPXLOPER TempActiveRow(Flag!"autoFree" autoFree = Yes.autoFree)(WORD rw)
{
	return TempActiveRef!autoFree(rw, rw, 0, 0xFF);
}

/**
   TempActiveRow12()

   Purpose:
        Creates a temporary reference to an entire row on the active
        sheet. Remember that the active sheet is the sheet that
        the user sees in front, not the sheet that is currently
        being calculated.

   Parameters:

        RW rw           (0 based) The row.

   Returns:

        LPXLOPER12      The temporary XLOPER12, or 0
                        if GetTempMemory() failed.

   Comments:

        Paramter type change to RW to accomodate the increase in sheet
        sizes introduced in Excel 2007
*/

LPXLOPER12 TempActiveRow12(Flag!"autoFree" autoFree = Yes.autoFree)(RW rw)
{
	return TempActiveRef12!autoFree(rw, rw, 0, 0x00003FFF);
}

/**
   TempActiveColumn()

   Purpose:
        Creates a temporary reference to an entire column on the active
        sheet. Remember that the active sheet is the sheet that
        the user sees in front, not the sheet that is currently
        being calculated.

   Parameters:

        LPSTR s         First string
        LPSTR t         Second string

   Returns:

        LPXLOPER        The temporary XLOPER, or 0
                        if GetTempMemory() failed.

   Comments:

*/

LPXLOPER TempActiveColumn(Flag!"autoFree" autoFree = Yes.autoFree)(BYTE col)
{
	return TempActiveRef!autoFree(0, 0xFFFF, col, col);
}

/**
   TempActiveColumn12()

   Purpose:
        Creates a temporary reference to an entire column on the active
        sheet. Remember that the active sheet is the sheet that
        the user sees in front, not the sheet that is currently
        being calculated.

   Parameters:

        COL col	        (0 based) The column.

   Returns:

        LPXLOPER12      The temporary XLOPER12, or 0
                        if GetTempMemory() failed.

   Comments:

        Paramter type change to COL to accomodate the increase in sheet
        sizes introduced in Excel 2007

*/

LPXLOPER12 TempActiveColumn12(Flag!"autoFree" autoFree = Yes.autoFree)(COL col)
{
	return TempActiveRef12!autoFree(0, 0x000FFFFF, col, col);
}


/**
   TempMissing()

   Purpose:
        This is used to simulate a missing argument when
        calling Excel(). It creates a temporary
        "missing" XLOPER.

   Parameters:

   Returns:

        LPXLOPER        The temporary XLOPER, or 0
                        if GetTempMemory() failed.

   Comments:

*/

LPXLOPER TempMissing(Flag!"autoFree" autoFree = Yes.autoFree)()
{
	LPXLOPER lpx;

	lpx = cast(LPXLOPER) GetTempMemory!autoFree(XLOPER.sizeof);

	if (!lpx)
	{
		return null;
	}

	lpx.xltype = xltypeMissing;

	return lpx;
}

/**
   TempMissing12()

   Purpose:
        This is used to simulate a missing argument when
        calling Excel12f(). It creates a temporary
        "missing" XLOPER12.

   Parameters:

   Returns:

        LPXLOPER12      The temporary XLOPER12, or 0
                        if GetTempMemory() failed.

   Comments:

*/

LPXLOPER12 TempMissing12(Flag!"autoFree" autoFree = Yes.autoFree)()
{
	LPXLOPER12 lpx;

	lpx = cast(LPXLOPER12) GetTempMemory!autoFree(XLOPER12.sizeof);

	if (!lpx)
	{
		return null;
	}

	lpx.xltype = xltypeMissing;

	return lpx;
}



/**
   ConvertXLRefToXLRef12()

   Purpose:
        Will attempt to convert an XLREF into the given XREF12

   Parameters:

        LPXLREF pxref       Pointer to the XLREF to copy
        LPXLREF12 pxref12   Pointer to the XLREF12 to copy into

   Returns:

        BOOL                true if the conversion succeeded, false otherwise

   Comments:

*/


BOOL ConvertXLRefToXLRef12(LPXLREF pxref, LPXLREF12 pxref12)
{
	if (pxref.rwLast >= pxref.rwFirst && pxref.colLast >= pxref.colFirst)
	{
		if (pxref.rwFirst >= 0 && pxref.colFirst >= 0)
		{
			pxref12.rwFirst = pxref.rwFirst;
			pxref12.rwLast = pxref.rwLast;
			pxref12.colFirst = pxref.colFirst;
			pxref12.colLast = pxref.colLast;
			return true;
		}
	}
	return false;
}

/**
   ConvertXLRef12ToXLRef()

   Purpose:
        Will attempt to convert an XLREF12 into the given XLREF

   Parameters:

        LPXLREF12 pxref12   Pointer to the XLREF12 to copy
        LPXLREF pxref       Pointer to the XLREF to copy into

   Returns:

        BOOL                true if the conversion succeeded, false otherwise

   Comments:

*/

BOOL ConvertXLRef12ToXLRef(LPXLREF12 pxref12, LPXLREF pxref)
{
	if (pxref12.rwLast >= pxref12.rwFirst && pxref12.colLast >= pxref12.colFirst)
	{
		if (pxref12.rwFirst >=0 && pxref12.colFirst >= 0)
		{
			if (pxref12.rwLast < rwMaxO8 && pxref12.colLast < colMaxO8)
			{
				pxref.rwFirst = cast(WORD)pxref12.rwFirst;
				pxref.rwLast = cast(WORD)pxref12.rwLast;
				pxref.colFirst = cast(BYTE)pxref12.colFirst;
				pxref.colLast = cast(BYTE)pxref12.colLast;
				return true;
			}
		}
	}
	return false;
}

/**
   XLOper12ToXLOper()

   Purpose:
        Conversion routine used to convert from the new XLOPER12
        to the old XLOPER.

   Parameters:

        LPXLOPER12 pxloper12    Pointer to the XLOPER12 to copy
        LPXLOPER pxloper        Pointer to the XLOPER to copy into

   Returns:

        BOOL                    true if the conversion succeeded, false otherwise

   Comments:

        - The caller is responsible for freeing any memory associated with
          the copy if the conversion is a success; FreeXOperT can be
          used, or it may be done by hand.

        - If the conversion fails, any memory the method needed to malloc
          up until the point of failure in the conversion will be freed
          by the method itself during cleanup.
*/

BOOL XLOper12ToXLOper(LPXLOPER12 pxloper12, LPXLOPER pxloper)
{
    import core.stdc.stdlib: malloc,free;

	BOOL fRet;
	BOOL fClean;
	//DWORD xltype;
	WORD cref;
	int cxloper12;
	RW crw;
	COL ccol;
	long cbyte;
	wchar *st;
	char *ast;
	int cch;
	char cach;
	BYTE *pbyte;
	LPXLMREF pmref;
	LPXLREF12 pref12;
	LPXLREF rgref;
	LPXLREF pref;
	LPXLOPER rgxloperConv;
	LPXLOPER pxloperConv;
	LPXLOPER12 pxloper12Conv;

	fClean = false;
	fRet = true;
	auto xltype = pxloper12.xltype;

	switch (xltype) with(XlType)
	{
	case xltypeNum:
		pxloper.val.num = pxloper12.val.num;
		break;
	case xltypeBool:
		pxloper.val.bool_ = cast(typeof(pxloper.val.bool_))pxloper12.val.bool_;
		break;
	case xltypeErr:
		if (pxloper12.val.err > MAXWORD)
		{
			fRet = false;
			// problem... overflow
		}
		else
		{
			pxloper.val.err = cast(WORD)pxloper12.val.err;
		}
		break;
	case xltypeMissing:
	case xltypeNil:
		break;
	case xltypeInt:
		if ((pxloper12.val.w + MAXSHORTINT + 1) >> 16)
		{
			pxloper.val.num = cast(float)pxloper12.val.w;
			xltype = xltypeNum;
		}
		else
		{
			pxloper.val.w = cast(short)pxloper12.val.w;
		}
		break;
	case xltypeStr:
		st = pxloper12.val.str;
		cch = st[0];
		cach = cast(BYTE)cch;

		if (cch > cchMaxStz || cch < 0)
		{
			fRet = false;
		}
		else
		{
			ast = cast(char*)malloc((cach + 2) * char.sizeof);
			if (ast is null)
			{
				fRet = false;
			}
			else
			{
				WideCharToMultiByte(CP_ACP, 0, st + 1, cch, ast + 1, cach, null, null);
				ast[0] = cach;
				ast[cach + 1] = '\0';
				pxloper.val.str = ast;
			}
		}
		break;
	case xltypeFlow:
		if (pxloper12.val.flow.rw > rwMaxO8 || pxloper12.val.flow.col > colMaxO8)
		{
			fRet = false;
		}
		else
		{
			pxloper.val.flow.rw = cast(WORD)pxloper12.val.flow.rw;
			pxloper.val.flow.col = cast(BYTE)pxloper12.val.flow.col;
			pxloper.val.flow.xlflow = pxloper12.val.flow.xlflow;
			pxloper.val.flow.valflow.idSheet = pxloper12.val.flow.valflow.idSheet;
		}
		break;
	case xltypeRef:
		if (pxloper12.val.mref.lpmref && pxloper12.val.mref.lpmref.count > 0)
		{
			pref12 = pxloper12.val.mref.lpmref.reftbl;
			cref = pxloper12.val.mref.lpmref.count;

			pmref = cast(LPXLMREF) (malloc(XLMREF.sizeof) + XLREF.sizeof*(cref-1));
			if (pmref is null)
			{
				fRet = false;
			}
			else
			{
				pmref.count = cref;
				rgref = pmref.reftbl;
				pref = rgref;
				while (cref > 0 && !fClean)
				{
					if (!ConvertXLRef12ToXLRef(pref12, pref))
					{
						fClean = true;
						cref = 0;
					}
					else
					{
						pref++;
						pref12++;
						cref--;
					}
				}
				if (fClean)
				{
					free(pmref);
					fRet = false;
				}
				else
				{
					pxloper.val.mref.lpmref = pmref;
					pxloper.val.mref.idSheet = pxloper12.val.mref.idSheet;
				}
			}
		}
		else
		{
			xltype = xltypeMissing;
		}
		break;
	case xltypeSRef:
		if (pxloper12.val.sref.count != 1)
		{
			fRet = false;
		}
		else if (ConvertXLRef12ToXLRef(&pxloper12.val.sref.ref_, &pxloper.val.sref.ref_))
		{
			pxloper.val.sref.count = 1;
		}
		else
		{
			fRet = false;
		}
		break;
	case xltypeMulti:
		crw = pxloper12.val.array.rows;
		ccol = pxloper12.val.array.columns;
		if (crw > rwMaxO8 || ccol > colMaxO8)
		{
			fRet = false;
		}
		else
		{
			cxloper12 = crw * ccol;
			if (cxloper12 == 0)
			{
				xltype = xltypeMissing;
			}
			else
			{
				rgxloperConv = cast(typeof(rgxloperConv))malloc(cxloper12 * XLOPER.sizeof);
				if (rgxloperConv is null)
				{
					fRet = false;
				}
				else
				{
					pxloperConv = rgxloperConv;
					pxloper12Conv = pxloper12.val.array.lparray;
					while (cxloper12 > 0 && !fClean)
					{
						if (!XLOper12ToXLOper(pxloper12Conv, pxloperConv))
						{
							fClean = true;
							cxloper12 = 0;
						}
						else
						{
							pxloperConv++;
							pxloper12Conv++;
							cxloper12--;
						}
					}
					if (fClean)
					{
						fRet = false;
						while (pxloperConv > rgxloperConv)
						{
							FreeXLOper(pxloperConv);
							pxloperConv--;
						}
						free(rgxloperConv);
					}
					else
					{
						pxloper.val.array.lparray = rgxloperConv;
						pxloper.val.array.rows = cast(typeof(pxloper.val.array.rows))crw;
						pxloper.val.array.columns = cast(typeof(pxloper.val.array.columns))ccol;
					}
				}
			}
		}
		break;
	case xltypeBigData:
		cbyte = pxloper12.val.bigdata.cbData;
		if (pxloper12.val.bigdata.h.lpbData !is null && cbyte > 0)
		{
			pbyte = cast(BYTE *)malloc(cast(uint)cbyte);
			if (pbyte !is null)
			{
				memcpy_s(pbyte, cast(uint)cbyte, pxloper12.val.bigdata.h.lpbData, cast(uint)cbyte);
				pxloper.val.bigdata.h.lpbData = pbyte;
				pxloper.val.bigdata.cbData = cbyte;
			}
			else
			{
				fRet = false;
			}
		}
		else
		{
			fRet = false;
		}
		break;
  default:
    break;
	}
	if (fRet)
	{
		pxloper.xltype = cast(WORD)xltype;
	}
	return fRet;
}

/**
   XLOperToXLOper12()

   Purpose:
        Conversion routine used to convert from the old XLOPER
        to the new XLOPER12.

   Parameters:

        LPXLOPER pxloper        Pointer to the XLOPER to copy
        LPXLOPER12 pxloper12    Pointer to the XLOPER12 to copy into

   Returns:

        BOOL                    true if the conversion succeeded, false otherwise

   Comments:

        - The caller is responsible for freeing any memory associated with
          the copy if the conversion is a success; FreeXLOper12T can be
          used, or it may be done by hand.

        - If the conversion fails, any memory the method needed to malloc
          up until the point of failure in the conversion will be freed
          by the method itself during cleanup.

*/


BOOL XLOperToXLOper12(LPXLOPER pxloper, LPXLOPER12 pxloper12)
{
    import core.stdc.stdlib: malloc,free;

	BOOL fRet;
	BOOL fClean;
	WORD crw;
	WORD ccol;
	WORD cxloper;
	WORD cref12;
	BYTE *pbyte;
	char *ast;
	wchar *st;
	int cach;
	int cch;
	long cbyte;
	LPXLREF pref;
	LPXLREF12 pref12;
	LPXLREF12 rgref12;
	LPXLMREF12 pmref12;
	LPXLOPER pxloperConv;
	LPXLOPER12 rgxloper12Conv;
	LPXLOPER12 pxloper12Conv;

	fClean = false;
	fRet = true;
	auto xltype = pxloper.xltype;

	switch (xltype) with(XlType)
	{
	case xltypeNum:
		pxloper12.val.num = pxloper.val.num;
		break;
	case xltypeBool:
		pxloper12.val.bool_ = pxloper.val.bool_;
		break;
	case xltypeErr:
		pxloper12.val.err = cast(int)pxloper.val.err;
		break;
	case xltypeMissing:
	case xltypeNil:
		break;
	case xltypeInt:
		pxloper12.val.w = pxloper.val.w;
                break;
	case xltypeStr:
		ast = pxloper.val.str;
		if (ast is null)
		{
			fRet = false;
		}
		else
		{
			cach = ast[0];
			cch = cach;
			if (cach > cchMaxStz || cach < 0)
			{
				fRet = false;
			}
			else
			{
				st = cast(wchar*) malloc((cch + 2) * wchar.sizeof);
				if (st is null)
				{
					fRet = false;
				}
				else
				{
					MultiByteToWideChar(CP_ACP, 0, ast + 1, cach, st + 1, cch);
					st[0] = cast(wchar) cch;
					st[cch + 1] = '\0';
					pxloper12.val.str = st;
				}
			}
		}
		break;
	case xltypeFlow:
		pxloper12.val.flow.rw = pxloper.val.flow.rw;
		pxloper12.val.flow.col = pxloper.val.flow.col;
		pxloper12.val.flow.xlflow = pxloper.val.flow.xlflow;
		break;
	case xltypeRef:
		if (pxloper.val.mref.lpmref && pxloper.val.mref.lpmref.count > 0)
		{
			pref = pxloper.val.mref.lpmref.reftbl;
			cref12 = pxloper.val.mref.lpmref.count;

			auto tmp = XLMREF12.sizeof + XLREF12.sizeof*(cref12-1);
      pmref12 = cast(LPXLMREF12) malloc(tmp)[0..tmp];
			if (pmref12 is cast(typeof(pmref12))0)
			{
				fRet = false;
			}
			else
			{
				pmref12.count = cref12;
				rgref12 = pmref12.reftbl;
				pref12 = rgref12;
				while (cref12 > 0 && !fClean)
				{
					if (!ConvertXLRefToXLRef12(pref, pref12))
					{
						fClean = true;
						cref12 = 0;
					}
					else
					{
						pref++;
						pref12++;
						cref12--;
					}
				}
				if (fClean)
				{
					free(cast(void*)pmref12);
					fRet = false;
				}
				else
				{
					pxloper12.val.mref.lpmref = cast(typeof(pxloper12.val.mref.lpmref ))pmref12;
					pxloper12.val.mref.idSheet = pxloper.val.mref.idSheet;
				}
			}
		}
		else
		{
			xltype = xltypeMissing;
		}
		break;
	case xltypeSRef:
		if (pxloper.val.sref.count != 1)
		{
			fRet = false;
		}
		else if (ConvertXLRefToXLRef12(&pxloper.val.sref.ref_, &pxloper12.val.sref.ref_))
		{
			pxloper12.val.sref.count = 1;
		}
		else
		{
			fRet = false;
		}
		break;
	case xltypeMulti:
		crw = pxloper.val.array.rows;
		ccol = pxloper.val.array.columns;
		if (crw > rwMaxO8 || ccol > colMaxO8)
		{
			fRet = false;
		}
		else
		{
			cxloper = cast(typeof(cxloper))(crw * ccol);
			if (cxloper == 0)
			{
				xltype = xltypeMissing;
			}
			else
			{
				rgxloper12Conv = cast(typeof(rgxloper12Conv))malloc(cast(uint)(cxloper * XLOPER12.sizeof))[0..cxloper*XLOPER12.sizeof];
				if (rgxloper12Conv is null)
				{
					fRet = false;
				}
				else
				{
					pxloper12Conv = rgxloper12Conv;
					pxloperConv = pxloper.val.array.lparray;
					while (cxloper > 0 && !fClean)
					{
						if (!XLOperToXLOper12(pxloperConv, pxloper12Conv))
						{
							fClean = true;
							cxloper = 0;
						}
						else
						{
							pxloperConv++;
							pxloper12Conv++;
							cxloper--;
						}
					}
					if (fClean)
					{
						fRet = false;
						while (pxloper12Conv > rgxloper12Conv)
						{
							FreeXLOper(pxloperConv);
							pxloperConv--;
						}
						free(rgxloper12Conv);
					}
					else
					{
						pxloper12.val.array.lparray = rgxloper12Conv;
						pxloper12.val.array.rows = crw;
						pxloper12.val.array.columns = ccol;
					}
				}
			}
		}
		break;
	case xltypeBigData:
		cbyte = pxloper.val.bigdata.cbData;
		if (pxloper.val.bigdata.h.lpbData !is null && cbyte > 0)
		{
			pbyte = cast(BYTE *)malloc(cast(uint)cbyte);
			if (pbyte !is null)
			{
				memcpy_s(cast(ubyte*)pbyte, cast(uint)cbyte, cast(ubyte*)pxloper.val.bigdata.h.lpbData, cast(uint)cbyte);
				pxloper12.val.bigdata.h.lpbData = pbyte;
				pxloper12.val.bigdata.cbData = cbyte;
			}
			else
			{
				fRet = false;
			}
		}
		else
		{
			fRet = false;
		}
		break;
  default:
    break;
	}

	if (fRet)
	{
            pxloper12.xltype = cast(XlType)xltype;
	}
	return fRet;
}



/**
memcpy_s.c - contains memcpy_s routine


Purpose:
       memcpy_s() copies a source memory buffer to a destination buffer.
       Overlapping buffers are not treated specially, so propagation may occur.
*******************************************************************************/


/***
*memcpy_s - Copy source buffer to destination buffer
*
*Purpose:
*       memcpy_s() copies a source memory buffer to a destination memory buffer.
*       This routine does NOT recognize overlapping buffers, and thus can lead
*       to propagation.
*
*       For cases where propagation must be avoided, memmove_s() must be used.
*
*Entry:
*       void *dst = pointer to destination buffer
*       size_t sizeInBytes = size in bytes of the destination buffer
*       const void *src = pointer to source buffer
*       size_t count = number of bytes to copy
*
*Exit:
*       Returns 0 if everything is ok, else return the error code.
*
*Exceptions:
*       Input parameters are validated. Refer to the validation section of the function.
*       On error, the error code is returned and the destination buffer is zeroed.
*
*******************************************************************************/

int memcpy_s(ubyte * dst, size_t sizeInBytes, const ubyte * src, size_t count)
{
    import core.stdc.string: memset, memcpy;

    if (count == 0)
        return 0;

    /* validation section */
    if(dst !is null)
      return -1;
    if (src is null || sizeInBytes < count)
    {
        memset(dst, 0, sizeInBytes);
        if(src is null)
          return -1;
        if(sizeInBytes>=count)
          return -1;
        return -1;
    }

    memcpy(dst, src, count);
    return 0;
}


int wmemcpy_s(wchar* dst, size_t numElements, const wchar* src, size_t count)
{
    import core.stdc.string: memset, memcpy;

  auto sizeInBytes=numElements*wchar.sizeof;
count=count*2;
    if (count == 0)
        return 0;

    /* validation section */
    if(dst is null)
      return -1;
    if (src is null || sizeInBytes < count)
    {
        memset(dst, 0, sizeInBytes);
        if(src is null)
          return -1;
        if(sizeInBytes>=count)
          return -1;
        return -1;
    }

    memcpy(dst, src, count);
    return 0;
}


extern(Windows) short CallerExample()
{
	XLOPER12 xRes;

	Excel12(xlfCaller, &xRes, []);
	Excel12(xlcSelect, cast(LPXLOPER12)0, [cast(LPXLOPER12)&xRes]);
	Excel12(xlFree, cast(LPXLOPER12)0,  [cast(LPXLOPER12)&xRes]);
	return 1;
}
