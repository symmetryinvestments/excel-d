/**
   This module provides D versions of xlf Excel functions
 */
module xlld.func.xlf;

import xlld.func.framework: excel12;
import xlld.sdk.xlcall: XLOPER12;

version(testingExcelD) {
    import xlld.test.util;
    import unit_threaded;
    import std.experimental.allocator.gc_allocator: GCAllocator;
    alias theGC = GCAllocator.instance;
}


// should be pure but can't due to calling Excel12
int year(double date) @safe @nogc nothrow {
    import xlld.sdk.xlcall: xlfYear;
    return datePart(date, xlfYear);
}

// should be pure but can't due to calling Excel12
int month(double date) @safe @nogc nothrow {
    import xlld.sdk.xlcall: xlfMonth;
    return datePart(date, xlfMonth);
}

// should be pure but can't due to calling Excel12
int day(double date) @safe @nogc nothrow {
    import xlld.sdk.xlcall: xlfDay;
    return datePart(date, xlfDay);
}

// should be pure but can't due to calling Excel12
int hour(double date) @safe @nogc nothrow {
    import xlld.sdk.xlcall: xlfHour;
    return datePart(date, xlfHour);
}

// should be pure but can't due to calling Excel12
int minute(double date) @safe @nogc nothrow {
    import xlld.sdk.xlcall: xlfMinute;
    return datePart(date, xlfMinute);
}

// should be pure but can't due to calling Excel12
int second(double date) @safe @nogc nothrow {
    import xlld.sdk.xlcall: xlfSecond;
    return datePart(date, xlfSecond);
}

private int datePart(double date, int xlfn) @safe @nogc nothrow {
    // the Excel APIs for some reason return a double
    try
        return cast(int)excel12!double(xlfn, date);
    catch(Exception ex)
        return 0;
}


double date(int year, int month, int day) @safe @nogc nothrow {
    import xlld.sdk.xlcall: xlfDate;
    try
        return excel12!double(xlfDate, year, month, day);
    catch(Exception ex)
        return 0;
}

double time(int year, int month, int day) @safe @nogc nothrow {
    import xlld.sdk.xlcall: xlfTime;
    try
        return excel12!double(xlfTime, year, month, day);
    catch(Exception ex)
        return 0;
}

int rtd(XLOPER12 comId,
        XLOPER12 server,
        XLOPER12 topic0 = XLOPER12(),
        XLOPER12 topic1 = XLOPER12(),
        XLOPER12 topic2 = XLOPER12(),
        XLOPER12 topic3 = XLOPER12(),
        XLOPER12 topic4 = XLOPER12(),
        XLOPER12 topic5 = XLOPER12(),
        XLOPER12 topic6 = XLOPER12(),
        XLOPER12 topic7 = XLOPER12(),
        XLOPER12 topic8 = XLOPER12(),
        XLOPER12 topic9 = XLOPER12())
{
    import xlld.sdk.xlcall: xlfRtd;
    import xlld.sdk.framework: Excel12f;

    XLOPER12 result;
    return Excel12f(xlfRtd, &result, comId, server,
                    topic0, topic1, topic2, topic3, topic4, topic5, topic6, topic7, topic8, topic9);
}

__gshared immutable callerException = new Exception("Error calling xlfCaller");

auto caller() @safe {
    import xlld.sdk.xlcall: xlfCaller, xlretSuccess;
    import xlld.sdk.framework: Excel12f;
    import xlld.func.xl: ScopedOper;

    XLOPER12 result;
    () @trusted {
        if(Excel12f(xlfCaller, &result) != xlretSuccess) {
            throw callerException;
        }
    }();

    return ScopedOper(result);
}

private auto callerCell() @safe {
    import xlld.sdk.xlcall: XlType;
    import xlld.func.xl: coerce, free, Coerced;

    auto oper = caller();

    if(oper.xltype != XlType.xltypeSRef)
        throw new Exception("Caller not a cell");

    return Coerced(oper);
}

@("callerCell throws if caller is string")
unittest {
    import xlld.sdk.xlcall: xlfCaller;
    import xlld.conv.to: toXlOper;

    with(MockXlFunction(xlfCaller, "foobar".toXlOper(theGC))) {
        callerCell.shouldThrowWithMessage("Caller not a cell");
    }
}


@("callerCell with SRef")
unittest {
    import xlld.sdk.xlcall: xlfCaller;

    with(MockXlFunction(xlfCaller, "foobar".toSRef(theGC))) {
        auto oper = callerCell;
        oper.shouldEqualDlang("foobar");
    }
}
