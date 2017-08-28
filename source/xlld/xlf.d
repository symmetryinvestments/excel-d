/**
   This module provides D versions of xlf Excel functions
 */
module xlld.xlf;

import xlld.framework: excel12;
import xlld.xlcall: XLOPER12;


// should be pure but can't due to calling Excel12
int year(double date) @safe @nogc nothrow {
    import xlld.xlcall: xlfYear;
    return datePart(date, xlfYear);
}

// should be pure but can't due to calling Excel12
int month(double date) @safe @nogc nothrow {
    import xlld.xlcall: xlfMonth;
    return datePart(date, xlfMonth);
}

// should be pure but can't due to calling Excel12
int day(double date) @safe @nogc nothrow {
    import xlld.xlcall: xlfDay;
    return datePart(date, xlfDay);
}

// should be pure but can't due to calling Excel12
int hour(double date) @safe @nogc nothrow {
    import xlld.xlcall: xlfHour;
    return datePart(date, xlfHour);
}

// should be pure but can't due to calling Excel12
int minute(double date) @safe @nogc nothrow {
    import xlld.xlcall: xlfMinute;
    return datePart(date, xlfMinute);
}

// should be pure but can't due to calling Excel12
int second(double date) @safe @nogc nothrow {
    import xlld.xlcall: xlfSecond;
    return datePart(date, xlfSecond);
}

private int datePart(double date, int xlfn) @safe @nogc nothrow {
    try
        return cast(int)excel12!double(xlfn, date);
    catch(Exception ex)
        return 0;
}


double date(int year, int month, int day) @safe @nogc nothrow {
    import xlld.xlcall: xlfDate;
    try
        return cast(int)excel12!double(xlfDate, year, month, day);
    catch(Exception ex)
        return 0;
}

double time(int year, int month, int day) @safe @nogc nothrow {
    import xlld.xlcall: xlfTime;
    try
        return cast(int)excel12!double(xlfTime, year, month, day);
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
    import xlld.xlcall: xlfRtd;
    import xlld.framework: Excel12f;

    XLOPER12 result;
    return Excel12f(xlfRtd, &result, comId, server,
                    topic0, topic1, topic2, topic3, topic4, topic5, topic6, topic7, topic8, topic9);
}
