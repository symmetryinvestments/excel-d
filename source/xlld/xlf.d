/**
   This module provides D versions of xlf Excel functions
 */
module xlld.xlf;


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

private int datePart(double date, int xlfn) @safe @nogc nothrow {
    import xlld.framework: excel12;

    try
        return cast(int)excel12!double(xlfn, date);
    catch(Exception ex)
        return 0;
}
