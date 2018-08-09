module ut.wrap.wrap;

import test;
import xlld.wrap.wrap;
import xlld.conv.to: toXlOper;


// this has to be a top-level function and can't be declared in the unittest
double twice(double d) { return d * 2; }

///
@Flaky
@("wrapAsync")
@system unittest {
    import xlld.test.util: asyncReturn, newAsyncHandle;
    import xlld.conv.from: fromXlOper;
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

@("xltypeNum can convert to array")
@safe unittest {
    import std.typecons: tuple;

    void fun(double[] arg) {}
    auto arg = 33.3.toSRef(theGC);
    auto argPtr = () @trusted { return &arg; }();
    toDArgs!fun(theGC, argPtr).shouldEqual(tuple([33.3]));
}

@("xltypeNil can convert to array")
@safe unittest {
    import xlld.sdk.xlcall: XlType;
    import std.typecons: tuple;

    void fun(double[] arg) {}
    XLOPER12 arg;
    arg.xltype = XlType.xltypeNil;
    double[] empty;
    auto argPtr = () @trusted { return &arg; }();
    toDArgs!fun(theGC, argPtr).shouldEqual(tuple(empty));
}


@("excelRet!double[] from row caller")
unittest {
    import xlld.sdk.xlcall: XlType, xlfCaller;
    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.memorymanager: autoFree;

    XLOPER12 caller;
    caller.xltype = XlType.xltypeSRef;
    caller.val.sref.ref_.rwFirst = 1;
    caller.val.sref.ref_.rwLast = 1;
    caller.val.sref.ref_.colFirst = 2;
    caller.val.sref.ref_.colLast = 4;

    with(MockXlFunction(xlfCaller, caller)) {
        auto doubles = [1.0, 2.0, 3.0, 4.0];
        auto oper = excelRet(doubles);
        scope(exit) autoFree(&oper);

        oper.shouldEqualDlang(doubles);
        oper.xltype.stripMemoryBitmask.shouldEqual(XlType.xltypeMulti);
        oper.val.array.rows.shouldEqual(1);
        oper.val.array.columns.shouldEqual(4);
    }
}

@("excelRet!double[] from column caller")
unittest {
    import xlld.sdk.xlcall: XlType, xlfCaller;
    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.memorymanager: autoFree;

    XLOPER12 caller;
    caller.xltype = XlType.xltypeSRef;
    caller.val.sref.ref_.rwFirst = 1;
    caller.val.sref.ref_.rwLast = 4;
    caller.val.sref.ref_.colFirst = 5;
    caller.val.sref.ref_.colLast = 5;

    with(MockXlFunction(xlfCaller, caller)) {
        auto doubles = [1.0, 2.0, 3.0, 4.0];
        auto oper = excelRet(doubles);
        scope(exit) autoFree(&oper);

        oper.shouldEqualDlang(doubles);
        oper.xltype.stripMemoryBitmask.shouldEqual(XlType.xltypeMulti);
        oper.val.array.rows.shouldEqual(4);
        oper.val.array.columns.shouldEqual(1);
    }
}

@("excelRet!double[] from other caller")
unittest {
    import xlld.sdk.xlcall: XlType, xlfCaller;
    import xlld.conv.misc: stripMemoryBitmask;
    import xlld.memorymanager: autoFree;

    XLOPER12 caller;
    caller.xltype = XlType.xltypeErr;

    with(MockXlFunction(xlfCaller, caller)) {
        auto doubles = [1.0, 2.0, 3.0, 4.0];
        auto oper = excelRet(doubles);
        scope(exit) autoFree(&oper);

        oper.shouldEqualDlang(doubles);
        oper.xltype.stripMemoryBitmask.shouldEqual(XlType.xltypeMulti);
        oper.val.array.rows.shouldEqual(1);
        oper.val.array.columns.shouldEqual(4);
    }
}

@("toDArgs optional arguments")
@safe unittest {
    import std.typecons: tuple;

    static int add(int i, int j = 42);

    XLOPER12 missing;
    missing.xltype = XlType.xltypeMissing;

    auto i = 2.toXlOper(theGC);
    auto j = 3.toXlOper(theGC);

    auto iPtr = () @trusted { return &i; }();
    auto jPtr = () @trusted { return &j; }();
    auto missingPtr = () @trusted { return &missing; }();

    toDArgs!add(theGC, iPtr, jPtr).shouldEqual(tuple(2, 3));
    toDArgs!add(theGC, iPtr, missingPtr).shouldEqual(tuple(2, 42));
}
