module ut.wrap.all;

import test;
import ut.wrap.wrapped;
import xlld.wrap;
import xlld.conv.to: toXlOper;

///
@("wrapAll function that returns Any[][]")
@safe unittest {
    import xlld.memorymanager: autoFree;

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
    import xlld.memorymanager: allocatorContext;

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
    import xlld.memorymanager: allocatorContext;
    import xlld.any: Any;

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
    import xlld.memorymanager: allocatorContext;
    import xlld.any: Any;

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
    auto double_ = (42.0).toXlOper(theGC);
    auto string_ = "foobar".toXlOper(theGC);
    static assert(!__traits(compiles, Overloaded(&double_).shouldEqualDlang(84.0)));
    static assert(!__traits(compiles, Overloaded(&string_).shouldEqualDlang(84.0)));
}

///
@("wrapAll bool -> int")
@safe unittest {
    auto string_ = "true".toXlOper(theGC);
    () @trusted { BoolToInt(&string_).shouldEqualDlang(1); }();
}

///
@("wrapAll FuncAddEverything")
unittest  {
    import xlld.memorymanager: allocator;

    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]], allocator);
    FuncAddEverything(&arg).shouldEqualDlang(60.0);
}



///
@("wrapAll FuncReturnArrayNoGC")
@safe unittest {
    import xlld.test.util: gTestAllocator;
    import xlld.memorymanager: gTempAllocator;

    // this is needed since gTestAllocator is global, so we can't rely
    // on its destructor
    scope(exit) gTestAllocator.verify;

    double[4] args = [1.0, 2.0, 3.0, 4.0];
    auto oper = args[].toSRef(gTempAllocator); // don't use gTestAllocator
    auto arg = () @trusted { return &oper; }();
    auto ret = () @safe @nogc { return FuncReturnArrayNoGc(arg); }();
    ret.shouldEqualDlang([2.0, 4.0, 6.0, 8.0]);
}
