module ut.conv.to;

import test;
import xlld.conv.to;

///
@("toExcelOper!int")
unittest {
    auto oper = 42.toXlOper(theGC);
    oper.xltype.shouldEqual(XlType.xltypeInt);
    oper.val.w.shouldEqual(42);
}


///
@("toExcelOper!double")
unittest {
    auto oper = (42.0).toXlOper(theGC);
    oper.xltype.shouldEqual(XlType.xltypeNum);
    oper.val.num.shouldEqual(42.0);
}

///
@("toXlOper!string utf8")
@system unittest {
    import xlld.memorymanager: allocator;
    import xlld.sdk.framework: freeXLOper;
    import std.conv: to;

    const str = "foo";
    auto oper = str.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);

    oper.xltype.shouldEqual(XlType.xltypeStr);
    (cast(int)oper.val.str[0]).shouldEqual(str.length);
    (cast(wchar*)oper.val.str)[1 .. str.length + 1].to!string.shouldEqual(str);
}


///
@("toXlOper!string utf16")
@system unittest {
    import xlld.memorymanager: allocator;
    import xlld.sdk.framework: freeXLOper;

    const str = "foo"w;
    auto oper = str.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);

    oper.xltype.shouldEqual(XlType.xltypeStr);
    (cast(int)oper.val.str[0]).shouldEqual(str.length);
    (cast(wchar*)oper.val.str)[1 .. str.length + 1].shouldEqual(str);
}

///
@("toXlOper!string TestAllocator")
@system unittest {
    import xlld.sdk.framework: freeXLOper;

    auto allocator = TestAllocator();
    auto oper = "foo".toXlOper(allocator);
    allocator.numAllocations.shouldEqual(1);
    freeXLOper(&oper, allocator);
}

///
@("toXlOper!string unicode")
@system unittest {
    import std.utf: byWchar;
    import std.array: array;

    "é".byWchar.array.length.shouldEqual(1);
    "é"w.byWchar.array.length.shouldEqual(1);

    auto oper = "é".toXlOper(theGC);
    const ushort length = oper.val.str[0];
    length.shouldEqual("é"w.length);
}

@("toXlOper!string failing allocator")
@safe unittest {
    import xlld.conv.misc: dup;
    auto allocator = FailingAllocator();
    "foo".toXlOper(theGC).dup(allocator).shouldThrowWithMessage("Failed to allocate memory in dup");
}


///
@("toXlOper string[][]")
@system unittest {
    import xlld.memorymanager: allocator;
    import xlld.sdk.framework: freeXLOper;

    auto oper = [["foo", "bar", "baz"], ["toto", "titi", "quux"]].toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);

    oper.xltype.shouldEqual(XlType.xltypeMulti);
    oper.val.array.rows.shouldEqual(2);
    oper.val.array.columns.shouldEqual(3);
    auto opers = oper.val.array.lparray[0 .. oper.val.array.rows * oper.val.array.columns];

    opers[0].shouldEqualDlang("foo");
    opers[3].shouldEqualDlang("toto");
    opers[5].shouldEqualDlang("quux");
}

///
@("toXlOper string[][] TestAllocator")
@system unittest {
    import xlld.sdk.framework: freeXLOper;

    TestAllocator allocator;
    auto oper = [["foo", "bar", "baz"], ["toto", "titi", "quux"]].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(7);
    freeXLOper(&oper, allocator);
}

///
@("toXlOper double[][]")
@system unittest {
    import xlld.sdk.framework: freeXLOper;

    TestAllocator allocator;
    auto oper = [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(1);
    freeXLOper(&oper, allocator);
}

@("toXlOper!double[][] failing allocation")
@safe unittest {
    auto allocator = FailingAllocator();
    [33.3].toXlOper(allocator).shouldThrowWithMessage("Failed to allocate memory for multi oper");
}

@("toXlOper!double[][] wrong shape")
@safe unittest {
    [[33.3], [1.0, 2.0]].toXlOper(theGC).shouldThrowWithMessage("# of columns must all be the same and aren't");
}

///
@("toXlOper string[]")
@system unittest {
    import xlld.sdk.framework: freeXLOper;

    TestAllocator allocator;
    auto oper = ["foo", "bar", "baz", "toto", "titi", "quux"].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(7);
    freeXLOper(&oper, allocator);
}

///
@("toXlOper any double")
unittest {
    any(5.0, theGC).toXlOper(theGC).shouldEqualDlang(5.0);
}

///
@("toXlOper any string")
unittest {
    any("foo", theGC).toXlOper(theGC).shouldEqualDlang("foo");
}

///
@("toXlOper any double[][]")
unittest {
    any([[1.0, 2.0], [3.0, 4.0]], theGC)
        .toXlOper(theGC).shouldEqualDlang([[1.0, 2.0], [3.0, 4.0]]);
}

///
@("toXlOper any string[][]")
unittest {
    any([["foo", "bar"], ["quux", "toto"]], theGC)
        .toXlOper(theGC).shouldEqualDlang([["foo", "bar"], ["quux", "toto"]]);
}


///
@("toXlOper any[]")
unittest {
    import xlld.memorymanager: allocatorContext;

    with(allocatorContext(theGC)) {
        auto oper = toXlOper([any(42.0), any("foo")]);
        oper.xltype.shouldEqual(XlType.xltypeMulti);
        oper.val.array.lparray[0].shouldEqualDlang(42.0);
        oper.val.array.lparray[1].shouldEqualDlang("foo");
    }
}


///
@("toXlOper mixed 1D array of any")
unittest {
    const a = any([any(1.0, theGC), any("foo", theGC)],
                  theGC);
    auto oper = a.toXlOper(theGC);
    oper.xltype.shouldEqual(XlType.xltypeMulti);

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;
    auto opers = oper.val.array.lparray[0 .. rows * cols];
    opers[0].shouldEqualDlang(1.0);
    opers[1].shouldEqualDlang("foo");
}

///
@("toXlOper any[][]")
unittest {
    import xlld.memorymanager: allocatorContext;

    with(allocatorContext(theGC)) {
        auto oper = toXlOper([[any(42.0), any("foo"), any("quux")], [any("bar"), any(7.0), any("toto")]]);
        oper.xltype.shouldEqual(XlType.xltypeMulti);
        oper.val.array.rows.shouldEqual(2);
        oper.val.array.columns.shouldEqual(3);
        oper.val.array.lparray[0].shouldEqualDlang(42.0);
        oper.val.array.lparray[1].shouldEqualDlang("foo");
        oper.val.array.lparray[2].shouldEqualDlang("quux");
        oper.val.array.lparray[3].shouldEqualDlang("bar");
        oper.val.array.lparray[4].shouldEqualDlang(7.0);
        oper.val.array.lparray[5].shouldEqualDlang("toto");
    }
}


///
@("toXlOper mixed 2D array of any")
unittest {
    const a = any([
                     [any(1.0, theGC), any(2.0, theGC)],
                     [any("foo", theGC), any("bar", theGC)]
                 ],
                 theGC);
    auto oper = a.toXlOper(theGC);
    oper.xltype.shouldEqual(XlType.xltypeMulti);

    const rows = oper.val.array.rows;
    const cols = oper.val.array.columns;
    auto opers = oper.val.array.lparray[0 .. rows * cols];
    opers[0].shouldEqualDlang(1.0);
    opers[1].shouldEqualDlang(2.0);
    opers[2].shouldEqualDlang("foo");
    opers[3].shouldEqualDlang("bar");
}

@("toXlOper!DateTime")
@safe unittest {

    import xlld.sdk.xlcall: xlfDate, xlfTime;

    const dateTime = DateTime(2017, 12, 31, 1, 2, 3);
    {
        auto mockDate = MockXlFunction(xlfDate, 0.1.toXlOper(theGC));
        auto mockTime = MockXlFunction(xlfTime, 0.2.toXlOper(theGC));

        auto oper = dateTime.toXlOper(theGC);

        oper.xltype.shouldEqual(XlType.xltypeNum);
        oper.val.num.shouldApproxEqual(0.3);
    }

    {
        auto mockDate = MockXlFunction(xlfDate, 1.1.toXlOper(theGC));
        auto mockTime = MockXlFunction(xlfTime, 1.2.toXlOper(theGC));

        auto oper = dateTime.toXlOper(theGC);

        oper.xltype.shouldEqual(XlType.xltypeNum);
        oper.val.num.shouldApproxEqual(2.3);
    }
}

///
@("toXlOper!bool when bool")
@system unittest {
    import xlld.sdk.xlcall: XlType;
    {
        const oper = true.toXlOper(theGC);
        oper.xltype.shouldEqual(XlType.xltypeBool);
        oper.val.bool_.shouldEqual(1);
    }

    {
        const oper = false.toXlOper(theGC);
        oper.xltype.shouldEqual(XlType.xltypeBool);
        oper.val.bool_.shouldEqual(0);
    }
}

@("toXlOper!enum")
@safe unittest {

    enum Enum {
        foo,
        bar,
        baz,
    }

    Enum.bar.toXlOper(theGC).shouldEqualDlang("bar");
}


@("toXlOper!struct")
@safe unittest {
    static struct Foo { int x, y; }
    Foo(2, 3).toXlOper(theGC).shouldEqualDlang("Foo(2, 3)");
}
