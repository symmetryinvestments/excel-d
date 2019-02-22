module ut.conv.from;

import test;
import xlld.conv.to: toXlOper;
import xlld.conv.from;


@("fromXlOper!double")
@system unittest {
    import xlld.sdk.framework: freeXLOper;

    TestAllocator allocator;
    auto num = 4.0;
    auto oper = num.toXlOper(allocator);
    auto back = oper.fromXlOper!double(allocator);
    back.shouldEqual(num);

    freeXLOper(&oper, allocator);
}


///
@("isNan for fromXlOper!double")
@system unittest {
    import std.math: isNaN;
    import xlld.memorymanager: allocator;
    XLOPER12 oper;
    oper.xltype = XlType.xltypeMissing;
    fromXlOper!double(&oper, allocator).isNaN.shouldBeTrue;
}


@("fromXlOper!double wrong oper type")
@system unittest {
    "foo".toXlOper(theGC).fromXlOper!double(theGC).shouldThrowWithMessage("Wrong type for fromXlOper!double");
}


///
@("fromXlOper!int")
@system unittest {
    42.toXlOper(theGC).fromXlOper!int(theGC).shouldEqual(42);
}

///
@("fromXlOper!int when given xltypeNum")
@system unittest {
    42.0.toXlOper(theGC).fromXlOper!int(theGC).shouldEqual(42);
}

///
@("0 for fromXlOper!int missing oper")
@system unittest {
    XLOPER12 oper;
    oper.xltype = XlType.xltypeMissing;
    oper.fromXlOper!int(theGC).shouldEqual(0);
}

@("fromXlOper!int wrong oper type")
@system unittest {
    "foo".toXlOper(theGC).fromXlOper!int(theGC).shouldThrowWithMessage("Wrong type for fromXlOper!int");
}

///
@("fromXlOper!string missing")
@system unittest {
    import xlld.memorymanager: allocator;
    XLOPER12 oper;
    oper.xltype = XlType.xltypeMissing;
    fromXlOper!string(&oper, allocator).shouldBeNull;
}

///
@("fromXlOper!string")
@system unittest {
    import std.experimental.allocator: dispose;
    import xlld.sdk.framework: freeXLOper;

    TestAllocator allocator;
    auto oper = "foo".toXlOper(allocator);
    auto str = fromXlOper!string(&oper, allocator);
    allocator.numAllocations.shouldEqual(2);

    freeXLOper(&oper, allocator);
    str.shouldEqual("foo");
    allocator.dispose(cast(void[])str);
}

///
@("fromXlOper!string unicode")
@system unittest {
    auto oper = "é".toXlOper(theGC);
    auto str = fromXlOper!string(&oper, theGC);
    str.shouldEqual("é");
}

@("fromXlOper!string allocation failure")
@system unittest {
    auto allocator = FailingAllocator();
    "foo".toXlOper(theGC).fromXlOper!string(allocator).shouldThrowWithMessage("Could not allocate memory for array of char");
}


@("fromXlOper!string wrong oper type")
@system unittest {
    42.toXlOper(theGC).fromXlOper!string(theGC).shouldThrowWithMessage("Wrong type for fromXlOper!string");
}

///
@("fromXlOper any double")
@system unittest {
    any(5.0, theGC).fromXlOper!Any(theGC).shouldEqual(any(5.0, theGC));
}

///
@("fromXlOper any string")
@system unittest {
    any("foo", theGC).fromXlOper!Any(theGC)._impl
        .fromXlOper!string(theGC).shouldEqual("foo");
}

///
@("fromXlOper!string[][]")
unittest {
    import xlld.memorymanager: allocator;
    import xlld.sdk.framework: freeXLOper;

    auto strings = [["foo", "bar", "baz"], ["toto", "titi", "quux"]];
    auto oper = strings.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);
    oper.fromXlOper!(string[][])(allocator).shouldEqual(strings);
}

///
@("fromXlOper!double[][]")
unittest {
    import xlld.memorymanager: allocator;
    import xlld.sdk.framework: freeXLOper;

    auto doubles = [[1.0, 2.0], [3.0, 4.0]];
    auto oper = doubles.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);
    oper.fromXlOper!(double[][])(allocator).shouldEqual(doubles);
}

///
@("fromXlOper!string[][] TestAllocator")
unittest {
    import std.experimental.allocator: disposeMultidimensionalArray;
    import xlld.sdk.framework: freeXLOper;

    TestAllocator allocator;
    auto strings = [["foo", "bar", "baz"], ["toto", "titi", "quux"]];
    auto oper = strings.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(string[][])(allocator);

    allocator.numAllocations.shouldEqual(16);

    freeXLOper(&oper, allocator);
    backAgain.shouldEqual(strings);
    allocator.disposeMultidimensionalArray(cast(void[][][])backAgain);
}

///
@("fromXlOper!string[][] when not all opers are strings")
unittest {
    import xlld.conv.misc: multi;
    import std.experimental.allocator.mallocator: Mallocator;
    alias allocator = theGC;

    const rows = 2;
    const cols = 3;
    auto array = multi(rows, cols, allocator);
    auto opers = array.val.array.lparray[0 .. rows*cols];
    const strings = ["foo", "bar", "baz"];
    const numbers = [1.0, 2.0, 3.0];

    int i;
    foreach(r; 0 .. rows) {
        foreach(c; 0 .. cols) {
            if(r == 0)
                opers[i++] = strings[c].toXlOper(allocator);
            else
                opers[i++] = numbers[c].toXlOper(allocator);
        }
    }

    opers[3].fromXlOper!string(allocator).shouldEqual("1.000000");
    // sanity checks
    opers[0].fromXlOper!string(allocator).shouldEqual("foo");
    opers[3].fromXlOper!double(allocator).shouldEqual(1.0);
    // the actual assertion
    array.fromXlOper!(string[][])(allocator).shouldEqual([["foo", "bar", "baz"],
                                                          ["1.000000", "2.000000", "3.000000"]]);
}


///
@("fromXlOper!double[][] TestAllocator")
unittest {
    import xlld.sdk.framework: freeXLOper;
    import std.experimental.allocator: disposeMultidimensionalArray;

    TestAllocator allocator;
    auto doubles = [[1.0, 2.0], [3.0, 4.0]];
    auto oper = doubles.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(double[][])(allocator);

    allocator.numAllocations.shouldEqual(4);

    freeXLOper(&oper, allocator);
    backAgain.shouldEqual(doubles);
    allocator.disposeMultidimensionalArray(backAgain);
}

///
@("fromXlOper!string[]")
unittest {
    import xlld.memorymanager: allocator;
    import xlld.sdk.framework: freeXLOper;

    auto strings = ["foo", "bar", "baz", "toto", "titi", "quux"];
    auto oper = strings.toXlOper(allocator);
    scope(exit) freeXLOper(&oper, allocator);
    oper.fromXlOper!(string[])(allocator).shouldEqual(strings);
}

///
@("fromXlOper!double[] from row")
unittest {
    import xlld.sdk.xlcall: xlfCaller;

    XLOPER12 caller;
    caller.xltype = XlType.xltypeSRef;
    caller.val.sref.ref_.rwFirst = 1;
    caller.val.sref.ref_.rwLast = 1;
    caller.val.sref.ref_.colFirst = 2;
    caller.val.sref.ref_.colLast = 4;

    with(MockXlFunction(xlfCaller, caller)) {
        auto doubles = [1.0, 2.0, 3.0, 4.0];
        auto oper = doubles.toXlOper(theGC);
        oper.shouldEqualDlang(doubles);
    }
}

///
@("fromXlOper!double[]")
unittest {
    auto doubles = [1.0, 2.0, 3.0, 4.0];
    doubles.toXlOper(theGC).fromXlOper!(double[])(theGC).shouldEqual(doubles);
}

@("fromXlOper!int[]")
unittest {
    auto ints = [1, 2, 3, 4];
    ints.toXlOper(theGC).fromXlOper!(int[])(theGC).shouldEqual(ints);
}


///
@("fromXlOper!string[] TestAllocator")
unittest {
    import std.experimental.allocator: disposeMultidimensionalArray;
    import xlld.sdk.framework: freeXLOper;

    TestAllocator allocator;
    auto strings = ["foo", "bar", "baz", "toto", "titi", "quux"];
    auto oper = strings.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(string[])(allocator);

    allocator.numAllocations.shouldEqual(14);

    backAgain.shouldEqual(strings);
    freeXLOper(&oper, allocator);
    allocator.disposeMultidimensionalArray(cast(void[][])backAgain);
}

///
@("fromXlOper!double[] TestAllocator")
unittest {
    import std.experimental.allocator: dispose;
    import xlld.sdk.framework: freeXLOper;

    TestAllocator allocator;
    auto doubles = [1.0, 2.0, 3.0, 4.0];
    auto oper = doubles.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(double[])(allocator);

    allocator.numAllocations.shouldEqual(2);

    backAgain.shouldEqual(doubles);
    freeXLOper(&oper, allocator);
    allocator.dispose(backAgain);
}

@("fromXlOper!double[][] nil")
@system unittest {
    XLOPER12 oper;
    oper.xltype = XlType.xltypeNil;
    double[][] empty;
    oper.fromXlOper!(double[][])(theGC).shouldEqual(empty);
}


@("fromXlOper any 1D array")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    with(allocatorContext(theGC)) {
        auto array = [any(1.0), any("foo")];
        auto oper = toXlOper(array);
        auto back = fromXlOper!(Any[])(oper);
        back.shouldEqual(array);
    }
}


///
@("fromXlOper Any 2D array")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    with(allocatorContext(theGC)) {
        auto array = [[any(1.0), any(2.0)], [any("foo"), any("bar")]];
        auto oper = toXlOper(array);
        auto back = fromXlOper!(Any[][])(oper);
        back.shouldEqual(array);
    }
}

///
@("fromXlOper!DateTime")
@system unittest {
    XLOPER12 oper;
    auto mockDateTime = MockDateTime(2017, 12, 31, 1, 2, 3);

    const dateTime = oper.fromXlOper!DateTime(theGC);

    dateTime.year.shouldEqual(2017);
    dateTime.month.shouldEqual(12);
    dateTime.day.shouldEqual(31);
    dateTime.hour.shouldEqual(1);
    dateTime.minute.shouldEqual(2);
    dateTime.second.shouldEqual(3);
}

@("fromXlOper!DateTime wrong oper type")
@system unittest {
    42.toXlOper(theGC).fromXlOper!DateTime(theGC).shouldThrowWithMessage(
        "Wrong type for fromXlOper!DateTime");
}


@("fromXlOper!bool when bool")
@system unittest {
    import xlld.sdk.xlcall: XLOPER12, XlType;
    XLOPER12 oper;
    oper.xltype = XlType.xltypeBool;
    oper.val.bool_ = 1;
    oper.fromXlOper!bool(theGC).shouldEqual(true);

    oper.val.bool_ = 0;
    oper.fromXlOper!bool(theGC).shouldEqual(false);

    oper.val.bool_ = 2;
    oper.fromXlOper!bool(theGC).shouldEqual(true);
}

@("fromXlOper!bool when int")
@system unittest {
    42.toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(true);
    0.toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(false);
}

@("fromXlOper!bool when double")
@system unittest {
    33.3.toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(true);
    0.0.toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(false);
}


@("fromXlOper!bool when string")
@system unittest {
    "true".toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(true);
    "True".toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(true);
    "TRUE".toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(true);
    "false".toXlOper(theGC).fromXlOper!bool(theGC).shouldEqual(false);
}

@("fromXlOper!enum")
@system unittest {
    enum Enum {
        foo, bar, baz,
    }

    "bar".toXlOper(theGC).fromXlOper!Enum(theGC).shouldEqual(Enum.bar);
    "quux".toXlOper(theGC).fromXlOper!Enum(theGC).shouldThrowWithMessage("Enum does not have a member named 'quux'");
}

@("fromXlOper!enum wrong type")
@system unittest {
    enum Enum { foo, bar, baz, }

    42.toXlOper(theGC).fromXlOper!Enum(theGC).shouldThrowWithMessage(
        "Wrong type for fromXlOper!Enum");
}


@("1D array to struct")
@system unittest {
    static struct Foo { int x, y; }
    [2, 3].toXlOper(theGC).fromXlOper!Foo(theGC).shouldEqual(Foo(2, 3));
}

@("wrong oper type to struct")
@system unittest {
    static struct Foo { int x, y; }

    2.toXlOper(theGC).fromXlOper!Foo(theGC).shouldThrowWithMessage(
        "Can only convert arrays to structs. Must be either 1xN, Nx1, 2xN or Nx2");
}

@("1D array to struct with wrong length")
@system unittest {

    static struct Foo { int x, y; }

    [2, 3, 4].toXlOper(theGC).fromXlOper!Foo(theGC).shouldThrowWithMessage(
        "1D array length must match number of members in Foo. Expected 2, got 3");

    [2].toXlOper(theGC).fromXlOper!Foo(theGC).shouldThrowWithMessage(
        "1D array length must match number of members in Foo. Expected 2, got 1");
}

@("1D array to struct with wrong type")
@system unittest {
    static struct Foo { int x, y; }

    ["foo", "bar"].toXlOper(theGC).fromXlOper!Foo(theGC).shouldThrowWithMessage(
        "Wrong type converting oper to Foo");
}

@("2D horizontal array to struct")
unittest {
    import xlld.memorymanager: allocatorContext;

    static struct Foo { int x, y, z; }

    with(allocatorContext(theGC)) {
        [[any("x"), any("y"), any("z")], [any(2), any(3), any(4)]].toFrom!Foo.shouldEqual(Foo(2, 3, 4));
    }
}

@("2D vertical array to struct")
unittest {
    import xlld.memorymanager: allocatorContext;

    static struct Foo { int x, y, z; }

    with(allocatorContext(theGC)) {
        [[any("x"), any(2)], [any("y"), any(3)], [any("z"), any(4)]].toFrom!Foo.shouldEqual(Foo(2, 3, 4));
    }
}


@("2D array wrong size")
unittest {
    import xlld.memorymanager: allocatorContext;

    static struct Foo { int x, y, z; }

    with(allocatorContext(theGC)) {
        [[any("x"), any(2)], [any("y"), any(3)], [any("z"), any(4)], [any("w"), any(5)]].toFrom!Foo.
            shouldThrowWithMessage("2D array must be 2x3 or 3x2 for Foo");
    }
}


@("PriceBar[]")
@system /*allocatorContext*/ unittest {

    import xlld.memorymanager: allocatorContext;

    static struct PriceBar {
        double open, high, low, close;
    }

    with(allocatorContext(theGC)) {
        auto array =
        [
            [any("open"), any("high"), any("low"), any("close")],
            [any(1.1),    any(2.2),    any(3.3),   any(4.4)],
            [any(2.1),    any(3.2),    any(4.3),   any(5.4)],
        ];

        array.toFrom!(PriceBar[]).should == [
            PriceBar(1.1, 2.2, 3.3, 4.4),
            PriceBar(2.1, 3.2, 4.3, 5.4),
        ];
    }
}


@("fromXlOperCoerce")
unittest {
    double[][] doubles = [[1, 2, 3, 4], [11, 12, 13, 14]];
    auto doublesOper = toSRef(doubles, theGC);
    doublesOper.fromXlOper!(double[][])(theGC).shouldThrowWithMessage(
        "fromXlOper: oper not of multi type");
    doublesOper.fromXlOperCoerce!(double[][]).shouldEqual(doubles);
}


private auto toFrom(R, T)(T val) {
    import std.experimental.allocator.gc_allocator: GCAllocator;
    return val.toXlOper(GCAllocator.instance).fromXlOper!R(GCAllocator.instance);
}


@("fromXlOper!(Tuple!(double, double))")
@system unittest {
    import std.typecons: tuple, Tuple;
    import xlld.conv.from: fromXlOper;
    tuple(22.2, 33.3)
        .toXlOper(theGC)
        .fromXlOper!(Tuple!(double, double))(theGC)
        .shouldEqual(tuple(22.2, 33.3));
}

@("fromXlOper!(Tuple!(int, int, int))")
@system unittest {
    import std.typecons: tuple, Tuple;
    import xlld.conv.from: fromXlOper;
    tuple(1, 2, 3)
        .toXlOper(theGC)
        .fromXlOper!(Tuple!(int, int, int))(theGC)
        .shouldEqual(tuple(1, 2, 3));
}
