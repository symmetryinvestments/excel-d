module ut.wrap.module_;

import xlld.wrap;
import xlld.test.util;
import xlld.sdk.xlcall;
import xlld.conv.to: toXlOper;
import unit_threaded;
import std.datetime;
import std.experimental.allocator.mallocator: Mallocator;
alias theMallocator = Mallocator.instance;


mixin("import xlld.wrap.traits: Async;\n" ~
      wrapModuleWorksheetFunctionsString!"test.d_funcs");


///
@("Wrap double[][] -> double")
@system unittest {
    import xlld.memorymanager: allocator;

    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]], allocator);
    FuncAddEverything(&arg).shouldEqualDlang(60.0);

    arg = toSRef(cast(double[][])[[0, 1, 2, 3], [4, 5, 6, 7]], allocator);
    FuncAddEverything(&arg).shouldEqualDlang(28.0);
}

///
@("Wrap double[][] -> double[][]")
@system unittest {
    import xlld.memorymanager: allocator;

    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]], allocator);
    FuncTripleEverything(&arg).shouldEqualDlang(cast(double[][])[[3, 6, 9, 12], [33, 36, 39, 42]]);

    arg = toSRef(cast(double[][])[[0, 1, 2, 3], [4, 5, 6, 7]], allocator);
    FuncTripleEverything(&arg).shouldEqualDlang(cast(double[][])[[0, 3, 6, 9], [12, 15, 18, 21]]);
}


///
@("Wrap string[][] -> double")
@system unittest {

    import xlld.memorymanager: allocator;

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]], allocator);
    FuncAllLengths(&arg).shouldEqualDlang(29.0);

    arg = toSRef([["", "", "", ""], ["", "", "", ""]], allocator);
    FuncAllLengths(&arg).shouldEqualDlang(0.0);
}

///
@("Wrap string[][] -> double[][]")
@system unittest {

    import xlld.memorymanager: allocator;

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]], allocator);
    FuncLengths(&arg).shouldEqualDlang(cast(double[][])[[3, 3, 3, 4], [4, 4, 4, 4]]);

    arg = toSRef([["", "", ""], ["", "", "huh"]], allocator);
    FuncLengths(&arg).shouldEqualDlang(cast(double[][])[[0, 0, 0], [0, 0, 3]]);
}

///
@("Wrap string[][] -> string[][]")
@system unittest {

    import xlld.memorymanager: allocator;

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]], allocator);
    FuncBob(&arg).shouldEqualDlang([["foobob", "barbob", "bazbob", "quuxbob"],
                                    ["totobob", "titibob", "tutubob", "tetebob"]]);
}

///
@("Wrap string[] -> double")
@system unittest {
    import xlld.memorymanager: allocator;

    auto arg = toSRef([["foo", "bar"], ["baz", "quux"]], allocator);
    FuncStringSlice(&arg).shouldEqualDlang(4.0);
}

///
@("Wrap double[] -> double")
@system unittest {
    import xlld.memorymanager: allocator;
    auto arg = toSRef([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], allocator);
    FuncDoubleSlice(&arg).shouldEqualDlang(6.0);
}

///
@("Wrap double[] -> double[]")
@system unittest {
    import xlld.memorymanager: allocator;
    auto arg = toSRef([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], allocator);
    FuncSliceTimes3(&arg).shouldEqualDlang([3.0, 6.0, 9.0, 12.0, 15.0, 18.0]);
}

///
@("Wrap string[] -> string[]")
@system unittest {
    import xlld.memorymanager: allocator;
    auto arg = toSRef(["quux", "toto"], allocator);
    StringsToStrings(&arg).shouldEqualDlang(["quuxfoo", "totofoo"]);
}

///
@("Wrap string[] -> string")
@system unittest {
    import xlld.memorymanager: allocator;
    auto arg = toSRef(["quux", "toto"], allocator);
    StringsToString(&arg).shouldEqualDlang("quux, toto");
}

///
@("Wrap string -> string")
@system unittest {
    import xlld.memorymanager: allocator;
    auto arg = toXlOper("foo", allocator);
    StringToString(&arg).shouldEqualDlang("foobar");
}

///
@("Wrap string, string, string -> string")
@system unittest {
    import xlld.memorymanager: allocator;
    auto arg0 = toXlOper("foo", allocator);
    auto arg1 = toXlOper("bar", allocator);
    auto arg2 = toXlOper("baz", allocator);
    ManyToString(&arg0, &arg1, &arg2).shouldEqualDlang("foobarbaz");
}

///
@("nothrow functions")
@system unittest {
    import xlld.memorymanager: allocator;
    auto arg = toXlOper(2.0, allocator);
    static assert(__traits(compiles, FuncThrows(&arg)));
}

///
@("FuncAddEverything wrapper is @nogc")
@system @nogc unittest {
    import std.experimental.allocator.mallocator: Mallocator;
    import xlld.sdk.framework: freeXLOper;

    auto arg = toXlOper(2.0, Mallocator.instance);
    scope(exit) freeXLOper(&arg, Mallocator.instance);
    FuncAddEverything(&arg);
}

///
@("Wrap a function that throws")
@system unittest {
    auto arg = toSRef(33.3, theGC);
    FuncThrows(&arg); // should not actually throw
}

///
@("Wrap a function that asserts")
@system unittest {
    auto arg = toSRef(33.3, theGC);
    FuncAsserts(&arg); // should not actually throw
}

///
@("Wrap a function that accepts DateTime")
@system unittest {
    import xlld.sdk.xlcall: XlType;
    import xlld.conv.misc: stripMemoryBitmask;
    import std.conv: text;

    // the argument doesn't matter since we mock extracting the year from it
    // but it does have to be of double type (DateTime for Excel)
    auto arg = 33.3.toXlOper(theGC);

    const year = 2017;
    const mock = MockDateTime(year, 1, 2, 3, 4, 5);
    const ret = DateTimeToDouble(&arg);

    try
        ret.xltype.stripMemoryBitmask.shouldEqual(XlType.xltypeNum);
    catch(Exception _)
        assert(false, text("Expected xltypeNum, got ", *ret));

    ret.val.num.shouldEqual(year * 2);
}

///
@("Wrap a function that accepts DateTime[]")
@system unittest {
    //the arguments don't matter since we mock extracting year, etc. from them
    //they just need to be double (DateTime to Excel)
    auto arg = [0.1, 0.2].toXlOper(theGC);

    auto mockDateTimes = MockDateTimes(DateTime(1, 1, 31),
                                       DateTime(1, 1, 30));

    auto ret = DateTimesToString(&arg);

    ret.shouldEqualDlang("31, 30");
}


@Serial
@("Wrap a function that takes an enum")
@safe unittest {
    import test.d_funcs: MyEnum;

    auto arg = MyEnum.baz.toXlOper(theGC);
    auto ret = () @trusted { return FuncMyEnumArg(&arg); }();
    ret.shouldEqualDlang("prefix_baz");
}

@Serial
@("Wrap a function that returns an enum")
@safe unittest {
    import test.d_funcs: MyEnum;

    auto arg = 1.toXlOper(theGC);
    auto ret = () @trusted { return FuncMyEnumRet(&arg); }();
    ret.shouldEqualDlang("bar");
}


@Serial
@("Register a custom to enum conversion")
@safe unittest {
    import std.conv: to;
    import test.d_funcs: MyEnum;
    import xlld.conv.from: registerConversionTo, unregisterConversionTo;

    registerConversionTo!MyEnum((str) => str[3 .. $].to!MyEnum);
    scope(exit) unregisterConversionTo!MyEnum;

    auto arg = "___baz".toXlOper(theGC);
    auto ret = () @trusted { return FuncMyEnumArg(&arg); }();

    ret.shouldEqualDlang("prefix_baz");
}

@Serial
@("Register a custom from enum conversion")
@safe unittest {

    import std.conv: text;
    import test.d_funcs: MyEnum;
    import xlld.conv: registerConversionFrom, unregisterConversionFrom;

    registerConversionFrom!MyEnum((val) => "___" ~ text(cast(MyEnum)val));
    scope(exit)unregisterConversionFrom!MyEnum;

    auto arg = 1.toXlOper(theGC);
    auto ret = () @trusted { return FuncMyEnumRet(&arg); }();

    ret.shouldEqualDlang("___bar");
}

@("Wrap a function that takes a struct using 1D array")
unittest {
    auto arg = [2, 3].toXlOper(theGC);
    auto ret = () @trusted { return FuncPointArg(&arg); }();

    ret.shouldEqualDlang(5);
}

@("Wrap a function that returns a struct")
unittest {
    auto arg1 = 2.toXlOper(theGC);
    auto arg2 = 3.toXlOper(theGC);
    auto ret = () @trusted { return FuncPointRet(&arg1, &arg2); }();

    ret.shouldEqualDlang("Point(2, 3)");
}


///
@("wrapModuleFunctionStr")
@system unittest {
    import xlld.wrap.worksheet;
    import std.traits: getUDAs;

    mixin(wrapModuleFunctionStr!("test.d_funcs", "FuncAddEverything"));
    alias registerAttrs = getUDAs!(FuncAddEverything, Register);
    static assert(registerAttrs[0].argumentText.value == "Array to add");
}


///
@("No memory allocation bugs in wrapModuleFunctionImpl for double return Mallocator")
@system unittest {
    import test.d_funcs: FuncAddEverything;
    import xlld.sdk.xlcall: xlbitDLLFree;
    import xlld.memorymanager: autoFree;

    TestAllocator allocator;
    auto arg = toSRef([1.0, 2.0], theMallocator);
    auto oper = wrapModuleFunctionImpl!FuncAddEverything(allocator, &arg);
    (oper.xltype & xlbitDLLFree).shouldBeTrue;
    allocator.numAllocations.shouldEqual(2);
    oper.shouldEqualDlang(3.0);
    autoFree(oper); // normally this is done by Excel
}

///
@("No memory allocation bugs in wrapModuleFunctionImpl for double[][] return Mallocator")
@system unittest {
    import test.d_funcs: FuncTripleEverything;
    import xlld.sdk.xlcall: xlbitDLLFree, XlType;
    import xlld.memorymanager: autoFree;

    TestAllocator allocator;
    auto arg = toSRef([1.0, 2.0, 3.0], theMallocator);
    auto oper = wrapModuleFunctionImpl!FuncTripleEverything(allocator, &arg);
    (oper.xltype & xlbitDLLFree).shouldBeTrue;
    (oper.xltype & ~xlbitDLLFree).shouldEqual(XlType.xltypeMulti);
    allocator.numAllocations.shouldEqual(2);
    oper.shouldEqualDlang([[3.0, 6.0, 9.0]]);
    autoFree(oper); // normally this is done by Excel
}

///
@("No memory allocation bugs in wrapModuleFunctionImpl for double[][] return pool")
@system unittest {
    import std.typecons: Ternary;
    import xlld.memorymanager: gTempAllocator, autoFree;
    import test.d_funcs: FuncTripleEverything;

    auto arg = toSRef([1.0, 2.0, 3.0], gTempAllocator);
    auto oper = wrapModuleFunctionImpl!FuncTripleEverything(gTempAllocator, &arg);
    gTempAllocator.empty.shouldEqual(Ternary.yes);
    oper.shouldEqualDlang([[3.0, 6.0, 9.0]]);
    autoFree(oper); // normally this is done by Excel
}

///
@("No memory allocation bugs in wrapModuleFunctionImpl for string")
@system unittest {
    import std.typecons: Ternary;
    import xlld.memorymanager: gTempAllocator;
    import test.d_funcs: StringToString;

    auto arg = "foo".toSRef(gTempAllocator);
    auto oper = wrapModuleFunctionImpl!StringToString(gTempAllocator, &arg);
    gTempAllocator.empty.shouldEqual(Ternary.yes);
    oper.shouldEqualDlang("foobar");
}

///
@("No memory allocation bugs in wrapModuleFunctionImpl for Any[][] -> Any[][] -> Any[][] mallocator")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    import test.d_funcs: FirstOfTwoAnyArrays;

    with(allocatorContext(theGC)) {
        auto dArg = [[any(1.0), any("foo"), any(3.0)], [any(4.0), any(5.0), any(6.0)]];
        auto arg = toXlOper(dArg);
        auto oper = wrapModuleFunctionImpl!FirstOfTwoAnyArrays(theMallocator, &arg, &arg);
        oper.shouldEqualDlang(dArg);
    }
}

///
@("No memory allocation bugs in wrapModuleFunctionImpl for Any[][] -> Any[][] -> Any[][] TestAllocator")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    import test.d_funcs: FirstOfTwoAnyArrays;

    auto testAllocator = TestAllocator();

    with(allocatorContext(theGC)) {
        auto dArg = [
            [ any(1.0), any("foo"), any(3.0) ],
            [ any(4.0), any(5.0),   any(6.0) ],
        ];
        auto arg = toXlOper(dArg);
        auto oper = wrapModuleFunctionImpl!FirstOfTwoAnyArrays(testAllocator, &arg, &arg);
        oper.shouldEqualDlang(dArg);
    }
}

///
@("Correct number of coercions and frees in wrapModuleFunctionImpl")
@system unittest {
    import test.d_funcs: FuncAddEverything;
    import xlld.test.util: gNumXlAllocated, gNumXlFree;

    const oldNumAllocated = gNumXlAllocated;
    const oldNumFree = gNumXlFree;

    auto arg = toSRef([1.0, 2.0], theGC);
    auto oper = wrapModuleFunctionImpl!FuncAddEverything(theGC, &arg);

    (gNumXlAllocated - oldNumAllocated).shouldEqual(1);
    (gNumXlFree   - oldNumFree).shouldEqual(1);
}


///
@("Can't return empty 1D array to Excel")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    import test.d_funcs: EmptyStrings1D;

    with(allocatorContext(theGC)) {
        auto dArg = any(1.0);
        auto arg = toXlOper(dArg);
        auto oper = wrapModuleFunctionImpl!EmptyStrings1D(theGC, &arg);
        oper.shouldEqualDlang("#ERROR: empty result");
    }
}


///
@("Can't return empty 2D array to Excel")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    import test.d_funcs: EmptyStrings2D;

    with(allocatorContext(theGC)) {
        auto dArg = any(1.0);
        auto arg = toXlOper(dArg);
        auto oper = wrapModuleFunctionImpl!EmptyStrings2D(theGC, &arg);
        oper.shouldEqualDlang("#ERROR: empty result");
    }
}

///
@("Can't return half empty 2D array to Excel")
@system unittest {
    import xlld.memorymanager: allocatorContext;
    import test.d_funcs: EmptyStringsHalfEmpty2D;

    with(allocatorContext(theGC)) {
        auto dArg = any(1.0);
        auto arg = toXlOper(dArg);
        auto oper = wrapModuleFunctionImpl!EmptyStringsHalfEmpty2D(theGC, &arg);
        oper.shouldEqualDlang("#ERROR: empty result");
    }
}

///
@("issue 25 - make sure to reserve memory for all dArgs")
@system unittest {
    import std.typecons: Ternary;
    import xlld.memorymanager: allocatorContext, MemoryPool;
    import test.d_funcs: FirstOfTwoAnyArrays;

    auto pool = MemoryPool();

    with(allocatorContext(theGC)) {
        auto dArg = [[any(1.0), any("foo"), any(3.0)], [any(4.0), any(5.0), any(6.0)]];
        auto arg = toSRef(dArg);
        auto oper = wrapModuleFunctionImpl!FirstOfTwoAnyArrays(pool, &arg, &arg);
    }

    pool.empty.shouldEqual(Ternary.yes); // deallocateAll in wrapImpl
}


///
@("wrapModuleFunctionStr function that returns Any[][]")
@safe unittest {
    mixin(wrapModuleFunctionStr!("test.d_funcs", "DoubleArrayToAnyArray"));

    auto oper = [[1.0, 2.0], [3.0, 4.0]].toSRef(theGC);
    auto arg = () @trusted { return &oper; }();
    auto ret = DoubleArrayToAnyArray(arg);

    auto opers = () @trusted { return ret.val.array.lparray[0 .. 4]; }();
    opers[0].shouldEqualDlang(2.0);
    opers[1].shouldEqualDlang(6.0);
    opers[2].shouldEqualDlang("3quux");
    opers[3].shouldEqualDlang("4toto");
}

///
@("wrapModuleFunctionStr int -> int")
@safe unittest {
    mixin(wrapModuleFunctionStr!("test.d_funcs", "Twice"));

    auto oper = 3.toSRef(theGC);
    auto arg = () @trusted { return &oper; }();
    Twice(arg).shouldEqualDlang(6);
}

///
@("issue 31 - D functions can have const arguments")
@safe unittest {
    mixin(wrapModuleFunctionStr!("test.d_funcs", "FuncConstDouble"));

    auto oper = (3.0).toSRef(theGC);
    auto arg = () @trusted { return &oper; }();
    FuncConstDouble(arg).shouldEqualDlang(3.0);
}


@Flaky
@("wrapModuleFunctionStr async double -> double")
unittest {
    import xlld.conv.from: fromXlOper;
    import xlld.wrap.traits: Async;
    import xlld.test.util: asyncReturn, newAsyncHandle;
    import core.time: MonoTime;
    import core.thread;

    mixin(wrapModuleFunctionStr!("test.d_funcs", "AsyncDoubleToDouble"));

    auto oper = (3.0).toXlOper(theGC);
    auto arg = () @trusted { return &oper; }();
    auto asyncHandle = newAsyncHandle;
    () @trusted { AsyncDoubleToDouble(arg, &asyncHandle); }();

    const start = MonoTime.currTime;
    const expected = 6.0;
    while(asyncReturn(asyncHandle).fromXlOper!double(theGC) != expected &&
          MonoTime.currTime - start < 1.seconds)
    {
        Thread.sleep(10.msecs);
    }
    asyncReturn(asyncHandle).shouldEqualDlang(expected);
}

@("wrapModuleFunctionStr () -> NaN")
unittest {
    mixin(wrapModuleFunctionStr!("test.d_funcs", "NaN"));
    NaN().shouldEqualDlang("#NaN");
}
