module ut.conv.misc;

import test;
import xlld.conv.misc;
import xlld.conv.to: toXlOper;

@("isUserStruct")
@safe pure unittest {
    import std.datetime: DateTime;
    import std.typecons: Tuple;
    import xlld.any: Any;

    static struct Foo {}

    static assert( isUserStruct!Foo);
    static assert(!isUserStruct!Any);
    static assert(!isUserStruct!DateTime);
    static assert(!isUserStruct!(Tuple!(int, int)));
    static assert(!isUserStruct!(Tuple!(int, string)));
    static assert(!isUserStruct!(Tuple!(int, Foo, double)));
}

///
@("dup")
@safe unittest {
    auto int_ = 42.toXlOper(theGC);
    int_.dup(theGC).shouldEqualDlang(42);

    auto double_ = (33.3).toXlOper(theGC);
    double_.dup(theGC).shouldEqualDlang(33.3);

    auto string_ = "foobar".toXlOper(theGC);
    string_.dup(theGC).shouldEqualDlang("foobar");

    auto array = () @trusted {
        return [
            ["foo", "bar", "baz"],
            ["quux", "toto", "brzz"]
        ]
        .toXlOper(theGC);
    }();

    array.dup(theGC).shouldEqualDlang(
        [
            ["foo", "bar", "baz"],
            ["quux", "toto", "brzz"],
        ]
    );
}

@("dup string allocator fails")
@safe unittest {
    auto allocator = FailingAllocator();
    "foo".toXlOper(theGC).dup(allocator).shouldThrowWithMessage("Failed to allocate memory in dup");
}

@("dup multi allocator fails")
@safe unittest {
    auto allocator = FailingAllocator();
    auto oper = () @trusted { return [33.3].toXlOper(theGC); }();
    oper.dup(allocator).shouldThrowWithMessage("Failed to allocate memory in dup");
}

///
@("operStringLength")
unittest {
    import std.experimental.allocator.mallocator: Mallocator;
    auto oper = "foobar".toXlOper(theGC);
    const length = () @nogc { return operStringLength(oper); }();
    length.shouldEqual(6);
}


///
@("multi")
@safe unittest {
    auto allocator = FailingAllocator();
    multi(2, 3, allocator).shouldThrowWithMessage("Failed to allocate memory for multi oper");
}
