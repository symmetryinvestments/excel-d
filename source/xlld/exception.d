/**
   This module implements utility code to throw exceptions in @nogc code.
 */
module xlld.exception;

version(unittest) import unit_threaded;
import std.traits: isScalarType, isPointer, isAssociativeArray, isAggregateType;
import std.range: isInputRange;

enum BUFFER_SIZE = 1024;


T enforce(size_t bufferSize = BUFFER_SIZE, string file = __FILE__, size_t line = __LINE__, T, Args...)(T value, auto ref Args args)
@trusted if (is(typeof({ if (!value) {} }))) {

    import std.conv: emplace;

    static void[__traits(classInstanceSize, NoGcException)] buffer = void;

    if (!value) {
        auto exception = emplace!NoGcException(buffer);
        exception.adjust!(bufferSize, file, line)(args);
        throw exception;
    }
    return value;
}

///
@("enforce")
@safe unittest {
    string msg, file;
    size_t line, expectedLine;
    () @nogc {
        try {
            expectedLine = __LINE__ + 1;
            enforce(false, "foo", 5, "bar");
        } catch(NoGcException ex) {
            msg = ex.msg;
            file = ex.file;
            line = ex.line;
        }
    }();

    msg.shouldEqual("foo5bar");
    file.shouldEqual(__FILE__);
    line.shouldEqual(expectedLine);
}

class NoGcException: Exception {

    this() @safe @nogc nothrow pure {
        super("");
    }

    ///
    @("exception can be constructed in @nogc code")
    @safe @nogc pure unittest {
        static const exception = new NoGcException();
    }

    void adjust(size_t bufferSize = BUFFER_SIZE, string file = __FILE__, size_t line = __LINE__, A...)(auto ref A args) {
        import core.stdc.stdio: snprintf;

        static char[bufferSize] buffer;

        this.file = file;
        this.line = line;

        int index;
        foreach(ref const arg; args) {
            index += () @trusted {
                return snprintf(&buffer[index], buffer.length - index, format(arg), value(arg));
            }();

            if(index >= buffer.length - 1) {
                msg = () @trusted { return cast(string)buffer[]; }();
                return;
            }
        }

        msg = () @trusted { return cast(string)buffer[0 .. index]; }();
    }

    ///
    @("adjust with only strings")
    @safe unittest {
        auto exception = new NoGcException();
        () @nogc { exception.adjust("foo", "bar"); }();
        exception.msg.shouldEqual("foobar");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with string and integer")
    @safe unittest {
        auto exception = new NoGcException();
        () @nogc { exception.adjust(1, "bar"); }();
        exception.msg.shouldEqual("1bar");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with string and uint")
    @safe unittest {
        auto exception = new NoGcException();
        () @nogc { exception.adjust(1u, "bar"); }();
        exception.msg.shouldEqual("1bar");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }


    @("adjust with string and long")
    @safe unittest {
        auto exception = new NoGcException();
        () @nogc { exception.adjust("foo", 7L); }();
        exception.msg.shouldEqual("foo7");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with string and ulong")
    @safe unittest {
        auto exception = new NoGcException();
        () @nogc { exception.adjust("foo", 7UL); }();
        exception.msg.shouldEqual("foo7");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with char")
    @safe unittest {
        auto exception = new NoGcException();
        () @nogc { exception.adjust("fo", 'o'); }();
        exception.msg.shouldEqual("foo");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with bool")
    @safe unittest {
        auto exception = new NoGcException();
        () @nogc { exception.adjust("it is ", false); }();
        exception.msg.shouldEqual("it is false");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with float")
    @safe unittest {
        auto exception = new NoGcException();
        () @nogc { exception.adjust("it is ", 3.0f); }();
        exception.msg.shouldEqual("it is 3.000000");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with double")
    @safe unittest {
        auto exception = new NoGcException();
        () @nogc { exception.adjust("it is ", 3.0); }();
        exception.msg.shouldEqual("it is 3.000000");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with enums")
    @safe unittest {
        enum Enum {
            quux,
            toto,
        }
        auto exception = new NoGcException();
        () @nogc { exception.adjust(Enum.quux, "_middle_", Enum.toto); }();
        exception.msg.shouldEqual("quux_middle_toto");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with pointer")
    @safe unittest {
        import std.conv: to;
        import std.string: toLower;

        auto exception = new NoGcException();
        const ptr = new int(42);
        const expected = "0x" ~ ptr.to!string.toLower;

        () @nogc { exception.adjust(ptr); }();

        exception.msg.shouldEqual(expected);
        exception.line.shouldEqual(__LINE__ - 3);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with int[]")
    @safe unittest {
        auto exception = new NoGcException();
        const array = [1, 2, 3];

        () @nogc { exception.adjust(array); }();

        exception.msg.shouldEqual("[1, 2, 3]");
        exception.line.shouldEqual(__LINE__ - 3);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with int[string]")
    @safe unittest {
        auto exception = new NoGcException();
        const aa = ["foo": 1, "bar": 2];

        () @nogc { exception.adjust(aa); }();

        // I hope the hash function doesn't change...
        exception.msg.shouldEqual(`[bar: 2, foo: 1]`);
        exception.line.shouldEqual(__LINE__ - 4);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with struct")
    @safe unittest {
        auto exception = new NoGcException();
        struct Struct {
            int i;
            string s;
        }

        () @nogc { exception.adjust(Struct(42, "foobar")); }();

        exception.msg.shouldEqual(`Struct(42, foobar)`);
        exception.line.shouldEqual(__LINE__ - 3);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with class")
    @safe unittest {
        auto exception = new NoGcException();
        class Class {
            int i;
            string s;
            this(int i, string s) @safe pure nothrow { this.i = i; this.s = s; }
        }
        auto obj = new Class(42, "foobar");

        () @nogc { exception.adjust(obj); }();

        exception.msg.shouldEqual(`Class(42, foobar)`);
        exception.line.shouldEqual(__LINE__ - 3);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with class with toString")
    @safe unittest {
        auto exception = new NoGcException();
        class Class {
            int i;
            string s;
            this(int i, string s) @safe pure nothrow { this.i = i; this.s = s; }
            override string toString() @safe @nogc pure const nothrow {
                return "always the same";
            }
        }
        auto obj = new Class(42, "foobar");

        () @nogc { exception.adjust(obj); }();

        exception.msg.shouldEqual(`always the same`);
        exception.line.shouldEqual(__LINE__ - 3);
        exception.file.shouldEqual(__FILE__);
    }

}


private const(char)* format(T)(ref const(T) arg) if(is(T == string)) {
    return &"%s"[0];
}

private const(char)* format(T)(ref const(T) arg) if(is(T == int) || is(T == short) || is(T == byte)) {
    return &"%d"[0];
}

private const(char)* format(T)(ref const(T) arg) if(is(T == uint) || is(T == ushort) || is(T == ubyte)) {
    return &"%u"[0];
}

private const(char)* format(T)(ref const(T) arg) if(is(T == long)) {
    return &"%ld"[0];
}

private const(char)* format(T)(ref const(T) arg) if(is(T == ulong)) {
    return &"%lu"[0];
}

private const(char)* format(T)(ref const(T) arg) if(is(T == char)) {
    return &"%c"[0];
}

private const(char)* format(T)(ref const(T) arg) if(is(T == float)) {
    return &"%f"[0];
}

private const(char)* format(T)(ref const(T) arg) if(is(T == double)) {
    return &"%lf"[0];
}

private const(char)* format(T)(ref const(T) arg)
    if(is(T == enum) || is(T == bool) || (isInputRange!T && !is(T == string)) || isAssociativeArray!T || isAggregateType!T) {
    return &"%s"[0];
}

private const(char)* format(T)(ref const(T) arg) if(isPointer!T) {
    return &"%p"[0];
}



private auto value(T)(ref const(T) arg) if((isScalarType!T || isPointer!T) && !is(T == enum) && !is(T == bool)) {
    return arg;
}

private auto value(T)(ref const(T) arg) if(is(T == enum)) {
    import std.traits: EnumMembers;
    import std.conv: to;

    string enumToString(in T arg) {
        return arg.to!string;
    }

    final switch(arg) {
        foreach(member; EnumMembers!T) {
        case member:
            mixin(`return &"` ~ member.to!string ~ `"[0];`);
        }
    }
}


private auto value(T)(ref const(T) arg) if(is(T == bool)) {
    return arg
        ? &"true"[0]
        : &"false"[0];
}


private auto value(T)(ref const(T) arg) if(is(T == string)) {
    static char[BUFFER_SIZE] buffer;
    if(arg.length > buffer.length - 1) return null;
    buffer[0 .. arg.length] = arg[];
    buffer[arg.length] = 0;
    return &buffer[0];
}

private auto value(T)(ref const(T) arg) if(isInputRange!T && !is(T == string)) {
    import core.stdc.string: strlen;
    import core.stdc.stdio: snprintf;

    static char[BUFFER_SIZE] buffer;

    if(arg.length > buffer.length - 1) return null;

    int index;
    buffer[index++] = '[';
    foreach(i, ref const elt; arg) {
        index += snprintf(&buffer[index], buffer.length - index, format(elt), value(elt));
        if(i != arg.length - 1) index += snprintf(&buffer[index], buffer.length - index, ", ");
    }

    buffer[index++] = ']';
    buffer[index++] = 0;

    return &buffer[0];
}

private auto value(T)(ref const(T) arg) if(isAssociativeArray!T) {
    import core.stdc.string: strlen;
    import core.stdc.stdio: snprintf;

    static char[BUFFER_SIZE] buffer;

    if(arg.length > buffer.length - 1) return null;

    int index;
    buffer[index++] = '[';
    int i;
    foreach(ref const elt; arg.byKeyValue) {
        index += snprintf(&buffer[index], buffer.length - index, format(elt.key), value(elt.key));
        index += snprintf(&buffer[index], buffer.length - index, ": ");
        index += snprintf(&buffer[index], buffer.length - index, format(elt.value), value(elt.value));
        if(i++ != arg.length - 1) index += snprintf(&buffer[index], buffer.length - index, ", ");
    }

    buffer[index++] = ']';
    buffer[index++] = 0;

    return &buffer[0];
}

private auto value(T)(ref const(T) arg) @nogc if(isAggregateType!T) {
    import core.stdc.string: strlen;
    import core.stdc.stdio: snprintf;
    import std.traits: hasMember;

    static char[BUFFER_SIZE] buffer;

    static if(__traits(compiles, callToString(arg))) {
        const repr = arg.toString;
        if(repr.length > buffer.length - 1) return null;
        buffer[0 .. repr.length] = repr[];
        buffer[repr.length] = 0;
        return &buffer[0];
    } else {

        int index;
        index += snprintf(&buffer[index], buffer.length - index, T.stringof);
        buffer[index++] = '(';
        foreach(i, ref const elt; arg.tupleof) {
            index += snprintf(&buffer[index], buffer.length - index, format(elt), value(elt));
            if(i != arg.tupleof.length - 1) index += snprintf(&buffer[index], buffer.length - index, ", ");
        }

        buffer[index++] = ')';
        buffer[index++] = 0;

        return &buffer[0];
    }
}

// helper function to avoid a closure
private string callToString(T)(ref const(T) arg) @nogc {
    return arg.toString;
}
