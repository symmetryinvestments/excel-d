/**
   This module implements utility code to throw exceptions in @nogc code.
 */
module xlld.exception;

version(unittest) import unit_threaded;
import std.traits: isScalarType;

enum BUFFER_SIZE = 1024;

T enforce(string file = __FILE__, size_t line = __LINE__, T, Args...)(T value, Args args)
@trusted if (is(typeof({ if (!value) {} }))) {

    import std.conv: emplace;

    static void[__traits(classInstanceSize, NoGcException)] buffer = void;

    if (!value) {
        auto exception = emplace!NoGcException(buffer);
        exception.adjust!(file, line)(args);
        throw exception;
    }
    return value;
}

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

    void adjust(string file = __FILE__, size_t line = __LINE__, A...)(auto ref A args) {
        import core.stdc.stdio: snprintf;

        static char[BUFFER_SIZE] buffer;

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

private const(char)* format(T)(ref const(T) arg) if(is(T == enum)) {
    return &"%s"[0];
}


private auto value(T)(ref const(T) arg) if(isScalarType!T && !is(T == enum)) {
    return arg;
}

private auto value(T)(ref const(T) arg) if(is(T == enum)) {
    import std.traits: EnumMembers;
    final switch(arg) {
        foreach(member; EnumMembers!T) {
        case member:
            mixin(`return &"` ~ enumToString(member) ~ `"[0];`);
        }
    }
}

private string enumToString(T)(in T arg) if(is(T == enum)) {
    if(!__ctfe) return "";
    import std.conv: to;
    return arg.to!string;
}

private auto value(T)(ref const(T) arg) if(is(T == string)) {
    static char[BUFFER_SIZE] buffer;
    if(arg.length > buffer.length - 1) return null;
    buffer[0 .. arg.length] = arg[];
    buffer[arg.length] = 0;
    return &buffer[0];
}
