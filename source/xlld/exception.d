/**
   This module implements utility code to throw exceptions in @nogc code.
 */
module xlld.exception;

version(unittest) import unit_threaded;

enum BUFFER_SIZE = 1024;


T enforce(size_t bufferSize = BUFFER_SIZE, string file = __FILE__, size_t line = __LINE__, T, Args...)(T value, auto ref Args args)
@trusted if (is(typeof({ if (!value) {} }))) {

    import std.conv: emplace;

    static void[__traits(classInstanceSize, NoGcException)] buffer = void;

    if (!value) {
        auto exception = emplace!NoGcException(buffer);
        () @trusted { exception.adjust!(bufferSize, file, line)(args); }();
        throw cast(const)exception;
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
        import xlld.conv: text;

        this.file = file;
        this.line = line;

        this.msg = text!bufferSize(args);
    }

    ///
    @("adjust with only strings")
    @system unittest {
        auto exception = new NoGcException();
        () @nogc nothrow { exception.adjust("foo", "bar"); }();
        exception.msg.shouldEqual("foobar");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with string and integer")
    @system unittest {
        auto exception = new NoGcException();
        () @nogc nothrow { exception.adjust(1, "bar"); }();
        exception.msg.shouldEqual("1bar");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with string and uint")
    @system unittest {
        auto exception = new NoGcException();
        () @nogc nothrow { exception.adjust(1u, "bar"); }();
        exception.msg.shouldEqual("1bar");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }


    @("adjust with string and long")
    @system unittest {
        auto exception = new NoGcException();
        () @nogc nothrow { exception.adjust("foo", 7L); }();
        exception.msg.shouldEqual("foo7");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with string and ulong")
    @system unittest {
        auto exception = new NoGcException();
        () @nogc nothrow { exception.adjust("foo", 7UL); }();
        exception.msg.shouldEqual("foo7");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with char")
    @system unittest {
        auto exception = new NoGcException();
        () @nogc nothrow { exception.adjust("fo", 'o'); }();
        exception.msg.shouldEqual("foo");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with bool")
    @system unittest {
        auto exception = new NoGcException();
        () @nogc nothrow { exception.adjust("it is ", false); }();
        exception.msg.shouldEqual("it is false");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with float")
    @system unittest {
        auto exception = new NoGcException();
        () @nogc nothrow { exception.adjust("it is ", 3.0f); }();
        exception.msg.shouldEqual("it is 3.000000");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with double")
    @system unittest {
        auto exception = new NoGcException();
        () @nogc nothrow { exception.adjust("it is ", 3.0); }();
        exception.msg.shouldEqual("it is 3.000000");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with enums")
    @system unittest {
        enum Enum {
            quux,
            toto,
        }
        auto exception = new NoGcException();
        () @nogc nothrow { exception.adjust(Enum.quux, "_middle_", Enum.toto); }();
        exception.msg.shouldEqual("quux_middle_toto");
        exception.line.shouldEqual(__LINE__ - 2);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with pointer")
    @system unittest {
        import std.conv: to;
        import std.string: toLower;

        auto exception = new NoGcException();
        const ptr = new int(42);
        const expected = "0x" ~ ptr.to!string.toLower;

        () @nogc nothrow { exception.adjust(ptr); }();

        exception.msg.shouldEqual(expected);
        exception.line.shouldEqual(__LINE__ - 3);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with int[]")
    @system unittest {
        auto exception = new NoGcException();
        const array = [1, 2, 3];

        () @nogc nothrow { exception.adjust(array); }();

        exception.msg.shouldEqual("[1, 2, 3]");
        exception.line.shouldEqual(__LINE__ - 3);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with int[string]")
    @system unittest {
        auto exception = new NoGcException();
        const aa = ["foo": 1, "bar": 2];

        () @nogc nothrow { exception.adjust(aa); }();

        // I hope the hash function doesn't change...
        exception.msg.shouldEqual(`[bar: 2, foo: 1]`);
        exception.line.shouldEqual(__LINE__ - 4);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with struct")
    @system unittest {
        auto exception = new NoGcException();
        struct Struct {
            int i;
            string s;
        }

        () @nogc nothrow { exception.adjust(Struct(42, "foobar")); }();

        exception.msg.shouldEqual(`Struct(42, foobar)`);
        exception.line.shouldEqual(__LINE__ - 3);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with class")
    @system unittest {
        auto exception = new NoGcException();
        class Class {
            int i;
            string s;
            this(int i, string s) @safe pure nothrow { this.i = i; this.s = s; }
        }
        auto obj = new Class(42, "foobar");

        () @nogc nothrow { exception.adjust(obj); }();

        exception.msg.shouldEqual(`Class(42, foobar)`);
        exception.line.shouldEqual(__LINE__ - 3);
        exception.file.shouldEqual(__FILE__);
    }

    @("adjust with class with toString")
    @system unittest {
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

        () @nogc nothrow { exception.adjust(obj); }();

        exception.msg.shouldEqual(`always the same`);
        exception.line.shouldEqual(__LINE__ - 3);
        exception.file.shouldEqual(__FILE__);
    }
}
