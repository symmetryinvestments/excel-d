/**
   Analogous to std.conv but @nogc
 */

module xlld.conv;

version(unittest) import unit_threaded;
import std.traits: isScalarType, isPointer, isAssociativeArray, isAggregateType;
import std.range: isInputRange;

enum BUFFER_SIZE = 1024;

string text(size_t bufferSize = BUFFER_SIZE, A...)(auto ref A args) {
    import core.stdc.stdio: snprintf;

    static char[bufferSize] buffer;

    int index;
    foreach(ref const arg; args) {
        index += snprintf(&buffer[index], buffer.length - index, format(arg), value(arg));

        if(index >= buffer.length - 1) {
            return cast(string)buffer[];
        }
    }

    return cast(string)buffer[0 .. index];
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
