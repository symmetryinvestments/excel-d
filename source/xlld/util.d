module xlld.util;


version(Windows) {
    extern(Windows) void OutputDebugStringW(const wchar* fmt) nothrow;
}


/**
   Polymorphic logging function.
   Prints to the console when unit testing and on Linux,
   otherwise uses the system logger on Windows.
 */
void log(A...)(auto ref A args) @trusted {
    try {
        version(unittest) {
            version(Have_unit_threaded) {
                import unit_threaded: writelnUt;
                writelnUt(args);
            } else {
                import std.stdio: writeln;
                writeln(args);
            }
        } else version(Windows) {
            import nogc.conv: text, toWStringz;
            OutputDebugStringW(text(args).toWStringz);
        } else {
            import std.experimental.logger: trace;
            trace(args);
        }
    } catch(Exception e) {
        import core.stdc.stdio: printf;
        printf("Error - could not log\n");
    }
}
