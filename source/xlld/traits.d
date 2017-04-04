/**
 This module implements the compile-time reflection machinery to
 automatically register all D functions that are eligible in a
 compile-time define list of modules to be called from Excel.

 Import this module from any module from your XLL build and:

 -----------
 import xlld;

 mixin(implGetWorksheetFunctionsString!("module1", "module2", "module3"));
 -----------

 All eligible functions in the 3 example modules above will automagically
 be accessible from Excel (assuming the built XLL is loaded as an add-in).
 */
module xlld.traits;

import xlld.worksheet;
import xlld.xlcall;
import std.traits: isSomeFunction, allSatisfy, isSomeString;

// import unit_threaded and introduce helper functions for testing
version(unittest) {
    import unit_threaded;

    // return a WorksheetFunction for a double function(double) with no
    // optional arguments
    WorksheetFunction makeWorksheetFunction(wstring name, wstring typeText) @safe pure nothrow {
        return
            WorksheetFunction(
                Procedure(name),
                TypeText(typeText),
                FunctionText(name),
                Optional(
                    ArgumentText(""w),
                    MacroType("1"w),
                    Category(""w),
                    ShortcutText(""w),
                    HelpTopic(""w),
                    FunctionHelp(""w),
                    ArgumentHelp([]),
                )
            );
    }

    WorksheetFunction doubleToDoubleFunction(wstring name) @safe pure nothrow {
        return makeWorksheetFunction(name, "BB"w);
    }

    WorksheetFunction FP12ToDoubleFunction(wstring name) @safe pure nothrow {
        return makeWorksheetFunction(name, "BK%"w);
    }

    WorksheetFunction operToOperFunction(wstring name) @safe pure nothrow {
        return makeWorksheetFunction(name, "UU"w);
    }
}

/**
 Take a D function as a compile-time parameter and returns a
 WorksheetFunction struct with the fields filled in accordingly.
 */
WorksheetFunction getWorksheetFunction(alias F)() if(isSomeFunction!F) {
    import std.traits: ReturnType, Parameters, getUDAs;
    import std.conv: text;

    alias R = ReturnType!F;
    alias T = Parameters!F;

    static if(!isWorksheetFunction!F) {
        throw new Exception("Unsupported function type " ~ R.stringof ~ T.stringof ~ " for " ~
                            __traits(identifier, F).stringof[1 .. $-1]);
    } else {

        WorksheetFunction ret;
        ret.procedure = Procedure(__traits(identifier, F));
        ret.functionText = FunctionText(__traits(identifier, F));
        ret.typeText = TypeText(getTypeText!F);

        // check to see if decorated with @Register
        alias registerAttrs = getUDAs!(F, Register);
        static if(registerAttrs.length > 0) {
            static assert(registerAttrs.length == 1,
                          text("Only 1 @Register allowed, found ", registerAttrs.length,
                               " on function ", __traits(identifier, F)));
            ret.optional = registerAttrs[0];
        }

        return ret;
    }
}

@("getWorksheetFunction for double -> double functions with no extra attributes")
@safe pure unittest {
    double foo(double) nothrow @nogc { return 0; }
    getWorksheetFunction!foo.shouldEqual(doubleToDoubleFunction("foo"));

    double bar(double) nothrow @nogc { return 0; }
    getWorksheetFunction!bar.shouldEqual(doubleToDoubleFunction("bar"));
}

@("getWorksheetFunction for double -> int functions should fail")
@safe pure unittest {
    double foo(int) { return 0; }
    getWorksheetFunction!foo.shouldThrowWithMessage("Unsupported function type double(int) for foo");
}

@("getworksheetFunction with @Register in order")
@safe pure unittest {

    @Register(ArgumentText("my arg txt"), MacroType("macro"))
    double foo(double) nothrow;

    auto expected = doubleToDoubleFunction("foo");
    expected.argumentText = ArgumentText("my arg txt");
    expected.macroType = MacroType("macro");

    getWorksheetFunction!foo.shouldEqual(expected);
}

@("getworksheetFunction with @Register out of order")
@safe pure unittest {

    @Register(HelpTopic("I need somebody"), ArgumentText("my arg txt"))
    double foo(double) nothrow;

    auto expected = doubleToDoubleFunction("foo");
    expected.argumentText = ArgumentText("my arg txt");
    expected.helpTopic = HelpTopic("I need somebody");

    getWorksheetFunction!foo.shouldEqual(expected);
}


private wstring getTypeText(alias F)() if(isSomeFunction!F) {
    import std.traits: ReturnType, Parameters;

    wstring typeToString(T)() {
        static if(is(T == double))
            return "B";
        else static if(is(T == FP12*))
            return "K%";
        else static if(is(T == LPXLOPER12))
            return "U";
        else
            static assert(false, "Unsupported type " ~ T.stringof);
    }

    auto retType = typeToString!(ReturnType!F);
    foreach(argType; Parameters!F)
        retType ~= typeToString!(argType);

    return retType;
}


@("getTypeText")
@safe pure unittest {
    import std.conv: to; // working around unit-threaded bug

    double foo(double);
    getTypeText!foo.to!string.shouldEqual("BB");

    double bar(FP12*);
    getTypeText!bar.to!string.shouldEqual("BK%");

    FP12* baz(FP12*);
    getTypeText!baz.to!string.shouldEqual("K%K%");

    FP12* qux(double);
    getTypeText!qux.to!string.shouldEqual("K%B");

    LPXLOPER12 fun(LPXLOPER12);
    getTypeText!fun.to!string.shouldEqual("UU");
}



// helper template for aliasing
private alias Identity(alias T) = T;


// whether or not this is a function that has the "right" types
template isSupportedFunction(alias F, T...) {
    import std.traits: isSomeFunction, ReturnType, Parameters, functionAttributes, FunctionAttribute;
    import std.meta: AliasSeq, allSatisfy;
    import std.typecons: Tuple;

    // trying to get a pointer to something is a good way of making sure we can
    // attempt to evaluate `isSomeFunction` - it's not always possible
    enum canGetPointerToIt = __traits(compiles, &F);
    enum isOneOfSupported(U) = isSupportedType!(U, T);

    static if(canGetPointerToIt) {
        static if(isSomeFunction!F) {

            enum isSupportedFunction =
                __traits(compiles, F(Tuple!(Parameters!F)().expand)) &&
                isOneOfSupported!(ReturnType!F) &&
                allSatisfy!(isOneOfSupported, Parameters!F) &&
                functionAttributes!F & FunctionAttribute.nothrow_;

            static if(!isSupportedFunction && !(functionAttributes!F & FunctionAttribute.nothrow_))
                pragma(msg, "Warning: Function '", __traits(identifier, F), "' not considered because it throws");

        } else
            enum isSupportedFunction = false;
    } else
        enum isSupportedFunction = false;
}


// if T is one of U
private template isSupportedType(T, U...) {
    static if(U.length == 0)
        enum isSupportedType = false;
    else
        enum isSupportedType = is(T == U[0]) || isSupportedType!(T, U[1..$]);
}

@safe pure unittest {
    static assert(isSupportedType!(int, int, int));
    static assert(!isSupportedType!(int, double, string));
}

// whether or not this is a function that can be called from Excel
private enum isWorksheetFunction(alias F) = isSupportedFunction!(F, double, FP12*, LPXLOPER12);

@safe pure unittest {
    double doubleToDouble(double) nothrow;
    static assert(isWorksheetFunction!doubleToDouble);

    LPXLOPER12 operToOper(LPXLOPER12) nothrow;
    static assert(isWorksheetFunction!operToOper);
}


/**
 Gets all Excel-callable functions in a given module
 */
WorksheetFunction[] getModuleWorksheetFunctions(string moduleName)() {
    mixin(`import ` ~ moduleName ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    WorksheetFunction[] ret;

    foreach(moduleMemberStr; __traits(allMembers, module_)) {

        alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));

        static if(isWorksheetFunction!moduleMember) {
            try
                ret ~= getWorksheetFunction!moduleMember;
            catch(Exception ex)
                assert(0); //can't happen
        }
    }

    return ret;
}

@("getWorksheetFunctions on test_xl_funcs")
@safe pure unittest {
    getModuleWorksheetFunctions!"xlld.test_xl_funcs".shouldEqual(
        [
            doubleToDoubleFunction("FuncMulByTwo"),
            FP12ToDoubleFunction("FuncFP12"),
            operToOperFunction("FuncFib"),
        ]
    );
}

/**
 Gets all Excel-callable functions from the given modules
 */
WorksheetFunction[] getAllWorksheetFunctions(Modules...)() pure @safe if(allSatisfy!(isSomeString, typeof(Modules))) {
    WorksheetFunction[] ret;

    foreach(module_; Modules) {
        ret ~= getModuleWorksheetFunctions!module_;
    }

    return ret;
}

/**
 Implements the getWorksheetFunctions function needed by xlld.xll in
 order to register the Excel-callable functions at runtime
 This used to be a template mixin but even using a string mixin inside
 fails to actually make it an extern(C) function.
 */
string implGetWorksheetFunctionsString(Modules...)() if(allSatisfy!(isSomeString, typeof(Modules))) {
    import std.array: join;

    string modulesString() {

        string[] modules;
        foreach(module_; Modules) {
            modules ~= `"` ~ module_ ~ `"`;
        }
        return modules.join(", ");
    }

    return
        [
            `extern(C) WorksheetFunction[] getWorksheetFunctions() @safe pure nothrow {`,
            `    return getAllWorksheetFunctions!(` ~ modulesString ~ `);`,
            `}`,
        ].join("\n");
}

@("template mixin for getWorkSheetFunctions for test_xl_funcs")
unittest {
    import xlld.traits;
    import xlld.worksheet;

    // mixin the function here then call it to see if it does what it's supposed to
    mixin(implGetWorksheetFunctionsString!"xlld.test_xl_funcs");
    getWorksheetFunctions.shouldEqual(
        [
            doubleToDoubleFunction("FuncMulByTwo"),
            FP12ToDoubleFunction("FuncFP12"),
            operToOperFunction("FuncFib"),
        ]
    );
}

struct DllDefFile {
    Statement[] statements;
}

struct Statement {
    string name;
    string[] args;

    this(string name, string[] args) @safe pure nothrow {
        this.name = name;
        this.args = args;
    }

    this(string name, string arg) @safe pure nothrow {
        this(name, [arg]);
    }

    string toString() @safe pure const {
        import std.array: join;
        import std.algorithm: map;

        if(name == "EXPORTS")
        {
            version(X86)
            {
                return name ~ "\n" ~ args.map!(a => "\t\t" ~ a).join("\n");
            }
            else version(X86_64)
            {
                import std.array:Appender;
                import std.conv:to;
                Appender!string ret;
                ret.put(name ~ "\n");
                foreach(i,arg;args)
                {
                    ret.put("\t\t" ~ arg ~ "\t@" ~ i.to!string~"\n");
                }
                return ret.data;
            }
            else static assert("unsupported version");
        }
        else
        {
            return name ~ "\t\t" ~ args.map!(a => stringify(name, a)).join(" ");
        }
    }

    static private string stringify(in string name, in string arg) @safe pure {
        if(name == "LIBRARY") return `"` ~ arg ~ `"`;
        if(name == "DESCRIPTION") return `'` ~ arg ~ `'`;
        return arg;
    }
}

/**
   Returns a structure descripting a Windows .def file.
   This allows the tests to not care about the specific formatting
   used when writing the information out.
   This encapsulates all the functions to be exported by the DLL/XLL.
 */
DllDefFile dllDefFile(Modules...)(string libName, string description)
if(allSatisfy!(isSomeString, typeof(Modules)))
{
    import std.conv: to;

    version(X86)
    {
        auto statements = [
            Statement("LIBRARY", libName),
            Statement("DESCRIPTION", description),
            Statement("EXETYPE", "NT"),
            Statement("CODE", "PRELOAD DISCARDABLE"),
            Statement("DATA", "PRELOAD MULTIPLE"),
        ];
    }
    else version(X86_64)
    {
        auto statements = [
            Statement("LIBRARY", libName),
        ];        
    }
    else static assert("unsupported target");

    string[] exports = ["xlAutoOpen", "xlAutoClose", "xlAutoFree12"];
    foreach(func; getAllWorksheetFunctions!Modules) {
        exports ~= func.procedure.to!string;
    }

    return DllDefFile(statements ~ Statement("EXPORTS", exports));
}

@("worksheet functions to .def file")
unittest {
    version(X86)
    {
        dllDefFile!"xlld.test_xl_funcs"("myxll32.dll", "Simple D add-in").shouldEqual(
            DllDefFile(
                [
                    Statement("LIBRARY", "myxll32.dll"),
                    Statement("DESCRIPTION", "Simple D add-in"),
                    Statement("EXETYPE", "NT"),
                    Statement("CODE", "PRELOAD DISCARDABLE"),
                    Statement("DATA", "PRELOAD MULTIPLE"),
                    Statement("EXPORTS", ["xlAutoOpen", "xlAutoClose", "xlAutoFree12", "FuncMulByTwo", "FuncFP12", "FuncFib"]),
                ]
            )
        );
    }
    else version(X86_64)
    {
        dllDefFile!"xlld.test_xl_funcs"("myxll64.dll", "Simple D add-in").shouldEqual(
            DllDefFile(
                [
                    Statement("LIBRARY", "myxll64.dll"),
                    Statement("EXPORTS", ["xlAutoOpen", "xlAutoClose", "xlAutoFree12", "FuncMulByTwo", "FuncFP12", "FuncFib"]),
                ]
            )
        );
    }
    else static assert("unsupported version");
}


mixin template GenerateDllDef(string module_ = __MODULE__) {
    version(exceldDef) {
        void main(string[] args) nothrow {
            try {
                import std.stdio: File;
                import std.exception: enforce;
                import std.path: stripExtension;

                enforce(args.length >= 2 && args.length <= 4,
                        "Usage: " ~ args[0] ~ " [file_name] <lib_name> <description>");

                immutable fileName = args[1];
                immutable libName = args.length > 2 ? args[2] : fileName.stripExtension ~ ".xll";
                immutable description = args.length > 3 ? args[3] : "Simple D add-in to Excel";

                auto file = File(fileName, "w");
                foreach(stmt; dllDefFile!module_(libName, description).statements)
                    file.writeln(stmt.toString);
            } catch(Exception ex) {
                import std.stdio: stderr;
                try
                    stderr.writeln("Error: ", ex.msg);
                catch(Exception ex2)
                    assert(0, "Program could not write exception message");
            }
        }
    }
}
