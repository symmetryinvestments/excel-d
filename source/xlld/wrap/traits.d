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
module xlld.wrap.traits;

import xlld.wrap.worksheet;
import xlld.sdk.xlcall;
import std.traits: isSomeFunction, isSomeString;
import std.meta: allSatisfy;
import std.typecons: Flag, No;

/// import unit_threaded and introduce helper functions for testing
version(testingExcelD) {
    import unit_threaded;
}

/**
 Take a D function as a compile-time parameter and returns a
 WorksheetFunction struct with the fields filled in accordingly.
 */
WorksheetFunction getWorksheetFunction(alias F)() if(isSomeFunction!F) {
    import xlld.wrap.wrap: pascalCase;
    import std.traits: ReturnType, Parameters, getUDAs;
    import std.conv: text, to;

    alias R = ReturnType!F;
    alias T = Parameters!F;

    static if(!isWorksheetFunction!F) {
        throw new Exception("Unsupported function type " ~ R.stringof ~ T.stringof ~ " for " ~
                            __traits(identifier, F).stringof[1 .. $-1]);
    } else {

        WorksheetFunction ret;
        auto name = __traits(identifier, F).pascalCase.to!wstring;
        ret.procedure = Procedure(name);
        ret.functionText = FunctionText(name);
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


wstring getTypeText(alias F)() if(isSomeFunction!F) {
    import std.traits: ReturnType, Parameters, Unqual, hasUDA;

    wstring typeToString(T)() {

        alias Type = Unqual!T;

        static if(is(Type == double))
            return "B";
        else static if(is(Type == FP12*))
            return "K%";
        else static if(is(Type == LPXLOPER12))
            return "U";
        else static if(is(Type == void))
            return ">";
        else
            static assert(false, "Unsupported type " ~ T.stringof);
    }

    auto retType = typeToString!(ReturnType!F);
    foreach(i, argType; Parameters!F) {
        static if(i == Parameters!F.length - 1 && hasUDA!(F, Async))
            retType ~= "X";
        else
            retType ~= typeToString!(argType);

    }

    return retType;
}



// helper template for aliasing
alias Identity(alias T) = T;


/**
   Is true if F is a callable function and functionTypePredicate is true
   for the return type and all parameter types of F.
 */
template isSupportedFunction(alias F, alias functionTypePredicate) {
    import std.traits: ReturnType, Parameters;
    import std.meta: allSatisfy;

    static if(isCallableFunction!F) {
        enum returnTypeOk = functionTypePredicate!(ReturnType!F) || is(ReturnType!F == void);
        enum paramTypesOk = allSatisfy!(functionTypePredicate, Parameters!F);
        enum isSupportedFunction = returnTypeOk && paramTypesOk;
    } else
        enum isSupportedFunction = false;
}


template isCallableFunction(alias F) {
    import std.traits: isSomeFunction, Parameters;
    import std.typecons: Tuple;

    /// trying to get a pointer to something is a good way of making sure we can
    /// attempt to evaluate `isSomeFunction` - it's not always possible
    enum canGetPointerToIt = __traits(compiles, &F);

    static if(canGetPointerToIt) {
        static if(isSomeFunction!F)
            enum isCallableFunction = __traits(compiles, F(Tuple!(Parameters!F)().expand));
         else
             enum isCallableFunction = false;
    } else
        enum isCallableFunction = false;
}


// if T is one of A
template isOneOf(T, A...) {
    static if(A.length == 0)
        enum isOneOf = false;
    else
        enum isOneOf = is(T == A[0]) || isOneOf!(T, A[1..$]);
}

@safe pure unittest {
    static assert(isOneOf!(int, int, int));
    static assert(!isOneOf!(int, double, string));
}

// whether or not this is a function that can be called from Excel
template isWorksheetFunction(alias F) {
    static if(isWorksheetFunctionModuloLinkage!F) {
        import std.traits: functionLinkage;
        enum isWorksheetFunction = functionLinkage!F == "Windows";
    } else
        enum isWorksheetFunction = false;
}

/// if the types match for a worksheet function but without checking the linkage
template isWorksheetFunctionModuloLinkage(alias F) {
    import std.traits: ReturnType, Parameters, isCallable;
    import std.meta: anySatisfy;

    static if(!isCallable!F)
        enum isWorksheetFunctionModuloLinkage = false;
    else {

        enum isEnum(T) = is(T == enum);
        enum isOneOfSupported(U) = isOneOf!(U, double, FP12*, LPXLOPER12);

        enum isWorksheetFunctionModuloLinkage =
            isSupportedFunction!(F, isOneOfSupported) &&
            !is(ReturnType!F == enum) &&
            !anySatisfy!(isEnum, Parameters!F);
    }
}


/**
 Gets all Excel-callable functions in a given module
 */
WorksheetFunction[] getModuleWorksheetFunctions(string moduleName)
                                               (Flag!"onlyExports" onlyExports = No.onlyExports)
{
    mixin(`import ` ~ moduleName ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    WorksheetFunction[] ret;

    foreach(moduleMemberStr; __traits(allMembers, module_)) {

        alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));

        static if(isWorksheetFunction!moduleMember) {
            try {
                const shouldWrap = onlyExports ? __traits(getProtection, moduleMember) == "export" : true;
                if(shouldWrap)
                    ret ~= getWorksheetFunction!(moduleMember);
            } catch(Exception ex)
                assert(0); //can't happen
        } else static if(isWorksheetFunctionModuloLinkage!moduleMember) {
            import std.traits: functionLinkage;
            pragma(msg, "!!!!! excel-d warning: function " ~ __traits(identifier, moduleMember) ~
                   " has the right types to be callable from Excel but isn't due to having " ~
                   functionLinkage!moduleMember ~ " linkage instead of the required 'Windows'");
        }
    }

    return ret;
}

/**
 Gets all Excel-callable functions from the given modules
 */
WorksheetFunction[] getAllWorksheetFunctions(Modules...)
                                            (Flag!"onlyExports" onlyExports = No.onlyExports)
    pure @safe if(allSatisfy!(isSomeString, typeof(Modules)))
{
    WorksheetFunction[] ret;

    foreach(module_; Modules) {
        ret ~= getModuleWorksheetFunctions!module_(onlyExports);
    }

    return ret;
}

/**
 Implements the getWorksheetFunctions function needed by xlld.sdk.xll in
 order to register the Excel-callable functions at runtime
 This used to be a template mixin but even using a string mixin inside
 fails to actually make it an extern(C) function.
 */
string implGetWorksheetFunctionsString(Modules...)() if(allSatisfy!(isSomeString, typeof(Modules))) {
    return implGetWorksheetFunctionsString(Modules);
}


string implGetWorksheetFunctionsString(string[] modules...) {
    import std.array: join;

    if(!__ctfe) {
        return "";
    }

    string modulesString() {

        string[] ret;
        foreach(module_; modules) {
            ret ~= `"` ~ module_ ~ `"`;
        }
        return ret.join(", ");
    }

    return
        [
            `extern(C) WorksheetFunction[] getWorksheetFunctions() @safe pure nothrow {`,
            `    return getAllWorksheetFunctions!(` ~ modulesString ~ `);`,
            `}`,
        ].join("\n");
}


///
struct DllDefFile {
    Statement[] statements;
}

///
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
            return name ~ "\n" ~ args.map!(a => "\t\t" ~ a).join("\n");
        else
            return name ~ "\t\t" ~ args.map!(a => stringify(name, a)).join(" ");
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
DllDefFile dllDefFile(Modules...)
                     (string libName,
                      string description,
                      Flag!"onlyExports" onlyExports = No.onlyExports)
    if(allSatisfy!(isSomeString, typeof(Modules)))
{
    import std.conv: to;

    auto statements = [
        Statement("LIBRARY", libName),
    ];

    string[] exports = ["xlAutoOpen", "xlAutoClose", "xlAutoFree12"];
    foreach(func; getAllWorksheetFunctions!Modules(onlyExports)) {
        exports ~= func.procedure.to!string;
    }

    return DllDefFile(statements ~ Statement("EXPORTS", exports));
}


///
mixin template GenerateDllDef(string module_ = __MODULE__) {
    version(exceldDef) {
        void main(string[] args) nothrow {
            import xlld.wrap.traits: generateDllDef;
            try {
                generateDllDef!module_(args);
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

///
void generateDllDef(string module_ = __MODULE__,
                    Flag!"onlyExports" onlyExports = No.onlyExports)
                   (string[] args)
{
    import std.stdio: File;
    import std.exception: enforce;
    import std.path: stripExtension;

    enforce(args.length >= 2 && args.length <= 4,
            "Usage: " ~ args[0] ~ " [file_name] <lib_name> <description>");

    immutable fileName = args[1];
    immutable libName = args.length > 2
        ? args[2]
        : fileName.stripExtension ~ ".xll";
    immutable description = args.length > 3
        ? args[3]
        : "Simple D add-in to Excel";

    auto file = File(fileName, "w");
    foreach(stmt; dllDefFile!module_(libName, description, onlyExports).statements)
        file.writeln(stmt.toString);
}

/**
   UDA for functions to be executed asynchronously
 */
enum Async;

version(unittest) {
// to link
    extern(C) auto getWorksheetFunctions() @safe pure nothrow {
        import xlld: WorksheetFunction;
        WorksheetFunction[] ret;
        return ret;
    }
}
