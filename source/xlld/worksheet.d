/**
   Interface for registering worksheet functions with Excel
 */
module xlld.worksheet;

/**
 Simple wrapper struct for a value. Provides a type-safe way
 of making sure positional arguments match the intended semantics,
 which is important given that nearly all of the arguments for
 worksheet function registration are of the same type: wstring
 ST is short for "SmallType".
 */
private mixin template ST(string name, T = wstring) {
    mixin(`struct ` ~ name ~ `{ T value; }`);
}

///
struct Procedure {
    wstring value;
    string toString() @safe pure const {
        import std.conv: to;
        return value.to!string;
    }
}

mixin ST!"TypeText";
mixin ST!"FunctionText";
mixin ST!"ArgumentText";
mixin ST!"MacroType";
mixin ST!"Category";
mixin ST!"ShortcutText";
mixin ST!"HelpTopic";
mixin ST!"FunctionHelp";
mixin ST!("ArgumentHelp", wstring[]);


/**
   The arguments used to register a worksheet function with the spreadsheet.
 */
struct WorksheetFunction {
    // the first few parameters have to be set, the others are optional
    Procedure procedure;
    TypeText typeText;
    FunctionText functionText;
    Optional optional;
    alias optional this; //for ease of use

    /**
       Returns an array suitable for use with spreadsheet registration
     */
    const(wstring)[] toStringArray() @safe pure const nothrow {
        return
            [
                procedure.value, typeText.value,
                functionText.value, argumentText.value,
                macroType.value, category.value,
                shortcutText.value, helpTopic.value, functionHelp.value
            ] ~ argumentHelp.value;
    }
}

// helper template to type-check variadic template constructor below
private alias toType(alias U) = typeof(U);

/**
   Optional arguments that can be set by a function author, but don't necessarily
   have to be.
 */
struct Optional {
    ArgumentText argumentText;
    MacroType macroType = MacroType("1"w);
    Category category;
    ShortcutText shortcutText;
    HelpTopic helpTopic;
    FunctionHelp functionHelp;
    ArgumentHelp argumentHelp;

    this(T...)(T args) {
        import std.meta: staticIndexOf, staticMap, allSatisfy, AliasSeq;
        import std.conv: text;

        static assert(T.length <= this.tupleof.length, "Too many arguments for Optional/Register");

        // myTypes: ArgumentText, MacroType, ...
        alias myTypes = staticMap!(toType, AliasSeq!(this.tupleof));
        enum isOneOfMyTypes(U) = staticIndexOf!(U, myTypes) != -1;
        static assert(allSatisfy!(isOneOfMyTypes, T),
                      text("Unknown types passed to Optional/Register constructor. ",
                           "Has to be one of:\n", myTypes.stringof));

        // loop over whatever was given and set each of our members based on the
        // type of the parameter instead of by position
        foreach(ref member; this.tupleof) {
            enum index = staticIndexOf!(typeof(member), T);
            static if(index != -1)
                member = args[index];
        }
    }
}

/**
    A user-facing name to use as an UDA to decorate D functions.
    Any arguments passed to its constructor will be used to register
    the function with the spreadsheet.
*/
alias Register = Optional;


///
struct Dispose(alias function_) {
    alias dispose = function_;
}
