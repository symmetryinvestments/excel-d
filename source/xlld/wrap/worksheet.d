/**
   Interface for registering worksheet functions with Excel
 */
module xlld.wrap.worksheet;

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

// see: https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1

///
struct Procedure {
    wstring value;
    string toString() @safe pure const {
        import std.conv: to;
        return value.to!string;
    }
}

mixin ST!"TypeText";
/++
	Function name as it will appear in the Function Wizard.
+/
mixin ST!"FunctionText";
/++
	Names of the arguments, semicolon separated. For example:

	foo;bar
+/
mixin ST!"ArgumentText";
mixin ST!"MacroType";
/++
	Category name it appears in in the wizard. Must be picked
	from the list of existing categories in Excel.
+/
mixin ST!"Category";
/++
	One-character, case-sensitive key. "A" assigns this command
	to ctrl+shift+A. Used only for commands.
+/
mixin ST!"ShortcutText";
/++
	Reference to the help file to display when the user clicks help.
	Form of `filepath!HelpContextID` or `URL!0`.

	The !0 is required if you provide a web link.
+/
mixin ST!"HelpTopic";
/++
	Describes your function in the Function Wizard.
++/
mixin ST!"FunctionHelp";
/++
	Array of text strings displayed in the function dialog
	in Excel to describe each arg.
+/
struct ArgumentHelp {
	wstring[] value;
	// allow @ArgumentHelp(x, y, x) too
	this(wstring[] txt...) pure nothrow @safe {
		// this is fine because below it is all copied
		// into GC memory anyway.
		this(txt[]);
	}
	this(scope wstring[] txt) pure nothrow @safe {
		// Excel has a bug that chops off the last
		// character of this, so adding a space here
		// works around that.
		//
		// Doing it here instead of below, at
		// registration time, means it is also CTFE'd
		foreach(t; txt)
			value ~= t ~ " ";
	}
}


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

    bool opEquals(const WorksheetFunction rhs) const @safe pure { return optional == rhs.optional; }
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
alias Excel = Optional;


///
struct Dispose(alias function_) {
    alias dispose = function_;
}
