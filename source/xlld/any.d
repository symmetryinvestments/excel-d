/**
   Any type
 */
module xlld.any;

///
struct Any {
    import xlld.sdk.xlcall: XLOPER12;

    ///
    XLOPER12 _impl;
    alias _impl this;

    ///
    bool opEquals(Any other) @trusted const {
        import xlld.sdk.xlcall: XlType;
        import xlld.conv.from: fromXlOper;

        switch(_impl.xltype) {

        default:
            import xlld.conv.misc: stripMemoryBitmask;

            if(_impl.xltype.stripMemoryBitmask != other._impl.xltype.stripMemoryBitmask)
                return false;

            XLOPER12 comp = _impl;
            comp.xltype = other._impl.xltype;

            return comp == other._impl;

        case XlType.xltypeStr:

            import std.experimental.allocator.gc_allocator: GCAllocator;
            return _impl.fromXlOper!string(GCAllocator.instance) ==
                other._impl.fromXlOper!string(GCAllocator.instance);

        case XlType.xltypeMulti:

            if(_impl.val.array.rows != other._impl.val.array.rows) return false;
            if(_impl.val.array.columns != other._impl.val.array.columns) return false;

            int i;
            foreach(r; 0 .. _impl.val.array.rows) {
                foreach(c; 0 .. _impl.val.array.columns) {
                    if(Any(cast(XLOPER12)_impl.val.array.lparray[i]) !=
                       Any(cast(XLOPER12)other._impl.val.array.lparray[i]))
                        return false;
                    ++i;
                }
            }

            return true;
        }
    }

    ///
    string toString() @safe const {
        import std.conv: text;
        import xlld.sdk.xlcall: XlType;
        import xlld.conv.from: fromXlOper;
        import xlld.sdk.xlcall: xlbitXLFree, xlbitDLLFree;
        import std.experimental.allocator.gc_allocator: GCAllocator;

        alias allocator = GCAllocator.instance;

        string ret = text("Any(", );
        const type = _impl.xltype & ~(xlbitXLFree | xlbitDLLFree);
        switch(type) {
        default:
            ret ~= type.text;
            break;
        case XlType.xltypeStr:
            ret ~= () @trusted { return text(`"`, _impl.fromXlOper!string(allocator), `"`); }();
            break;
        case XlType.xltypeNum:
            ret ~= () @trusted { return _impl.fromXlOper!double(allocator).text; }();
            break;
        case XlType.xltypeInt:
            ret ~= () @trusted { return _impl.fromXlOper!int(allocator).text; }();
            break;
        case XlType.xltypeMulti:
            int i;
            ret ~= `[`;
            const rows = () @trusted { return _impl.val.array.rows; }();
            const cols = () @trusted { return _impl.val.array.columns; }();
            foreach(r; 0 .. rows) {
                ret ~= `[`;
                foreach(c; 0 .. cols) {
                    auto oper = () @trusted { return _impl.val.array.lparray[i++]; }();
                    ret ~= text(Any(cast(XLOPER12)oper), `, `);
                }
                ret ~= `], `;
            }
            ret ~= `]`;
            break;
        }
        return ret ~ ")";
    }

}


///
auto any(T, A)(auto ref T value, auto ref A allocator, in string file = __FILE__, in size_t line = __LINE__) @trusted {
    import xlld.conv: toXlOper;
    static if(__traits(compiles, Any(value.toXlOper(allocator, file, line))))
        return Any(value.toXlOper(allocator, file, line));
    else
        return Any(value.toXlOper(allocator));
}
