/**
   Any type
 */
module xlld.any;


///
struct Any {
    import xlld.xlcall: XLOPER12;

    ///
    XLOPER12 _impl;
    alias _impl this;

    version(unittest) {

        ///
        bool opEquals(Any other) @trusted const {
            import xlld.xlcall: XlType;
            import xlld.wrap: fromXlOper;

            switch(_impl.xltype) {

            default:
                return _impl == other._impl;

            case XlType.xltypeStr:

                import std.experimental.allocator.gc_allocator: GCAllocator;
                return _impl.fromXlOper!(string)(GCAllocator.instance) ==
                    other._impl.fromXlOper!(string)(GCAllocator.instance);

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
    }


    ///
    string toString() @safe const {
        import std.conv: text, to;
        import xlld.xlcall: XlType;
        import xlld.wrap: fromXlOper;
        import xlld.xlcall: xlbitXLFree, xlbitDLLFree;
        import std.experimental.allocator.gc_allocator: GCAllocator;

        alias allocator = GCAllocator.instance;

        string ret = text("Any(", );
        const type = _impl.xltype & ~(xlbitXLFree | xlbitDLLFree);
        switch(type) {
        default:
            ret ~= type.to!string;
            break;
        case XlType.xltypeStr:
            ret ~= () @trusted { return text(`"`, _impl.fromXlOper!string(allocator), `"`); }();
            break;
        case XlType.xltypeNum:
            ret ~= () @trusted { return _impl.fromXlOper!double(allocator).to!string; }();
            break;
        case XlType.xltypeInt:
            ret ~= () @trusted { return _impl.fromXlOper!int(allocator).to!string; }();
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
                ret ~= `]`;
            }
            ret ~= `]`;
            break;
        }
        return ret ~ ")";
    }
}


///
auto any(T, A)(auto ref T value, auto ref A allocator) @trusted {
    import xlld.wrap: toXlOper;
    return Any(value.toXlOper(allocator));
}
