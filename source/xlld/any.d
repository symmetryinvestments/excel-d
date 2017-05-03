module xlld.any;


struct Any {
    import xlld.xlcall: XLOPER12;

    XLOPER12 _impl;
    alias _impl this;

    version(unittest) {

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


    string toString() @trusted const {
        import std.conv: text, to;
        import xlld.xlcall: XlType;
        import xlld.wrap: fromXlOper;
        import xlld.memorymanager: gMemoryPool;

        scope(exit) gMemoryPool.deallocateAll;

        string ret = text("Any(", );
        switch(_impl.xltype) {
        default:
            ret ~= _impl.xltype.to!string;
            break;
        case XlType.xltypeStr:
            ret ~= text(`"`, _impl.fromXlOper!string(gMemoryPool), `"`);
            break;
        case XlType.xltypeNum:
            ret ~= _impl.fromXlOper!double(gMemoryPool).to!string;
            break;
        case XlType.xltypeMulti:
            int i;
            ret ~= `[`;
            foreach(r; 0 .. _impl.val.array.rows) {
                ret ~= `[`;
                foreach(c; 0 .. _impl.val.array.columns) {
                    ret ~= text(Any(cast(XLOPER12)_impl.val.array.lparray[i++]), `, `);
                }
                ret ~= `]`;
            }
            ret ~= `]`;
            break;
        }
        return ret ~ ")";
    }
}


auto any(T, A)(auto ref T value, auto ref A allocator) @trusted {
    import xlld.wrap: toXlOper;
    return Any(value.toXlOper(allocator));
}
