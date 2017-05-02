module xlld.any;


struct Any {
    import xlld.xlcall: XLOPER12;

    package XLOPER12 _impl;
    alias _impl this;
}


auto any(T, A)(auto ref T value, auto ref A allocator) {
    import xlld.wrap: toXlOper;
    return Any(value.toXlOper(allocator));
}
