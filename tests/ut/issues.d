module ut.issues;


import test;
import xlld;


@("74")
@safe unittest {
    auto oper = () @trusted { return [[1, 2, 3], [4, 5, 6]].toXlOper(theGC); }();
    auto back = oper.fromXlOper!(Any[][])(theGC);
}


@("89")
@safe unittest {
    import ut.wrap.wrapped: appendFoo;
    import xlld.memorymanager: allocator;
    auto arg = toXlOper("quux", allocator);
    appendFoo(&arg).shouldEqualDlang("quux_foo");
}


@("92")
@safe pure unittest {
    wrapModuleMember!("test.d_funcs", "core.stdc.stdio").should == "";
}
