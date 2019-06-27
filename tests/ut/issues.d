module ut.issues;


import test;
import xlld;


@("74")
@safe unittest {
    auto oper = () @trusted { return [[1, 2, 3], [4, 5, 6]].toXlOper(theGC); }();
    auto back = oper.fromXlOper!(Any[][])(theGC);
}
