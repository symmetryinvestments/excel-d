module ut.func.xlf;

import test;
import xlld.func.xlf;

@("callerCell throws if caller is string")
unittest {
    import xlld.sdk.xlcall: xlfCaller;
    import xlld.conv.to: toXlOper;

    with(MockXlFunction(xlfCaller, "foobar".toXlOper(theGC))) {
        callerCell.shouldThrowWithMessage("Caller not a cell");
    }
}


@("callerCell with SRef")
unittest {
    import xlld.sdk.xlcall: xlfCaller;

    with(MockXlFunction(xlfCaller, "foobar".toSRef(theGC))) {
        auto oper = callerCell;
        oper.shouldEqualDlang("foobar");
    }
}
