/**
   This module exists to house the static constructors/destructors to prevent cycles.
   No module imports this one.
 */

module xlld.static_;

static this() {
    version(unittest) {
        import xlld.xlcallcpp: SetExcel12EntryPt;
        import xlld.test_util: excel12UnitTest;
        // this effectively "implements" the Excel12v function
        // so that the code can be unit tested without needing to link
        // with the Excel SDK
        SetExcel12EntryPt(&excel12UnitTest);
    }
}

version(unittest) {
    static ~this() {
        import xlld.test_util: gCoerced, gNumXlCoerce, gFreed, gNumXlFree;
        import unit_threaded.should: shouldBeSameSetAs;
        gCoerced[0 .. gNumXlCoerce].shouldBeSameSetAs(gFreed[0 .. gNumXlFree]);
    }
}
