/**
 Deals with linking issues such as on Windows without having to link to the
 real implementations or dependent packages' unittest builds.
 Only for testing.
 */
module xlld.dummy;



version(XllDummyGetter) {
    // to be able to link
    extern(C) auto getWorksheetFunctions() @safe pure nothrow {
        import xlld: WorksheetFunction;
        WorksheetFunction[] ret;
        return ret;
    }
}


version(testingExcelD)
    enum useDummy = true;
else version(exceldDef)
    enum useDummy = true;
else
    enum useDummy = false;

version(Windows):
static if(useDummy) {

    import xlld.sdk.xlcall;

    extern(System) int Excel4v(int xlfn, LPXLOPER operRes, int count, LPXLOPER* opers) {
        return 0;
    }

    extern(System) int Excel4(int xlfn, LPXLOPER operRes, int count,... ) {
        return 0;
    }
}
