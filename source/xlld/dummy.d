/**
 Enables linking on  Windows without having to link to the real implementations
 Only for unit testing.
 */
module xlld.dummy;

version(unittest)
    enum useDummy = true;
else version(exceldDef)
    enum useDummy = true;
else
    enum useDummy = false;

version(Windows):
static if(useDummy) {

    import xlld.xlcall;

    extern(System) int Excel4v(int xlfn, LPXLOPER operRes, int count, LPXLOPER* opers) {
        return 0;
    }

    extern(System) int Excel4(int xlfn, LPXLOPER operRes, int count,... ) {
        return 0;
    }
}
