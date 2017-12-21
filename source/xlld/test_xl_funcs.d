
/**
 Only exists to test the module reflection functionality
 Contains functions with types that the spreadsheet knows about
 */

module xlld.test_xl_funcs;

version(unittest):

import xlld.xlcall;
import xlld.worksheet;
import xlld.traits: Async;

extern(Windows) double FuncMulByTwo(double n) nothrow {
    return n * 2;
}

extern(Windows) double FuncFP12(FP12* cells) nothrow {
    return 0;
}


extern(Windows) LPXLOPER12 FuncFib (LPXLOPER12 n) nothrow {
    return LPXLOPER12.init;
}

@Async
extern(Windows) void FuncAsync(LPXLOPER12 n, LPXLOPER12 asyncHandle) nothrow {

}
