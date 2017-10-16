
/**
 Only exists to test the module reflection functionality
 Contains functions with types that the spreadsheet knows about
 */

module xlld.test_xl_funcs;

version(unittest):

import xlld.xlcall;
import xlld.worksheet;

/// extern(C) export means it doesn't have to be explicitly
/// added to the .def file
extern(C) export double FuncMulByTwo(double n) nothrow {
    return n * 2;
}

///
extern(C) export double FuncFP12(FP12* cells) nothrow {
    return 0;
}


///
extern(C) export LPXLOPER12 FuncFib (LPXLOPER12 n) nothrow {
    return LPXLOPER12.init;
}
