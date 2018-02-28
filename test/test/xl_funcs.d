
/**
 Only exists to test the module reflection functionality
 Contains functions with types that the spreadsheet knows about
 */

module xlld.test.xl_funcs;

version(unittest):

import xlld.from;

import xlld.sdk.xlcall;
import xlld.wrap.worksheet;
import xlld.wrap.traits: Async;

extern(Windows) double FuncMulByTwo(double n) nothrow {
    return n * 2;
}

extern(Windows) double FuncFP12(FP12* cells) nothrow {
    return 0;
}


extern(Windows) from!"xlld.sdk.xlcall".LPXLOPER12 FuncFib (from!"xlld.sdk.xlcall".LPXLOPER12 n) nothrow {
    return null;
}

@Async
extern(Windows) void FuncAsync(from!"xlld.sdk.xlcall".LPXLOPER12 n,
                               from!"xlld.sdk.xlcall".LPXLOPER12 asyncHandle)
    nothrow
{

}
