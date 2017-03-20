/**
 Wraps calls to xlXXX "functions" via the Excel4/Excel12 functions
 */
module xlld.xl;

import xlld.xlcall: XLOPER12, LPXLOPER12;

version(unittest) {

    // this version(unittest) block effectively "implements" the Excel12v function
    // so that the code can be unit tested without needing to link the the Excel SDK
    import xlld.xlcallcpp: EXCEL12PROC, SetExcel12EntryPt;

    static this() {
        SetExcel12EntryPt(&excel12UnitTest);
    }

    static ~this() {
        import xlld.wrap: gNumXlFree, gNumXlCoerce, gCoerced, gFreed;
        import unit_threaded;
        gCoerced[0 .. gNumXlCoerce].shouldBeSameSetAs(gFreed[0 .. gNumXlFree]);
    }


    extern(Windows) int excel12UnitTest (int xlfn, int numOpers, LPXLOPER12 *opers, LPXLOPER12 result) nothrow @nogc {

        import xlld.xlcall: XlType, xlretFailed, xlretSuccess, xlFree, xlCoerce;
        import xlld.wrap: gReferencedType, gNumXlFree, gNumXlCoerce, gCoerced, gFreed, toXlOper;

        switch(xlfn) {

        default:
            return xlretFailed;

        case xlFree:
            assert(numOpers == 1);
            auto oper = opers[0];

            gFreed[gNumXlFree++] = oper.val.str;

            if(oper.xltype == XlType.xltypeStr)
                *oper = "".toXlOper;

            return xlretSuccess;

        case xlCoerce:
            assert(numOpers == 1);

            auto oper = opers[0];
            gCoerced[gNumXlCoerce++] = oper.val.str;
            *result = *oper;

            switch(oper.xltype) with(XlType) {

            case xltypeSRef:
                result.xltype = gReferencedType;
                break;

            case xltypeNum:
            case xltypeStr:
                result.xltype = oper.xltype;
                break;

            case xltypeMissing:
                result.xltype = xltypeNil;
                break;

            default:
            }

            return xlretSuccess;
        }
    }
}

XLOPER12 coerce(LPXLOPER12 oper) nothrow @nogc {
    import xlld.framework: Excel12f;
    import xlld.xlcall: xlCoerce;

    XLOPER12 coerced;
    LPXLOPER12[1] arg = [oper];
    Excel12f(xlCoerce, &coerced, arg);
    return coerced;
}

void free(LPXLOPER12 oper) nothrow @nogc {
    import xlld.framework: Excel12f;
    import xlld.xlcall: xlFree;

    LPXLOPER12[1] arg = [oper];
    Excel12f(xlFree, null, arg);
}
