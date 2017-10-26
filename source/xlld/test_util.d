/**
    Utility test functions
*/
module xlld.test_util;

version(unittest):

import xlld.xlcall: LPXLOPER12, XLOPER12, XlType;
import unit_threaded;
import std.range: isInputRange;

///
TestAllocator gTestAllocator;
/// emulates SRef types by storing what the referenced type actually is
XlType gReferencedType;

// tracks calls to `coerce` and `free` to make sure memory allocations/deallocations match
int gNumXlCoerce;
///
int gNumXlFree;
///
enum maxCoerce = 1000;
///
void*[maxCoerce] gCoerced;
///
void*[maxCoerce] gFreed;
///
double[] gDates, gTimes, gYears, gMonths, gDays, gHours, gMinutes, gSeconds;

///
static this() {
    import xlld.xlcallcpp: SetExcel12EntryPt;
    // this effectively "implements" the Excel12v function
    // so that the code can be unit tested without needing to link
    // with the Excel SDK
    SetExcel12EntryPt(&excel12UnitTest);
}

///
static ~this() {
    import unit_threaded.should: shouldBeSameSetAs;
    gCoerced[0 .. gNumXlCoerce].shouldBeSameSetAs(gFreed[0 .. gNumXlFree]);
}

///
extern(Windows) int excel12UnitTest(int xlfn, int numOpers, LPXLOPER12 *opers, LPXLOPER12 result) nothrow @nogc {

    import xlld.xlcall;
    import xlld.wrap: toXlOper;
    import std.experimental.allocator.mallocator: Mallocator;
    import std.array: front, popFront, empty;

    switch(xlfn) {

    default:
        return xlretFailed;

    case xlFree:
        assert(numOpers == 1);
        auto oper = opers[0];

        gFreed[gNumXlFree++] = oper.val.str;

        if(oper.xltype == XlType.xltypeStr) {
            try
                *oper = "".toXlOper(Mallocator.instance);
            catch(Exception _) {
                assert(false, "Error converting in excel12UnitTest");
            }
        }

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

    case xlfDate:
        return returnGlobalMockFrom(gDates, result);

    case xlfTime:
        return returnGlobalMockFrom(gTimes, result);

    case xlfYear:
        return returnGlobalMockFrom(gYears, result);

    case xlfMonth:
        return returnGlobalMockFrom(gMonths, result);

    case xlfDay:
        return returnGlobalMockFrom(gDays, result);

    case xlfHour:
        return returnGlobalMockFrom(gHours, result);

    case xlfMinute:
        return returnGlobalMockFrom(gMinutes, result);

    case xlfSecond:
        return returnGlobalMockFrom(gSeconds, result);
    }
}

private int returnGlobalMockFrom(R)(R values, LPXLOPER12 result) if(isInputRange!R) {
    import xlld.wrap: toXlOper;
    import xlld.xlcall: xlretSuccess;
    import std.array: front, popFront, empty;
    import std.experimental.allocator.mallocator: Mallocator;
    import std.range: ElementType;

    const ret = values.empty ? ElementType!R.init : values.front;
    if(!values.empty) values.popFront;

    *result = ret.toXlOper(Mallocator.instance);
    return xlretSuccess;

}

/// automatically converts from oper to compare with a D type
void shouldEqualDlang(U)(LPXLOPER12 actual, U expected, string file = __FILE__, size_t line = __LINE__) @trusted {
    import xlld.memorymanager: allocator;
    import xlld.wrap: fromXlOper, stripMemoryBitmask;
    import xlld.xlcall: XlType;
    import std.traits: Unqual;
    import std.conv: text;
    import std.experimental.allocator.gc_allocator: GCAllocator;

    actual.shouldNotBeNull;

    const type = actual.xltype.stripMemoryBitmask;

    if(type == XlType.xltypeErr)
        fail("XLOPER is of error type", file, line);

    static if(!is(Unqual!U == string))
        if(type == XlType.xltypeStr)
            fail(text("XLOPER is of string type. Value: ", actual.fromXlOper!string(GCAllocator.instance)), file, line);

    actual.fromXlOper!U(allocator).shouldEqual(expected, file, line);
}

/// automatically converts from oper to compare with a D type
void shouldEqualDlang(U)(ref XLOPER12 actual, U expected, string file = __FILE__, size_t line = __LINE__) @trusted {
    shouldEqualDlang(&actual, expected, file, line);
}

/// automatically converts from oper to compare with a D type
void shouldEqualDlang(U)(XLOPER12 actual, U expected, string file = __FILE__, size_t line = __LINE__) @trusted {
    shouldEqualDlang(actual, expected, file, line);
}


///
XLOPER12 toSRef(T, A)(T val, ref A allocator) @trusted {
    import xlld.wrap: toXlOper;

    auto ret = toXlOper(val, allocator);
    //hide real type somewhere to retrieve it
    gReferencedType = ret.xltype;
    ret.xltype = XlType.xltypeSRef;
    return ret;
}

/// Mimics Excel calling a particular D function, including freeing memory
void fromExcel(alias F, A...)(auto ref A args) {
    import xlld.wrap: wrapModuleFunctionImpl;
    import xlld.memorymanager: gTempAllocator, autoFree;

    auto oper = wrapModuleFunctionImpl!F(gTempAllocator, args);
    autoFree(oper);
}


/// tracks allocations and throws in the destructor if there is a memory leak
/// it also throws when there is an attempt to deallocate memory that wasn't
/// allocated
struct TestAllocator {

    import std.experimental.allocator.common: platformAlignment;
    import std.experimental.allocator.mallocator: Mallocator;

    ///
    alias allocator = Mallocator.instance;

    private static struct ByteRange {
        void* ptr;
        size_t length;
        inout(void)[] opSlice() @trusted @nogc inout nothrow {
            return ptr[0 .. length];
        }
    }

    ///
    bool debug_;
    private ByteRange[] _allocations;
    private ByteRange[] _deallocations;
    private int _numAllocations;

    ///
    enum uint alignment = platformAlignment;

    ///
    void[] allocate(size_t numBytes) @trusted @nogc {
        import std.experimental.allocator: makeArray, expandArray;
        import core.stdc.stdio: printf;

        static __gshared immutable exception = new Exception("Allocation failed");

        ++_numAllocations;

        auto ret = allocator.allocate(numBytes);
        if(numBytes > 0 && ret.length == 0)
            throw exception;

        if(debug_) () @trusted { printf("+   allocate: %p for %d bytes\n", &ret[0], cast(int)ret.length); }();
        auto newEntry = ByteRange(&ret[0], ret.length);

        if(_allocations is null)
            _allocations = allocator.makeArray(1, newEntry);
        else
            () @trusted { allocator.expandArray(_allocations, 1, newEntry); }();

        return ret;
    }

    ///
    bool deallocate(void[] bytes) @trusted @nogc nothrow {
        import std.experimental.allocator: makeArray, expandArray;
        import std.algorithm: remove, find, canFind;
        import std.array: front, empty;
        import core.stdc.stdio: printf, sprintf;

        bool pred(ByteRange other) { return other.ptr == bytes.ptr && other.length == bytes.length; }

        static char[1024] buffer;

        if(debug_) printf("- deallocate: %p for %d bytes\n", &bytes[0], cast(int)bytes.length);

        auto findAllocations = _allocations.find!pred;

        if(findAllocations.empty) {
            if(_deallocations.canFind!pred) {
                auto index = sprintf(&buffer[0],
                                     "Double free on byte range Ptr: %p, length: %ld, allocations:\n",
                                     &bytes[0], bytes.length);
                index = printAllocations(buffer, index);
                assert(false, buffer[0 .. index]);

            } else {
                auto index = sprintf(&buffer[0],
                                     "Unknown deallocate byte range. Ptr: %p, length: %ld, allocations:\n",
                                     &bytes[0], bytes.length);
                index = printAllocations(buffer, index);
                assert(false, buffer[0 .. index]);
            }
        }

        if(_deallocations is null)
            _deallocations = allocator.makeArray(1, findAllocations.front);
        else
            () @trusted { allocator.expandArray(_allocations, 1, findAllocations.front); }();


        _allocations = _allocations.remove!pred;

        return () @trusted { return allocator.deallocate(bytes); }();
    }

    ///
    bool deallocateAll() @safe @nogc nothrow {
        import std.array: empty, front;

        while(!_allocations.empty) {
            auto allocation = _allocations.front;
            deallocate(allocation[]);
        }
        return true;
    }

    ///
    auto numAllocations() @safe @nogc pure nothrow const {
        return _numAllocations;
    }

    ///
    ~this() @safe @nogc nothrow {
        verify;
    }

    ///
    void verify() @trusted @nogc nothrow {

        static char[1024] buffer;

        if(_allocations.length) {
            import core.stdc.stdio: sprintf;
            auto index = sprintf(&buffer[0], "Memory leak in TestAllocator. Allocations:\n");
            index = printAllocations(buffer, index);
            assert(false, buffer[0 .. index]);
        }
    }

    ///
    int printAllocations(int N)(ref char[N] buffer, int index = 0) @trusted @nogc const nothrow {
        import core.stdc.stdio: sprintf;
        index += sprintf(&buffer[index], "[\n");
        foreach(ref allocation; _allocations) {
            index += sprintf(&buffer[index], "    ByteRange(%p, %ld),\n",
                             allocation.ptr, allocation.length);
        }

        index += sprintf(&buffer[index], "]");
        buffer[index++] = 0; // null terminate
        return index;
    }
}
