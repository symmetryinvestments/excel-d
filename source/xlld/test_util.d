/**
    Utility test functions
*/
module xlld.test_util;

version(unittest):

import xlld.xlcall: LPXLOPER12, XLOPER12, XlType;
import unit_threaded;
import std.range: isInputRange;
import std.experimental.allocator.gc_allocator: GCAllocator;
alias theGC = GCAllocator.instance;


///
TestAllocator gTestAllocator;
/// emulates SRef types by storing what the referenced type actually is
XlType gReferencedType;

// tracks calls to `coerce` and `free` to make sure memory allocations/deallocations match
int gNumXlAllocated;
///
int gNumXlFree;
///
enum maxAllocTrack = 1000;
///
const(void)*[maxAllocTrack] gAllocated;
///
const(void)*[maxAllocTrack] gFreed;

alias XlFunction = int;

/**
   This stores what Excel12f should "return" for each integer XL "function"
 */
XLOPER12[][XlFunction] gXlFuncResults;

private shared AA!(XLOPER12, XLOPER12) gAsyncReturns = void;


XLOPER12 asyncReturn(XLOPER12 asyncHandle) @safe {
    return gAsyncReturns[asyncHandle];
}

private void fakeAllocate(XLOPER12 oper) @nogc nothrow {
    fakeAllocate(&oper);
}

private void fakeAllocate(XLOPER12* oper) @nogc nothrow {
    gAllocated[gNumXlAllocated++] = oper.val.str;
}


private void fakeFree(XLOPER12* oper) @nogc nothrow {
    gFreed[gNumXlFree++] = oper.val.str;
}

///
extern(Windows) int excel12UnitTest(int xlfn, int numOpers, LPXLOPER12 *opers, LPXLOPER12 result)
    nothrow @nogc
{

    import xlld.xlcall;
    import xlld.conv: toXlOper, stripMemoryBitmask;
    import std.experimental.allocator.gc_allocator: GCAllocator;
    import std.experimental.allocator.mallocator: Mallocator;
    import std.array: front, popFront, empty;

    if(auto xlfnResults = xlfn in gXlFuncResults) {

        assert(!(*xlfnResults).empty, "No results to return for xlfn");

        auto mockResult = (*xlfnResults).front;
        (*xlfnResults).popFront;

        if(xlfn == xlfCaller) {
            fakeAllocate(mockResult);
        }
        *result = mockResult;
        return xlretSuccess;
    }

    switch(xlfn) {

    default:
        return xlretFailed;

    case xlFree:
        assert(numOpers == 1);
        auto oper = opers[0];

        fakeFree(oper);

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
        fakeAllocate(oper);
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

    case xlAsyncReturn:
        assert(numOpers == 2);
        gAsyncReturns[*opers[0]] = *opers[1];
        return xlretSuccess;
    }
}


/// automatically converts from oper to compare with a D type
void shouldEqualDlang(U)
                     (XLOPER12* actual,
                      U expected,
                      string file = __FILE__,
                      size_t line = __LINE__)
    @trusted
{
    import xlld.memorymanager: allocator;
    import xlld.conv: fromXlOper, stripMemoryBitmask;
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
    import xlld.conv: toXlOper;

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


/**
   @nogc associative array
 */
struct AA(K, V, int N = 100) {
    import core.thread: Mutex;

    Entry[N] entries;
    size_t index;
    Mutex mutex;

    static struct Entry {
        K key;
        V value;
    }

    @disable this();

    static shared(AA) create() @trusted {
        shared AA aa = void;
        aa.index = 0;
        aa.mutex = new shared Mutex;
        foreach(ref entry; aa.entries[]) {
            entry = typeof(entry).init;
        }
        return aa;
    }

    V opIndex(in K key) shared {
        import std.algorithm: find;
        import std.array: front, empty;

        mutex.lock_nothrow;
        scope(exit) mutex.unlock_nothrow;

        auto fromKey = () @trusted { return entries[].find!(e => cast(K)e.key == key); }();
        return fromKey.empty
            ? V.init
            : cast(V)fromKey.front.value;
    }

    void opIndexAssign(V value, K key) shared {
        import core.atomic: atomicOp;

        mutex.lock_nothrow;
        scope(exit) mutex.unlock_nothrow;

        assert(index < N - 1, "No more space");
        entries[index] = Entry(key, value);
        index.atomicOp!"+="(1);
    }
}

@("AA")
@safe unittest {
    import core.exception: AssertError;
    auto aa = AA!(string, int).create;
    aa["foo"] = 5;
    aa["foo"].shouldEqual(5);
    aa["bar"].shouldEqual(0);
}


XLOPER12 newAsyncHandle() @safe nothrow {
    import xlld.xlcall: XlType;
    import core.atomic: atomicOp;

    static shared typeof(XLOPER12.val.w) index = 1;
    XLOPER12 asyncHandle;

    asyncHandle.xltype = XlType.xltypeBigData;
    () @trusted { asyncHandle.val.bigdata.h.hdata = cast(void*)index; }();
    index.atomicOp!"+="(1);

    return asyncHandle;
}

struct MockXlFunction {

    XlFunction xlFunction;
    typeof(gXlFuncResults) oldResults;

    this(int xlFunction, XLOPER12 result) @safe {
        this(xlFunction, [result]);
    }

    this(int xlFunction, XLOPER12[] results) @safe {
        this.xlFunction = xlFunction;
        this.oldResults = gXlFuncResults.dup;
        gXlFuncResults[xlFunction] ~= results;
    }

    ~this() @safe {
        gXlFuncResults = oldResults;
    }
}

struct MockDateTime {

    import xlld.xlcall: xlfYear, xlfMonth, xlfDay, xlfHour, xlfMinute, xlfSecond;

    MockXlFunction year, month, day, hour, minute, second;

    this(int year, int month, int day, int hour, int minute, int second) @safe {
        import xlld.conv: toXlOper;

        this.year   = MockXlFunction(xlfYear,   double(year).toXlOper(theGC));
        this.month  = MockXlFunction(xlfMonth,  double(month).toXlOper(theGC));
        this.day    = MockXlFunction(xlfDay,    double(day).toXlOper(theGC));
        this.hour   = MockXlFunction(xlfHour,   double(hour).toXlOper(theGC));
        this.minute = MockXlFunction(xlfMinute, double(minute).toXlOper(theGC));
        this.second = MockXlFunction(xlfSecond, double(second).toXlOper(theGC));
    }
}

struct MockDateTimes {

    import std.datetime: DateTime;

    MockDateTime[] mocks;

    this(DateTime[] dateTimes...) @safe {
        foreach(dateTime; dateTimes)
            mocks ~= MockDateTime(dateTime.year, dateTime.month, dateTime.day,
                                  dateTime.hour, dateTime.minute, dateTime.second);
    }
}


struct FailingAllocator {
    void[] allocate(size_t numBytes) @safe @nogc pure nothrow {
        return null;
    }

    bool deallocate(void[] bytes) @safe @nogc pure nothrow {
        assert(false);
    }
}
