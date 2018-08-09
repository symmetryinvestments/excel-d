# excel-d

[![Build Status](https://travis-ci.org/kaleidicassociates/excel-d.png?branch=master)](https://travis-ci.org/kaleidicassociates/excel-d)
[![Coverage](https://codecov.io/gh/kaleidicassociates/excel-d/branch/master/graph/badge.svg)](https://codecov.io/gh/kaleidicassociates/excel-d)

Excel API bindings and wrapper API for D

This dub package contains D declarations for the [Excel SDK](https://msdn.microsoft.com/en-us/library/office/bb687883.aspx)
 as well as a D wrapper API. This allows programmers to write Excel worksheet functions in D.

Motivation and background for the project can be found [here](https://dlang.org/blog/2017/05/31/project-highlight-excel-d/).
See also the [DConf 2017 lightning talk](https://youtu.be/xJy6ifCekCE?list=PL3jwVPmk_PRxo23yyoc0Ip_cP3-rCm7eB) about excel-d.

Generated documentation - a work in progress - is available at
[dpldocs](http://excel-d.dpldocs.info/index.html).

A working XLL example can be found in the [`example`](example)
directory. Running `dub build` there will create an XLL
(`myxll32.xll`) that can be loaded in Excel making all of the
functions in `test/xlld/test_d_funcs.d` available to be used in Excel
cells. The types are automatically converted between D native types
and Excel ones.  To build the example: `dub build -c example [--arch=x86_mscoff|--arch=x86_64]`.

For this package to build you will need the Excel SDK `xlcall32.lib`
that can be downloaded [from Microsoft](http://go.microsoft.com/fwlink/?LinkID=251082&clcid=0x409).
Copying it to the build directory should be sufficient
(i.e. when building the example, to the `example` directory).
The library file should be useable as-is, as long as on 32-bit Excel `dub build` is run with
`--arch=x86_mscoff` to use Microsoft's binary format. If linking with optlink, the file must
be converted first.  We recommend using `link.exe` to not need the conversion.  On 64 bit Excel
just use `--arch=x86_64` - no questions of different library formats.

As part of the build a `.def` file is generated with all functions to be exported by the XLL.

Excel won't load the XLL automatically: this must be done manually in File->Tools->Add-Ins.
Click on "Go" for "Excel Add-Ins" (the default) and select your XLL there after clicking on
"Browse".

The only difference between building for 32-bit or 64-bit Excel is the `arch=` option passed
to dub. A 32-bit XLL will only work on 32-bit Excel and similarly for 64-bit. You will also
need the appropriate 32/64 xlcall32.lib from the Excel SDK to link.

Sample code (see the [example](example) directory for more):

```d
    import xlld;

    @Excel(ArgumentText("Array to add"),
           HelpTopic("Adds all cells in an array"),
           FunctionHelp("Adds all cells in an array"),
           ArgumentHelp(["The array to add"]))
    double FuncAddEverything(double[][] args) nothrow @nogc { // nothrow and @nogc are optional
        import std.algorithm: fold;
        import std.math: isNaN;

        double ret = 0;
        foreach(row; args)
            ret += row.fold!((a, b) => b.isNaN ? a : a + b)(0.0);
        return ret;
    }
```

and then in Excel:

`=FuncAddEverything(A1:D20)`

Future functionality will include creating menu items and dialogue boxes.  Pull requests welcomed.


WARNING: Memory for parameters passed to D functions
---------------------------------------------------

Any parameters with indirections (pointers, slices) should NOT be escaped. The memory for those
parameters WILL be reused and might cause crashes.

There is support to fail at compile-time if user-written D functions attempt to escape their
arguments but unfortunately given the current D defaults requires user intervention. Annotate
all D code to be called by Excel with `@safe` and compile with `-dip1000` - all parameters will
then need to be `scope` or the code will compile.

It is *strongly* advised to compile with `-dip1000` and to make all your functions `@safe`,
or your add-ins could cause Excel to crash.


Function spelling
------------------

excel-d will always convert the first character in the D function being wrapped to uppercase
since that is the Excel convention.


Variant type `Any`
---------------------

Sometimes it is useful for a D function to take in any type that Excel supports. Typically
this will happen when receiving a matrix of values where the types might differ
(e.g. the columns are date, string, double). To get the expected D type from an `Any` value,
use `xlld.wrap.fromXlOper`. An example:

```d
double Func(Any[][] values) {
    import xlld.wrap: fromXlOper;
    import std.experimental.allocator: theAllocator;
    foreach(row; values) {
        auto date = row[0].fromXlOper!DateTime(theAllocator);
        auto string_ = row[1].fromXlOper!DateTime(theAllocator);
        auto double_ = row[2].fromXlOper!double(theAllocator);
        // ...
    }
    return ret;
}
```


Asynchronous functions
----------------------

A D function can be decorated with the `@Async` UDA and will be executed asynchronously:

```d
@Async
double AsyncFunc(double d) {
    // long-running task
}
```

Please see [the Microsoft documentation](https://msdn.microsoft.com/en-us/library/office/ff796219(v=office.14).aspx).

Custom enum coversions
----------------------

If the usual conversions between strings and enums don't work for the user, it is possible to register
custom coversions by calling the functions `registerConversionTo` and `registerConversionFrom`.

Structs
--------

D structs can be returned by functions. They are transformed into a string representation.

D structs can also be passed to functions. To do so, pass in a 1D array with the same number
of elements as the struct in question.


Optional custom memory allocation and `@nogc`
---------------------------------------------

If you are not familiar with questions of memory allocation, the below may seem intimidating.
However it's entirely optional and unless performance and latency are critical to you (or
possibly if you are interfacing with C or C++ code) then you do not need to worry about the
extra complexity introduced by using allocators.  The code in the previous section will simply
work.

excel-d uses a custom allocator for all allocations that are needed when doing the conversions
between D and Excel types. It uses a different one for allocations of XLOPER12s that are
returned to Excel, which are then freed in xlAutoFree12 with the same allocator. D functions
that are `@nogc` are wrapped by `@nogc` Excel functions and similarly for `@safe`. However,
if returning a value that is dynamically allocated from a D function and not using the GC
(such as an array of doubles), it is necessary to specify how that memory is to be freed.
An example:

```d
// @Dispose is used to tell the framework how to free memory that is dynamically
// allocated by the D function. After returning, the value is converted to an
// Excel type and the D value is freed using the lambda defined here.
@Dispose!((ret) {
    import std.experimental.allocator.mallocator: Mallocator;
    import std.experimental.allocator: dispose;
    Mallocator.instance.dispose(ret);
})
double[] FuncReturnArrayNoGc(double[] numbers) @nogc @safe nothrow {
    import std.experimental.allocator.mallocator: Mallocator;
    import std.experimental.allocator: makeArray;
    import std.algorithm: map;

    try {
        // Allocate memory here in order to return an array of doubles.
        // The memory will be freed after the call by calling the
        // function in `@Dispose` above
        return Mallocator.instance.makeArray(numbers.map!(a => a * 2));
    } catch(Exception _) {
        return [];
    }
}
```

This allows for `@nogc` functions to be called from Excel without memory leaks.


Registering code to run when the XLL is unloaded
------------------------------------------------

Since this library automatically writes `xlAutoClose` it is not possible to use it to
run custom code at XLL unloading. As an alternative XLL writers can use
`xlld.xll.registerAutoCloseFunc` passing it a function or a delegate to be executed
when `xlAutoClose` is called.


About Kaleidic Associates
-------------------------
We are a boutique consultancy that advises a small number of hedge fund clients.  We are
not accepting new clients currently, but if you are interested in working either remotely
or locally in London or Hong Kong, and if you are a talented hacker with a moral compass
who aspires to excellence then feel free to drop me a line: laeeth at kaleidic.io

We work with our partner Symmetry Investments, and some background on the firm can be
found here:

http://symmetryinvestments.com/about-us/
