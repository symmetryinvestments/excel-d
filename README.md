# excel-d

Excel API bindings and wrapper API for D

This dub package contains D declarations for the [Excel SDK](https://msdn.microsoft.com/en-us/library/office/bb687883.aspx)
 as well as a D wrapper API. This allows programmers to write Excel worksheet functions in D.

Generated documentation is available at [Kaleidic Open Source Documentation for Excel-D](http://excel-d.code.kaleidic.io).

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

Excel won't load the XLL automatically: this must be done manually in File->Tools->Add-Ins.
Click on "Go" for "Excel Add-Ins" (the default) and select your XLL there after clicking on
"Browse".

The only difference between building for 32-bit or 64-bit Excel is the `arch=` option passed
to dub. A 32-bit XLL will only work on 32-bit Excel and similarly for 64-bit. You will also
need the appropriate 32/64 xlcall32.lib from the Excel SDK to link.

Sample code (see the [example](example) directory for more):

```d
    import xlld;

    @Register(ArgumentText("Array to add"),
              HelpTopic("Adds all cells in an array"),
              FunctionHelp("Adds all cells in an array"),
              ArgumentHelp(["The array to add"]))
    double FuncAddEverything(double[][] args) nothrow @nogc { // must be nothrow, @nogc optional
        import std.algorithm: fold;
        import std.math: isNaN;

        double ret = 0;
        foreach(row; args)
            ret += row.fold!((a, b) => b.isNaN ? 0.0 : a + b)(0.0);
        return ret;
    }
```d

and then in Excel:

`=FuncAddEverything(A1:D20)`

Future functionality will include creating menu items and dialogue boxes.  Pull requests welcomed.


Memory allocation and `@nogc`
-----------------------------

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
