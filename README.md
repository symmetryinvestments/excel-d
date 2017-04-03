# excel-d

Excel API bindings and wrapper API for D

This dub package contains D declarations for the [Excel SDK](https://msdn.microsoft.com/en-us/library/office/bb687883.aspx)
 as well as a D wrapper API. This allows programmers to write Excel functions in D.

A working XLL example can be found in the [`example`](example)
directory. Running `dub build` there will create an XLL
(`myxll32.xll`) that can be loaded in Excel making all of the
functions in `test/xlld/test_d_funcs.d` available to be used in Excel
cells. The types are automatically converted between D native types
and Excel ones.  To build the example: `dub build -c example`.

For this package to build you will need the Excel SDK `xlcall32.lib`
that can be downloaded [from Microsoft](http://go.microsoft.com/fwlink/?LinkID=251082&clcid=0x409).
Copying it to the build directory should be sufficient
(i.e. when building the example, to the `example` directory).
The library file should be useable as-is, as long as `dub build` is run with `--arch=x86_mscoff`
to use Microsoft's binary format. If linking with optlink, the file must be converted first.
We recommed using `link.exe` to not need the conversion.

Excel won't load the XLL automatically: this must be done manually in File->Tools->Add-Ins.
Click on "Go" for "Excel Add-Ins" (the default) and select your XLL there after clicking on
"Browse".

Currently only tested with 32-bit Excel 2013 on 64-bit Windows.  See add64bit branch for untested 64 bit
compatible version.
