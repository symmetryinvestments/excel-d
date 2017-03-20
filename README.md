# excel-d

Excel API bindings and wrapper API for D

This dub package contains D declarations for the [Excel SDK](https://msdn.microsoft.com/en-us/library/office/bb687883.aspx)
 as well as a D wrapper API. This allows programmers to write Excel functions in D.

A working XLL example can be found in the [`example`](example) directory. Compiling this file will result in
an XLL that can be loaded in Excel making all of the functions in `test/xlld/test_d_funcs.d` available
to be used in Excel cells. The types are automatically converted between D native types and Excel ones.
To build the example: `dub build -c example`.

For this package to build you will need the Excel SDK `xlcall32.lib`
that can be downloaded [from Microsoft](http://go.microsoft.com/fwlink/?LinkID=251082&clcid=0x409).
Copying it to the repository's top directory should be sufficient.

Excel won't load the XLL automatically: this must be done manually in File->Tools->Add-Ins.
Click on "Go" for "Excel Add-Ins" (the default) and select your XLL there.

Currently only tested with 32-bit Excel 2013 on 64-bit Windows.
