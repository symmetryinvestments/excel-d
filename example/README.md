# Example XLL

The dub package in this directory will build an XLL that Excel can load.
The functions exported aren't very useful can be found in
[excel-d's source tree](../source/xlld/test_d_funcs).

Either copy the appropriate (32/64 bit) `xlcall32.lib` from the Excel SDK
in this directory to build (then either `dub build --arch=x86_mscoff` or `dub build --arch=x86_64`)
or make sure it's somewhere that `link.exe` can find.
