# Example XLL

The dub package in this directory will build an XLL that Excel can load.
The functions exported aren't very useful can be found in
[excel-d's source tree](../source/xlld/test_d_funcs).

Remember to build with `dub build --arch=x86_mscoff` for 32-bits in order
to be able to link with Microsoft's Excel SDK `xlcall32.lib`.
