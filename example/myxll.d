module example.myxll;

import xlld: wrapAll;

// wrapAll takes a list of modules, in this case we're only wrapping
// one, but there's no limit
// The current file will then contain all the code to correctly
// start an XLL with wrapped Excel functions of all D functions
// in the listed modules
mixin(wrapAll!("d_funcs"));
