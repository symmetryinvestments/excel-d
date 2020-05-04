/**
   This module exists so that the instantiation of the wrapper functions
   for Excel only happens once. From LDC 1.9.0, two `extern(Windows)` functions
   with the same name can't exist since they mangle the same.

   Instead we create them here.
 */
module ut.wrap.wrapped;

import xlld.wrap;
import xlld.wrap.traits: Async;
mixin(wrapAll!"test.d_funcs");
