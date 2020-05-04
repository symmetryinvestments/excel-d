module ut.wrap.wrapped_no_uppercase;

import std.typecons: No;

import xlld.wrap;
import xlld.wrap.traits: Async;

mixin(wrapAll!"test.d_funcs"(No.onlyExports, No.pascalCase));
