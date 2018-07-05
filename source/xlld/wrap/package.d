/**
   Autowrapping of D functions to be called from Excel. Generally this will be enough:

   ------
   import xlld;
   mixin wrapAll!("mymodule1", "mymodule2", ...);

   All public functions in the modules in the call to `wrapAll` will then be available
   to be called from Excel.
   ------
 */
module xlld.wrap;

public import xlld.wrap.wrap;
public import xlld.wrap.traits;
public import xlld.wrap.worksheet;
