/**
   Utility to avoid top-level imports
 */
module xlld.from;

/**
   Local imports everywhere.
 */
template from(string moduleName) {
    mixin("import from = " ~ moduleName ~ ";");
}
