import unit_threaded;

int main(string[] args)
{
    return args.runTests!(
        "xlld.any",
        "xlld.memorymanager",
        "xlld.wrap.worksheet",
        "xlld.wrap.traits",
        "xlld.wrap.wrap",
        "xlld.sdk.xll",
        "xlld.test.util",
        "xlld.func.xl",
        "xlld.func.xlf",
        "xlld.conv.misc",
        "wrap.module_",
        "wrap.all",
    );
}
