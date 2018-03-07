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
        "ut.wrap.module_",
        "ut.wrap.all",
        "ut.conv.from",
        "ut.conv.to",
    );
}
