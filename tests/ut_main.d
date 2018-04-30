import unit_threaded;

int main(string[] args)
{
    return args.runTests!(
        "xlld.any",
        "xlld.wrap.wrap",
        "xlld.wrap.traits",
        "xlld.sdk.xll",
        "xlld.test.util",
        "ut.wrap.module_",
        "ut.wrap.all",
        "ut.conv.from",
        "ut.conv.to",
        "ut.conv.misc",
        "ut.func.xlf",
        "ut.wrap.wrap",
        "ut.wrap.traits",
        "ut.misc",
    );
}
