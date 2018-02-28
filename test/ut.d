import unit_threaded;

int main(string[] args)
{
    return args.runTests!(
        "xlld.worksheet",
        "xlld.traits",
        "xlld.wrap",
        "xlld.xl",
        "xlld.sdk.xll",
        "xlld.memorymanager",
        "xlld.test.util",
        "xlld.xlf",
        "xlld.conv.misc",
        "xlld.any",
    );
}
