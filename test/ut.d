import unit_threaded;

int main(string[] args)
{
    return args.runTests!(
        "xlld.worksheet",
        "xlld.traits",
        "xlld.wrap",
        "xlld.any",
        "xlld.sdk.xll",
        "xlld.memorymanager",
        "xlld.test.util",
        "xlld.func.xl",
        "xlld.func.xlf",
        "xlld.conv.misc",
    );
}
