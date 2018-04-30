module ut.misc;

import test;
import xlld.any;
import xlld.sdk.xll;
import unit_threaded: DontTest;

@("opEquals str")
unittest {
    any("foo", theGC).shouldEqual(any("foo", theGC));
    any("foo", theGC).shouldNotEqual(any("bar", theGC));
    any("foo", theGC).shouldNotEqual(any(33.3, theGC));
}

@("opEquals multi")
unittest {
    any([1.0, 2.0], theGC).shouldEqual(any([1.0, 2.0], theGC));
    any([1.0, 2.0], theGC).shouldNotEqual(any("foos", theGC));
    any([1.0, 2.0], theGC).shouldNotEqual(any([2.0, 2.0], theGC));
}


@("registerAutoClose delegate")
unittest {
    int i;
    registerAutoCloseFunc({ ++i; });
    callRegisteredAutoCloseFuncs();
    i.shouldEqual(1);
}

@("registerAutoClose function")
unittest {
    const old = gAutoCloseCounter;
    registerAutoCloseFunc(&testAutoCloseFunc);
    callRegisteredAutoCloseFuncs();
    (gAutoCloseCounter - old).shouldEqual(1);
}


int gAutoCloseCounter;


@DontTest
void testAutoCloseFunc() nothrow {
    ++gAutoCloseCounter;
}
