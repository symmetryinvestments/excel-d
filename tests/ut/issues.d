module ut.issues;

import test;
import ut.wrap.wrapped;

@("58")
unittest {
    foreach(i; 0 .. 100)
        Leaker;
}
