import reggae;
import std.typecons;

alias ut = dubTestTarget!(CompilerFlags("-g -debug"),
                          LinkerFlags(),
                          CompilationMode.package_);
mixin build!(ut);
