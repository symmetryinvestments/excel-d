import reggae;
import std.typecons;

alias ut = dubTestTarget!(CompilerFlags("-g -debug"),
                          LinkerFlags(),
                          CompilationMode.package_);
alias utl = dubConfigurationTarget!(Configuration("utl"),
                                    CompilerFlags("-unittest -version=unitThreadedLight -g -debug"),
                                    LinkerFlags(),
                                    Yes.main,
                                    CompilationMode.package_);
mixin build!(ut, utl.optional);
