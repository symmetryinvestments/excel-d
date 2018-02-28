import reggae;
import std.typecons;

alias ut = dubTestTarget!(CompilerFlags("-g -debug"),
                          LinkerFlags(),
                          No.allTogether);
alias utl = dubConfigurationTarget!(Configuration("ut"),
                                    CompilerFlags("-unittest -version=unitThreadedLight -g -debug"),
                                    LinkerFlags(),
                                    Yes.main,
                                    No.allTogether);
mixin build!(ut, utl.optional);
