import reggae;
import std.typecons;

alias ut = dubTestTarget!(CompilerFlags("-g -debug"));
alias utl = dubConfigurationTarget!(Configuration("ut"),
                                    CompilerFlags("-unittest -version=unitThreadedLight -g -debug"));
mixin build!(ut, utl.optional);
