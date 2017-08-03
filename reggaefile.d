import reggae;
import std.typecons;

alias ut = dubTestTarget!(Flags("-g -debug"));
alias utl = dubConfigurationTarget!(Configuration("ut"),
                                    Flags("-unittest -version=unitThreadedLight -g -debug"));
mixin build!(ut, utl);
