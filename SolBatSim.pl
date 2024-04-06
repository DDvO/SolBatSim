#!/usr/bin/perl
################################################################################
# Eigenverbrauchs-Simulation mit stündlichen PV-Daten und einem mindestens
# stündlichen Lastprofil, optional mit Begrenzung/Drosselung des Wechselrichters
# und optional mit Stromspeicher (Batterie o.ä.)
#
# Nutzung: SolBatSim.pl <Lastprofil-Datei> [<Jahresverbrauch in kWh>]
#   (<PV-Datei> [direct] [<Nennleistung Wp> [<WR-Eingangsbegrenzung in W>]])+
#   [-only <*|Jahr[..Jahr]>[-<*|Mon[..Mon]>[-<*|Tag[..Tag]>[:<*|Std[..Std]]]]]
#   [-dist <relative Lastverteilung über den Tag pro Stunde 0,..,23>
#   [-bend <Lastverzerrungsfaktoren tgl. pro Std. 0,..,23, sonst 1>
#   [-load [min] <konstante Last, auf gleicher Skala wie in der PV-Daten-Datei>
#          [<Zahl der Tage pro Woche ab Montag, sonst 5>:<von Uhrzeit,
#           sonst 8 Uhr>..<bis Uhrzeit, sonst 16 Uhr, auch über Mitternacht>]]
#   [-avg_hour] [-dump_avg <PVTool-Verbrauchsdatei>] [-verbose]
#   [-peff <PV-System-Wirkungsgrad in %, ansonsten von PV-Daten-Datei>]
#   [-ac | -dc] [-capacity <Speicherkapazität Wh, ansonsten 0 (kein Batterie)>]
#   [-pass [spill] <konstante Speicher-Umgehung in W zusätzlich zu 'direct' PV,
#                   mit 'spill' oder AC-Kopplung auch bei vollem Speicher>]
#   [-feed ( lim <Begrenzung der lastoptimierten Entladung aus Speicher in W>
#         | comp <Einspeiseziel in W bis zu dem der Speicher die PV kompensiert>
#         | excl <Grenzwert in W zur Speicherentladung statt PV-Nutzung>
#         | <konstante Entladung aus Speicher in W> )
#         [<von Uhrzeit, sonst 0 Uhr>..
#          <bis Uhrzeit, sonst 24 Uhr, auch über Mitternacht>]]
#   [-max_charge <Ladehöhe in %, sonst 90> [<max Laderate, sonst 1 C>]]
#   [-max_discharge <Entladetiefe in %, sonst 90> [<Rate, sonst 1 C>]]
#   [-ceff <Lade-Wirkungsgrad in %, ansonsten 94>]
#   [-seff <Speicher-Wirkungsgrad in %, ansonsten 95>]
#   [-ieff <Wechselrichter-Wirkungsgrad in %, ansonsten 94>]
#   [-ieff2 <Wirkungsgrad des Entlade-Wechselrichters in %, Standard von -ieff>]
#   [-debug] [-test <Lastpunkte pro Stunde, für Prüfberechnung über 24 Stunden>]
#   [-en] [-tmy] [-curb <Wechselrichter-Ausgangs-Drosselung in W>]
#   [-hour <Statistik-Datei>] [-day <Stat.Datei>] [-week <Stat.Datei>]
#   [-month <Stat.Datei>] [-season <Stat.Datei>] [-max <Stat.Datei>]
#
# Alle Uhrzeiten sind in lokaler Winterzeit (MEZ, GMT/UTC + 1 ohne Sommerzeit).
# Mit "-en" erfolgen die Textausgaben auf Englisch. Fehlertexte sind englisch.
#
# Mit "-avg_hour" wird nur der stündliche Durchschnitt der Lastdaten verwendet.
# Mit "-dump_avg", was "-avg_hour" impliziert, wird eine Last-Eingabe-Datei fürs
# PVTool im CSV-Format erzeugt mit UTC-Zeitangaben und stündlichen Lastdaten.
#
# Wenn PV-Daten für mehrere Jahre gegeben sind, wird der Durchschnitt berechnet
# oder mit Option "-tmy" Monate für ein typisches meteorologisches Jahr gewählt.
#
# Beim Speicher ist AC-Kopplung Standard. Dabei Verluste durch zweimal WR, aber
# kein Überlauf. DC-gekoppelte Ladung umgeht den ersten WR und seine Limits.
#
# Mit den Optionen "-hour"/"-day"/"-week"/"-month" wird jeweils eine CSV-Datei
# mit gegebenen Namen mit Statistik-Daten pro Stunde/Tag/Woche/Monat erzeugt.
#
# Beispiel:
# SolBatSim.pl load_4673_kWh.csv 3000 yield_1274_kWh.csv 1000 -curb 600 -tmy
#
# PV-Daten können bezogen werden von https://re.jrc.ec.europa.eu/pvg_tools/de/
# Wähle den Standort und "DATEN PRO STUNDE", setze Häkchen bei "PV-Leistung".
# Optional "Installierte maximale PV-Leistung" und "Systemverlust" anpassen.
# Bei Nutzung von "-tmy" Startjahr 2008 oder früher, Endjahr 2020 oder später.
# Dann den Download-Knopf "csv" drücken.
#
################################################################################
# Simulation of actual own consumption of photovoltaic power output according
# to load profiles with a resolution of at least one hour, typically per minute.
# Optionally takes into account input limit and output crop of solar inverter.
# Optionally with energy storage (using a battery or the like).
#
# Usage: SolBatSim.pl <load profile file> [<consumption per year in kWh>]
#   (<PV file> [direct] [<nominal power in Wp> [<inverter input limit in W]])+
#   [-only <*|year[..year]>[-<*|mon[..mon]>[-<*|day[..day]>[:<*|hour[..hour]]]]]
#   [-dist <relative load distribution over each day, per hour 0,..,23>
#   [-bend <load distort factors for hour 0,..,23 each day, default 1>
#   [-load [min] <constant load, at same scale as in PV data file>
#          [<count of days per week starting Monday, default 5>:<from hour,
#           default 8 o'clock>..<to hour, default 16, also across midnight>]]
#   [-avg_hour] [-dump_avg <PVTool load file>] [-verbose]
#   [-peff <PV system efficiency in %, default from PV data file(s)>]
#   [-ac | -dc] [-capacity <storage capacity in Wh, default 0 (no battery)>]
#   [-pass [spill] <constant storage bypass in W in addition to 'direct PV',
#                   with 'spill' or AC coupling also when storage is full>]
#   [-feed ( lim <limit of optimal feed (according to load) from storage in W >
#         | comp <target rate in W up to which storage compensates PV output>
#         | excl <threshold in W for discharging storage instead of using PV>
#         | <constant feed-in from storage in W> )
#         [<from hour, default 0>..<to hour, default 24, also over midnight>]]
#   [-max_charge <SoC in %, default 90> [<max charge rate, default 1 C>]]
#   [-max_discharge <DoD in %, default 90> [<max rate, default 1 C>]]
#   [-ceff <charging efficiency in %, default 94>]
#   [-seff <storage efficiency in %, default 95>]
#   [-ieff <inverter efficiency in %, default 94>]
#   [-ieff2 <efficiency of discharge inverter in %, default from -ieff>]
#   [-debug] [-test <load points per hour, for using test data over 24 hours>]
#   [-en] [-tmy] [-curb <inverter output power limitation in W>]
#   [-hour <statistics file>] [-day <stat file>] [-week <stat file>]
#   [-month <stat file>] [-season <file>] [-max <stat file>]
#
# All times (hours) are in local winter time (CET, GMT/UTC +1, no daylight sv.).
# Use "-en" for text output in English. Error messages are all in English.
#
# With "-avg_hour", only the average of load items per hour are used.
# With `-dump_avg`, which implies `-avg_hour`, produce PVTool load input file
# in CSV format with UTC timestamps and hourly average load data.
#
# When PV data for more than one year is given, the average is computed, while
# with the option "-tmy" months for a typical meteorological year are selected.
#
# With storage, AC-coupled charging is the default. It has extra inverter loss,
# but no spill loss. DC-coupled charging bypasses first inverter and its limits.
#
# With each the options "-hour"/"-day"/"-week"/"-month" a CSV file is produced
# with the given name containing with statistical data per hour/day/week/month.
#
# Example:
# SolBatSim.pl load_4673_kWh.csv 3000 yield_1274_kWh.csv 1000 -curb 600 -tmy -en
#
# PV data can be obtained from https://re.jrc.ec.europa.eu/pvg_tools/
# Select location, click "HOURLY DATA", and set the check mark at "PV power".
# May adapt "Installed peak PV power" and "System loss" already here.
# For using TMY data, choose Start year 2008 or earlier, End year 2020 or later.
# Then press the download button marked "csv".
#
# (c) 2022-2023 David von Oheimb - License: MIT - Version 2.3
################################################################################

use strict;
use warnings;

# deliberately not using any extra packages like Math
sub min { return !defined $_[0] || $_[0] > $_[1] ? $_[1] : $_[0]; }
sub max { return !defined $_[0] || $_[0] < $_[1] ? $_[1] : $_[0]; }
sub round { return int(($_[0] < 0 ? -.5 : .5) + $_[0]); }

my $test         = 0;   # unless 0, number of test load points per hour
my $debug        = 0;   # turn on debug output, implied by $test != 0
die "Missing command line arguments" if $#ARGV < 0;
my $load_profile = shift @ARGV unless $ARGV[0] =~ m/^-/; # file name
my $consumption  = shift @ARGV # kWh/year, default is implicit from load profile
    if $#ARGV >= 0 && $ARGV[0] =~ m/^[\d\.]+$/;

my @load_dist;          # if set, relative load distribution per hour each day
my @load_factors;       # load distortion factors per hour, on top of @load_dist
my $load_min      = 0;  # modifies $load_const to mean constant minimal load
my $load_const;         # constant load in W, during certain times as follows:
my $avg_hour      = 0;  # use only the average of load items per hour
my ($DA, $dump_avg);    # file to dump average load per hour, implies $avg_hour
my $verbose       = 0;  # verbose output, including averages/day for each hour
my $load_days    =  5;  # count of days per week with constant load
my $load_from    =  8;  # hour of constant load begin
my $load_to      = 16;  # hour of constant load end
my $first_weekday = 4;  # HTW load profiles start on Friday in 2010, Monday == 0

use constant YearHours => 24 * 365;
use constant TimeZone => 1; # CET/MEZ

my @PV_files;         # PV data input file(s), one per string
my @PV_direct;        # with storage, bypass it for respective PV output string
my @PV_nomin;         # nominal PV output(s), default from PV data file(s)
my @PV_limit;         # power limit at inverter input, default '*' (none)
my ($lat, $lon);      # optional, from PV data file(s)
my $pvsys_eff;        # PV system efficiency, default from PV data file(s)
my $inverter_eff;     # PV inverter efficiency; default see below
my $inverter2_eff;    # discharging inverter efficiency; default from above
my $capacity;         # nominal storage capacity in Wh on average degradation
my $bypass;           # direct const feed-in to inverter in W, bypassing storage
my $bypass_spill;     # bypass storage when it is full, implied by $AC_coupled
my $max_feed;         # maximal feed-in in W from storage
my $comp_feed = 0;    # $max_feed target power for storage compensating PV out,
                      # relevant only if defined $max_feed
my $excl_feed = 0;    # $max_feed is threshold in W for discharging storage
                      # instead of using PV, relevant only if defined $bypass
my $const_feed  =  0; # constant feed-in, relevant only if defined $max_feed
my $feed_from   =  0; # hour of constant feed-in begin
my $feed_to     = 24; # hour of constant feed-in end
my $AC_coupled  = -1; # by default, charging is AC-coupled (via inverter)
my $max_soc     = .9; # maximal state of charge (SoC); default 90%
my $max_dod     = .9; # maximal depth of discharge (DoD); default 90%
my $max_chgrate =  1; # maximal charge rate; default 1 C
my $max_disrate =  1; # maximal discharge rate; default 1 C
my $charge_eff;       # charge efficiency; default see below
my $storage_eff;      # storage efficiency; default see below

while ($#ARGV >= 0 && $ARGV[0] =~ m/^\s*[^-]/) {
    push @PV_files, shift @ARGV; # PV data file
    push @PV_direct,$#ARGV >= 0 && $ARGV[0] =~ m/^direct$/  ? shift @ARGV : 0;
    push @PV_nomin, $#ARGV >= 0 && $ARGV[0] =~ m/^[\d\.]+$/ ? shift @ARGV : undef;
    push @PV_limit, $#ARGV >= 0 && $ARGV[0] =~ m/^[\d\.]+$/ ? shift @ARGV : '*';
}

sub no_arg {
    shift @ARGV;
    return 1;
}
sub num_arg {
    my $opt = $ARGV[0];
    die "Missing number argument for $opt option"
        unless $#ARGV >= 1 && $ARGV[1] =~ m/^-?[\d\.]+$/ && $ARGV[1] ne ".";
    shift @ARGV;
    my $arg = shift @ARGV;
    die "Numeric argument $arg for $opt option is negative"
        if $arg < 0;
    return $arg;
}
sub eff_arg {
    my $opt = $ARGV[0];
    my $eff = num_arg();
    die "Percentage argument $eff for $opt option is out of range 0..100"
        unless 0 <= round($eff) && round($eff) <= 100;
    return $eff / 100;
}
sub str_arg {
    die "Missing arg for $ARGV[0] option" if $#ARGV < 1;
    shift @ARGV;
    return shift @ARGV;
}
sub array_arg { my ($opt, $arg, $min, $max, $default) = @_;
    $arg =~ tr/ /,/ if defined $arg;

    my @result;
    my @items = defined $arg ? split ",", $arg : ();
    my $n = $#items;
    my $j = 0;
    for (my $i = $min; $i <= $max; $i++) {
        my $val = $default;
        if ($j <= $n) {
            $items[$j] =~ m/^\s*((\d+)\s*:)?(.*)$/;
            if (defined $2) {
                die "Index $2 in -$opt option argument < $i" if $2 < $i;
                die "Index $2 in -$opt option argument > $max" if $2 > $max;
            }
            if (!defined $2 || $2 == $i) {
                $val = $3;
                $j++;
            }
            die "Item value '$val' of -$opt option argument is not a number"
                unless $val =~ m/^\s*-?[\d\.]+\s*$/ && $val ne ".";
        }
        $result[$i] = $val;
    }
    die (($n + 1 - $j)." extra items in -$opt option argument") if $j <= $n;
    return @result;
}

my $load_dist;
my $load_dist_sum = 0;
my $load_factors;
my ($en, $only, $tmy, $curb,
    $max, $hourly, $daily, $weekly,  $monthly, $seasonly);
while ($#ARGV >= 0) {
    if      ($ARGV[0] eq "-test"    ) { $test         = num_arg();
    } elsif ($ARGV[0] eq "-debug"   ) { $debug        =  no_arg();
    } elsif ($ARGV[0] eq "-en"      ) { $en           =  no_arg();
    } elsif ($ARGV[0] eq "-only"    ) { $only         = str_arg();
    } elsif ($ARGV[0] eq "-dist"    ) { $load_dist    = str_arg();
    } elsif ($ARGV[0] eq "-bend"    ) { $load_factors = str_arg();
    } elsif ($ARGV[0] eq "-load"    ) { $load_min     = 1 if $#ARGV >= 1 &&
                                            $ARGV[1] eq "min" && shift @ARGV;
                                        $load_const   = num_arg();
                                       ($load_days, $load_from, $load_to)
                                           = ($1, $2, $3) if $#ARGV >= 0 &&
                                           $ARGV[0] =~ m/^(\d+):(\d+)\.\.(\d+)$/
                                           && shift @ARGV;
    } elsif ($ARGV[0] eq "-avg_hour") { $avg_hour     =  no_arg();
    } elsif ($ARGV[0] eq "-dump_avg") { $dump_avg     = str_arg(); $avg_hour=1;
    } elsif ($ARGV[0] eq "-verbose" ) { $verbose      =  no_arg();
    } elsif ($ARGV[0] eq "-peff"    ) { $pvsys_eff    = eff_arg();
    } elsif ($ARGV[0] eq "-tmy"     ) { $tmy          =  no_arg();
    } elsif ($ARGV[0] eq "-curb"    ) { $curb         = num_arg();
    } elsif ($ARGV[0] eq "-capacity") { $capacity     = num_arg();
    } elsif ($ARGV[0] eq "-ac"      ) { $AC_coupled   =  no_arg();
    } elsif ($ARGV[0] eq "-dc"      ) { $AC_coupled   =  no_arg() * 0;
    } elsif ($ARGV[0] eq "-pass"    ) { $bypass_spill = 1 if $#ARGV >= 1 &&
                                            $ARGV[1] eq "spill" && shift @ARGV;
                                        $bypass       = num_arg();
    } elsif ($ARGV[0] eq "-feed"    ) { my $mod = $ARGV[1];
                                        if ($#ARGV >= 1 && ($mod eq "lim" ||
                                                            $mod eq "comp" ||
                                                            $mod eq "excl")) {
                                            $comp_feed = 1 if $mod eq "comp";
                                            $excl_feed = 1 if $mod eq "excl";
                                            shift @ARGV;
                                        } else {
                                            $const_feed = 1;
                                        }
                                        $max_feed     = num_arg();
                                        ($feed_from, $feed_to) = ($1, $2)
                                            if $#ARGV >= 0
                                            && $ARGV[0] =~ m/^(\d+)\.\.(\d+)$/
                                            && shift @ARGV;
    } elsif ($ARGV[0] eq "-max_charge") { $max_soc    = eff_arg();
                                        $max_chgrate =shift @ARGV if $#ARGV >= 0
                                            && $ARGV[0] =~ m/^([\d\.]+)$/;
    } elsif ($ARGV[0] eq "-max_discharge") { $max_dod = eff_arg();
                                        $max_disrate =shift @ARGV if $#ARGV >= 0
                                            && $ARGV[0] =~ m/^([\d\.]+)$/;
    } elsif ($ARGV[0] eq "-ceff"    ) { $charge_eff   = eff_arg();
    } elsif ($ARGV[0] eq "-seff"    ) { $storage_eff  = eff_arg();
    } elsif ($ARGV[0] eq "-ieff"    ) { $inverter_eff = eff_arg();
    } elsif ($ARGV[0] eq "-ieff2"   ) { $inverter2_eff= eff_arg();
    } elsif ($ARGV[0] eq "-max"     ) { $max          = str_arg();
    } elsif ($ARGV[0] eq "-hour"    ) { $hourly       = str_arg();
    } elsif ($ARGV[0] eq "-day"     ) { $daily        = str_arg();
    } elsif ($ARGV[0] eq "-week"    ) { $weekly       = str_arg();
    } elsif ($ARGV[0] eq "-month"   ) { $monthly      = str_arg();
    } elsif ($ARGV[0] eq "-season"  ) { $seasonly     = str_arg();
    } else { die "Invalid option: $ARGV[0]";
    }
}

use constant TEST_START      =>  6; # load starts in the morning
use constant TEST_LENGTH     => 24; # until 6 AM on the next day
use constant TEST_END        => TEST_START + TEST_LENGTH;
use constant TEST_DROP_START => 10; # start hour of load drop
use constant TEST_DROP_LEN   =>  4; # length of load drop around noon
use constant TEST_DROP_END   => TEST_DROP_START + TEST_DROP_LEN;
use constant TEST_PV_START   =>  7; # start hour of PV yield
use constant TEST_PV_END     => 17; # end   hour of PV yield
use constant TEST_PV_NOMIN   => 1100; # in Wp
use constant TEST_PV_LIMIT   =>  '*'; # in W
use constant TEST_LOAD       =>  900; # in W
use constant TEST_LOAD_LEN   =>
    TEST_END > TEST_DROP_END ? TEST_LENGTH - TEST_DROP_LEN
    : (TEST_END <= TEST_DROP_START ? TEST_END - TEST_START
       : TEST_DROP_START - TEST_START);
my $test_drop_const = defined $load_const && $load_from == TEST_DROP_START
    && $load_to == TEST_DROP_END ? $load_const * TEST_DROP_LEN : 0;
my $test_load = defined $consumption ?
    ($consumption * 1000 - $test_drop_const) / TEST_LOAD_LEN : TEST_LOAD
    if $test;
if ($test) {
    $debug         = 1;
    $load_profile  = "test load data";
    my $pv_nomin   = $#PV_nomin < 0 ? TEST_PV_NOMIN : $PV_nomin[0]; # in W
    my $pv_limit   = $#PV_limit < 0 ? TEST_PV_LIMIT : $PV_limit[0]; # in W
    my $pv_power   = defined $pvsys_eff ? $pv_nomin * $pvsys_eff :
        sprintf("%.0e", $pv_nomin); # in W
    # after PV system losses, which may be derived as follows:
    $pvsys_eff     = $pv_nomin ? $pv_power / $pv_nomin : 0.92
        unless defined $pvsys_eff;
    $inverter_eff  =  0.8 unless defined $inverter_eff;
    # $consumption = $test_load/1000 * TEST_LOAD_LEN unless defined $consumption;
    push @PV_files, "test PV data" if $#PV_files < 0;
    push @PV_direct, 0             if $#PV_nomin < 0;
    push @PV_nomin, $pv_nomin      if $#PV_nomin < 0;
    push @PV_limit, $pv_limit      if $#PV_limit < 0;
    $charge_eff    =  0.9 if defined $capacity && !defined $charge_eff;
    my $pv_net_bat =  600; # in W, after inverter losses
    $storage_eff = $pv_net_bat / never_0($pv_power)
        / never_0($charge_eff) / never_0($inverter_eff)
        if defined $capacity && !defined $storage_eff;
}

die "Missing PV data file name" if $#PV_nomin < 0;
die "-only option argument does not have form (*|YYYY[..YYYY])[-(*|MM[..MM])".
    "[-(*|DD[..DD])[:(*|HH[..HH])]]]"
    if $only &&
    !($only =~ m/^(\*|\d{4})(\.\.(\d{4}))?(-(\*|\d\d?)(\.\.(\d\d?))?
                (-(\*|\d\d?)(\.\.(\d\d?))?(:(\*|\d\d?)(\.\.(\d\d?))?)?)?)?$/x);
my ($sel_year, $sel_year2, $sel_month, $sel_month2,
    $sel_day, $sel_day2, $sel_hour, $sel_hour2) =
    ($1, $3, $5, $7, $9, $11, $13, $15) if $only;
undef $sel_year  if defined $sel_year  && $sel_year  eq "*";
undef $sel_month if defined $sel_month && $sel_month eq "*";
undef $sel_day   if defined $sel_day   && $sel_day   eq "*";
undef $sel_hour  if defined $sel_hour  && $sel_hour  eq "*";
die "With -tmy, the year given with the -only option must be '*'"
    if $tmy && defined $sel_year;
if (defined $sel_year2) {
    die "Second year value in the -only option must not be given ".
        "as the first one is '*'" if !defined $sel_year;
    die "Second year $sel_year2 in the -only option must be at least $sel_year"
        if $sel_year2 < $sel_year;
}
sub check_range { my ($desc, $val, $start, $end) = @_;
    die "Second $desc value in the -only option must not be given ".
        "as the first one is '*'" if defined $val && !defined $start;
    die "$desc $val given with -only option is out of range $start..$end"
        if defined $val &&
        ($val < (!defined $start ? ($end == 23 ? 0 : 1) : $start)
         || $val > $end);
}
my @days_per_month = (0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
my $max_day = defined $sel_month
    && (!defined $sel_month2 || $sel_month2 == $sel_month)
    ? $days_per_month[$sel_month] : 31;
check_range("Month", $sel_month , 1         ,       12);
check_range("month", $sel_month2, $sel_month,       12);
check_range("Day"  , $sel_day   , 1         , $max_day);
check_range("day"  , $sel_day2  , $sel_day  , $max_day);
check_range("Hour" , $sel_hour  , 0         ,       23);
check_range("hour" , $sel_hour2 , $sel_hour ,       23);
$sel_year2  = $sel_year  unless defined $sel_year2;
$sel_month2 = $sel_month unless defined $sel_month2;
$sel_day2   = $sel_day   unless defined $sel_day2;
$sel_hour2  = $sel_hour  unless defined $sel_hour2;

die "Missing load profile file name - should be first CLI argument"
    unless defined $load_profile;
@load_dist    = array_arg("dist", $load_dist   , 0, 23, 100);
@load_factors = array_arg("bend", $load_factors, 0, 23, 1);
for (my $i = 0; $i < 24; $i++) {
    $load_factors[$i] *= 100; # for printing them as percent values
}

if (defined $load_dist) {
    for (my $i = 0; $i < 24; $i++) {
        $load_dist_sum += $load_dist[$i];
    }
    die "Sum of -dist argument items is 0" if $load_dist_sum == 0;
}
if (defined $load_const) {
    die "Days count for -load option must be in range 0..7"  if $load_days > 7;
    die "Begin hour for -load option must be in range 0..24" if $load_from > 24;
    die "End hour for -load option must be in range 0..24"   if $load_to > 24;
}
$inverter_eff = 0.94    unless defined $inverter_eff;
if (defined $capacity) {
    $charge_eff  = 0.94 unless defined $charge_eff;
    $storage_eff = 0.95 unless defined $storage_eff;
    $inverter2_eff = $inverter_eff unless defined $inverter2_eff;
    die "Begin hour for -feed option must be in range 0..24"
        if  $feed_from > 24;
    die "End hour for -feed option must be in range 0..24"
        if  $feed_to > 24;
    die "-feed excl requires -bypass option"
        if defined $max_feed && $excl_feed && !defined $bypass;
    die "-feed excl requires -dc option"
        if defined $max_feed && $excl_feed && $AC_coupled;
} else {
    die   "-ac or -dc option requires -capacity option" if $AC_coupled >= 0;
    die "-pass option requires -capacity option" if defined $bypass;
    die "-feed option requires -capacity option" if defined $max_feed;
    die "-ceff option requires -capacity option" if defined $charge_eff;
    die "-seff option requires -capacity option" if defined $storage_eff;
    die "-ieff2 option requires -capacity option" if defined $inverter2_eff;
    $AC_coupled = 0;
}
my $DC_coupled = defined $capacity && !$AC_coupled;
$bypass_spill = 1 if $AC_coupled;

sub never_0 { return $_[0] == 0 ? 1 : $_[0]; }
my $pvsys_eff_never_0;
my  $inverter_eff_never_0 = never_0( $inverter_eff);
my    $charge_eff_never_0 = never_0(   $charge_eff) if defined $charge_eff;
my   $storage_eff_never_0 = never_0(  $storage_eff) if defined $storage_eff;
my $inverter2_eff_never_0 = never_0($inverter2_eff) if defined $storage_eff;
my $max_feed_scaled_by_eff = $max_feed /
    ($storage_eff_never_0 * $inverter2_eff_never_0) if defined $max_feed;
my $excl_threshold = $max_feed if $excl_feed;

sub check_consistency { my ($actual, $expected, $name, $file) = @_;
    return if $test;
    die "Got $actual $name rather than $expected from $file"
        if $actual != $expected;
}

my $none_txt    = $en ? "none"                  : "keine";
my $no_time_txt = $en ? "at no time"            : "zu keiner Zeit";

sub date_str { my ($year, $month, $day) = @_;
    $year = "*" if $year eq "0";
    my $at_txt = $en ? "at" : "am";
    # to support switching language of error output must not evaluate $en before
    return "$at_txt $year-".sprintf("%02d", $month)."-".sprintf("%02d", $day);
}
sub date_hour_str { my ($year, $month, $day, $hour) = @_;
    # to support switching language of error output must not evaluate $en before
    return date_str($year, $month, $day).
        " ".($en ? "at" : "um")." ".sprintf("%02d", $hour);
}
sub minute_str { my ($item, $items) = @_;
    my $str = sprintf(":%02d", int( 60 * $item / $items));
    $str .=   sprintf(":%02d", int((60 * $item % $items) / $items * 60))
        unless (60 % $items == 0);
    return $str
}
sub time_str {
    return date_hour_str(@_).minute_str($_[4], $_[5]);
}

sub round_1000 { return round(shift() / 1000); }
sub kWh     {
    my $val = shift;
    return sprintf("%5.2f kWh", $val / 1000) if $val != 0 && $val < 10000;
    return sprintf("%5d kWh", round_1000($val));
}
sub W       { return sprintf("%5d W", round(shift)); }
sub percent {
    my $val = shift() * 100;
    my $res = round($val);
    die "Percentage value $val is out of range 0..100"
        unless 0 <= $res && $res <= 100;
    return $res;
}
sub round_percent { return percent(shift) / 100; }
sub print_arr { my ($msg, $arr_ref, $sum, $start, $end, $inc) = @_;
    return if $sum == 0;
    print $msg;
    for (my $i = $start; $i <= $end; $i += $inc) {
        my $values = 0;
        for (my $j = 0; $j < $inc; $j++) {
            $values += $arr_ref->[$i + $j];
        }
        if ($sum > 0) {
            printf "%2d%%", percent($values / $sum);
        } else {
            printf "%3d", $values;
        }
        print $i < $end ? " " : "\n";
    }
}
sub print_arr_day { print_arr(@_, , -1, 0, 23, 1) };

sub selected { my ($m, $d, $h) = @_;
    return (!defined $sel_month || $sel_month <= $m && $m <= $sel_month2)
        && (!defined $sel_day   || $sel_day   <= $d && $d <= $sel_day2  )
        && (!defined $sel_hour  || $sel_hour  <= $h && $h <= $sel_hour2 );
}

# all hours according to local time without switching to daylight saving time
use constant NIGHT_START =>  0; # at night (with just base load)
use constant NIGHT_END   =>  6;
use constant BRIGHT_START =>  8; # daily period of possibly bright sunshine
use constant BRIGHT_END   => 16;

my $sum_items = 0;
my ($sel_items, $sel_hours) = 0; # items/hours a year selected by -only
my $base_load;    # left undefined
my $load_min_time = $no_time_txt;
my $load_max      = 0;
my $load_max_time = $no_time_txt;
my $night_sum = 0;
my $bright_sum = 0;
my ($load_sum, $sel_load_sum) = (0, 0);

my ($month, $day, $hour) = (1, 1, 0);
sub adjust_day_month {
    return if $hour % 24 != 0;
    if ($day == $days_per_month[$month]) {
        $day = 1;
        $month++;
    } else {
        $day++;
    }
}


################################################################################
# read load profile data

my $items_per_hour;
my @items_by_hour;
my @load_item;
my @load;
my @hours = ( 0,  1,  2,  3,  4,  5,  6,  7,  8,  9, 10, 11,
             12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23);
my @load_per_hour;
my @load_by_hour;
my @load_min_hour;
my @load_max_hour = (0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
my @load_by_weekday = (0,0,0,0,0,0,0);
my @load_by_month = (0,0,0,0,0,0,0,0,0,0,0,0,0); # months starting at index 1

sub get_profile {
    my @lines;

    my $hours = 0;
    my $file = shift;
    if ($test) {
        while ($hours < TEST_END) {
            my $load = $hours < TEST_START ||
                ($hours % 24 >= TEST_DROP_START &&
                 $hours % 24 <  TEST_DROP_END)
                ? 0 : $test_load;
            my $line = "date".(",$load") x $test;
            $lines[$hours++] = $line;
        }
    } else {
        open(my $IN, '<', $file)
            or die "Could not open profile file $file: $!\n";
        while (my $line = <$IN>) {
            chomp $line;
            next if $line =~ m/^\s*#/; # skip comment line
            # skip header line of PVTool individual input or similar CSV file:
            next if $line =~ m/[A-Za-z]{4,}/; # e.g., date", "time", "power"
            $lines[$hours++] = $line;
        }
        close $IN unless $test;
        # by default, assuming one line per hour

        if ($hours == 365) {
            # handle one line per day, as for German BDEW profiles
            my $items = $lines[0] =~ tr/,//;
            die ("Load data in $file contains $items data items in 1st line, ".
                 "not a multiple of 24 hours per day")
                if $items % 24 != 0;
            $items /= 24;
            my $hour24 = 23;
            for (my $hour = YearHours - 1; $hour >= 0; $hour--) {
                my $line;
                for (my $j = 0; $j < $items; $j++) {
                    my @sources = split ",", $lines[int($hour / 24)];
                    $line = "$sources[0]:$hour24" if $j == 0;
                    $line .= ",".$sources[$items * $hour24 + $j + 1];
                }
                $lines[$hour] = $line;
                $hour24 = 23 if --$hour24 < 0;
            }
            $hours = YearHours;

        } elsif ($hours > YearHours) {
            # handle multiple lines per hour, assuming load in 2nd column
            die ("Load data in $file contains $hours data lines, ".
                 "not a multiple of ".YearHours." hours per year")
                if $hours % YearHours != 0;
            my $items = $hours / YearHours;
            for (my $hour = 0; $hour < YearHours; $hour++) {
                my $line;
                # my $load = 0;
                for (my $j = 0; $j < $items; $j++) {
                    my @sources = split ",", $lines[$items * $hour + $j];
                    $line = $sources[0] if $j == 0;
                    $line .= ",$sources[1]";
                    # $load += $sources[1];
                }
                # $line .= ",$load";
                $lines[$hour] = $line;
            }
            $hours = YearHours;
        }
        my $first_date = (split ",", $lines[0])[0];
        if ($first_date =~ /(19\d\d)/ || $first_date =~ /(20\d\d)/) {
            my $leap_years = $1 == 1900 ? 0 : int(($1 - 1901) / 4);
            $first_weekday = ($1 - 1900 + $leap_years) % 7;
        } else {
            print "Warning: cannot find year in load profile $file; ".
                "assuming that Jan 1st is a Friday (as for 2010)\n";
        }
    }

    my $rest = $hours % 24;
    die "Load data in $file does not cover full day; last day has $rest hours"
        if $hours == 0 || !$test && $rest != 0;
    print "Warning: load data in $file covers ".($hours / 24)." day(s); ".
        "will be repeated for the rest of the year\n"
        if !$test && $hours != YearHours;

    my $warned_just_before = 0;
    my $weekday = $first_weekday;
    my $items = 0; # number of load measure points in current hour
    my $day_load = 0;
    my $num_hours = $test ? TEST_END : YearHours;
    for (my $hour_per_year = 0; $hour_per_year < $num_hours; $hour_per_year++) {
        my @sources = split ",", $lines[$hour_per_year % $hours];
        my $n = $#sources;
        my $year = $sources[0] =~ /((19|20)\d\d)/ ? $1 : "0";
        shift @sources; # ignore month, day, and time info
        if ($items > 0 && $items != $n) {
            print "Warning: unequal number of items per line: $n vs. $items in "
                ."$file\n";
            $items = -1;
        }
        $items = $n if $items >= 0;
        $sum_items += ($items_by_hour[$month][$day][$hour] = $n);
        my $hload = 0;
        for (my $item = 0; $item < $n; $item++) {
            my $load = $sources[$item];
            die "Error parsing load item: '$load' in $file hour $hour_per_year"
                unless $load =~ m/^\s*-?\d+(\.\d+)?\s*$/;
            $hload += $load;
            $load_item[$month][$day][$hour][$item] = $load;
            if ($load <= 0) {
                my $lang = $en;
                $en = 1;
                print "Warning: load ".date_hour_str(0, $month, $day, $hour)
                    .sprintf("h, item %4d", $item + 1)." = $load\n"
                    unless $test || $warned_just_before;
                $en = $lang;
                $warned_just_before = 1;
            } else {
                $warned_just_before = 0;
            }
        }
        $hload /= $n;
        $load[$month][$day][$hour] = $hload;
        if ($avg_hour) {
            $items_by_hour[$month][$day][$hour] = 1;
            $load_item[$month][$day][$hour][0] = $hload;
        }
        $day_load += $hload;

        my $reached_end = $hour_per_year == $num_hours - 1; # needed for test
        if (++$hour == 24 || $test && $reached_end) {
            # adapt day according to @load_dist, @load_factors, and $load_const
            my $hour_end = $test && $reached_end ? TEST_END % 24 : 24;
            for ($hour = 0; $hour < $hour_end; $hour++) {
                my $sel = $test ? ($hour >= TEST_START || $day >= 2)
                                : selected($month, $day, $hour);
                $hload = $load[$month][$day][$hour];
                my $load_const_sel = defined $load_const && $weekday < $load_days &&
                    ($load_from > $load_to
                     ? ($load_from <= $hour || $hour < $load_to)
                     : ($load_from <= $hour && $hour < $load_to));
                if ($load_const_sel && !$load_min) {
                    $items_by_hour[$month][$day][$hour] = 1;
                    $load_item[$month][$day][$hour][0] = $hload = $load_const;
                }
                my $orig_hload = $hload;
                my $ref_hload =
                    $load[$month][$day][$hour] = $load_factors[$hour] / 100 *
                    (!defined $load_dist ? $hload :
                     $load_dist[$hour] * $day_load / $load_dist_sum);
                my $items = $items_by_hour[$month][$day][$hour];
                $sel_items += $items if $sel;
                $hload = 0; # need re-calculation at least with $load_min
                for (my $item = 0; $item < $items; $item++) {
                    my $load = $load_const;
                    if (!$load_const_sel || $load_min) {
                        $load = $load_item[$month][$day][$hour][$item]
                            * $load_factors[$hour] / 100;
                        $load *= $ref_hload / $orig_hload
                            if defined $load_dist && $orig_hload != 0;
                        $load = $load_const if (!$test || $hour >= TEST_START)
                            && $load_const_sel && $load_min && $load < $load_const;
                        $load_item[$month][$day][$hour][$item] = $load;
                    }
                    if ($sel) {
                        if (!defined $base_load || $load < $base_load) {
                            $base_load = $load;
                            $load_min_time = time_str($year, $month, $day,
                                                      $hour, $item, $items);
                        }
                        if ($load > $load_max) {
                            $load_max = $load;
                            $load_max_time = time_str($year, $month, $day,
                                                      $hour, $item, $items);
                        }
                        $load_min_hour[$hour]=min($load_min_hour[$hour], $load);
                        $load_max_hour[$hour]=max($load_max_hour[$hour], $load);
                    }
                    $hload += $load;
                }
                $hload /= $items;
                $load[$month][$day][$hour] = $hload;
                $load_sum += $hload;
                $hload = 0 unless $sel;
                $night_sum += $hload
                    if NIGHT_START <= $hour && $hour < NIGHT_END;
                $bright_sum += $hload
                    if BRIGHT_START <= $hour && $hour < BRIGHT_END;
                $load_by_hour   [$hour   ] += $hload;
                $load_by_weekday[$weekday] += $hload;
                $load_by_month  [$month  ] += $hload;
                $sel_load_sum += $hload;
                $sel_hours++ if $sel;
            }
            $day_load = 0;
            $hour = 0;
            if ($first_weekday==2 && $lines[$hour_per_year] =~ m/^31.05.1997/) {
                print "Warning: assuming $file is a BDEW profile of 1996..97\n";
                # staying on Sat when switching from 31 May 1997 to 1 Jan 1996
            } else {
                $weekday = 0 if ++$weekday == 7;
            }
        }
        adjust_day_month();
    }
    $items_per_hour = $sum_items / $num_hours;
}

get_profile($load_profile);
my $orig_load_sum = $load_sum;
my $load_scale = defined $consumption && $load_sum != 0
    ? 1000 * $consumption / $load_sum : 1;
my $load_scale_never_0 = never_0($load_scale);
sub rls { return round($_[0] * $load_scale); };

my $hours_a_day = defined $sel_hour ? $sel_hour2 - $sel_hour + 1 : 24;
my $sn = $sel_hours ? 1 / $sel_hours * $hours_a_day : 0;
for (my $hour = 0; $hour < 24; $hour++) {
    $load_per_hour[$hour] = rls($load_by_hour [$hour] * $sn);
    $load_min_hour[$hour] = !defined $load_min_hour[$hour] ? 0
                          : rls($load_min_hour[$hour]);
    $load_max_hour[$hour] = rls($load_max_hour[$hour]);
}
$load_const = rls($load_const) if defined $load_const;
$load_sum = rls($load_sum);
$base_load= rls($base_load) if defined $base_load;
$load_max = rls($load_max);

my $profile_txt = $en ? "load profile file"     : "Lastprofil-Datei";
my $pv_data_txt = $en ? "PV data file"          : "PV-Daten-Datei";
my $plural_txt  = $en ? "(s)"                   : "(en)";
my $direct_txt  = $en ? "direct"                : "direkt";
my $limit_txt   = $en ? "inverter input limit"  : "WR-Eingangs-Begrenzung";
my $none0_txt   = $en ? "('*' means none)"      : "('*' bedeutet keine)";
my $slope_txt   = $en ? "slope"                 : "Neigungswinkel";
my $azimuth_txt = $en ? "azimuth"               : "Azimut";
my $yearly_txt  = $en ? "over a year"           : "über ein Jahr";
my $values_txt  = $en ? "values"                : "Werte";
my $only_txt    = $en ? "only"                  : "nur";
my $during_txt  = $en ? "during"                : "während";
my $from_txt    = $en ? "from"                  : "von";
my $to_txt      = $en ? "to"                    : "bis";
my $hrs_txt     = $en ? "hrs"                   : "Uhr";
my $hour_txt    = $en ? "hour"                  : "Stunde";
my $TMY         = "TMY (2008..2020)";
my $simul_year  = $en ? "simulated PV year"     : "Simuliertes PV-Jahr";
my $energy_txt  = $en ? "energy values are"     : "Energiewerte sind";
my $p_txt = $en ? "load data points per hour  " : "Last-Datenpunkte pro Stunde";
my $eh          = $en ? "each hour"             : "je Std";
my $D_txt       = $en ? "rel. load distr. $eh"  : "Rel. Lastverteilung $eh.";
my $d_txt       = $en ? "load distortions $eh"  : "Last-Verzerrung je Stunde";
my $lph_txt     = $en ? "load $eh"              : "Last je Stunde";
my $l_txt       = $en ? "average $lph_txt"      : "Mittlere $lph_txt";
my $l_min_txt   = $en ? "minimal $lph_txt"      : "Minimale $lph_txt";
my $l_max_txt   = $en ? "maximal $lph_txt"      : "Maximale $lph_txt";
my $t_txt = $en ? "consumption acc. to profile" : "Verbrauch gemäß Lastprofil ";
my $consumpt_txt= $en ? "consumption" : "Verbrauch";
my $L_txt = $en ? "load portion"                : "Last-Anteil";
my $V_txt = $en ? "PV portion"                  : "PV-Anteil";
my $per3  = $en ? "per 3 hours"                 : "pro 3 Stunden";
my $per_m = $en ? "per month"                   : "pro Monat";
my $W_txt = $en ? "portion per weekday (Mo-Su)" :"Anteil pro Wochentag (Mo-So)";
my $b_txt = $en ? "nightly load on average"     :"Nächtliche Durchschnittslast";
my $B_txt = $en ? "daytime load on average"     :"Tagsüber Durchschnittslast";
my $m_txt = $en ? "min load (base load)"        : "Minimallast (Grundlast)";
my $M_txt = $en ? "max load"                    : "Maximallast";
my $en1 = $en ? " "   : "";
my $en2 = $en ? "  "  : "";
my $en3 = $en ? "   " : "";
my $en4 = $en2.$en2;
my $en5 = $en2.$en3;
my $en6 = $en3.$en3;
my $de1 = $en ? ""    : " ";
my $de2 = $en ? ""    : "  ";
my $de3 = $en ? ""    : "   ";
my $s10   = "          "; 
my $s13   = "             ";
my $lhs_spaces = " " x length("$W_txt$en1");

sub from_to_hours_str {
    my ($from, $to) = @_;
    return $from == 0 && $to == 24 ? ""
        : " $from_txt ".int($from)." $to_txt ".int($to)." $hrs_txt";
}
my $only_during = $only ? " $only_txt $during_txt $only" : "";
# $only =~ s/^[^-]*-// if $only; # chop year restriction
my $timeframe = $only ? $only_during : " $yearly_txt";

print "load_scale                  = $load_scale\n" if $debug && !$test;
print "$profile_txt$de1$s10 : $load_profile\n" unless $test;
if ($verbose) {
    print "$p_txt = ".sprintf("%4d", $items_per_hour)."\n";
    print "$t_txt =".kWh($orig_load_sum)."\n";
    print "$consumpt_txt $de2                =".kWh($load_sum)." $yearly_txt\n";
    print "$consumpt_txt $de2                =".kWh(rls($sel_load_sum)).
        "$timeframe\n" if $only;
}
my $night_avg_load = $night_sum * $sn;
$night_avg_load /= (NIGHT_END - NIGHT_START) if $hours_a_day == 24;
my $bright_avg_load = $bright_sum * $sn;
$bright_avg_load /= (BRIGHT_END - BRIGHT_START) if $hours_a_day == 24;
print "$b_txt$en5=".W(rls($night_avg_load)).
    from_to_hours_str(NIGHT_START, NIGHT_END)."\n";
print "$B_txt$en3  =".W(rls($bright_avg_load)).
    from_to_hours_str(BRIGHT_START, BRIGHT_END)."\n";
print "$m_txt $en3    =".(defined $base_load ? W($base_load) : " $none_txt")
                                                          ." $load_min_time\n";
print          "$M_txt $en3                =".W($load_max)." $load_max_time\n";
print_arr_day( "$D_txt $en1= ", \@load_dist   ) if defined $load_dist;
print_arr_day("$d_txt $de1 = ", \@load_factors) if defined $load_factors;
if ($verbose) {
    print_arr_day("$hour_txt                     $en2 = ", \@hours);
    print_arr_day(    "$l_txt $en1    = ", \@load_per_hour);
    print_arr_day("$l_min_txt $en1    = ", \@load_min_hour);
    print_arr_day("$l_max_txt $en1    = ", \@load_max_hour);
    print_arr("$L_txt $per3 $en1  = ", \@load_by_hour, $sel_load_sum, 0, 21, 3);
    print_arr("$L_txt $per_m $de1     = ",
                              \@load_by_month  , $sel_load_sum, 1, 12, 1);
    print_arr("$W_txt$en1= ", \@load_by_weekday, $sel_load_sum, 0,  6, 1);
}
$sel_load_sum *= $load_scale;
print "\n";

################################################################################
# read PV production data

my $nominal_power_sum = 0;
my $DC_input_losses = 0; # due to inverter/charger DC input limits on PV power
my @PV_gross;
my @PV_gross_across; # sum over years, only where selected
my @PV_net;
my @PV_net_loss;
my @PV_net_direct;
my @PV_net_across; # sum over years, only where selected
my ($start_year, $years);
my $garbled_hours = 0;
sub get_power { my ($file, $nominal_power, $limit, $direct) = @_;
    open(my $IN, '<', $file) or die "Could not open PV data file $file: $!\n"
        unless $test;
    print "$pv_data_txt $en2$s13: $file" unless $test;

    # my $sum_needed = 0;
    my ($slope, $azimuth);
    my $current_years = 0;
    my $nominal_power_deflt; # PVGIS default: 1 kWp
    my       $sys_eff_deflt; # PVGIS default: 0.86
    my     $pvsys_eff_deflt; # PVGIS default: 0.86 / $inverter_eff_never_0
    my $power_provided = 1;  # PVGIS default: none
    my ($power_rate, $gross_rate, $net_rate);
    my ($months, $hours) = (0, 0);
    my $net_limit = $limit * $inverter_eff if $limit ne '*';
    while ($_ = !$test ? <$IN>
           : "2000010".(1 + int($hours / 24)).":"
           .sprintf("%02d", $hours % 24)."00, "
           .(TEST_PV_START <= $hours && $hours < TEST_PV_END ?
             $nominal_power * $pvsys_eff * $inverter_eff : 0)."\n") {
        chomp;
        if (m/^Latitude \(decimal degrees\):[\s,]*(-?\d+([\.,]\d+)?)/) {
            check_consistency($1, $lat, "latitude", $file) if defined $lat;
            $lat = $1;
        }
        if (m/^Longitude \(decimal degrees\):[\s,]*(-?\d+([\.,]\d+)?)/) {
            check_consistency($1, $lon, "longitude", $file) if $lon;
            $lon = $1;
        }
        $slope   = $1 if (!$slope   && m/^Slope:\s*(-?[\d\.]+ deg[^,\s]*)/);
        $azimuth = $1 if (!$azimuth && m/^Azimuth:\s*(-?[\d\.]+ deg[^,\s]*)/);
        $nominal_power_deflt = $1 * 1000
            if (!defined $nominal_power_deflt
                && m/Nominal power.*? \(kWp\):[\s,]*([\d\.]+)/);
        if (!$pvsys_eff_deflt
            && m/System losses \(%\):[\s,]*(\d+([\.,]\d+)?)/) {
            $sys_eff_deflt = 1 - $1 / 100;
            # See section "System loss" in
            # ​https://joint-research-centre.ec.europa.eu/pvgis-online-tool/getting-started-pvgis/pvgis-user-manual_en
            $pvsys_eff_deflt = $sys_eff_deflt / $inverter_eff_never_0;
            my $eff = $sys_eff_deflt * 100;
            print ", ".($en ?
                        "contained system efficiency $eff% was overridden" :
                        "enthaltene System-Effizienz $eff% wurde übersteuert")
                if !$test && $pvsys_eff && $pvsys_eff != $pvsys_eff_deflt
                && abs($sys_eff_deflt - $pvsys_eff * $inverter_eff) * 100 >= .5;
        }
        $power_provided = 0 if m/^time,/ && !(m/^time,P,/);

        # work around CSV lines at hour 00 garbled by load&save with LibreOffice
        $garbled_hours++ if s/^\d+:10:00,0,/20000000:0010,0,/;
        next unless m/^20\d\d\d\d\d\d:\d\d\d\d,/;

        unless (defined $power_rate || $test) {
            print "\n" # close line started with print "$pv_data_txt..."
                unless $lat && $lon && $slope && $azimuth && $power_provided;
            # die "Missing latitude in $file"  unless $lat;
            # die "Missing longitude in $file" unless $lon;
            # die "Missing slope in $file"     unless $slope;
            # die "Missing azimuth in $file"   unless $azimuth;
            die "Missing PV power output data in $file" unless $power_provided;

            unless ($nominal_power_deflt) {
                print "\n"; # close line started with print "$pv_data_txt..."
                die "Missing nominal PV power in command line option or $file"
                    unless $nominal_power;
                print "Warning: cannot find nominal PV power in $file, ".
                    "taking value $nominal_power from command line\n";
                $nominal_power_deflt = $nominal_power;
            }
            $nominal_power = $nominal_power_deflt unless defined $nominal_power;

            unless ($pvsys_eff || defined $pvsys_eff_deflt) {
                print "\nWarning: cannot find system efficiency in $file and ".
                    "no -peff option given, assuming system efficiency 86%\n";
                $pvsys_eff = 0.86 / $inverter_eff_never_0;
            }
            $pvsys_eff = $pvsys_eff_deflt unless defined $pvsys_eff;
            if ($pvsys_eff < 0 || $pvsys_eff > 1) {
                print "\n"; # close line started with print "$pv_data_txt..."
                die "Unreasonable PV system efficiency ".($pvsys_eff * 100)
                    ." % - have -peff and -ieff been used properly?";
            }
            print ", assuming that PV data is net (after PV system and "
                ."inverter losses)" unless defined $pvsys_eff_deflt;
        }
        unless (defined $power_rate) {
            $nominal_power_sum += $nominal_power;
            $power_rate = $test ? 1 : $nominal_power / $nominal_power_deflt;
            $pvsys_eff_never_0 = never_0($pvsys_eff);
            $gross_rate = 1 / $pvsys_eff_never_0 / $inverter_eff_never_0;
            $net_rate = $pvsys_eff * $inverter_eff;
        }

        next if m/^20\d\d0229:/; # skip data of Feb 29th (in leap years)
        $start_year = $1 if (!$start_year && m/^(\d\d\d\d)/ && $1 != 2000);
        if ($tmy && !$test) {
            # typical metereological year
            my $selected_month = 0;
            if ($lat eq "48.215" && $lon eq "11.727") {
                $selected_month = $1 if m/^2007(01)/;
                $selected_month = $1 if m/^2018(02)/;
                $selected_month = $1 if m/^2020(03)/;
                $selected_month = $1 if m/^2008(04)/;
                $selected_month = $1 if m/^2011(05)/;
                $selected_month = $1 if m/^2010(06)/;
                $selected_month = $1 if m/^2012(07)/;
                $selected_month = $1 if m/^2013(08)/;
                $selected_month = $1 if m/^2012(09)/;
                $selected_month = $1 if m/^2017(10)/;
                $selected_month = $1 if m/^2014(11)/;
                $selected_month = $1 if m/^2018(12)/;
            } else {
                $selected_month = $1 if m/^2012(01)/;
                $selected_month = $1 if m/^2013(02)/;
                $selected_month = $1 if m/^2010(03)/;
                $selected_month = $1 if m/^2019(04)/; # + 2 kWh
                $selected_month = $1 if m/^2016(05)/; # + 1 kWh
                $selected_month = $1 if m/^2015(06)/;
                $selected_month = $1 if m/^2007(07)/;
                $selected_month = $1 if m/^2008(08)/;
                $selected_month = $1 if m/^2010(09)/;
                $selected_month = $1 if m/^2019(10)/;
                $selected_month = $1 if m/^2008(11)/;
                $selected_month = $1 if m/^2020(12)/;
            }
            next unless $selected_month;
        }
        # matching hour 01 rather than 00 due to potentially garbled CSV lines:
        $current_years++ if m/^20\d\d0101:01/;
        $months++ if m/^20....01:01/;

        die "Missing PV power data in $file line $_"
            unless m/^(\d\d\d\d)(\d\d)(\d\d):(\d\d)(\d\d),\s?([\d\.]+)/;
        my $hour_offset = $test ? 0 : TimeZone;
        my ($year, $month, $day, $hour, $minute_unused, $net_power) =
            ($tmy ? 0 : $1-$start_year, $2, $3, ($4+$hour_offset) % 24, $5, $6);
        # for simplicity, attributing hours wrapped via time zone to same day
        # Note: for garbled CSV lines, $year, $month, and $day are wrong,
        # which leads to missing (undefined) PV power data entries in @PV_gross

        $net_power *= $power_rate;
        my $gross_power = $net_power *
            (defined $sys_eff_deflt ? 1 / $sys_eff_deflt : $gross_rate);
        $PV_gross[$year][$month][$day][$hour] += $gross_power;
        $net_power = $gross_power * $net_rate;
        if ($limit ne '*' &&  $net_power > $net_limit) {
            $PV_net_loss[$year][$month][$day][$hour] += $net_power - $net_limit;
            $net_power = $net_limit;
        }
        $PV_net       [$year][$month][$day][$hour] += $net_power;
        $PV_net_direct[$year][$month][$day][$hour] += $direct ? $net_power : 0;

        $hours++;
        last if $test && $hours == TEST_END;
    }
    close $IN unless $test;
    print "\n" unless $test; # close line started with print "$pv_data_txt..."

    check_consistency($years, $current_years, "years", $file) if $years;
    $years = $current_years;
    die "Number of years detected is $years in $file"
        if $years < 1 || $years > 100;
    check_consistency($months,       12 * $years, "months", $file);
    check_consistency($hours, YearHours * $years, "hours", $file)
        if $garbled_hours == 0;
    if ($test) {
        $years = 1;
        return $nominal_power;
    }

    if (defined $slope && defined $azimuth) {
        $slope   =~ s/ deg\./°/;
        $azimuth =~ s/ deg\./°/;
        $slope   =~ s/optimum/opt./;
        $azimuth =~ s/optimum/opt./;
        print "$slope_txt, $azimuth_txt$en3$en5      = $slope, $azimuth\n";
    }
    return $nominal_power;
}

my ($total_nomin, $total_limit) = (0, 0);
for (my $i = 0; $i <= $#PV_files; $i++) {
    my $nomin =
        get_power($PV_files[$i], $PV_nomin[$i], $PV_limit[$i], $PV_direct[$i]);
    $total_nomin += $nomin;
    $total_limit += $PV_limit[$i] ne '*' ? $PV_limit[$i] : $nomin;
    $PV_nomin[$i] = ($PV_direct[$i] ? "$direct_txt:" : "").$nomin;
}
my $PV_nomin = join("+", @PV_nomin);
my $PV_limit = join("+", @PV_limit);

my $lat_txt        = $en ? "latitude"             : "Breitengrad";
my $lon_txt        = $en ? "longitude"            : "Längengrad";
print "$lat_txt, $lon_txt $en4    = $lat, $lon\n"
if defined $lat && defined $lon && !$test;

################################################################################
# PV usage simulation

my @PV_per_hour;
my @PV_max_hour = (0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
my @PV_by_hour  = (0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
my @PV_by_month = (0,0,0,0,0,0,0,0,0,0,0,0,0); # months starting at index 1
my $PV_gross_max = 0;
my $PV_gross_max_time = $no_time_txt;
my $PV_gross_sum = 0;

my @PV_loss           if $curb;
my $PV_losses     = 0 if $curb;
my $PV_loss_hours = 0 if $curb;
my $PV_net_max = 0;
my $PV_net_max_time = $no_time_txt;
my $PV_net_sum = 0;
my $PV_net_discarded_sum = 0;

# set only if respective component of @PV_loss != 0:
my @PV_use_loss_by_item   if $curb && $max;
my @PV_use_loss           if $curb;

my $PV_use_loss_sum   = 0 if $curb;
my $PV_use_loss_hours = 0 if $curb;
my @PV_used_by_item;
my @PV_used;
my $PV_used_sum = 0;
my @PV_excess_by_item if $max;
my @PV_excess;
my $PV_excess_sum = 0;
my $PV_excess_max = 0; # per day
my $PV_excess_max_time = $no_time_txt;
my $PV_used_via_storage = 0;

my @grid_feed_by_item if $max;
my @grid_feed_per_hour = (0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
my @grid_feed_max_hour = (0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
my @grid_feed;
my $grid_feed_sum = 0;
my $dis_feed_sum = 0
    if defined $capacity; # grid feed-in via inefficient storage discharge

my $soc_max = $capacity *      $max_soc  if defined $capacity;
my $soc_min = $capacity * (1 - $max_dod) if defined $capacity;
my $eff_capacity = $soc_max - $soc_min   if defined $capacity;
my $soc_max_reached = $soc_min           if defined $capacity;
my $soc_max_time = $no_time_txt;
my $max_chgpower = $max_chgrate * $capacity / $charge_eff_never_0
    / $load_scale_never_0 if defined $capacity;
my $max_dispower = $max_disrate * $capacity / $storage_eff_never_0
    / $load_scale_never_0 if defined $capacity;
my @soc               if defined $capacity; # state of charge on avg over years
my @soc_by_item       if defined $capacity;
my $soc               if defined $capacity; # state of charge of the battery
my @charge            if defined $capacity; # charge delta on average over years
my @charge_by_item    if defined $capacity && $max;
my @charge_per_hour = (0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
                      if defined $capacity;
my @charge_max_hour = (0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
                      if defined $capacity;
my $charge_sum    = 0 if defined $capacity;
my @dischg            if defined $capacity; # dischg delta on average over years
my @dischg_by_item    if defined $capacity && $max;
my @dischg_per_hour = (0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
                      if defined $capacity;
my @dischg_max_hour = (0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
                      if defined $capacity;
my $dischg_sum    = 0 if defined $capacity;
my $DC_feed_loss  = 0 if defined $capacity; # due to inverter before storage
my $spill_loss    = 0 if defined $capacity;
my $charging_loss = 0 if defined $capacity;
my $storage_loss  = 0 if defined $capacity;
my $coupling_loss = 0 if defined $capacity;

sub simulate_charge {
    my ($pv_used, $grid_feed_in, $maybe_loss, $power_needed,
        $unused_bypass, $excess_power,
        $year_str, $month, $day, $hour, $item, $items,
        $trace, $PV_loss) = @_;
    my $charge_delta = 0;

    my $capacity_to_fill = $soc_max - $soc;
    $capacity_to_fill = 0 if $capacity_to_fill < 0;
    my $charge_input = 0;
    my $trace_blank_no_surplus = " " x
        (defined $bypass ? ($bypass_spill ? 28 : 11): 18)
        if $trace;
    my $chg_limited = "" if $trace;
    if ($excess_power > 0) {
        # $excess_power is the power available for charging

        my $need_for_fill = $capacity_to_fill / $charge_eff_never_0;
        my $limited_fill = $need_for_fill;
        $limited_fill = $max_chgpower if $limited_fill > $max_chgpower;

        # optimal charge: exactly as much as unused and fits in
        $charge_input = $excess_power;
        # will become min($excess_power, $limited_fill);
        my $surplus = $excess_power - $limited_fill;
        printf("[excess=%4d,tofill=%6d,surplus=%4d] ", rls($excess_power),
               rls($need_for_fill), rls(max($surplus, 0))) if $trace;
        if ($surplus > 0) {
            $chg_limited =  $limited_fill < $need_for_fill ? "(rate limit) "
                : "" if $trace; # not showing "(SoC limit) "
            $charge_input = $limited_fill;
            # TODO properly handle simultaneous charge and discharge
            # which is relevant with non-optimal charging (-pass)

            my $surplus_net = $surplus;
            # when DC-coupled, need to transform surplus back to net
            $surplus_net *= $inverter_eff if $DC_coupled;

            if (!defined $bypass) { # i.e., on optimal charge
                $grid_feed_in += $surplus_net;
                printf("surplus feed=%4d ", rls($surplus_net))
                    if $trace;
            } elsif ($bypass_spill) {
                my $remaining_surplus = $surplus_net -$power_needed;
                my $used_surplus = $power_needed;
                if ($remaining_surplus < 0) {
                    $used_surplus = $surplus_net;
                    $remaining_surplus = 0;
                }
                $pv_used += $used_surplus;
                $power_needed -= $used_surplus;
                $grid_feed_in += $remaining_surplus;
                printf("surplus feed=%4d,used=%4d ", rls($remaining_surplus),
                       rls($used_surplus)) if $trace;
            } else {
                # defined $bypass && !$bypass_spill && $DC_coupled
                $spill_loss += $surplus;
                printf("spill=%4d ", rls($surplus)) if $trace;
            }
        } elsif ($trace) {
            printf($trace_blank_no_surplus);
        }

        # add reduced charging due to curb to potential usage losses
        # - well, this is just approximate:
        $maybe_loss +=
            $capacity_to_fill * $storage_eff * $inverter2_eff
            if $AC_coupled && $PV_loss != 0 # implies $curb
               && $unused_bypass == 0;

        $charge_delta = $charge_input * $charge_eff;
        $soc += $charge_delta;
        if ($soc > $soc_max_reached) {
            $soc_max_reached = $soc;
            $soc_max_time = time_str($year_str, $month, $day, $hour, $item, $items);
        }
        $charging_loss += $charge_input - $charge_delta;

    } elsif ($trace) {
        printf(" " x 41); # no $excess_power
        printf($trace_blank_no_surplus);
    }
    my $trace_charge = $trace &&
        ($excess_power > 0 || $soc > $soc_min);
    printf("chrg loss=%4d dischrg needed=%4d [SoC %6d + %4d %s",
           rls($charge_input - $charge_delta),
           rls($power_needed), rls($soc - $charge_delta),
           rls($charge_delta), $chg_limited) if $trace_charge;

    return ($charge_delta, $pv_used, $grid_feed_in, $maybe_loss, $power_needed,
            $trace_charge);
}

sub simulate_item {
    my ($year_str, $month, $day, $hour,
        $gross_power, $pvnet_power, $pvnet_direct,
        $item, $items, $trace, $PV_loss, $PV_loss_capa) = @_;
    my $load = $load_item[$month][$day][$hour][$item];
    # $gross_power is used only with DC-coupled charging
    # $pvnet_{power,direct} and $load are the main input for simulation
    # $PV_loss is upper limit for PV net usage loss computation on $curb

    die "Internal error: load_item[$month][$day][$hour][$item] is undefined"
        unless defined $load;
    if ($dump_avg) {
        # adapt CET to UTC
        my ($m, $d, $h) = ($month, $day, $hour + 1);
        ($d, $h) = ($d + 1, 0) if $h > 23;
        ($m, $d) = ($m == 12 ? 1 : $m + 1, 1) if $d > $days_per_month[$m];
        printf $DA "\"Datetime\",\"Power\"\r\n"
            if $month == 1 && $day == 1 && $hour == 0;
        printf $DA "%s%02d%02d:%02d,%d\r\n", $year_str, $month, $day, $hour,
            rls($load_item[$m][$d][$h][0]);
    }
    if ($trace) {
        printf("%s-%02d-%02d ", $year_str, $month, $day) unless $test;
        printf("%02d".minute_str($item, $items)." load=%4d PV net=%4d ",
               $hour, rls($load), rls($pvnet_power));
    }
    # $needed += $load;
    # load will be reduced by $bypass or $bypass_spill or $pvnet_direct

    my $pv_use_loss = 0 if $curb;
    my $grid_feed_in = 0; # locally accumulates grid feed
    my ($charge_delta, $dischg_delta, $coupling_loss) = (0, 0, 0)
        if defined $capacity;

    my $pv_taken = # value set here is preliminary, may be adapted below
        defined $bypass ? $bypass + $pvnet_direct # constant and direct bypass
        : ($load >= $pvnet_direct ? $load # only as needed (optimal charge)
           : $pvnet_direct); # direct feed from PV array exceeding load
    my $maybe_loss = 0; # potential usage losses due to curb
    my $excess_power = $pvnet_power - $pv_taken;
    if ($excess_power < 0) {
        $maybe_loss = -$excess_power if $curb;
        $pv_taken = $pvnet_power;
        $excess_power = 0;
    }

    # $pv_used locally accumulates PV own consumption
    my $pv_used = $pv_taken; # value here is preliminary, may be adapted below
    my $unused_bypass = 0;
    if (defined $bypass || $load < $pvnet_direct) {
        $unused_bypass = $pv_taken - $load;
        if ($unused_bypass > 0) {
            $pv_used = $load;
            $grid_feed_in = $unused_bypass;
            $maybe_loss = 0;
        } else {
            $unused_bypass = 0;
        }
    }
    if ($trace) {
        if (defined $bypass) {
            printf("bypass%s used=%4d,unused=%4d ",
                   $pvnet_direct > 0 ? "+direct" : "",
                   rls($pv_used), rls($unused_bypass))
        } elsif ($pvnet_direct > 0) {
            printf("direct used=%4d,unused=%4d ",
                   rls($pvnet_direct - $unused_bypass), rls($unused_bypass))
        }
    }
    my $power_needed = $load - $pv_used;

    my $dis_feed_in = 0; # grid feed-in via inefficient storage discharge
    my $PV_loss_curr = $PV_loss if $PV_loss != 0;
    if (defined $capacity) { # storage present
        $PV_loss_curr = $PV_loss_capa if $PV_loss != 0 && $maybe_loss == 0;
        # when charging is DC-coupled, no loss through inverter:
        my $exc = $gross_power * $pvsys_eff
            - ($pv_used + $grid_feed_in) / $inverter_eff_never_0;

        ($charge_delta, $pv_used, $grid_feed_in, $maybe_loss, $power_needed,
         my $trace_charge) =
            simulate_charge($pv_used, $grid_feed_in, $maybe_loss, $power_needed,
                            $unused_bypass, $DC_coupled ? $exc : $excess_power,
                            $year_str, $month, $day, $hour, $item, $items,
                            $trace, $PV_loss);


        ## add reduced discharging due to curb to potential usage losses
        ## - well, this would be just approximate:
        #$maybe_loss +=
        #    $power_needed / $charge_eff_never_0 / $storage_eff_never_0
        #    if $unused_bypass == 0 && $PV_loss != 0; # implies $curb

        my $dischg_loss = 0;
        if ($soc > $soc_min) { # storage not empty
            # optimal discharge: exactly as much as currently needed
            # $discharge = min($power_needed, $soc)
            my $discharge = $power_needed /
                ($storage_eff_never_0 * $inverter2_eff_never_0);
            if (defined $max_feed_scaled_by_eff) {
                my $active = $feed_from > $feed_to
                    ? ($feed_from <= $hour || $hour < $feed_to)
                    : ($feed_from <= $hour && $hour < $feed_to);
                if ($comp_feed || $excl_feed || $const_feed) {
                    $discharge = 0;
                    if ($const_feed) {
                        $discharge = $max_feed_scaled_by_eff if $active;
                    } elsif ($active) { # for $comp_feed || $excl_feed
                        my $pv_power = $pvnet_power - $pvnet_direct;
                        $pv_power = $bypass if $comp_feed &&
                            defined $bypass && $bypass < $pv_power;
                        my $target_power = $excl_feed ? $bypass :
                            $max_feed / $load_scale_never_0;
                        my $compensation = $target_power - $pv_power;
                        if ($comp_feed) {
                            $discharge = $compensation  /
                                ($storage_eff_never_0 * $inverter2_eff_never_0)
                                if $compensation > 0;
                        } elsif ($pv_power <= $target_power &&
                                 $pv_power <= $excl_threshold &&
                                 $compensation > $excl_threshold) {
                            $discharge = $target_power /
                                ($storage_eff_never_0 * $inverter2_eff_never_0);
                            # unfortunately, exclusive feed discards $pv_power:
                            $PV_net_discarded_sum +=
                                $pv_power * $load_scale / $items;
                            $pv_used -= $pv_power;
                            if ($pv_used < 0) {
                                $grid_feed_in += $pv_used;
                                $pv_used = 0;
                            }
                        }
                    }
                } elsif ($active) { # limit the optimal feed if active
                    $discharge = $max_feed_scaled_by_eff
                        if $discharge > $max_feed_scaled_by_eff;
                }
            }
            my $dis_limited = "" if $trace;
            my $charge_available = $soc - $soc_min;
            if ($discharge > $charge_available) {
                # $dis_limited .= " (DoD limit)" if $trace;
                $discharge = $charge_available;
            }
            if ($discharge > $max_dispower) {
                $dis_limited .= " (rate limit)" if $trace;
                $discharge = $max_dispower;
            }
            printf("- %4d%s - lost:%4d", rls($discharge * $storage_eff),
                   $dis_limited, rls($discharge * (1 - $storage_eff)))
                if $trace;
            if ($discharge != 0) {
                $soc -= $discharge; # includes storage loss
                $discharge *= $storage_eff;
                $dischg_delta = $discharge; # after storage loss
                my $discharge_net = $discharge * $inverter2_eff;
                $dischg_loss = $discharge - $discharge_net;
                $coupling_loss = $dischg_loss; # not really for DC-coupled
                if (defined $max_feed_scaled_by_eff &&
                    ($comp_feed || $excl_feed || $const_feed)) {
                    # $feed_sum += $discharge;
                    $dis_feed_in = $discharge_net - $power_needed;
                    if ($dis_feed_in > 0) {
                        $grid_feed_in  += $dis_feed_in;
                        $discharge_net -= $dis_feed_in;
                    } else {
                        $dis_feed_in = 0;
                    }
                }
                $pv_used += $discharge_net;
                $PV_used_via_storage += $discharge_net;
                # reduce potential usage losses by discharge
                # - well, this is be just approximate:
                $maybe_loss -= $discharge_net
                    if $unused_bypass == 0 && $PV_loss != 0; # implies $curb
            }
        } else {
            print "                   " if $trace;
        }
        # printf("= %4d] ", rls($soc)) if $trace_charge;
        printf("] ") if $trace_charge;
        printf("dischg loss=%4d ", rls($dischg_loss)) if $trace_charge;
        if ($max) {
            $charge_by_item[$month][$day][$hour][$item] += $charge_delta;
            $dischg_by_item[$month][$day][$hour][$item] += $dischg_delta;
        }
        $soc_by_item[$month][$day][$hour][$item] += $soc;
        printf(" " x (55 + 17)) if $trace && !$trace_charge;
    } elsif ($excess_power > 0) {
        $grid_feed_in = $excess_power;
    }

    if ($trace) {
        printf("used=%4d feed=%4d", rls($pv_used), rls($grid_feed_in));
        printf(" missing=%4d", rls($maybe_loss)) if $PV_loss != 0;
    }
    if ($max) {
        $PV_used_by_item  [$month][$day][$hour][$item] += $pv_used;
        $PV_excess_by_item[$month][$day][$hour][$item] += $excess_power;
        $grid_feed_by_item[$month][$day][$hour][$item] += $grid_feed_in;
    }

    if ($PV_loss != 0 && $maybe_loss > 0) { # implies $curb
        # just approximate if defined $capacity
        $pv_use_loss = min($PV_loss_curr, $maybe_loss);
        $PV_use_loss_by_item[$month][$day][$hour][$item]+= $pv_use_loss if $max;
        $PV_use_loss_hours++; # will be normalized by $sel_items
        printf(" curb loss=%4d", rls($pv_use_loss)) if $trace;
    } elsif ($PV_loss != 0 && $max) {
        $PV_use_loss_by_item[$month][$day][$hour][$item] += 0;
    }
    printf "\n" if $trace;

    return ($pv_used, $pv_use_loss, $excess_power, $grid_feed_in, $dis_feed_in,
            $charge_delta, $dischg_delta, $coupling_loss);

}

sub simulate_hour {
    my ($year_str, $year, $month, $day, $hour,
        $gross_power, $pvnet_power, $pvnet_direct) = @_;

    my ($hpv_used, $hexcess, $hgrid_feed ) = (0, 0, 0);
    my $hpv_use_loss = 0 if $curb;
    my ($hdis_feed, $hcharge_delta, $hdischg_delta, $hcpl_loss) = (0, 0, 0, 0)
        if defined $capacity;
    # my $needed = 0;

    # calculate statistics on PV gross power
    $PV_by_hour [$hour ] += $gross_power;
    $PV_max_hour[$hour ] = max($PV_max_hour[$hour], $gross_power);
    $PV_by_month[$month] += $gross_power;
    $PV_gross_across[$month][$day][$hour] += $gross_power;
    if ($gross_power > $PV_gross_max) {
        $PV_gross_max = $gross_power;
        $PV_gross_max_time = date_hour_str($year_str, $month, $day, $hour)."h";
    }
    $PV_gross_sum += $gross_power;

    my $PV_loss = 0;
    if ($curb && $pvnet_power > $curb) {
        $PV_loss = $pvnet_power - $curb;
        $PV_net[$year][$month][$day][$hour] = $pvnet_power = $curb;
        # print "".date_hour_str($year_str, $month, $day, $hour)."h".
        #"\tPV=".round($pvnet_power)."\tcurb=".round($curb).
        #"\tloss=".round($PV_losses)."\t$_\n";
    }
    $PV_net_across[$month][$day][$hour] += $pvnet_power;
    if ($pvnet_power > $PV_net_max) {
        $PV_net_max = $pvnet_power;
        $PV_net_max_time = date_hour_str($year_str, $month, $day, $hour)."h";
    }
    $PV_net_sum += $pvnet_power;

    # factor out $load_scale for optimizing the inner loop
    $gross_power  /= $load_scale_never_0 if $DC_coupled;
    $pvnet_power  /= $load_scale_never_0;
    $pvnet_direct /= $load_scale_never_0;
    $PV_loss      /= $load_scale_never_0 if $PV_loss != 0; # implies $curb

    my $PV_loss_capa;
    if (defined $capacity && $PV_loss != 0) {
        $PV_loss_capa = $PV_loss * $charge_eff * $storage_eff;
        $PV_loss_capa *= $inverter_eff if $AC_coupled;

        $PV_loss *= $charge_eff * $storage_eff * $inverter2_eff
            if defined $bypass && $bypass < $PV_loss;
    }

    my $items = $items_by_hour[$month][$day][$hour];

    # factor out $items for optimizing the inner loop
    if (defined $capacity && $items != 1) {
        $capacity *= $items;
        $soc_max *= $items;
        $soc_min *= $items;
        $soc_max_reached *= $items;
        $soc *= $items;
        $charging_loss *= $items;
        $spill_loss *= $items;
        $PV_used_via_storage *= $items;
    }

    my $trace = $test ? ($day - 1) * 24 + $hour >= TEST_START : $debug;
    # my $feed_sum = 0 if defined $max_feed_scaled_by_eff;
    for (my $item = 0; $item < $items; $item++) {
        my ($pvu, $pul, $excess_power, $gfi, $dfi, $chg, $dis, $cpl) =
        simulate_item($year_str, $month, $day, $hour,
                      $gross_power, $pvnet_power, $pvnet_direct,
                      $item, $items, $trace, $PV_loss, $PV_loss_capa);
        ($hpv_used += $pvu, $hexcess += $excess_power, $hgrid_feed += $gfi);
        $hpv_use_loss += $pul if $curb;
        $grid_feed_max_hour[$hour] = max($grid_feed_max_hour[$hour], $gfi);
        if (defined $capacity) {
            $hdis_feed += $dfi;
            $charge_max_hour[$hour] = max($charge_max_hour[$hour], $chg);
            $dischg_max_hour[$hour] = max($dischg_max_hour[$hour], $dis);
            $hcharge_delta += $chg;
            $hdischg_delta += $dis;
            $hcpl_loss += $cpl;
        }
    }
    # $spill_loss += ($pvnet_power - $feed_sum / $items) if defined $bypass;

    if ($PV_loss != 0) { # implies $curb
        $hpv_use_loss /= $items;
        $PV_use_loss[$month][$day][$hour] += $hpv_use_loss;
        $PV_use_loss_sum += $hpv_use_loss;
        $PV_loss = $hpv_use_loss if $DC_coupled;
        if ($PV_loss != 0) {
            $PV_loss[$month][$day][$hour] += $PV_loss;
            $PV_losses += $PV_loss;
            $PV_loss_hours++;
        }
    }
    # $sum_needed += $needed / $items; # per hour
    # print "".date_hour_str($year_str, $month, $day, $hour)."h\t".
    # "PV=".round($pvnet_power)."\tPN=".round($needed)."\tPU=".round($hpv_used).
    # "\t$_\n" if $pvnet_power != 0 && m/^20160214:1010/; # m/^20....02:12/;

    if ($items != 1) {
        # revert factoring out $items for optimizing the inner loop
        $hpv_used /= $items;
        $hexcess /= $items;
        $hgrid_feed /= $items;
    }
    $PV_used_sum   += $hpv_used;
    $PV_excess_sum += $hexcess;
    $grid_feed_sum += $hgrid_feed;

    if (defined $capacity) {
        if ($items != 1) {
            # revert factoring out $items for optimizing the inner loop
            $capacity /= $items;
            $soc_max /= $items;
            $soc_min /= $items;
            $soc_max_reached /= $items;
            $hdis_feed /= $items;
            $soc /= $items;
            $charging_loss /= $items;
            $hcpl_loss /= $items;
            $spill_loss /= $items;
            $PV_used_via_storage /= $items;
            $hcharge_delta /= $items;
            $hdischg_delta /= $items;
        }
        $dis_feed_sum  += $hdis_feed;
        $charge_sum += $hcharge_delta;
        $dischg_sum += $hdischg_delta;
        $coupling_loss += $hcpl_loss;
    }

    $PV_used   [$month][$day][$hour] += $hpv_used;
    $PV_excess [$month][$day][$hour] += $hexcess;
    $grid_feed_per_hour      [$hour] += $hgrid_feed;
    $grid_feed [$month][$day][$hour] += $hgrid_feed;
    if (defined $capacity) {
        $charge_per_hour     [$hour] += $hcharge_delta;
        $dischg_per_hour     [$hour] += $hdischg_delta;
        $charge[$month][$day][$hour] += $hcharge_delta;
        $dischg[$month][$day][$hour] += $hdischg_delta;
        $soc   [$month][$day][$hour] += $soc;
    }
    return $hexcess;
}

sub simulate()
{
    my $year = 0;
    ($month, $day, $hour) = (1, 1, 0);

    if (defined $capacity) {
        # factor out $load_scale for optimizing the inner loop
        $capacity /= $load_scale_never_0;
        $bypass   /= $load_scale_never_0 if defined $bypass;
        $max_feed_scaled_by_eff /= $load_scale_never_0 if defined $max_feed;
        $excl_threshold /= $load_scale_never_0 if defined $excl_threshold;
        $soc_max  /= $load_scale_never_0;
        $soc_min  /= $load_scale_never_0;
        $soc_max_reached /= $load_scale_never_0;
        $soc = $soc_min;
    }

    my $end_year = $years;
    # restrict simulation to year optionally given by the -only option
    if (defined $sel_year) {
        my $max = $start_year + $years - 1;
        $year = $sel_year - $start_year;
        die "Year given with -only option must be in range $start_year..$max"
            if $year < 0 || $year >= $years;
        die "Second year given with -only option must be in range "
            ."$sel_year..$max" if defined $sel_year2
            && ($sel_year2 < $sel_year || $sel_year2 - $start_year >= $years);
        $years = 1 + (defined $sel_year2 ? $sel_year2 - $sel_year : 0);
        $end_year = $year + $years;
    }
    my $first_year = $year;

    STDOUT->autoflush(1);
    print "$simul_year$en2         :"
        unless $test;
    my $excess = 0;
    while ($year < $end_year) {
        my $year_str = $tmy ? "TMY" : $start_year + $year;
        my $year_str_TMY = " ".($tmy ? "$TMY" : $year_str);
        print "$year_str_TMY\n" if !$test && $debug
            && $month == 1 && $day == 1 && $hour == 0;

        $PV_loss[$month][$day][$hour] = 0 if $curb && $year == $first_year;

        # restrict simulation to month, day, and hour given by the -only option
        if (selected($month, $day, $hour)) {
            my $gross_power  = $PV_gross       [$year][$month][$day][$hour];
            my $pvnet_power  = $PV_net         [$year][$month][$day][$hour];
            my $pvnet_direct = $PV_net_direct  [$year][$month][$day][$hour];
            my $DC_loss =      $PV_net_loss    [$year][$month][$day][$hour];
            $DC_input_losses += $DC_loss if defined $DC_loss;
            my $DC_power = ($pvnet_power-$pvnet_direct) / $inverter_eff_never_0;

            if (!defined $gross_power) {
                if ($hour == 0 && $garbled_hours) { # likely just garbled hour
                    $gross_power = $DC_power = $pvnet_power = $pvnet_direct = 0;
                } else {
                    $en = 1;
                    die "No PV power data ".
                        date_hour_str($year_str, $month, $day, $hour)."h";
                }
            }

            $excess += simulate_hour($year_str, $year, $month, $day, $hour,
                                     $gross_power, $pvnet_power, $pvnet_direct);
        }
        if (++$hour == 24) {
            if ($PV_excess_max < $excess) {
                $PV_excess_max = $excess;
                $PV_excess_max_time = date_str($year_str, $month, $day);
            }
            $excess = 0;
            $hour = 0;
        }
        adjust_day_month();
        if ($month > 12) {
            print $year_str_TMY if !$test && !$debug;
            ($year, $month, $day) = ($year + 1, 1, 1);
        }
        last if $test && ($day - 1) * 24 + $hour == TEST_END;
    }
    print "\n" if $test || !$debug;
    STDOUT->autoflush(0);

    # average sums over $years:
    $PV_gross_sum /= $years;
    $DC_input_losses /= $years;
    $PV_losses /= $years     if $curb;
    $PV_loss_hours /= $years if $curb;
    $PV_net_sum /= $years;
    $PV_net_discarded_sum /= $years;
    $PV_use_loss_hours /= ($sel_items / YearHours * $years) if $curb;
    # also undo: factor out $load_scale for optimizing the inner loop
    $PV_use_loss_sum *= $load_scale / $years if $curb;
    $PV_used_sum   *= $load_scale / $years;
    $PV_excess_sum *= $load_scale / $years;
    $PV_excess_max *= $load_scale;
    $grid_feed_sum *= $load_scale / $years;
    # $sum_needed = rls($sum_needed);
    # die "Internal error: load sum = $load_sum vs. needed = $sum_needed"
    #     if $load_sum != $sum_needed;

    if (defined $capacity) {
        # undo: factor out $load_scale for optimizing the inner loop
        $capacity *= $load_scale_never_0;
        $bypass *= $load_scale_never_0 if defined $bypass;
        $soc -= $soc_min;
        $soc_max *= $load_scale_never_0;
        $soc_min *= $load_scale_never_0;
        $soc_max_reached *= $load_scale_never_0;
        $soc     *= $load_scale;
        # also average storage-related sums over $years:
        $dis_feed_sum        *= $load_scale / $years;
        $charge_sum          *= $load_scale / $years;
        $dischg_sum          *= $load_scale / $years;
        $charging_loss       *= $load_scale / $years;
        $spill_loss          *= $load_scale / $years;
        $coupling_loss       *= $load_scale / $years;
        $PV_used_via_storage *= $load_scale / $years;
    }
}

open($DA,'>', $dump_avg) || die "cannot open '$dump_avg' for writing: $!"
    if defined $dump_avg;
simulate();
close $DA if defined $dump_avg;

my $pny = $sel_hours / $hours_a_day * $years;
my $sny = $sn / $years;
for (my $hour = 0; $hour < 24; $hour++) {
    $PV_per_hour        [$hour] = round( $PV_by_hour[$hour] / never_0($pny));
    $PV_max_hour        [$hour] = round($PV_max_hour[$hour]);
    $grid_feed_per_hour [$hour] = rls($grid_feed_per_hour[$hour] * $sny);
    $grid_feed_max_hour [$hour] = rls($grid_feed_max_hour[$hour]);
    if (defined $capacity) {
        $charge_per_hour[$hour] = rls(   $charge_per_hour[$hour] * $sny);
        $dischg_per_hour[$hour] = rls(   $dischg_per_hour[$hour] * $sny);
        $charge_max_hour[$hour] = rls(   $charge_max_hour[$hour] * $sny);
        $dischg_max_hour[$hour] = rls(   $dischg_max_hour[$hour] * $sny);
    }
}

################################################################################
# statistics output

my $nominal_power_txt= $en ? "nominal PV power"     : "PV-Nennleistung";
my $due_to           = $en ? "due to"               : "durch";
my $because          = $en ? "because"              : "weil";
my $Adapted_txt      = $en ? "adapted"              : "Adaptierte";
my $gross_max_txt    = $en ? "max gross PV power"   : "Max. PV-Bruttoleistung";
my $net_max_txt      = $en ? "max net PV power"     : "Max. PV-Nettoleistung";
my $curb_txt         = $en ? "inverter output curb" : "WR-Ausgangs-Drosselung";
my $pvsys_eff_txt    = $en ? "PV system efficiency" : "PV-System-Wirkungsgrad";
my $own_txt          = $en ? "PV own use"           : "PV-Eigenverbrauch";
my $own_ratio_txt    = $own_txt .  ($en ? " ratio"  : "santeil");
my $own_storage_txt  = $en ?"PV own use via storage":"PV-Nutzung über Speicher";
my $load_cover_txt   = $en ? "use coverage ratio"   : "Eigendeckungsanteil";
my $PV_gross_txt     = $en ? "PV gross yield"       : "PV-Bruttoertrag";
my $DC_limit_loss_txt= $en ? "loss by $limit_txt":"Verlust durch DC-Begrenzung";
my $PV_DC_txt        = $en ? "PV DC yield"          : "PV-DC-Ertrag";
my $PV_net_txt       = $en ? "PV net yield"         : "PV-Nettoertrag";
my $PV_txt   = $DC_coupled ? $PV_DC_txt             : $PV_net_txt;
my $of_yield         = $en ? "of $PV_txt"     :"des $PV_txt"."s (Nutzungsgrad)";
my $of_consumption   = $en ? "of consumption" : "des Verbrauchs (Autarkiegrad)";
my $PV_loss_txt      = $en ? "PV yield net loss"    : "PV-Netto-Ertragsverlust";
my $load_adapted_txt = "$Adapted_txt ".
                      ($en ? ($load_min ? "minimal" : "constant")." load"
                           : ($load_min ? "minimale": "konstante")." Last".
                             ($load_min ? " " : ""));
my @weekdays         = $en ? ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
                           : ("Mo", "Di", "Mi", "Do", "Fr", "Sa", "So");
my $dx = $load_days - 1;
my $load_during_txt  = ($load_days == 7 ? "" :
           ($en ? " on Mon".($load_days == 1 ? "days"   : "..$weekdays[$dx]")
                : " an Mo" .($load_days == 1 ? "ntagen" : "..$weekdays[$dx]"))).
                       from_to_hours_str($load_from, $load_to)
                       if defined $load_const;
my $usage_loss_txt   = $en ? "$own_txt net loss"    :  $own_txt."sverlust";
my $net_appr         = $en ? "net"                  : "netto";
$net_appr           .=($en ?" - just approximately -":" - nur näherungsweise -")
    if defined $capacity
    && ($AC_coupled || (!defined $bypass || ($bypass != 0 || $bypass_spill)));
my $by_txt           = $en ? "by"                   : "durch";
my $by_curb = "$by_txt $curb_txt";
my $with             = $en ? "with"                 : "mit";
my $without          = $en ? "without"              : "ohne";
my $PV_excess_txt    = $en ? "PV excess"            : "PV-Überschuss";
my $Max_txt          = $en ? "max"                  : "Max.";
my $grid_feed_txt    = $en ? "grid feed-in"         : "Netzeinspeisung";
my $dis_feed_txt     =
    $grid_feed_txt   .($en ? " via storage"         : " via Speicher");
my $charge_txt       = $en ? "charge (before storage losses)"
                           : "Ladung (vor Speicherverlusten)";
my $dischg_txt       = $en ? "discharge"            : "Entladung";
my $after            = $en ? "after"                : "nach";
my $sum_txt          = $en ? "sums"                 : "Summen";
my $each             = $en ? "each"                 : "alle";
my $on_average_txt   = $en ? "on average over $years years"
                           : "im Durchschnitt über $years Jahre";
my $capacity_txt     = $en ? "storage capacity"     : "Speicherkapazität";
my $soc_txt          = $en ? "state of charge"      : "Ladezustand";
my $max_soc_txt      = $en ? "max SoC"              : "max. Ladehöhe";
my $max_dod_txt      = $en ? "max DoD"              : "max. Entladetiefe";
my $coupled_txt = ($AC_coupled ? "AC" : "DC").($en ? " coupled": "-gekoppelt");
my $optimal_charge   = $en ? "optimal charging strategy (power not consumed)"
                           :"Optimale Ladestrategie (nicht gebrauchte Energie)";
my $bypass_txt       = $en ? "storage bypass"       : "Speicher-Umgehung";
my $spill_txt        = !$bypass_spill ? "" :
                      ($en ?  " and for surplus"    : " und für Überschuss");
my $max_chgrate_txt  = $en ? "max charge rate"      : "max. Laderate";
my $optimal_discharge= $en ? "optimal discharging strategy (as much as needed)"
                           :"Optimale Entladestrategie (so viel wie gebraucht)";
my $feed_txt        = $excl_feed
                    ? ($en ? "exclusive feed decision at"
                           : "Exkl. Einspeisung Grenzwert")
                    : ($const_feed
                    ? ($en ? "constant "            : "Konstant")
                    :  $comp_feed
                    ? ($en ? "compensation "        : "Kompensations")
                    : ($en ?  "limited "            : "Maximal"))
                     .($en ? "feed-in"              : "einspeisung");
my $feed_during_txt= $const_feed ? from_to_hours_str($feed_from, $feed_to) : "";
my $max_disrate_txt  = $en ? "max discharge rate"   : "max. Entladerate";
my $ceff_txt         = $en ? "charging efficiency"  : "Lade-Wirkungsgrad";
my $seff_txt         = $en ? "storage efficiency"   : "Speicher-Wirkungsgrad";
my $ieff_txt         = $en ?"inverter efficiency":"Wechselrichter-Wirkungsgrad";
my $ieff2_txt        = $en ? "discharge inverter efficiency"
                           : "Entlade-WR-Wirkungsgrad";
my $stored_txt       = $en ? "buffered energy"      : "Zwischenspeicherung";
my $spill_loss_txt   = $en ? "loss by spill"        : "Verlust durch Überlauf";
my $AC_coupl_loss_txt= $en ? "loss with AC coupling":"Verlust mit AC-Kopplung";
my $DC_coupl_loss_txt= $en ?"loss during discharge":"Verlust während Entladung";
my $PV_discarded_txt = $en ? "PV power discarded"   : "Verworfene PV-Leistung";
my $charging_loss_txt= $en ? "charging loss"        : "Ladeverlust";
my $storage_loss_txt = $en ? "storage loss"         : "Speicherverlust";
my $cycles_txt       = $en ? "full cycles" : "Vollzyklen";
# Vollzyklen: Kapazitätsdurchgänge, Kapazitätsdurchsatz
my $of_eff_cap_txt   = $en ? "of effective capacity"
                           : "der effektiven Kapazität";
my $ph               = $en ? "p. hour"              : "je Std";
my $c_txt            = $en?"average charge power $ph":"Mittlere Ladeleist. $ph";
my $c_Txt            = $en ? $c_txt               : "Mittlere Ladeleistung $ph";
my $c_max_txt        = $en?"maximal charge power $ph":"Maximale Ladeleist. $ph";
my $c_max_Txt        = $en ? $c_max_txt           : "Maximale Ladeleistung $ph";
my $C_txt            = $en?"avg discharge power $ph":"Mittl. Entladeleist. $ph";
my $C_Txt            = $en ? $C_txt            : "Mittlere Entladeleistung $ph";
my $C_max_txt        = $en?"max discharge power $ph":"Maxim. Entladeleist. $ph";
my $C_max_Txt        = $en ? $C_max_txt        : "Maximale Entladeleistung $ph";
my $P_txt            = $en ? "average PV power $ph" :"Mittlere PV-Leistung $ph";
my $P_max_txt        = $en ? "maximal PV power $ph" :"Maximale PV-Leistung $ph";
my $F_txt            = $en?"avg grid feed power $ph":"Mit. Einspeiseleist. $ph";
my $F_Txt            = $en ? $F_txt          : "Mittlere Einspeiseleistung $ph";
my $F_max_txt        = $en?"max grid feed power $ph":"Max. Einspeiseleist. $ph";
my $F_max_Txt        = $en ? $F_max_txt      : "Maximale Einspeiseleistung $ph";
my $dischg_after_txt = "$dischg_txt $after $storage_loss_txt";

my $PV_DC_sum = $PV_gross_sum * $pvsys_eff;
my $PV_sum = !$DC_coupled ? $PV_net_sum - $PV_net_discarded_sum :
    $PV_DC_sum - $PV_net_discarded_sum / $inverter_eff_never_0;
my $coupl_loss_txt = $AC_coupled ? $AC_coupl_loss_txt : $DC_coupl_loss_txt;
my $own_ratio = round_percent($PV_sum ? $PV_used_sum / $PV_sum : 1);
my $load_cover= round_percent($sel_load_sum ? $PV_used_sum / $sel_load_sum : 1);
my $cycles = 0;
my $of_eff_cap = sprintf("($of_eff_cap_txt %d Wh)", round($eff_capacity))
    if defined $capacity;
if (defined $capacity) { # also for future loss when discharging the rest:
    my $chg_loss_alt = ($charge_eff ? $charge_sum * (1 / $charge_eff - 1) : 1);
    if ($charge_eff) {
        my $discrepancy = $charging_loss - $chg_loss_alt;
        die "Internal error: charging loss calculation discrepancy ".
            "$discrepancy: $charging_loss vs. $chg_loss_alt"
            if abs($discrepancy) > 0.001; # 1 mWh
    }
    $storage_loss = $charge_sum * (1 - $storage_eff); # including future dischg
    $cycles = round($dischg_sum / $storage_eff_never_0 / # really storage_eff ?
                    $eff_capacity) if $eff_capacity != 0;
    if ($DC_coupled) { # not really for DC-coupled, as inverter needed anyway
        $DC_feed_loss =
            ($PV_used_sum + $grid_feed_sum - $dischg_sum * $inverter2_eff) *
            ($inverter_eff == 0 ? 1 : 1 / $inverter_eff - 1);
        $coupling_loss += $DC_feed_loss;
    }
    # future losses on discharging:
    $soc *= $storage_eff;
    $coupling_loss += $soc * (1 - $inverter2_eff);
    $soc *= $inverter2_eff;
}

sub save_statistics {
    my ($file, $type, $max, $hourly, $daily, $weekly, $monthly, $season) = @_;
    return unless $file;
    open(my $OU, '>',$file) or die "Could not open statistics file $file: $!\n";

    my $nominal_sum = $#PV_nomin == 0 ? $PV_nomin[0] : "$PV_nomin";
    my $limits_sum  = $#PV_limit == 0 ? $PV_limit[0] : "$PV_limit";
    print $OU "$consumpt_txt$timeframe in kWh,$profile_txt,";
    print $OU "$b_txt in W".from_to_hours_str(NIGHT_START, NIGHT_END).",";
    print $OU "$B_txt in W".from_to_hours_str(BRIGHT_START, BRIGHT_END).",";
    print $OU "$load_adapted_txt in W$load_during_txt," if defined $load_const;
    print $OU "$m_txt in W $load_min_time,";
    print $OU "$M_txt in W $load_max_time,";
    print $OU "$pv_data_txt$plural_txt".($tmy ? " $during_txt $TMY" : "").
        "$only_during,$pvsys_eff_txt,$ieff_txt\n";
    print $OU round_1000($sel_load_sum).",$load_profile,"
        .rls($night_avg_load).",".rls($bright_avg_load).",";
    print $OU "$load_const," if defined $load_const;
    print $OU "$base_load,$load_max,".join(",", @PV_files).","
        .round_percent($pvsys_eff).",$inverter_eff\n";

    print $OU "$nominal_power_txt in Wp,$limit_txt in W $none0_txt,"
        ."$gross_max_txt in W $PV_gross_max_time,"
        ."$net_max_txt in W $PV_net_max_time,"
        ."$curb_txt in W,$PV_DC_txt,$DC_limit_loss_txt,"
        ."$Max_txt $PV_excess_txt in Wh $PV_excess_max_time,"
        ."$own_ratio_txt $of_yield,$load_cover_txt $of_consumption,";
    print $OU ",$D_txt:,".join(",", @load_dist) if defined $load_dist;
    print $OU "\n$nominal_sum,$limits_sum,".round($PV_gross_max).","
        .round($PV_net_max).",".($curb ? $curb : $en ? "none" : "keine").","
        .round_1000($PV_DC_sum).",".round_1000($DC_input_losses).","
        .round($PV_excess_max).","."$own_ratio,$load_cover,";
    print $OU ",$d_txt:,".join(",", @load_factors) if defined $load_factors;
    print $OU "\n";

    if (defined $capacity) {
        print $OU "$capacity_txt $coupled_txt in Wh,"
            .(defined $bypass ? "$bypass_txt in W$spill_txt" : $optimal_charge)
            .",$max_chgrate_txt in C,"
            .(defined $max_feed ? "$feed_txt in W$feed_during_txt"
                                : $optimal_discharge).",$max_disrate_txt in C,"
            ."$ceff_txt,$ieff2_txt,$max_soc_txt,$max_dod_txt,$seff_txt\n";
        print $OU "$capacity,"
            .(defined $bypass ? $bypass : "").",$max_chgrate,"
            .(defined $max_feed ? $max_feed : "").",$max_disrate,"
            ."$charge_eff,$inverter2_eff,"
            .(percent($max_soc) / 100).",".(percent($max_dod) / 100).","
            .(percent($storage_eff) / 100)."\n";
        print $OU "$spill_loss_txt in kWh,"
            ."$charging_loss_txt in kWh,$storage_loss_txt in kWh,"
            ."$coupl_loss_txt in kWh,$dis_feed_txt in kWh,"
            .($excl_feed ? "$PV_discarded_txt in kWh," : "")
            ."$own_storage_txt in kWh,$max_soc_txt in Wh $PV_net_max_time,"
            ."$stored_txt in kWh,$cycles_txt $of_eff_cap\n";
        print $OU "".round_1000($spill_loss * $inverter_eff).","
            .round_1000($charging_loss).",".round_1000($storage_loss).","
            .round_1000($coupling_loss).",".round_1000($dis_feed_sum).","
            .($excl_feed ? round_1000($PV_net_discarded_sum)."," : "")
            .round_1000($PV_used_via_storage).",".round($soc_max_reached)
            .",".round_1000($charge_sum).",$cycles\n";
    }

    print $OU "\n";
    print $OU       "$hour_txt:,".join(",", @hours)."\n";
    print $OU     "$l_txt in W:,".join(",", @load_per_hour)."\n";
    print $OU "$l_min_txt in W:,".join(",", @load_min_hour)."\n";
    print $OU "$l_max_txt in W:,".join(",", @load_max_hour)."\n";
    print $OU     "$P_txt in W:,".join(",", @PV_per_hour)."\n";
    print $OU "$P_max_txt in W:,".join(",", @PV_max_hour)."\n";
    print $OU     "$F_Txt in W:,".join(",", @grid_feed_per_hour)."\n";
    print $OU "$F_max_Txt in W:,".join(",", @grid_feed_max_hour)."\n";
    if (defined $capacity) {
        print $OU     "$c_Txt in W:,".join(",", @charge_per_hour)."\n";
        print $OU "$c_max_Txt in W:,".join(",", @charge_max_hour)."\n";
        print $OU     "$C_Txt in W:,".join(",", @dischg_per_hour)."\n";
        print $OU "$C_max_Txt in W," .join(",", @dischg_max_hour)."\n";
    }
    print $OU "\n";

    my $sum_avg = $sum_txt . ($years > 1 ? " $on_average_txt" : "");
    print $OU "$values_txt$only_during,$PV_gross_txt,"
        .($curb ? "$PV_loss_txt $by_curb," : "")."$PV_net_txt,"
        .($curb ? "$usage_loss_txt $net_appr $by_curb,": "")
        ."$own_txt,$PV_excess_txt,$grid_feed_txt,$consumpt_txt"
        .(defined $capacity ? ",$charge_txt,$dischg_after_txt" : "")
        ."\n";
    print $OU "$sum_avg,".
        round_1000($PV_gross_sum).",".
        ($curb ? round_1000($PV_losses  )."," : "").
        round_1000($PV_net_sum).",".
        ($curb ? round_1000($PV_use_loss_sum)."," : "").
        round_1000($PV_used_sum).",".round_1000($PV_excess_sum).",".
        round_1000($grid_feed_sum ).",".round_1000($sel_load_sum).",".
        (defined $capacity ?
         round_1000(   $charge_sum).",".round_1000($dischg_sum).",," : "").
        "$each in kWh\n";

    my $i = 21 + (defined $capacity ? 6 : 0);
    my $j = $i - 1 + ($max ? $sel_items
                      : $hourly ? $sel_hours
                      : $daily ? 365 : $weekly ? 52 : $monthly ? 12 : 4);
    my ($I, $J) = $curb ? ("J", "K") : ("H", "I");
    print $OU "$type,$PV_gross_txt,"
        ."$PV_net_txt".($curb ? " $without $curb_txt" : "").","
        .($curb ?       "$PV_net_txt $with $curb_txt,": "")
        .($curb ? "$own_txt $without $curb_txt," : "")
        ."$own_txt".($curb ? " $with $curb_txt"  : "").","
        ."$PV_excess_txt,$grid_feed_txt,$consumpt_txt"
        .(defined $capacity ? ",$charge_txt,$dischg_after_txt,$soc_txt" : "")
        ."\n";
    sub SUM { my ($I, $i, $j) = @_;
        return "=ROUND(SUM($I$i:$I$j))";
    }
    print $OU "$sum_avg,".SUM("B", $i, $j).",".SUM("C", $i, $j)
        .",".SUM("D", $i, $j).",".SUM("E", $i, $j).",".SUM("F", $i, $j)
        .",".SUM("G", $i, $j)
        .($curb ? ",".SUM("H", $i, $j).",".SUM("I", $i, $j) : "")
        .(defined $capacity ? ",".SUM($I, $i, $j).",".SUM($J, $i, $j)."," : "")
        .",$each in ".($max ? "m" : "")."Wh\n";

    (my $week, my $days, $hour) = (1, 0, 0);
    ($month, $day) = $season && !$test ? (2, 5) : (1, 1);
    my ($g0, $gross, $p0, $ploss, $n0, $net) = (0, 0, 0, 0, 0, 0);
    my ($h0, $hload, $l0,  $loss, $u0, $used) = (0, 0, 0, 0, 0, 0);
    my ($e0, $exc, $f0, $feed) = (0, 0, 0, 0);
    my ($c0, $chg  , $d0,   $dis, $soc) = (0, 0, 0, 0) if defined $capacity;
    while ($days < 365) {
        my $sel = selected($month, $day, $hour);
        my $tim;

        my $PV_loss = 0;
        if ($sel) {
            if ($weekly) {
                $tim = $week;
            } elsif ($season) {
                $tim = $season== 1 ? ($en ? "spring" : "Frühjahr")
                    : $season == 2 ? ($en ? "summer" : "Sommer")
                    : $season == 3 ? ($en ? "autumn" : "Herbst")
                    :                ($en ? "winter" : "Winter");
            } else {
                $tim = sprintf("%02d", $month);
                $tim = $tim."-".sprintf("%02d", $day )
                    if $daily || $hourly || $max;
                $tim = $tim." ".sprintf("%02d", $hour) if $hourly || $max;
            }
            $gross += $PV_gross_across[$month][$day][$hour];
            $net   += $PV_net_across  [$month][$day][$hour];
            if ($curb) {
                $PV_loss = $PV_loss   [$month][$day][$hour];
                $ploss += $PV_loss;
            }
            if (!$max) {
                $hload   +=        $load[$month][$day][$hour];
                $loss    += $PV_use_loss[$month][$day][$hour] if $PV_loss != 0;
                $used    += $PV_used    [$month][$day][$hour];
                $exc     += $PV_excess  [$month][$day][$hour];
                $feed    += $grid_feed  [$month][$day][$hour];
                if (defined $capacity) {
                    $chg += $charge     [$month][$day][$hour];
                    $dis += $dischg     [$month][$day][$hour];
                    $soc  = $soc        [$month][$day][$hour];
                }
            }
        }

        my $items = $items_by_hour[$month][$day][$hour];
        my $fact  = ($max ? 1000 : 1) / ($max ? $items : 1) / $years;
        my $fact_s = $fact * $load_scale;
        my ($mon_, $day_, $hour_) = ($month, $day, $hour);

        if (++$hour == 24) {
            $hour = 0;
            $days++;
            $week++ if $days % 7 == 0
                && $days < 364; # count Dec-31 as part of week 52
            $season++ if $season && $days % (364 / 4) == 0
                && $days < 364; # count last day as part of winter (1 day more)
        }
        adjust_day_month();

        if ($max || $hourly || $test && $days * 24 + $hour == TEST_END ||
            ($hour == 0
             && ($days == 365 ||
                 ($daily ||
                  $weekly && $days < 364 && $days % 7 == 0 ||
                  $monthly && $day == 1 ||
                  $season && $days < 364 && $days % (364 / 4) == 0)))) {
            for (my $i = 0; $sel && $i < ($max ? $items : 1); $i++) {
                my $minute = $hourly ? ":00" : "";
                if ($max) {
                    $minute = minute_str($i, $items);
                    $hload =          $load_item[$mon_][$day_][$hour_][$i];
                    $loss = $PV_loss == 0 ? 0 :
                            $PV_use_loss_by_item[$mon_][$day_][$hour_][$i];
                    $used =     $PV_used_by_item[$mon_][$day_][$hour_][$i];
                    $exc  =   $PV_excess_by_item[$mon_][$day_][$hour_][$i];
                    $feed =   $grid_feed_by_item[$mon_][$day_][$hour_][$i];
                    if (defined $capacity) {
                        $chg =   $charge_by_item[$mon_][$day_][$hour_][$i];
                        $dis =   $dischg_by_item[$mon_][$day_][$hour_][$i];
                        $soc =      $soc_by_item[$mon_][$day_][$hour_][$i];
                    }
                }
                $g0 += (my $g = $fact   * $gross);
                $p0 += (my $p = $fact   * ($net + $ploss)) if $curb;
                $n0 += (my $n = $fact   *  $net);
                $l0 += (my $l = $fact_s * ($used + $loss)) if $curb;
                $u0 += (my $u = $fact_s *  $used);
                $e0 += (my $e = $fact_s *  $exc );
                $f0 += (my $f = $fact_s *  $feed);
                $h0 += (my $h = $fact_s *  $hload * $years);
                $c0 += (my $c = $fact_s * $chg) if defined $capacity;
                $d0 += (my $d = $fact_s * $dis) if defined $capacity;
                print $OU "$tim$minute, ".round($g).","
                    .($curb ? round($p)."," : "")
                    .round($n).", "
                    .($curb ? round($l)."," : "")
                    .round($u).",".round($e).",".round($f).",".round($h)
                    .(defined $capacity ? ",".round($c).",".round($d).
                                          ",".round($fact_s * $soc) : "")."\n";
            }
            ($gross, $ploss, $net, $exc ) = (0, 0, 0, 0),
            ($hload, $loss, $used, $feed) = (0, 0, 0, 0);
            ($chg, $dis)                  = (0, 0) if defined $capacity;
        }
        ($month, $day) = (1, 1) if $month > 12;
        last if $test && $days * 24 + $hour == TEST_END;
    }
    close $OU;
    sub check_consistent { my ($desc, $ref, $sum, $max) = @_;
        $sum /= 1000 if $max;
        die "Internal error: statistics output sum calculation discrepancy ".
            "for $desc: $ref vs. $sum" if abs($ref - $sum) > 0.001;
    }
    check_consistent($PV_gross_txt  , $PV_gross_sum    , $g0, $max);
    check_consistent($PV_loss_txt   , $PV_losses , $p0 - $n0, $max) if $curb;
    check_consistent($PV_net_txt    , $PV_net_sum      , $n0, $max);
    check_consistent($usage_loss_txt, $PV_use_loss_sum , $l0-$u0, $max) if $curb;
    check_consistent($consumpt_txt  , $sel_load_sum    , $h0, $max);
    check_consistent($own_txt       , $PV_used_sum     , $u0, $max);
    check_consistent($PV_excess_txt , $PV_excess_sum   , $e0, $max);
    check_consistent($grid_feed_txt , $grid_feed_sum   , $f0, $max);
    check_consistent($charge_txt, $charge_sum, $c0, $max) if defined $capacity;
    check_consistent($dischg_txt, $dischg_sum, $d0, $max) if defined $capacity;
}


my $DC_in_losses = $DC_coupled ? $DC_input_losses / $inverter_eff_never_0 : 0;
my $grid_feed_sum_alt = $PV_sum - $PV_used_sum;
$grid_feed_sum_alt -= $DC_in_losses + $coupling_loss + $spill_loss
    + $charging_loss + $storage_loss + $soc if defined $capacity;
my $discrepancy = $grid_feed_sum - $grid_feed_sum_alt;
my $cpl_loss2 = round($coupling_loss - $DC_feed_loss) if defined $capacity;
die "Internal error: overall (loss?) calculation discrepancy $discrepancy: ".
    "grid feed-in $grid_feed_sum vs. $grid_feed_sum_alt =\n".
    "PV ".($DC_coupled ? "DC" : "net")." sum "
    ."$PV_sum - PV used $PV_used_sum".(defined $capacity ?
        " - DC-coupled input limit loss $DC_in_losses".
        " - DC feed loss $DC_feed_loss - loss by spill $spill_loss".
        " - charging loss $charging_loss - storage loss $storage_loss".
        " - coupling loss by 2nd inverter $cpl_loss2".
        " - SoC ".round($soc) : "") if abs($discrepancy) > 0.001; # 1 mWh
$coupling_loss = 0 if $DC_coupled; # as inverter needed for PV anyway

my $max_txt = $items_per_hour == 60
             ? ($en ? "minute" : "Minute")
             : ($en ? "point in time" : "Zeitpunkt");
my $date_txt  = $en ? "date"   : "Datum";
my $week_txt  = $en ? "week"   : "Woche";
my $month_txt = $en ? "month"  : "Monat";
my $season_txt= $en ? "season" : "Saison";
save_statistics($max     , $max_txt   , 1, 0, 0, 0, 0, 0);
save_statistics($hourly  , $hour_txt  , 0, 1, 0, 0, 0, 0);
save_statistics($daily   , $date_txt  , 0, 0, 1, 0, 0, 0);
save_statistics($weekly  , $week_txt  , 0, 0, 0, 1, 0, 0);
save_statistics($monthly , $month_txt , 0, 0, 0, 0, 1, 0);
save_statistics($seasonly, $season_txt, 0, 0, 0, 0, 0, 1);

my $at             = $en ? "with"                 : "bei";
my $and            = $en ? "and"                  : "und";
my $by_curb_at     = $en ? "$by_curb at"          : "$by_curb auf";
my $yield_portion  = $en ? "yield portion"        : "Ertragsanteil";
# PV-Abregelungsverlust"
my $nominal_sum = $#PV_nomin == 0 ? "" : " = $PV_nomin Wp";
my $limits_sum = $total_limit == $total_nomin ? "" :
    ", $limit_txt: ".$PV_limit." W".($#PV_limit == 0 ? "" : " $none0_txt");

print "$energy_txt $on_average_txt\n" if $years > 1;
print "\n" unless $test;
print "$nominal_power_txt $de1           =" .W($nominal_power_sum)."p$nominal_sum".
    "$only_during$limits_sum\n";
print "$gross_max_txt $en4     =".W($PV_gross_max)." $PV_gross_max_time\n";
if ($verbose) {
    print_arr_day(    "$P_txt $en1= ", \@PV_per_hour);
    print_arr_day("$P_max_txt $en1= ", \@PV_max_hour);
    print_arr    ("$V_txt $per3$en1     = ", \@PV_by_hour,
                   $PV_gross_sum * $years, 0, 21, 3);
    print_arr("$V_txt $per_m$de1        = ", \@PV_by_month,
                   $PV_gross_sum * $years, 1, 12, 1);
}
print "$PV_gross_txt $en1            =".kWh($PV_gross_sum)."\n";
print "$PV_DC_txt $en1               =" .kWh($PV_DC_sum).
    ", $pvsys_eff_txt ".percent($pvsys_eff)."%\n";
print "$DC_limit_loss_txt$en5 =".kWh($DC_input_losses)."\n" if $total_limit != $total_nomin;
print "$PV_loss_txt $en2 $en2 $en2  ="       .kWh($PV_losses).
    " $during_txt ".round($PV_loss_hours)." h $by_curb_at $curb W\n" if $curb;
print "$PV_net_txt $en2             =" .kWh($PV_net_sum).
    " $at $ieff_txt ".percent($inverter_eff)."%\n";
print "$net_max_txt $en4$en1      =".W($PV_net_max)." $PV_net_max_time\n";
#print "$yield_portion $daytime  =   $yield_daytime %\n";
#my $yield_daytime =
#    percent($PV_net_sum ? $PV_net_bright_sum / $PV_net_sum : 0);

print "\n";
print "$consumpt_txt $de2                =".kWh($sel_load_sum)."$timeframe\n";
print "$load_adapted_txt $en1 ".($load_min ? "$en1" : "")."$en3 ="
    .W($load_const)."$load_during_txt\n" if defined $load_const;
if (defined $capacity) {
    print "\n".
        "$capacity_txt $en1          =" .W($capacity)."h $with"
        ." $max_soc_txt ".percent($max_soc)."%,"
        ." $max_dod_txt ".percent($max_dod)."%"
        .", $coupled_txt\n";
    print "$optimal_charge" unless defined $bypass;
    print"$bypass_txt $en3          =".W($bypass).$spill_txt if defined $bypass;
    print ", $max_chgrate_txt $max_chgrate C\n";
    print "$optimal_discharge" unless defined $max_feed;
    print "$feed_txt ".($excl_feed ? "$en1" : "$en3   "
        .($comp_feed ? "$en1" : "     ".($const_feed ? "" : " ")))
        ."=".W($max_feed)."$feed_during_txt" if defined $max_feed;
    print ", $max_disrate_txt $max_disrate C\n";
    print "$spill_loss_txt $en3 $en3 $en3   =".
        kWh($spill_loss * $inverter_eff).
        ($AC_coupled ? " $because $coupled_txt" : "")."\n";
    print "$charging_loss_txt $de2              =".kWh($charging_loss)
        ." $due_to $ceff_txt ".percent($charge_eff)."%\n";
    print "$storage_loss_txt $en3            =".kWh($storage_loss)
        ." $due_to $seff_txt ".percent($storage_eff)."%\n";
    print "$coupl_loss_txt $en3 $en1 ".($AC_coupled ? "$de2" : "")."=".
       kWh($coupling_loss)." $due_to $ieff2_txt ".percent($inverter2_eff)."%\n";
    print "$PV_discarded_txt $en4     =".
        kWh($PV_net_discarded_sum)."\n" if $excl_feed;
    print "$own_storage_txt $en2   =".kWh($PV_used_via_storage)."\n";
    print "$dis_feed_txt$en4=".kWh($dis_feed_sum)."\n";
    if ($verbose && defined $capacity) {
        print_arr_day("$hour_txt                     $en2 = ", \@hours);
        print_arr_day(    "$c_txt$de2= ", \@charge_per_hour);
        print_arr_day("$c_max_txt$de2= ", \@charge_max_hour);
        print_arr_day(    "$C_txt"." = ", \@dischg_per_hour);
        print_arr_day("$C_max_txt"." = ", \@dischg_max_hour);
    }
    printf "$max_soc_txt $en6              =%5d Wh $soc_max_time\n",
        round($soc_max_reached);
    print "$stored_txt $en4        =" .kWh($charge_sum)." $after $charging_loss_txt\n";
    printf "$cycles_txt $de1                =  %3d $of_eff_cap\n", round($cycles);
    print "\n";
}

print "$own_txt $en4 $en3         =" .kWh($PV_used_sum)."\n";
print "$usage_loss_txt $en3$en3  =" .kWh($PV_use_loss_sum)
    ." $net_appr $during_txt ".round($PV_use_loss_hours)
    ." h $by_curb_at $curb W\n" if $curb;
print "$PV_excess_txt $en4              =" .kWh($PV_excess_sum)."\n";
print "$Max_txt $PV_excess_txt $en5         =" .kWh($PV_excess_max)
    ." $PV_excess_max_time\n";
print "$grid_feed_txt $en3            =" .kWh($grid_feed_sum)."\n";
print_arr_day(     "$F_txt = ", \@grid_feed_per_hour) if $verbose;
print_arr_day( "$F_max_txt = ", \@grid_feed_max_hour) if $verbose;
print "$own_ratio_txt $en4 $en4  =  ".sprintf("%3d", percent($own_ratio))
    ." % $of_yield\n";
my $load_cover_str = sprintf("%3d", percent($load_cover));
print "$load_cover_txt $en1        =  $load_cover_str % $of_consumption\n";
