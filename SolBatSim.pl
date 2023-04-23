#!/usr/bin/perl
################################################################################
# Eigenverbrauchs-Simulation mit stündlichen PV-Daten und einem mindestens
# stündlichen Lastprofil, optional mit Begrenzung/Drosselung des Wechselrichters
# und optional mit Stromspeicher (Batterie o.ä.)
#
# Nutzung: Solar.pl <Lastprofil-Datei> [<Jahresverbrauch in kWh>]
#   (<PV-Daten-Datei> [<Nominalleistung in Wp> [<WR-Eingangsbegrenzung in W>]])+
#   [-only <Jahr>[-<Monat>[-<Tag>[:<Stunde>]]]]
#   [-dist <relative Lastverteilung über den Tag pro Stunde 0,..,23>
#   [-bend <Lastverzerrungsfaktoren tgl. pro Std. 0,..,23, sonst 1>
#   [-load <konstante Last in W> [<Zahl der Tage pro Woche, sonst 5>:
#          <von Uhrzeit, sonst 8 Uhr>..<bis Uhrzeit, sonst 16 Uhr>]]
#   [-avg_hour] [-verbose]
#   [-peff <PV-System-Wirkungsgrad in %, ansonsten von PV-Daten-Datei>]
#   [-capacity <Speicherkapazität Wh, ansonsten 0 (kein Batterie)>]
#   [-ac] {AC-gekoppelter Speicher, ohne Verlust durch Überlauf}
#   [-pass [spill] <Speicher-Umgehung in W, opt. auch bei Überfluss>]
#   [-feed (max <begrenzte bedarfsgerechte Entladung aus Speicher in W>
#           | [<von Uhrzeit, sonst 0 Uhr>..<bis Uhrzeit, sonst 24 Uhr>]
#             <konstante Entladung aus Speicher in W> )]
#   [-max_charge <Ladehöhe in %, sonst 90> [<max Laderate, sonst 1 C>]]
#   [-max_discharge <Entladetiefe in %, sonst 90> [<Rate, sonst 1 C>]]
#   [-ceff <Lade-Wirkungsgrad in %, ansonsten 94>]
#   [-seff <Speicher-Wirkungsgrad in %, ansonsten 95>]
#   [-ieff <Wechselrichter-Wirkungsgrad in %, ansonsten 94>]
#   [-test <Lastpunkte pro Stunde, für Prüfberechnung über 24 Stunden>]
#   [-en] [-tmy] [-curb <Wechselrichter-Ausgangs-Drosselung in W>]
#   [-hour <Statistik-Datei>] [-day <Stat.Datei>] [-week <Stat.Datei>]
#   [-month <Stat.Datei>] [-season <Stat.Datei>] [-max <Stat.Datei>]
# Alle Uhrzeiten sind in lokaler Winterzeit (MEZ, GMT+1/UTC+1).
# Mit "-en" erfolgen die Textausgaben auf Englisch. Fehlertexte sind englisch.
# Wenn PV-Daten für mehrere Jahre gegeben sind, wird der Durchschnitt berechnet
# oder mit Option "-tmy" Monate für ein typisches meteorologisches Jahr gewählt.
# Mit den Optionen "-hour"/"-day"/"-week"/"-month" wird jeweils eine CSV-Datei
# mit gegebenen Namen mit Statistik-Daten pro Stunde/Tag/Woche/Monat erzeugt.
#
# Beispiel:
# Solar.pl Lastprofil.csv 3000 Solardaten_1215_kWh.csv 1000 -curb 600 -tmy
#
# PV-Daten können bezogen werden von https://re.jrc.ec.europa.eu/pvg_tools/de/
# Wähle den Standort und "DATEN PRO STUNDE", setze Häkchen bei "PV-Leistung".
# Optional "Installierte maximale PV-Leistung" und "Systemverlust" anpassen.
# Bei Nutzung von "-tmy" Startjahr 2008 oder früher, Endjahr 2018 oder später.
# Dann den Download-Knopf "csv" drücken.
#
################################################################################
# Simulation of actual own consumption of photovoltaic power output according
# to load profiles with a resolution of at least one hour, typically per minute.
# Optionally takes into account input limit and output crop of solar inverter.
# Optionally with energy storage (using a battery or the like).
#
# Usage: Solar.pl <load profile file> [<consumption per year in kWh>]
#          (<PV datafile> [<nominal power in Wp> [<inverter input limit in W]])+
#          [-only <year>[-<month>[-<day>[:<hour]]]]
#          [-dist <relative load distribution over each day, per hour 0,..,23>
#          [-bend <load distort factors for hour 0,..,23 each day, default 1>
#          [-load <constant load in W> [<count of days per week, default 5>:
#                 <from hour, default 8 o'clock>..<to hour, default 16>]]
#          [-avg_hour] [-verbose]
#          [-peff <PV system efficiency in %, default from PV data file(s)>]
#          [-capacity <storage capacity in Wh, default 0 (no battery)>]
#          [-ac] {AC-coupled charging after inverter, without loss by spill}
#          [-pass [spill] <storage bypass in W, optionally also on surplus>]
#          [-feed (max <limited feed-in from storage in W according to load>
#                  | [<von Uhrzeit, sonst 0 Uhr>..<bis Uhrzeit, sonst 24 Uhr>]
#                    <constant feed-in from storage in W> )]
#          [-max_charge <SoC in %, default 90> [<max charge rate, default 1 C>]]
#          [-max_discharge <DoD in %, default 90> [<max rate, default 1 C>]]
#          [-ceff <charging efficiency in %, default 94>]
#          [-seff <storage efficiency in %, default 95>]
#          [-ieff <inverter efficiency in %, default 94>]
#          [-test <load points per hour, for debug calculation over 24 hours>]
#          [-en] [-tmy] [-curb <inverter output power limitation in W>]
#          [-hour <statistics file>] [-day <stat file>] [-week <stat file>]
#          [-month <stat file>] [-season <file>] [-max <stat file>]
# All times (hours) are in local winter time (CET, GMT+1/UTC+1).
# Use "-en" for text output in English. Error messages are all in English.
# When PV data for more than one year is given, the average is computed, while
# with the option "-tmy" months for a typical meteorological year are selected.
# With each the options "-hour"/"-day"/"-week"/"-month" a CSV file is produced
# with the given name containing with statistical data per hour/day/week/month.
#
# Example:
# Solar.pl loadprofile.csv 3000 solardata_1215_kWh.csv 1000 -curb 600 -tmy -en
#
# PV data can be obtained from https://re.jrc.ec.europa.eu/pvg_tools/
# Select location, click "HOURLY DATA", and set the check mark at "PV power".
# Optionally may adapt "Installed peak PV power" and "System loss" already here.
# For using TMY data, choose Start year 2008 or earlier, End year 2018 or later.
# Then press the download button marked "csv".
#
# (c) 2022-2023 David von Oheimb - License: MIT - Version 2.3
################################################################################

use strict;
use warnings;

my $test         = 0;   # unless 0, number of test load points per houra
die "Missing command line arguments" if $#ARGV < 0;
my $load_profile = shift @ARGV unless $ARGV[0] =~ m/^-/; # file name
my $consumption  = shift @ARGV # kWh/year, default is implicit from load profile
    if $#ARGV >= 0 && $ARGV[0] =~ m/^[\d\.]+$/;

my @load_dist;          # if set, relative load distribution per hour each day
my @load_factors;       # load distortion factors per hour, on top of @load_dist
my $load_const;         # constant load in W, during certain times as follows:
my $avg_hour      = 0;  # use only the average of load items per hour
my $verbose       = 0;  # verbose output, including averages/day for each hour
my $load_days    =  5;  # count of days per week with constant load
my $load_from    =  8;  # hour of constant load begin
my $load_to      = 16;  # hour of constant load end
my $first_weekday = 4;  # HTW load profiles start on Friday in 2010, Monday == 0

use constant YearHours => 24 * 365;
use constant TimeZone => 1; # CET/MEZ

my @PV_files;
my @PV_nomin;       # nominal/maximal PV output(s), default from PV data file(s)
my @PV_limit;       # power limit at inverter input, default 0 (none)
my ($lat, $lon);    # optional, from PV data file(s)
my $pvsys_eff;      # PV system efficiency, default from PV data file(s)
my $inverter_eff;   # inverter efficiency; default see below
my $capacity;       # nominal storage capacity in Wh on average degradation
my $bypass_spill;   # bypass storage on surplus (i.e., when storge is full)
my $bypass;         # direct feed to inverter in W, bypassing storage
my $max_feed;       # maximal feed-in in W from storage
my $const_feed = 1; # constant feed-in, relevant only if defined $max_feed
my $feed_from = 0;  # hour of constant feed-in begin
my $feed_to   = 24; # hour of constant feed-in end
my $AC_coupled;     # by default, charging is DC-coupled (w/o inverter)
my $soc       = .9; # maximal state of charge (SoC); default 90%
my $dod       = .9; # maximal depth of discharge (DoD); default 90%
my $max_charge = 1; # maximal charge rate; default 1 C
my $max_dischg = 1; # maximal discharge rate; default 1 C
my $charge_eff;     # charge efficiency; default see below
my $storage_eff;    # storage efficiency; default see below
# storage efficiency and discharge efficiency could have been combined
my $nominal_power_sum = 0;

while ($#ARGV >= 0 && $ARGV[0] =~ m/^\s*[^-]/) {
    push @PV_files, shift @ARGV; # PV data file
    push @PV_nomin, $#ARGV >= 0 && $ARGV[0] =~ m/^[\d\.]+$/ ? shift @ARGV : 0;
    push @PV_limit, $#ARGV >= 0 && $ARGV[0] =~ m/^[\d\.]+$/ ? shift @ARGV : 0;
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
        unless 0 <= $eff && $eff <= 100;
    return $eff / 100;
}
sub str_arg {
    die "Missing arg for $ARGV[0] option" if $#ARGV < 1;
    shift @ARGV;
    return shift @ARGV;
}
sub array_arg {
    my $opt = shift;
    my $arg = shift; # may be undefined
    $arg =~ tr/ /,/ if defined $arg;
    my $min = shift;
    my $max = shift;
    my $default = shift;

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
my ($en, $date, $tmy, $curb,
    $max, $hourly, $daily, $weekly,  $monthly, $seasonly);
while ($#ARGV >= 0) {
    if      ($ARGV[0] eq "-test"    ) { $test         = num_arg();
    } elsif ($ARGV[0] eq "-en"      ) { $en           =  no_arg();
    } elsif ($ARGV[0] eq "-only"    ) { $date         = str_arg();
    } elsif ($ARGV[0] eq "-dist"    ) { $load_dist    = str_arg();
    } elsif ($ARGV[0] eq "-bend"    ) { $load_factors = str_arg();
    } elsif ($ARGV[0] eq "-load"    ) { $load_const   = num_arg();
                                       ($load_days, $load_from, $load_to)
                                           = ($1, $2, $3) if $#ARGV >= 0 &&
                                           $ARGV[0] =~ m/^(\d+):(\d+)\.\.(\d+)$/
                                           && shift @ARGV;
    } elsif ($ARGV[0] eq "-avg_hour") { $avg_hour     = no_arg();
    } elsif ($ARGV[0] eq "-verbose" ) { $verbose      = no_arg();
    } elsif ($ARGV[0] eq "-peff"    ) { $pvsys_eff    = eff_arg();
    } elsif ($ARGV[0] eq "-tmy"     ) { $tmy          =  no_arg();
    } elsif ($ARGV[0] eq "-curb"    ) { $curb         = num_arg();
    } elsif ($ARGV[0] eq "-capacity") { $capacity     = num_arg();
    } elsif ($ARGV[0] eq "-ac"      ) { $AC_coupled   =
                                        $bypass_spill =  no_arg();
    } elsif ($ARGV[0] eq "-pass"    ) { $bypass_spill = 1 if $#ARGV >= 1 &&
                                            $ARGV[1] eq "spill" && shift @ARGV;
                                        $bypass       = num_arg();
    } elsif ($ARGV[0] eq "-feed"    ) { $const_feed   = 0 if $#ARGV >= 1 &&
                                            $ARGV[1] eq "max" && shift @ARGV;
                                       ($feed_from, $feed_to)
                                           = ($1, $2) if $#ARGV >= 1 &&
                                           $ARGV[1] =~ m/^(\d+)\.\.(\d+)$/
                                           && shift @ARGV;
                                        $max_feed     = num_arg();
    } elsif ($ARGV[0] eq "-max_charge") { $soc        = eff_arg();
                                        $max_charge = shift @ARGV if $#ARGV >= 0
                                            && $ARGV[0] =~ m/^([\d\.]+)$/;
    } elsif ($ARGV[0] eq "-max_discharge") { $dod     = eff_arg();
                                        $max_dischg = shift @ARGV if $#ARGV >= 0
                                            && $ARGV[0] =~ m/^([\d\.]+)$/;
    } elsif ($ARGV[0] eq "-ceff"    ) { $charge_eff   = eff_arg();
    } elsif ($ARGV[0] eq "-seff"    ) { $storage_eff  = eff_arg();
    } elsif ($ARGV[0] eq "-ieff"    ) { $inverter_eff = eff_arg();
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
use constant TEST_PV_LIMIT   =>    0; # in W
use constant TEST_LOAD       =>  900; # in W
use constant TEST_LOAD_LEN   =>
    TEST_END > TEST_DROP_END ? TEST_LENGTH - TEST_DROP_LEN
    : (TEST_END <= TEST_DROP_START ? TEST_END - TEST_START
       : TEST_DROP_START - TEST_START);
my $test_load = defined $consumption ?
    $consumption * 1000 / TEST_LOAD_LEN : TEST_LOAD if $test;
if ($test) {
    $load_profile  = "test load data";
    my $pv_nomin   = $#PV_nomin < 0 ? TEST_PV_NOMIN : $PV_nomin[0]; # in W
    my $pv_limit   = $#PV_limit < 0 ? TEST_PV_LIMIT : $PV_limit[0]; # in W
    my $pv_power   = defined $pvsys_eff ? $pv_nomin * $pvsys_eff :
        sprintf("%.0e", $pv_nomin); # in W
    # after PV system losses, which may be derived as follows:
    $pvsys_eff     = $pv_nomin ? $pv_power / $pv_nomin : 0.92
        unless defined $pvsys_eff;
    $inverter_eff  =  0.8 unless defined $inverter_eff;
    $consumption = $test_load/1000 * TEST_LOAD_LEN unless defined $consumption;
    push @PV_files, "test PV data" if $#PV_files < 0;
    push @PV_nomin, $pv_nomin      if $#PV_nomin < 0;
    push @PV_limit, $pv_limit      if $#PV_limit < 0;
    $charge_eff    =  0.9 if defined $capacity && !defined $charge_eff;
    my $pv_net_bat =  600; # in W, after inverter losses
    $storage_eff   =  $pv_net_bat / ($pv_power * $charge_eff * $inverter_eff)
        if defined $capacity && !defined $storage_eff;
}

die "Missing PV data file name" if $#PV_nomin < 0;
die "-only option does not have form (*|YYYY)[-(*|MM)[-(*|DD)[:(*|HH)]]]"
    if $date &&
    !($date =~ m/^(\*|\d\d?\d?\d?)(-(\*|\d\d?)(-(\*|\d\d?)(:(\*|\d\d?))?)?)?$/);
my ($sel_year, $sel_month, $sel_day, $sel_hour) = ($1, $3, $5, $7) if $date;
die "with -tmy, the year given with the -only option must be '*'"
    if $tmy && defined $sel_year && $sel_year ne "*";

die "Missing load profile file name - should be first CLI argument"
    unless defined $load_profile;
@load_dist    = array_arg("dist", $load_dist   , 0, 23, 100);
@load_factors = array_arg("bend", $load_factors, 0, 23, 1);
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
    die "Begin hour for -feed option must be in range 0..24"
        if  $feed_from > 24;
    die "End hour for -feed option must be in range 0..24"
        if  $feed_to > 24;
} else {
    die   "-ac option requires -capacity option" if defined $AC_coupled;
    die "-pass option requires -capacity option" if defined $bypass;
    die "-feed option requires -capacity option" if defined $max_feed;
    die "-ceff option requires -capacity option" if defined $charge_eff;
    die "-seff option requires -capacity option" if defined $storage_eff;
}

sub never_0 { return $_[0] == 0 ? 1 : $_[0]; }
my $pvsys_eff_never_0;
my $inverter_eff_never_0 = never_0($inverter_eff);
my   $charge_eff_never_0 = never_0($charge_eff)  if defined $charge_eff;
my  $storage_eff_never_0 = never_0($storage_eff) if defined $storage_eff;

# deliberately not using any extra packages like Math
sub min { return $_[0] < $_[1] ? $_[0] : $_[1]; }
sub max { return $_[0] > $_[1] ? $_[0] : $_[1]; }
sub round { return int(.5 + shift); }
sub check_consistency {
    return if $test;
    my ($actual, $expected, $name, $file) = (shift, shift, shift, shift);
    die "Got $actual $name rather than $expected from $file"
        if $actual != $expected;
}

sub date_string {
    my $year = shift;
    $year = "YYYY" if $year eq "0";
    return "$year-".sprintf("%02d", shift)."-".sprintf("%02d", shift).
        " ".($en ? "at" : "um")." ".sprintf("%02d", shift);
}
sub time_string {
    return date_string(shift,shift,shift,shift).sprintf(":%02d h", int(shift));
}
sub minute_string {
    my ($item, $items) = (shift, shift);
    my $minute = sprintf(":%02d", int( 60 * $item / $items));
    $minute .=   sprintf(":%02d", int((60 * $item % $items) / $items * 60))
        unless (60 % $items == 0);
    return $minute
}

sub round_1000 { return round(shift() / 1000); }
sub kWh     { return sprintf("%5.2f kWh", shift() / 1000) if $test;
              return sprintf("%5d kWh", round_1000(shift())); }
sub W       { return sprintf("%5d W"  , round(shift())); }
sub percent { return round(shift() *  100); }
sub round_percent { return percent(shift()) / 100; }
sub SUM     { my ($I, $i, $j) = (shift, shift, shift);
              my $div = $max ? "/1000" : "";
              return "=INT(SUM($I$i:$I$j)$div)"; }
sub print_arr_perc {
    my $msg = shift;
    my $arr_ref = shift;
    my ($sum, $start, $end, $inc) = (shift, shift, shift, shift);
    return if $sum == 0;
    print $msg;
    for (my $i = $start; $i <= $end; $i += $inc) {
        my $values = 0;
        for (my $j = 0; $j < $inc; $j++) {
            $values += $arr_ref->[$i + $j];
        }
        printf "%2d%%", percent($values / $sum);
        print $i < $end ? " " : "\n";
    }
}

# all hours according to local time without switching for daylight saving
use constant NIGHT_START =>  0; # at night (with just basic load)
use constant NIGHT_END   =>  6;

my $sum_items = 0;
my $load_max = 0;
my $load_max_time;
my $load_sum = 0;
my $night_sum = 0;

my ($month, $day, $hour) = (1, 1, 0);
sub adjust_day_month {
    return if $hour % 24 != 0;
    if ($day == 28 && $month == 2 ||
        $day == 30 && ($month == 4 || $month == 6 ||
                       $month == 9 || $month == 11) ||
        $day == 31) {
        $day = 1;
        $month++;
    } else {
        $day++;
    }
}


################################################################################
# read load profile data

sub max_load {
    my $load = shift;
    if ($load > $load_max) {
        $load_max = $load;
        $load_max_time = time_string(shift, shift, shift, shift, shift);
    }
}


my $items_per_hour;
my @items_by_hour;
my @load_item;
my @load;
my @load_per_hour;
my @load_by_hour;
my @load_by_weekday;
my @load_by_month;
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
        my $year = $sources[0] =~ /((19|20)\d\d)/ ? $1 : "YYYY";
        shift @sources; # ignore date & time info
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
                print "Warning: load on ".date_string(0, $month, $day, $hour)
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
        $day_load += $hload;

        # adapt according to @load_dist, @load_factors, and $load_const
        # $hour_per_year == $num_hours - 1 needed for $test
        if (++$hour == 24 || $hour_per_year == $num_hours - 1) {
            my $hour_end =
                $test && $hour_per_year == $num_hours - 1 ? TEST_END % 24 : 24;
            for ($hour = 0; $hour < $hour_end; $hour++) {
                $hload = $load[$month][$day][$hour];
                if ((defined $load_const && $weekday < $load_days &&
                     ($load_from > $load_to
                      ? ($load_from <= $hour || $hour < $load_to)
                      : ($load_from <= $hour && $hour < $load_to)))
                    || $avg_hour) {
                    $items_by_hour[$month][$day][$hour] = 1;
                    $hload = $load_const if defined $load_const;
                    $load_item[$month][$day][$hour][0] = $hload;
                    max_load($hload, $year, $month, $day, $hour, 0);
                } else {
                    my $orig_hload = $hload;
                    $hload = $load[$month][$day][$hour] = $load_factors[$hour] *
                        (!defined $load_dist ? $hload :
                         $load_dist[$hour] * $day_load / $load_dist_sum);
                    my $n = $items_by_hour[$month][$day][$hour];
                    for (my $item = 0; $item < $n; $item++) {
                        my $load = $load_item[$month][$day][$hour][$item]
                            * $load_factors[$hour];
                        $load *= $hload / $orig_hload
                            if defined $load_dist && $orig_hload != 0;
                        $load_item[$month][$day][$hour][$item] = $load;
                        max_load($load,$year,$month,$day, $hour,60* $item / $n);
                    }
                }
                $load_by_hour   [$hour   ] += $hload;
                $load_by_weekday[$weekday] += $hload;
                $load_by_month  [$month  ] += $hload;
                $load_sum += $hload;
                $night_sum += $hload
                    if NIGHT_START <= $hour && $hour < NIGHT_END;
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
my $load_scale = defined $consumption && $load_sum != 0
    ? 1000 * $consumption / $load_sum : 1;
my $load_scale_never_0 = never_0($load_scale);
my $n_days = $test ? TEST_LENGTH / 24 : 365;
my $sn = $load_scale / $n_days;
for (my $hour = 0; $hour < 24; $hour++) {
    $load_per_hour[$hour] = round($load_by_hour[$hour] * $sn);
}
$load_max *= $load_scale;

my $profile_txt = $en ? "load profile file"     : "Lastprofil-Datei";
my $pv_data_txt = $en ? "PV data file"          : "PV-Daten-Datei";
my $plural_txt  = $en ? "(s)"                   : "(en)";
my $limit_txt   = $en ? "inverter input limit"  : "WR-Eingangs-Begrenzung";
my $none_txt    = $en ? "(0 = none)"            : "(0 = keine)";
my $slope_txt   = $en ? "slope"                 : "Neigungswinkel";
my $azimuth_txt = $en ? "azimuth"               : "Azimut";
my $p_txt = $en ? "load data points per hour  " : "Last-Datenpunkte pro Stunde";
my $D_txt = $en ? "rel. load distr. each hour"  : "Rel. Lastverteilung je Std.";
my $d_txt = $en ? "load distortions each hour"  : "Last-Verzerrung je Stunde";
my $l_txt = $en ? "average load each hour"      : "Mittlere Last je Stunde";
my $t_txt = $en ? "total cons. acc. to profile" : "Verbrauch gemäß Lastprofil ";
my $consumpt_txt= $en ? "consumption by household" : "Verbrauch durch Haushalt";
my $L_txt = $en ? "load portion"                : "Last-Anteil";
my $V_txt = $en ? "PV portion"                  : "PV-Anteil";
my $per3  = $en ? "per 3 hours"                 : "pro 3 Stunden";
my $per_m = $en ? "per month"                   : "pro Monat";
my $W_txt = $en ? "portion per weekday (Mo-Su)" :"Anteil pro Wochentag (Mo-So)";
my $b_txt = $en ? "basic load                 " : "Grundlast                  ";
my $M_txt = $en ? "max load"                    : "Maximallast";
my $on    = $en ? "on" : "am";
my $en1 = $en ? " "   : "";
my $en2 = $en ? "  "  : "";
my $en3 = $en ? "   " : "";
my $en4 = $en ? "    ": "";
my $de1 = $en ? ""    : " ";
my $de2 = $en ? ""    : "  ";
my $de3 = $en ? ""    : "   ";
my $s10   = "          "; 
my $s13   = "             ";
my $lhs_spaces = " " x length("$W_txt$en1");
print "$profile_txt$de1$s10 : $load_profile\n" unless $test;
if ($verbose) {
    print "$p_txt = ".sprintf("%4d", $items_per_hour)."\n";
    print "$t_txt =".kWh($load_sum)."\n";
    print "$consumpt_txt    =" .kWh($load_sum * $load_scale)."\n";
}
print "$b_txt =".W($night_sum * $load_scale / (NIGHT_END - NIGHT_START)
                   / $n_days)."\n";
print "$M_txt $en3                =".W($load_max)." $on $load_max_time\n";
print "$D_txt $en1= @load_dist[0..23]\n" if defined $load_dist;
print "$d_txt $de1 = @load_factors\n"    if defined $load_factors;
if ($verbose) {
    print "$l_txt $en1    = @load_per_hour[0..23]\n";
    print_arr_perc("$L_txt $per3 $en1  = ", \@load_by_hour,$load_sum, 0, 21, 3);
    if (!$test) {
        print_arr_perc("$L_txt $per_m $de1     = ",
                                       \@load_by_month  ,$load_sum, 1, 12, 1);
        print_arr_perc("$W_txt$en1= ", \@load_by_weekday,$load_sum, 0,  6, 1);
    }
}
print "\n";

################################################################################
# read PV production data

my @PV_gross;
my @PV_net;
my ($start_year, $years);
my $garbled_hours = 0;
sub get_power {
    my ($file, $nominal_power, $limit) = (shift, shift, shift);
    $limit *= $inverter_eff; # limit at inverter input converted to net output
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
            my $eff = percent($sys_eff_deflt);
            print ", ".($en ?
                        "contained system efficiency $eff% was overridden" :
                        "enthaltene System-Effizienz $eff% wurde übersteuert")
                if !$test && $pvsys_eff && $pvsys_eff != $pvsys_eff_deflt
                && abs($sys_eff_deflt - $pvsys_eff * $inverter_eff) * 100 >= .5;
        }
        $power_provided = 0 if m/^time,/ && !(m/^time,P,/);

        # work around CSV lines at hour 00 garbled by load&save with LibreOffice
        $garbled_hours++ if s/^\d+:10:00,0,/20000000:0000,0,/;
        next unless m/^20\d\d\d\d\d\d:\d\d\d\d,/;

        unless (defined $power_rate || $test) {
            print "\n" # close line started with print "$pv_data_txt..."
                unless $lat && $lon && $slope && $azimuth && $power_provided;
            # die "Missing latitude in $file"  unless $lat;
            # die "Missing longitude in $file" unless $lon;
            # die "Missing slope in $file"     unless $slope;
            # die "Missing azimuth in $file"   unless $azimuth;
            die "Missing PV power output data in $file" unless $power_provided;

            unless (defined $nominal_power_deflt) {
                print "\n"; # close line started with print "$pv_data_txt..."
                die "Missing nominal PV power in command line option or $file"
                    unless $nominal_power;
                print "Warning: cannot find nominal PV power in $file, ".
                    "taking value $nominal_power from command line\n";
                $nominal_power_deflt = $nominal_power;
            }
            $nominal_power = $nominal_power_deflt unless $nominal_power;

            unless ($pvsys_eff || defined $pvsys_eff_deflt) {
                print "\nWarning: cannot find system efficiency in $file and ".
                    "no -peff option given, assuming system efficiency 86%\n";
                $pvsys_eff = 0.86 / $inverter_eff_never_0;
            }
            $pvsys_eff = $pvsys_eff_deflt unless defined $pvsys_eff;
            if ($pvsys_eff < 0 || $pvsys_eff > 1) {
                print "\n"; # close line started with print "$pv_data_txt..."
                die "unreasonable PV system efficiency ".percent($pvsys_eff)
                    ."% - have -peff and -ieff been used properly?";
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
        $start_year = $1 if (!$start_year && m/^(\d\d\d\d)/);
        if ($tmy && !$test) {
            # typical metereological year
            my $selected_month = 0;
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
            next unless $selected_month;
        }
        # matching hour 01 rather than 00 due to potentially garbled CSV lines:
        $current_years++ if m/^20\d\d0101:01/;
        $months++ if m/^20....01:01/;

        die "Missing power data in $file line $_"
            unless m/^(\d\d\d\d)(\d\d)(\d\d):(\d\d)(\d\d),\s?([\d\.]+)/;
        my $hour_offset = $test ? 0 : TimeZone;
        my ($year, $month, $day, $hour, $minute, $net_power) =
            ($tmy ? 0 : $1-$start_year, $2, $3, ($4+$hour_offset) % 24, $5, $6);
        # for simplicity, attributing hours wrapped via time zone to same day

        $net_power *= $power_rate;
        my $gross_power = $net_power *
            (defined $sys_eff_deflt ? 1 / $sys_eff_deflt : $gross_rate);
        $PV_gross[$year][$month][$day][$hour] += $gross_power;
        $net_power = $gross_power * $net_rate;
        $net_power = $limit if $limit != 0 && $net_power > $limit;
        $PV_net[$year][$month][$day][$hour] += $net_power;

        $hours++;
        last if $test && $hours == TEST_END;
    }
    close $IN unless $test;
    print "\n" unless $test; # close line started with print "$pv_data_txt..."

    check_consistency($years, $current_years, "years", $file) if $years;
    $years = $current_years;
    die "number of years detected is $years in $file"
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
        print "$slope_txt, $azimuth_txt$en3$en3$en2      = $slope, $azimuth\n";
    }
    return $nominal_power;
}

my $total_limit = 0;
for (my $i = 0; $i <= $#PV_files; $i++) {
    $total_limit += $PV_limit[$i];
    $PV_nomin[$i] = get_power($PV_files[$i], $PV_nomin[$i], $PV_limit[$i]);
}
my $PV_nomin = join("+", @PV_nomin);
my $PV_limit = join("+", @PV_limit);

################################################################################
# PV usage simulation

my @PV_per_hour;
my @PV_by_hour;
my @PV_by_month;
my $PV_gross_max = 0;
my $PV_gross_max_time;
my $PV_gross_sum = 0;

my @PV_net_loss;
my $PV_net_losses = 0;
my $PV_net_loss_hours = 0;
my $PV_net_max = 0;
my $PV_net_max_time;
my $PV_net_sum = 0;

my @PV_use_loss_by_item;
my @PV_use_loss;
my $PV_use_loss_sum = 0;
my $PV_use_loss_hours = 0;
my @PV_used_by_item;
my @PV_used;
my $PV_used_sum = 0;
my $PV_used_via_storage = 0;

my @grid_feed_by_item if $max;
my @grid_feed_per_hour;
my @grid_feed;
my $grid_feed_sum = 0;

my $charge_max = $capacity *      $soc  if defined $capacity;
my $charge_min = $capacity * (1 - $dod) if defined $capacity;
my $charge            if defined $capacity; # charge state of the battery
my @charge_by_item    if defined $capacity && $max;
my @charge_per_hour   if defined $capacity;
my @charge            if defined $capacity;
my $charge_sum    = 0 if defined $capacity;
my @dischg_by_item    if defined $capacity && $max;
my @dischg_per_hour   if defined $capacity;
my @dischg            if defined $capacity;
my $dischg_sum = 0    if defined $capacity;
my $charging_loss = 0 if defined $capacity;
my $spill_loss    = 0 if defined $capacity;
my $AC_coupling_losses = 0 if defined $capacity;

sub simulate()
{
    my $year = 0;
    ($month, $day, $hour) = (1, 1, 0);
    my $minute = 0; # currently fixed

    # factor out $load_scale for optimizing the inner loop
    if ($capacity) {
        $capacity /= $load_scale_never_0;
        $charge_max /= $load_scale_never_0;
        $charge_min /= $load_scale_never_0;
        $charge = $charge_min;
        # convert from C scale:
        $max_charge *= $capacity / $charge_eff_never_0 if $capacity;
        $max_dischg *= $capacity / $storage_eff_never_0 if $capacity;
    }
    $bypass /= $load_scale_never_0 if defined $bypass;
    $max_feed /= ($load_scale_never_0 * $storage_eff_never_0)
        if defined $max_feed;

    my $end_year = $years;
    if (defined $sel_year && $sel_year ne "*") {
        $year = $sel_year - $start_year;
        die "year given with -only option must be in range $start_year..".
            ($start_year + $years - 1) if $year < 0 || $year >= $years;
        ($years, $end_year) = (1, $year + 1);
    }
    while ($year < $end_year) {
        my $year_ = $tmy ? "TMY" : $start_year + $year;
        $PV_net_loss[$month][$day][$hour] = 0;
        $PV_use_loss[$month][$day][$hour] = 0;
        if ((defined $sel_month) && $sel_month ne "*" && $month != $sel_month ||
            (defined $sel_day  ) && $sel_day   ne "*" && $day   != $sel_day ||
            (defined $sel_hour ) && $sel_hour  ne "*" && $hour  != $sel_hour) {
            $PV_gross[$year][$month][$day][$hour] = 0;
            $PV_net  [$year][$month][$day][$hour] = 0;
            goto NEXT;
        }

        my $gross_power = $PV_gross[$year][$month][$day][$hour];
        my $pvnet_power = $PV_net  [$year][$month][$day][$hour];
        if (!defined $gross_power) {
            if ($hour == 0 && $garbled_hours != 0) { # likely just garbled hour
                $gross_power = 0;
                $pvnet_power = 0;
            } else {
                $en = 1;
                die "No power data for ".
                    time_string($year_, $month, $day, $hour, $minute);
            }
        }
        $PV_by_hour [$hour ] += $gross_power;
        $PV_by_month[$month] += $gross_power;
        if ($gross_power > $PV_gross_max) {
            $PV_gross_max = $gross_power;
            $PV_gross_max_time = time_string($year_,$month,$day,$hour,$minute);
        }
        $PV_gross_sum += $gross_power;

        my $PV_loss = 0;
        if ($curb && $pvnet_power > $curb) { # TODO adapt to storage use
            $PV_loss = $pvnet_power - $curb;
            $PV_net_loss[$month][$day][$hour] += $PV_loss;
            $PV_net_losses += $PV_loss;
            $PV_net_loss_hours++;
            $PV_net[$year][$month][$day][$hour] = $pvnet_power = $curb;
            # print "$year-".time_string($year_, $month, $day, $hour, $minute).
            #"\tPV=".round($pvnet_power)."\tcurb=".round($curb).
            #"\tloss=".round($PV_net_losses)."\t$_\n";
        }
        if ($pvnet_power > $PV_net_max) {
            $PV_net_max = $pvnet_power;
            $PV_net_max_time = time_string($year_,$month,$day,$hour,$minute);
        }
        $PV_net_sum += $pvnet_power;

        ##### at this point, $pvnet_power is the main input for simulation ####

        # factor out $load_scale for optimizing the inner loop
        $pvnet_power /= $load_scale_never_0;
        $PV_loss /= $load_scale_never_0;
        # my $needed = 0;
        my $usages = 0;
        my $hgrid_feed = 0;
        my ($hcharge_delta, $hdischg_delta) = (0, 0) if $capacity;
        my $curb_losses_this_hour = 0;
        my $AC_coupling_losses_this_hour = 0; # TODO calc AC losses on charge side
        my $items = $items_by_hour[$month][$day][$hour];

        # factor out $items for optimizing the inner loop
        if (defined $capacity && $items != 1) {
            $capacity *= $items;
            $charge_max *= $items;
            $charge_min *= $items;
            $charge *= $items;
            # $charging_loss *= $items;
            $spill_loss *= $items;
            $PV_used_via_storage *= $items;
        }

        my $test_started = $test && ($day - 1) * 24 + $hour >= TEST_START;
        # my $feed_sum = 0 if defined $max_feed;
        for (my $item = 0; $item < $items; $item++) {
            my $loss = 0;
            my $power_needed = $load_item[$month][$day][$hour][$item];
            die "Internal error: load_item[$month][$day][$hour][$item] ".
                "is undefined" unless defined $power_needed;
            printf("%02d".minute_string($item, $items)." load=%4d PV net=%4d ",
                   $hour, $power_needed, $pvnet_power) if $test_started;
            # $needed += $power_needed;
            # load will be reduced by constant $bypass or $bypass_spill

            # $grid_feed locally accumulates feed to grid
            my $grid_feed_in = 0;

            # $pv_used locally accumulates PV own consumption
            # feed by constant bypass or just as much as used (optimal charge)
            my $pv_used = defined $bypass ? $bypass * $inverter_eff
                : $power_needed; # preliminary
            my $excess_power = $pvnet_power - $pv_used;
            if ($excess_power < 0) {
                $pv_used = $pvnet_power;
                # == min($pvnet_power,
                #        $bypass ? $bypass * $inverter_eff : $power_needed)
            }
            if (defined $bypass) {
                my $unused_bypass = $pv_used - $power_needed;
                if ($unused_bypass > 0) {
                    $pv_used -= $unused_bypass;
                    $grid_feed_in += $unused_bypass;
                }
                printf("bypass feed=%4d,used=%4d ",
                       max($unused_bypass, 0), $pv_used) if $test_started;
            }
            $power_needed -= $pv_used;

            if (defined $capacity) { # storage present
                my $capacity_to_fill = $charge_max - $charge;
                $capacity_to_fill = 0 if $capacity_to_fill < 0;
                my ($charge_input, $charge_delta) = (0, 0);
                if ($excess_power > 0) {
                    # when charging is DC-coupled, no loss through inverter
                    $excess_power /= $inverter_eff_never_0 unless $AC_coupled;
                    # $excess_power is the power available for charging
                    my $need_for_fill = $capacity_to_fill / $charge_eff_never_0;
                    $need_for_fill = $max_charge
                        if $need_for_fill > $max_charge;
                    # optimal charge: exactly as much as unused and fits in
                    $charge_input = $excess_power;
                        # will become min($excess_power, $need_for_fill);
                    my $surplus = $excess_power - $need_for_fill;
                    printf("[excess=%4d,surplus=%4d] ", $excess_power,
                           max($surplus, 0) + .5) if $test_started;
                    if ($surplus > 0) {
                        $charge_input = $need_for_fill; # TODO check AC-coupled
                        my $surplus_net = $surplus;
                        $surplus_net *= $inverter_eff unless $AC_coupled;
                        if (!defined $bypass) {
                            $grid_feed_in += $surplus_net; # on optimal charge
                            printf("(grid feed=%4d) ", $surplus_net)
                                if $test_started;
                        } elsif ($bypass_spill) { # implied by $AC_coupled
                            my $remaining_surplus = $surplus_net - $power_needed;
                            my $used_surplus = $power_needed;
                            if ($remaining_surplus < 0) {
                                $used_surplus = $surplus_net;
                                $remaining_surplus = 0;
                            }
                            $pv_used += $used_surplus;
                            $power_needed -= $used_surplus;
                            $grid_feed_in += $remaining_surplus;
                            printf("spill feed=%4d,used=%4d ",
                                   max($remaining_surplus, 0),
                                   $used_surplus + .5) if $test_started;
                        } else { # defined $bypass && !$bypass_spill
                            $spill_loss += $surplus_net;
                            printf("spill loss=%4d", $surplus_net + .5)
                                if $test_started;
                        }
                    } elsif ($test_started) {
                        printf("                          "); # no surplus
                    }

                    # The following adds reduced charging due to curb
                    # to the usage losses, which is not entirely correct
                    # in case of constant feed (non-optimal discharge).
                    $loss += min($PV_loss, $capacity_to_fill
                                 * $storage_eff * $inverter_eff)
                        if $AC_coupled && $PV_loss != 0;

                    $charge_delta = $charge_input * $charge_eff;
                    $charge += $charge_delta;
                    # $charging_loss += $charge_input - $charge_delta;
                } elsif ($test_started) {
                    printf("                           "); # no $excess_power
                    printf("                          "); # no surplus
                }
                my $print_charge = $test_started &&
                    ($excess_power > 0 || $charge > $charge_min);
                printf("chrg loss=%4d dischrg needed=%4d [charge %4d + %4d ",
                       ($charge_input - $charge_delta) *
                       ($AC_coupled ? 1 : $inverter_eff) + .5,$power_needed +.5,
                       $charge - $charge_delta, $charge_delta) if $print_charge;

                my $dischg_delta = 0;
                my $AC_loss = 0;
                if ($charge > $charge_min) { # storage not empty
                    # optimal discharge: exactly as much as currently needed
                    # $discharge = min($power_needed, $charge)
                    my $discharge = $power_needed /
                        ($storage_eff_never_0 * $inverter_eff_never_0);
                    if (defined $max_feed) {
                        if ($const_feed) {
                            $discharge = 0;
                            $discharge = $max_feed
                                if $feed_from > $feed_to
                                   ? ($feed_from <= $hour || $hour < $feed_to)
                                   : ($feed_from <= $hour && $hour < $feed_to);
                        } else {
                            $discharge = $max_feed # optimal but limited feed
                                if $discharge > $max_feed;
                        }
                    }
                    $discharge = $max_dischg if $discharge > $max_dischg;
                    $discharge = $charge - $charge_min
                        if $discharge > $charge - $charge_min;
                    printf("- lost=%4d - %4d ", $discharge * (1 - $storage_eff)
                           + .5, $discharge*$storage_eff + .5) if $test_started;
                    if ($discharge != 0) {
                        $charge -= $discharge; # includes storage loss
                        $discharge *= $storage_eff;
                        $dischg_delta = $discharge; # after storage loss
                        my $discharge_net = $discharge * $inverter_eff;
                        if ($AC_coupled) {
                            $AC_loss = $discharge - $discharge_net;
                            $AC_coupling_losses_this_hour += $AC_loss;
                        }
                        if (defined $max_feed && $const_feed) {
                            # $feed_sum += $discharge;
                            my $dis_feed_in = $discharge_net - $power_needed;
                            if ($dis_feed_in > 0) {
                                $grid_feed_in += $dis_feed_in;
                                $discharge_net -= $dis_feed_in;
                            }
                        }
                        $pv_used += $discharge_net;
                        $PV_used_via_storage += $discharge_net;
                    }
                } else {
                    print "                   " if $test_started;
                }
                printf("= %4d] ", $charge) if $print_charge;
                printf("AC coupling loss=%4d ", $AC_loss + .5)
                    if $print_charge && $AC_coupled;
                $hcharge_delta += $charge_delta;
                $hdischg_delta += $dischg_delta;
                if ($max) {
                    $charge_by_item[$month][$day][$hour][$item]+= $charge_delta;
                    $dischg_by_item[$month][$day][$hour][$item]+= $dischg_delta;
                }
                printf(" " x 85) if $test_started && !$print_charge;
            } else {
                $grid_feed_in = $excess_power if $excess_power > 0;
            }
            printf("used=%4d\n", $pv_used + .5) if $test_started;
            $usages += $pv_used;
            $hgrid_feed += $grid_feed_in;
            if ($max) {
                $PV_used_by_item[$month][$day][$hour][$item] += $pv_used;
                $grid_feed_by_item[$month][$day][$hour][$item]+= $grid_feed_in;
            }

            if ($PV_loss != 0 && $power_needed > 0) {
                $loss += min($PV_loss, $power_needed);
            }
            $PV_use_loss_by_item[$month][$day][$hour][$item] += $loss if $max;
            if ($loss != 0) {
                $curb_losses_this_hour += $loss;
                $PV_use_loss_hours++; # will be normalized by $items
            }
        }
        # $spill_loss += ($pvnet_power - $feed_sum / $items) * $load_scale
        #     if defined $bypass;

        $curb_losses_this_hour *= $load_scale / $items;
        $PV_use_loss[$month][$day][$hour] += $curb_losses_this_hour;
        $PV_use_loss_sum += $curb_losses_this_hour;
        $AC_coupling_losses_this_hour *= $load_scale / $items;
        $AC_coupling_losses += $AC_coupling_losses_this_hour;
        # $sum_needed += $needed * $load_scale / $items; # per hour
        # print "$year-".time_string($year_, $month, $day, $hour, $minute).
        # "\tPV=".round($pvnet_power)."\tPN=".round($needed)."\tPU=".round($usages).
        # "\t$_\n" if $pvnet_power != 0 && m/^20160214:1010/; # m/^20....02:12/;
        $usages *= $load_scale / $items;
        $PV_used[$month][$day][$hour] += $usages;
        $PV_used_sum += $usages;
        $hgrid_feed /= $items;
        $grid_feed_sum += $hgrid_feed;
        $grid_feed_per_hour     [$hour] += $hgrid_feed;
        $grid_feed[$month][$day][$hour] += $hgrid_feed * $load_scale;

        if (defined $capacity) {
            if ($items != 1) {
                # revert factoring out $items for optimizing the inner loop
                $capacity /= $items;
                $charge_max /= $items;
                $charge_min /= $items;
                $charge /= $items;
                # $charging_loss /= $items;
                $spill_loss /= $items;
                $PV_used_via_storage /= $items;
                $hcharge_delta /= $items;
                $hdischg_delta /= $items;
            }
            $charge_per_hour     [$hour] += $hcharge_delta;
            $dischg_per_hour     [$hour] += $hdischg_delta;
            $charge[$month][$day][$hour] += $hcharge_delta;
            $dischg[$month][$day][$hour] += $hdischg_delta;
            $charge_sum += $hcharge_delta;
            $dischg_sum += $hdischg_delta;
        }

      NEXT:
        $hour = 0 if ++$hour == 24;
        adjust_day_month();
        ($year, $month, $day) = ($year + 1, 1, 1) if $month > 12;
        last if $test && ($day - 1) * 24 + $hour == TEST_END;
    }
    print "\n" if $test;

    $load_sum *= $load_scale;
    $PV_gross_sum /= $years;
    $PV_net_losses /= $years;
    $PV_net_loss_hours /= $years;
    $PV_net_sum /= $years;
    $PV_used_sum /= $years;
    $PV_use_loss_sum /= $years;
    $PV_use_loss_hours /= ($sum_items / YearHours * $years);
    # die "Inconsistent load calculation: sum = $sum vs. needed = $sum_needed"
    #     if round($sum) != round($sum_needed);
    $grid_feed_sum *= $load_scale / $years;

    if (defined $capacity) {
        $charge -= $charge_min;
        $max_charge *= $charge_eff / $capacity if $capacity;
        $max_dischg *= $storage_eff / $capacity if $capacity;
        $capacity *= $load_scale_never_0;
        $charge_max *= $load_scale_never_0;
        $charge_min *= $load_scale_never_0;
        $bypass *= $load_scale_never_0 if defined $bypass;
        $max_feed *= $load_scale_never_0 * $storage_eff_never_0
            if defined $max_feed;
        $charge     *= $load_scale;
        $charge_sum *= $load_scale / $years;
        $dischg_sum *= $load_scale / $years;
        # $charging_loss *= $load_scale / $years;
        $spill_loss *= $load_scale / $years;
        $PV_used_via_storage *= $load_scale / $years;
    }
}

simulate();

die "PV power data all zero" unless $PV_gross_max_time;
my $sny = $sn / $years;
for (my $hour = 0; $hour < 24; $hour++) {
    $PV_per_hour[$hour] = round($PV_by_hour[$hour] / $n_days / $years);
    $grid_feed_per_hour[$hour] = round($grid_feed_per_hour[$hour] * $sny);
    if (defined $capacity) {
        $charge_per_hour[$hour] = round($charge_per_hour[$hour] * $sny);
        $dischg_per_hour[$hour] = round($dischg_per_hour[$hour] * $sny);
    }
}

################################################################################
# statistics output

my $nominal_txt      = $en ? "nominal PV power"     : "PV-Nominalleistung";
my $only             = $en ? "only"                 : "nur";
my $during           = $en ? "during"               : "während";
my $gross_max_txt    = $en ? "max gross PV power"   : "Max. PV-Bruttoleistung";
my $net_max_txt      = $en ? "max net PV power"     : "Max. PV-Nettoleistung";
my $curb_txt    = $en ? "inverter ouput power curb" : "WR-Ausgangs-Drosselung";
my $pvsys_eff_txt    = $en ? "PV system efficiency" : "PV-System-Wirkungsgrad";
my $own_txt          = $en ? "PV own use"           : "PV-Eigenverbrauch";
my $own_ratio_txt    = $own_txt . ($en ? " ratio"  : "santeil");
my $own_storage_txt  = $en ?"PV own use via storage":"PV-Nutzung über Speicher";
my $load_cover_txt   = $en ? "use coverage ratio"   : "Eigendeckungsanteil";
my $PV_gross_txt     = $en ? "PV gross yield"       : "PV-Bruttoertrag";
my $PV_net_txt       = $en ? "PV net yield"         : "PV-Nettoertrag";
my $PV_loss_txt      = $en ? "PV yield net loss"    : "PV-Netto-Ertragsverlust";
my $load_const_txt   = $en ? "constant load"        : "Konstante Last";
my $load_during_txt  =($load_days == 7 ? "" :
                       ($en ? "on the first $load_days days a week"
                            : "an den ersten $load_days Tagen der Woche")).
                      (($load_from == 0 && $load_to == 24) ? "" :
                       ($en ? " from $load_from to $load_to h"
                           : " von $load_from bis $load_to Uhr"))
                      if defined $load_const;
my $usage_loss_txt   = $en ? "$own_txt net loss"    :  $own_txt."sverlust";
my $net_de           = $en ? ""                     : " netto";
my $by_curb   = $en ? "by inverter output curb" : "durch WR-Ausgangsdrosselung";
my $with             = $en ? "with"                 : "mit";
my $without          = $en ? "without"              : "ohne";
my $grid_feed_txt    = $en ? "grid feed-in"         : "Netzeinspeisung";
my $charge_txt       = $en ? "charge (before storage losses)"
                           : "Ladung (vor Speicherverlusten)";
my $dischg_txt       = $en ? "discharge"            : "Entladung";
my $after            = $en ? "after"                : "nach";
my $sum_txt          = $en ? "sums"                 : "Summen";
my $each             = $en ? "each"                 : "alle";
my $yearly_txt       = $en ? "over the year"        : "übers Jahr";
my $on_average_txt   = $en ? "on average over $years years"
                           : "im Durchschnitt über $years Jahre";
my $capacity_txt     = $en ? "storage capacity"     : "Speicherkapazität";
my $soc_txt          = $en ? "SoC"                  : "Ladehöhe";
my $dod_txt          = $en ? "DoD"                  : "Entladetiefe";
my $coupled_txt = ($AC_coupled ? "AC" : "DC").($en ? " coupled": "-gekoppelt");
my $optimal_charge   = $en ? "optimal charging strategy (power not consumed)"
                           :"Optimale Ladestrategie (nicht gebrauchte Energie)";
my $bypass_txt       = $en ? "storage bypass"       : "Speicher-Umgehung";
my $spill_txt        = !$bypass_spill ? "" :
                      ($en ?  " and for surplus"    : " und für Überschuss");
my $max_charge_txt   = $en ? "max charge rate"      : "max. Laderate";
my $optimal_discharge= $en ? "optimal discharging strategy (as much as needed)"
                           :"Optimale Entladestrategie (so viel wie gebraucht)";
my $feed_txt        = ($const_feed
                    ? ($en ? "constant "            : "Konstant")
                    : ($en ?  "maximal "            : "Maximal"))
                     .($en ? "feed-in"              : "einspeisung");
my $feed_during_txt  =!$const_feed || ($feed_from == 0 && $feed_to == 24) ? "" :
                       $en ? " from $feed_from to $feed_to h"
                           : " von $feed_from bis $feed_to Uhr";
my $max_dischg_txt   = $en ? "max discharge rate"   : "max. Entladerate";
my $ceff_txt         = $en ? "charging efficiency"  : "Lade-Wirkungsgrad";
my $seff_txt         = $en ? "storage efficiency"   : "Speicher-Wirkungsgrad";
my $ieff_txt         = $en ?"inverter efficiency":"Wechselrichter-Wirkungsgrad";
my $stored_txt       = $en ? "buffered energy"      : "Zwischenspeicherung";
my $spill_loss_txt   = $en ? "loss by spill"        : "Verlust durch Überlauf";
my $AC_coupl_loss_txt= $en ? "loss by AC coupling" :"Verlust durch AC-Kopplung";
my $charging_loss_txt= $en ? "charging loss"        : "Ladeverlust";
my $storage_loss_txt = $en ? "storage loss"         : "Speicherverlust";
my $cycles_txt       = $en ? "full cycles per year" : "Vollzyklen pro Jahr";
my $of_eff_cap_txt   = $en ? "of effective capacity":"der effektiven Kapazität";
my $c_txt = $en ? "average charge/day each hour": "Mittlere Ladung/Tag je Std";
my $C_txt = $en ? "avg discharge/day each hour" : "Mittl. Endladung/Tag je Std";
my $P_txt = $en ? "average PV power each hour"  : "Mittlere PV-Leistung je Std";
my $F_txt = $en ? "average grid feed each hour" : "Mittlere Einspeisung je Std";
my $dischg_after_txt = "$dischg_txt $after $storage_loss_txt";

my $only_during = defined $date ? " $only $during $date" : "";
my $own_ratio = round_percent($PV_net_sum ? $PV_used_sum/$PV_net_sum:0);
my $load_coverage = round_percent($load_sum ? $PV_used_sum / $load_sum : 0);
my $storage_loss;
my $cycles = 0;
if (defined $capacity) { # also for future loss when discharging the rest:
    $charging_loss = ($charge_eff ? $charge_sum * (1 / $charge_eff - 1) : 0);
    $charging_loss *= $inverter_eff unless $AC_coupled;
    $storage_loss = $charge_sum * (1 - $storage_eff);
    $storage_loss *= $inverter_eff unless $AC_coupled;
    $cycles = round($charge_sum / ($charge_max - $charge_min)) if $capacity !=0;
    # future losses:
    $charge *= $storage_eff;
    $AC_coupling_losses += $charge * (1 - $inverter_eff) if $AC_coupled;
    $charge *= $inverter_eff;
}

$load_const = round($load_const * $load_scale) if defined $load_const;

sub save_statistics {
    my $file = shift;
    return unless $file;
    my $type_txt= shift;
    my $max     = shift;
    my $hourly  = shift;
    my $daily   = shift;
    my $weekly  = shift;
    my $monthly = shift;
    my $season  = shift;
    open(my $OU, '>',$file) or die "Could not open statistics file $file: $!\n";

    my $nominal_sum = $#PV_nomin == 0 ? $PV_nomin[0] : "$PV_nomin";
    my $limits_sum  = $#PV_limit == 0 ? $PV_limit[0] : "$PV_limit";
    print $OU "$consumpt_txt in kWh,";
    print $OU "$load_const_txt in W $load_during_txt," if defined $load_const;
    print $OU "$profile_txt,$M_txt in W $on $load_max_time,";
    print $OU "$pv_data_txt$plural_txt$only_during\n";
    print $OU "".round_1000($load_sum).",";
    print $OU "$load_const," if defined $load_const;
    print $OU "$load_profile,".int($load_max).",".join(",", @PV_files)."\n";

    $pvsys_eff = round_percent($pvsys_eff);
    print $OU "$nominal_txt in Wp,$limit_txt in W $none_txt,"
        ."$gross_max_txt in W $on $PV_gross_max_time,"
        ."$net_max_txt in W $on $PV_net_max_time,"
        ."$curb_txt in W,$pvsys_eff_txt,$ieff_txt,"
        ."$own_ratio_txt,$load_cover_txt,";
    print $OU ",$D_txt:,".join(",", @load_dist) if defined $load_dist;
    print $OU "\n$nominal_sum,$limits_sum,".round($PV_gross_max).","
        .round($PV_net_max).",".($curb ? $curb : $en ? "none" : "keine").
        ",$pvsys_eff,$inverter_eff,$own_ratio,$load_coverage,";
    print $OU ",$d_txt:,".join(",", @load_factors) if defined $load_factors;
    print $OU "\n";

    if (defined $capacity) {
        print $OU "$capacity_txt in Wh,"
            .(defined $bypass ? "$bypass_txt in W$spill_txt" : $optimal_charge)
            .",$max_charge_txt in C,"
            .(defined $max_feed ? "$feed_txt in W$feed_during_txt"
                                : $optimal_discharge).",$max_dischg_txt in C,"
            ."$ceff_txt,$seff_txt,$soc_txt,$dod_txt\n";
        print $OU "$capacity,"
            .(defined $bypass ? $bypass : "").",$max_charge,"
            .(defined $max_feed ? $max_feed : "").",$max_dischg,"
            ."$charge_eff,$storage_eff,"
            .(percent($soc) / 100).",".(percent($dod) / 100)."\n";
        print $OU "".($AC_coupled ? "$AC_coupl_loss_txt in kWh": $coupled_txt)
            .",$spill_loss_txt in kWh,"
            ."$charging_loss_txt in kWh,$storage_loss_txt in kWh,"
            ."$own_storage_txt in kWh,$stored_txt in kWh,"
            ."$cycles_txt $of_eff_cap_txt\n";
        print $OU "".($AC_coupled ? round_1000($AC_coupling_losses) : "").","
            .round_1000($spill_loss).","
            .round_1000($charging_loss).",".round_1000($storage_loss).","
            .round_1000($PV_used_via_storage).",".round_1000($charge_sum)
            .",$cycles\n";
    }

    print $OU "\n";
    print $OU "$l_txt in W:," .join(",", @load_per_hour)."\n";
    print $OU "$P_txt in W:," .join(",", @PV_per_hour)."\n";
    print $OU "$F_txt in Wh:,".join(",", @grid_feed_per_hour)."\n";
    if (defined $capacity) {
        print $OU "$c_txt in Wh:,".join(",", @charge_per_hour)."\n";
        print $OU "$C_txt in Wh:,".join(",", @dischg_per_hour)."\n";
    }
    print $OU "\n";

    my $sum_avg = $sum_txt . ($years > 1 ? " $on_average_txt" : "");
    print $OU "$yearly_txt,$PV_gross_txt,"
        .($curb ? "$PV_loss_txt $by_curb," : "")."$PV_net_txt,"
        .($curb ? "$usage_loss_txt$net_de $by_curb,": "")
        ."$own_txt,$grid_feed_txt,$consumpt_txt"
        .(defined $capacity ? ",$charge_txt,$dischg_after_txt" : "")
        ."\n";
    print $OU "$sum_avg,".
        round_1000($PV_gross_sum).",".
        ($curb ? round_1000($PV_net_losses  )."," : "").
        round_1000($PV_net_sum).",".
        ($curb ? round_1000($PV_use_loss_sum)."," : "").
        round_1000($PV_used_sum).",".
        round_1000($grid_feed_sum ).",".round_1000($load_sum).",".
        (defined $capacity ?
         round_1000(   $charge_sum).",".round_1000($dischg_sum)."," : "").
        "$each in kWh\n";

    my $i = 14 + (defined $capacity ? 6 : 0);
    my $j = $i - 1 + ($max ? $sum_items : $hourly ? YearHours
                      : $daily ? 365 : $weekly ? 52 : $monthly ? 12 : 4);
    my ($I, $J) = $curb ? ("I", "J") : ("G", "H");
    print $OU "$type_txt,$PV_gross_txt,"
        ."$PV_net_txt".($curb ? " $without $curb_txt" : "").","
        .($curb ?       "$PV_net_txt $with $curb_txt,": "")
        .($curb ? "$own_txt $without $curb_txt," : "")
        ."$own_txt".($curb ? " $with $curb_txt"  : "").","
        ."$grid_feed_txt,$consumpt_txt"
        .(defined $capacity ? ",$charge_txt,$dischg_after_txt" : "")."\n";
    print $OU "$sum_avg,".SUM("B", $i, $j).",".SUM("C", $i, $j)
        .",".SUM("D", $i, $j).",".SUM("E", $i, $j).",".SUM("F", $i, $j)
        .($curb ? ",".SUM("G", $i, $j).",".SUM("H", $i, $j) : "")
        .(defined $capacity ? ",".SUM($I, $i, $j).",".SUM($J, $i, $j) : "")
        .",$each in ".($max ? "m" : "")."Wh\n";
    (my $week, my $days, $hour) = (1, 0, 0);
    ($month, $day) = $season && !$test ? (2, 5) : (1, 1);
    my ($gross, $PV_loss, $net, $hload, $loss, $used, $feed) = (0,0,0,0,0,0,0);
    my ($chg, $dis) = (0, 0) if defined $capacity;
    while ($days < 365) {
        my $tim;
        if ($weekly) {
            $tim = $week;
        } elsif ($season) {
            $tim = $season== 1 ? ($en ? "spring" : "Frühjahr")
                : $season == 2 ? ($en ? "summer" : "Sommer")
                : $season == 3 ? ($en ? "autumn" : "Herbst")
                :                ($en ? "winter" : "Winter");
        } else {
            $tim = sprintf("%02d", $month);
            $tim = $tim."-".sprintf("%02d", $day ) if $daily || $hourly || $max;
            $tim = $tim." ".sprintf("%02d", $hour) if $hourly || $max;
        }
        my ($gross_across, $net_across) = (0, 0);
        for (my $year = 0; $year < $years; $year++) {
            $year = $sel_year - $start_year if defined $sel_year && $sel_year ne "*";
            my $PV_gross = $PV_gross[$year][$month][$day][$hour];
            die "Internal error: PV_gross[$year][$month][$day] is undefined"
                . " on day $days" unless defined $PV_gross;
            $gross_across += $PV_gross;
            $net_across += $PV_net[$year][$month][$day][$hour];
        }
        my $items = $items_by_hour[$month][$day][$hour];
        my $f     = ($max ? 1000 : 1) / $years;
        my $fact  = $f / ($max ? $items : 1);
        $gross   += round($fact * $gross_across);
        $PV_loss += round($fact * $PV_net_loss[$month][$day][$hour]);
        $net     += round($fact * $net_across);
        if (!$max) {
            $hload   += round($load       [$month][$day][$hour] * $f *
                              $load_scale * $years);
            $loss    += round($PV_use_loss[$month][$day][$hour] * $f);
            $used    += round($PV_used    [$month][$day][$hour] * $f);
            $feed    += round($grid_feed  [$month][$day][$hour] * $f);
            if (defined $capacity) {
                $chg += round($charge[$month][$day][$hour] * $f * $load_scale);
                $dis += round($dischg[$month][$day][$hour] * $f * $load_scale);
            }
        }
        my ($m, $d, $h) = ($month, $day, $hour);
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
            for (my $i = 0; $i < ($max ? $items : 1); $i++) {
                my $minute = $hourly ? ":00" : "";
                if ($max) {
                    $minute = minute_string($i, $items);
                    my $s = $fact * $load_scale;
                    $hload = round(  $years * $load_item[$m][$d][$h][$i] * $s);
                    if ($net_across != 0) {
                        $loss=round($PV_use_loss_by_item[$m][$d][$h][$i] * $s);
                        $used=round($PV_used_by_item    [$m][$d][$h][$i] * $s);
                        $feed=round(  $grid_feed_by_item[$m][$d][$h][$i] * $s);
                    } else {
                        ($loss, $used, $feed) = (0, 0, 0);
                    }
                    if (defined $capacity) {
                        $chg = round($charge_by_item[$m][$d][$h][$i] * $s);
                        $dis = round($dischg_by_item[$m][$d][$h][$i] * $s);
                    }
                }
                print $OU "$tim$minute, $gross,".($net + $PV_loss).", "
                    .($curb ?             "$net," : "")
                    .($curb ? ($used + $loss)."," : "")
                    ."$used,$feed,$hload"
                    .(defined $capacity ? ",$chg,$dis" : "")."\n";
            }
            ($gross, $PV_loss, $net, $hload, $loss, $used, $feed) =
                (0, 0, 0, 0, 0, 0, 0);
            ($chg, $dis) = (0, 0) if defined $capacity;
        }
        ($month, $day) = (1, 1) if $month > 12;
        last if $test && $days * 24 + $hour == TEST_END;
    }
    close $OU;
}

my $max_txt = $items_per_hour == 60
             ? ($en ? "minute" : "Minute")
             : ($en ? "point in time" : "Zeitpunkt");
my $hour_txt  = $en ? "hour"   : "Stunde";
my $date_txt  = $en ? "date"   : "Datum";
my $week_txt  = $en ? "week"   : "Woche";
my $month_txt = $en ? "month"  : "Monat";
my $season_txt= $en ? "season" : "Saison";
save_statistics($max    , $max_txt  , 1, 0, 0, 0, 0, 0);
save_statistics($hourly , $hour_txt , 0, 1, 0, 0, 0, 0);
save_statistics($daily  , $date_txt , 0, 0, 1, 0, 0, 0);
save_statistics($weekly , $week_txt , 0, 0, 0, 1, 0, 0);
save_statistics($monthly, $month_txt, 0, 0, 0, 0, 1, 0);
save_statistics($seasonly,$season_txt,0, 0, 0, 0, 0, 1);

my $at             = $en ? "with"                 : "bei";
my $and            = $en ? "and"                  : "und";
my $due_to         = $en ? "due to"               : "durch";
my $by_curb_at     = $en ? "$by_curb at"          : "$by_curb auf";
my $yield          = $en ? "yield portion"        : "Ertragsanteil";
my $of_yield       = $en ? "of net yield"   : "des Nettoertrags (Nutzungsgrad)";
my $of_consumption = $en ? "of consumption" : "des Verbrauchs (Autarkiegrad)";
# PV-Abregelungsverlust"
my $lat_txt        = $en ? "latitude"             : "Breitengrad";
my $lon_txt        = $en ? "longitude"            : "Längengrad";
if (defined $lat && defined $lon && !$test) {
    print "$lat_txt, $lon_txt $en4    = $lat, $lon\n";
    print "\n";
}
my $nominal_sum = $#PV_nomin == 0 ? "" : " = $PV_nomin Wp";
my $limits_sum = $total_limit == 0 ? "" :
    ", $limit_txt: ".$PV_limit." W".($#PV_limit == 0 ? "" : " $none_txt");
print "$nominal_txt $en2         =" .W($nominal_power_sum)."p$nominal_sum".
    "$only_during$limits_sum\n";
print "$gross_max_txt $en4     =".W($PV_gross_max)." $on $PV_gross_max_time\n";
print "$PV_gross_txt $en1            =".kWh($PV_gross_sum).
    ", $pvsys_eff_txt ".percent($pvsys_eff)."%\n";
if ($verbose) {
    print "$P_txt $en1= @PV_per_hour[0..23]\n";
    print_arr_perc("$V_txt $per3$en1     = ", \@PV_by_hour,
                   $PV_gross_sum * $years, 0, 21, 3);
    print_arr_perc("$V_txt $per_m$de1        = ", \@PV_by_month,
                   $PV_gross_sum * $years, 1, 12, 1) unless $test;
}
print "$net_max_txt $en4$en1      =".W($PV_net_max)." $on $PV_net_max_time\n";
print "$PV_loss_txt $en2 $en2 $en2  ="       .kWh($PV_net_losses).
    " $during ".round($PV_net_loss_hours)." h $by_curb_at $curb W\n" if $curb;
print "$PV_net_txt $en2             =" .kWh($PV_net_sum).
    " $at $ieff_txt ".percent($inverter_eff)."%\n";
#print "$yield $daytime  =   $yield_daytime %\n";
#my $yield_daytime =
#    percent($PV_net_sum ? $PV_net_bright_sum / $PV_net_sum : 0);

print "\n";
print "$consumpt_txt    =" .kWh($load_sum)."\n";
print "$load_const_txt $en1             =".W($load_const)."  $load_during_txt\n"
    if defined $load_const;
if (defined $capacity) {
    print "\n".
        "$capacity_txt $en1          =" .W($capacity)."h"
        ." $with $soc_txt ".percent($soc)."%, $dod_txt ".percent($dod)."%"
        .", $coupled_txt\n";
    print "$optimal_charge" unless defined $bypass;
    print "$bypass_txt $en3          =".W($bypass)
        .($bypass_spill ? "  " : "")."$spill_txt" if defined $bypass;
    print ", $max_charge_txt $max_charge C\n";
    print "$optimal_discharge" unless defined $max_feed;
    print "$feed_txt $en3        ".($const_feed ? "" : " ")
        ."=".W($max_feed)."$feed_during_txt" if defined $max_feed;
    print ", $max_dischg_txt $max_dischg C\n";
    print "$AC_coupl_loss_txt $en3$en3  =" .kWh($AC_coupling_losses)."\n"
        if $AC_coupled;
    print "$spill_loss_txt $en3 $en3 $en3   =".kWh($spill_loss)."\n";
    print "$charging_loss_txt $de2              =".kWh($charging_loss)
        ." $due_to $ceff_txt ".percent($charge_eff)."%\n";
    print "$storage_loss_txt $en3            =".kWh($storage_loss)
        ." $due_to $seff_txt ".percent($storage_eff)."%\n";
    print "$own_storage_txt $en2   =".kWh($PV_used_via_storage)."\n";
    if ($verbose && defined $capacity) {
        print "$c_txt$de2= @charge_per_hour[0..23]\n";
        print "$C_txt". " = @dischg_per_hour[0..23]\n";
    }
    print "$stored_txt $en2$en2        =" .kWh($charge_sum)." ($after $charging_loss_txt)\n";
    # Vollzyklen, Kapazitätsdurchgänge pro Jahr Kapazitätsdurchsatz:
    printf "$cycles_txt $de1       =  %3d $of_eff_cap_txt\n", $cycles;

    my $grid_feed_sum_alt = $PV_net_sum - $PV_used_sum - $AC_coupling_losses
        - $spill_loss - $charging_loss - $storage_loss - $charge;
    my $discrepancy = $grid_feed_sum - $grid_feed_sum_alt;
    die "Internal error: grid feed-in calculation discrepancy $discrepancy: ".
        "grid feed-in $grid_feed_sum vs. ".int($grid_feed_sum_alt + .5)." =\n".
        "PV net sum $PV_net_sum - PV used $PV_used_sum".
        " - AC coupling loss $AC_coupling_losses - loss by spill $spill_loss".
        " - charging loss $charging_loss - storage loss $storage_loss".
        " - charge $charge" if abs($discrepancy) > 0.001; # 1 mWh
    print "\n";
}

print "$own_txt $en4 $en3         =" .kWh($PV_used_sum)."\n";
print "$usage_loss_txt $en3$en3  =" .kWh($PV_use_loss_sum)."$net_de $during "
    .round($PV_use_loss_hours)." h $by_curb_at $curb W\n" if $curb;
print "$grid_feed_txt $en3            =" .kWh($grid_feed_sum)."\n";
print "$F_txt = @grid_feed_per_hour[0..23]\n" if $verbose;
print "$own_ratio_txt $en4 $en4  =  ".sprintf("%3d", percent($own_ratio))
    ." % $of_yield\n";
my $load_coverage_str = sprintf("%3d", percent($load_coverage));
print "$load_cover_txt $en1        =  $load_coverage_str % $of_consumption\n";
