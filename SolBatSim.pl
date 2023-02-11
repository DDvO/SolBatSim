#!/usr/bin/perl
################################################################################
# Eigenverbrauchs-Simulation mit stündlichen PV-Daten und einem mindestens
# stündlichen Lastprofil, optional mit Ausgangs-Drosselung des Wechselrichters
# und optional mit Stromspeicher (Batterie o.ä.)
#
# Nutzung: Solar.pl <Lastprofil-Datei> [<Jahresverbrauch in kWh>]
#            (<PV-Daten-Datei> [<Brutto-Nennleistung in Wp>])+
#            [-load <konstante Last in W> [<Zahl der Tage pro Woche, sonst 5>:
#                   <von Uhrzeit, sonst 8 Uhr>..<bis Uhrzeit, sonst 16 Uhr>]]
#            [-peff <PV-System-Wirkungsgrad in %, ansonsten von PV-Daten-Datei>]
#            [-capacity <Speicherkapazität Wh, ansonsten 0 (kein Batterie)>]
#            [-ac] {AC-gekoppelter Speicher, ohne Verlust durch Überlauf}
#            [-pass [spill] <Speicher-Umgehung in W, opt. auch bei Überfluss>]
#            [-feed (max <begrenzte bedarfsgerechte Entladung aus Speicher in W>
#                    | [<von Uhrzeit, sonst 0 Uhr>..<bis Uhrzeit, sonst 24 Uhr>]
#                      <konstante Entladung aus Speicher in W> )]
#            [-ceff <Lade-Wirkungsgrad in %, ansonsten 94]
#            [-seff <Speicher-Wirkungsgrad in %, ansonsten 95]
#            [-ieff <Wechselrichter-Wirkungsgrad in %, ansonsten 94]
#            [-test <Lastpunkte pro Stunde, für Prüfberechnung über 24 Stunden>]
#            [-en] [-tmy] [-curb <Wechselrichter-Ausgangsleistungs-Limit in W>]
#            [-hour <Datei>] [-day <Datei>] [-week <Datei>] [-month <Datei>]
# Mit "-en" erfolgen die Textausgaben auf Englisch. Fehlertexte sind englisch.
# Wenn PV-Daten für mehrere Jahre gegeben sind, wird der Durchschnitt berechnet
# oder mit Option "-tmy" Monate für ein typisches meteorologisches Jahr gewählt.
# Mit den Optionen "-hour"/"-day"/"-week"/"-month" wird jeweils eine CSV-Datei
# mit gegebenen Namen mit Statistik-Daten pro Stunde/Tag/Woche/Monat erzeugt.
#
# Beispiel:
# Solar.pl Lastprofil.csv 3000 Solardaten_1215_kWh.csv 1000 -curb 600 -tmy
#
# Beziehe Solardaten für Standort von https://re.jrc.ec.europa.eu/pvg_tools/de/
# Wähle den Standort und "DATEN PRO STUNDE", setze Häkchen bei "PV-Leistung".
# Optional "Installierte maximale PV-Leistung" und "Systemverlust" anpassen.
# Bei Nutzung von "-tmy" Startjahr 2008 oder früher, Endjahr 2018 oder später.
# Dann den Download-Knopf "csv" drücken.
#
################################################################################
# Simulation of actual own consumption of photovoltaic power output according
# to load profiles with a resolution of at least one hour, typically per minute.
# Optionally takes into account power output cropping by solar inverter.
# Optionally with energy storage (using a battery or the like).
#
# Usage: Solar.pl <load profile file> [<consumption per year in kWh>]
#          (<PV data file> [<nominal gross power in Wp>])+
#          [-load <constant load in W> [<count of days per week, default 5>:
#                 <from hour, default 8 o'clock>..<to hour, default 16>]]
#          [-peff <PV system efficiency in %, default from PV data file(s)>]
#          [-capacity <storage capacity in Wh, default 0 (no battery)>]
#          [-ac] {AC-coupled charging after inverter, without loss by spill}
#          [-pass [spill] <storage bypass in W, optionally also on surplus>]
#          [-feed (max <limited feed-in from storage in W according to load>
#                  | [<von Uhrzeit, sonst 0 Uhr>..<bis Uhrzeit, sonst 24 Uhr>]
#                    <constant feed-in from storage in W> )]
#          [-ceff <charging efficiency in %, default 94]
#          [-seff <storage efficiency in %, default 95]
#          [-ieff <inverter efficiency in %, default 94]
#          [-test <load points per hour, for debug calculation over 24 hours>]
#          [-en] [-tmy] [-curb <inverter output power limit in W>]
#          [-hour <file>] [-day <file>] [-week <file>] [-month <file>]
# Use "-en" for text output in English. Error messages are all in English.
# When PV data for more than one year is given, the average is computed, while
# with the option "-tmy" months for a typical meteorological year are selected.
# With each the options "-hour"/"-day"/"-week"/"-month" a CSV file is produced
# with the given name containing with statistical data per hour/day/week/month.
#
# Example:
# Solar.pl loadprofile.csv 3000 solardata_1215_kWh.csv 1000 -curb 600 -tmy -en
#
# Take solar data from https://re.jrc.ec.europa.eu/pvg_tools/
# Select location, click "HOURLY DATA", and set the check mark at "PV power".
# Optionally may adapt "Installed peak PV power" and "System loss" already here.
# For using TMY data, choose Start year 2008 or earlier, End year 2018 or later.
# Then press the download button marked "csv".
#
# (c) 2022-2023 David von Oheimb - License: MIT - Version 2.3
################################################################################

use strict;
use warnings;

die "Missing command line arguments" if $#ARGV < 0;
my $test         = 0; # unless 0, number of test load points per hour
my $load_profile = shift @ARGV unless $ARGV[0] =~ m/^-/; # file name
my $consumption  = shift @ARGV # kWh/year, default is implicit from load profile
    if $#ARGV >= 0 && $ARGV[0] =~ m/^[\d\.]+$/;
my $load_const;     # constant load in W, during certain times:
my $load_days    =  5;  # count of days per week with constant load
my $load_from    =  8;  # hour of constant load begin
my $load_to      = 16;  # hour of constant load end

use constant YearHours => 24 * 365;

my @PV_files;
my @PV_peaks;       # nominal/maximal PV output(s), default from PV data file(s)
my ($lat, $lon);    # from PV data file(s)
my $pvsys_eff;      # PV system efficiency, default from PV data file(s)
my $inverter_eff;   # inverter efficiency; default see below
my $capacity;       # usable storage capacity in Wh on average degradation
my $bypass_spill;   # bypass storage on surplus (i.e., when storge is full)
my $bypass;         # direct feed to inverter in W, bypassing storage
my $max_feed;       # maximal feed-in in W from storage
my $const_feed = 1; # constant feed-in, relevant only if defined $max_feed
my $feed_from = 0;  # hour of constant feed-in begin
my $feed_to   = 24; # hour of constant feed-in end
my $AC_coupled;     # by default, charging is DC-coupled (w/o inverter)
my $charge_eff;     # charge efficiency; default see below
my $storage_eff;    # storage efficiency; default see below
# storage efficiency and discharge efficiency could have been combined
my $nominal_power_sum = 0;

while ($#ARGV >= 0 && $ARGV[0] =~ m/^\s*[^-]/) {
    push @PV_files, shift @ARGV; # PV data file
    push @PV_peaks, $#ARGV >= 0 && $ARGV[0] =~ m/^[\d\.]+$/ ? shift @ARGV : 0;
}

sub no_arg {
    shift @ARGV;
    return 1;
}
sub num_arg {
    my $opt = $ARGV[0];
    die "Missing number argument for $opt option"
        unless $#ARGV >= 1 && $ARGV[1] =~ m/^[-\d\.]+$/;
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

my ($en, $tmy, $curb, $max, $hourly, $daily, $weekly,  $monthly);
while ($#ARGV >= 0) {
    if      ($ARGV[0] eq "-test"    ) { $test         = num_arg();
    } elsif ($ARGV[0] eq "-en"      ) { $en           =  no_arg();
    } elsif ($ARGV[0] eq "-load"    ) { $load_const   = num_arg();
                                       ($load_days, $load_from, $load_to)
                                           = ($1, $2, $3) if $#ARGV >= 0 &&
                                           $ARGV[0] =~ m/^(\d+):(\d+)\.\.(\d+)$/
                                           && shift @ARGV;
    } elsif ($ARGV[0] eq "-peff"    ) { $pvsys_eff    = eff_arg();
    } elsif ($ARGV[0] eq "-tmy"     ) { $tmy          =  no_arg();
    } elsif ($ARGV[0] eq "-curb"    ) { $curb         = num_arg();
    } elsif ($ARGV[0] eq "-max"     ) { $max          = str_arg();
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
    } elsif ($ARGV[0] eq "-ceff"    ) { $charge_eff   = eff_arg();
    } elsif ($ARGV[0] eq "-seff"    ) { $storage_eff  = eff_arg();
    } elsif ($ARGV[0] eq "-ieff"    ) { $inverter_eff = eff_arg();
    } elsif ($ARGV[0] eq "-hour"    ) { $hourly       = str_arg();
    } elsif ($ARGV[0] eq "-day"     ) { $daily        = str_arg();
    } elsif ($ARGV[0] eq "-week"    ) { $weekly       = str_arg();
    } elsif ($ARGV[0] eq "-month"   ) { $monthly      = str_arg();
    } else { die "Invalid option: $ARGV[0]";
    }
}

die "Numeric argument of -test option must not be 0" if defined $test && !$test;
use constant TEST_START      =>  6; # load starts in the morning
use constant TEST_LENGTH     => 24; # until 6 AM on the next day
use constant TEST_END        => TEST_START + TEST_LENGTH;
use constant TEST_DROP_START => 10; # start hour of load drop
use constant TEST_DROP_LEN   =>  4; # length of load drop around noon
use constant TEST_DROP_END   => TEST_DROP_START + TEST_DROP_LEN;
use constant TEST_PV_START   =>  7; # start hour of PV yield
use constant TEST_PV_END     => 17; # end   hour of PV yield
use constant TEST_PV_GROSS   => 1100; # in Wp
use constant TEST_LOAD       => 900; # in W
use constant TEST_LOAD_LEN   =>
    TEST_END > TEST_DROP_END ? TEST_LENGTH - TEST_DROP_LEN
    : (TEST_END <= TEST_DROP_START ? TEST_END - TEST_START
       : TEST_DROP_START - TEST_START);
my $test_load = defined $consumption ?
    $consumption * 1000 / TEST_LOAD_LEN : TEST_LOAD if $test;
if ($test) {
    $load_profile  = "test load data";
    my $pv_gross   = $#PV_peaks < 0 ? TEST_PV_GROSS : $PV_peaks[0]; # in W
    my $pv_power   = defined $pvsys_eff ? $pv_gross * $pvsys_eff : 1000; # in W
    # after PV system losses, which may be derived as follows:
    $pvsys_eff     = $pv_gross ? $pv_power / $pv_gross : 0.92
        unless defined $pvsys_eff;
    $inverter_eff  =  0.8 unless defined $inverter_eff;
    $consumption = $test_load/1000 * TEST_LOAD_LEN unless defined $consumption;
    push @PV_files, "test PV data" if $#PV_files < 0;
    push @PV_peaks, $pv_gross      if $#PV_peaks < 0;
    $charge_eff    =  0.9 if defined $capacity && !defined $charge_eff;
    my $pv_net_bat =  600; # in W, after inverter losses
    $storage_eff   =  $pv_net_bat / ($pv_power * $charge_eff * $inverter_eff)
        if defined $capacity && !defined $storage_eff;
}

die "Missing load profile file name - should be first CLI argument"
    unless defined $load_profile;
die "Missing PV data file name" if $#PV_peaks < 0;
if (defined $load_const) {
    die "days count for -load option must be in range 0..7"
        if  $load_days > 7;
    die "begin hour for -load option must be in range 0..24"
        if  $load_from > 24;
    die "end hour for -load option must be in range 0..24"
        if  $load_to > 24;
}
$inverter_eff = 0.94    unless defined $inverter_eff;
if (defined $capacity) {
    $charge_eff  = 0.94 unless defined $charge_eff;
    $storage_eff = 0.95 unless defined $storage_eff;
    die "begin hour for -feed option must be in range 0..24"
        if  $feed_from > 24;
    die "end hour for -feed option must be in range 0..24"
        if  $feed_to > 24;
} else {
    die   "-ac option requires -capacity option" if defined $AC_coupled;
    die "-pass option requires -capacity option" if defined $bypass;
    die "-feed option requires -capacity option" if defined $max_feed;
    die "-ceff option requires -capacity option" if defined $charge_eff;
    die "-seff option requires -capacity option" if defined $storage_eff;
}

sub never_0 { return $_[0] == 0 ? 1 : $_[0]; }
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
    return sprintf("%02d", shift)."-".sprintf("%02d", shift).
        " ".($en ? "at" : "um")." ".sprintf("%02d", shift);
}
sub time_string {
    return date_string(shift, shift, shift).sprintf(":%02d h", int(shift));
}
sub minute_string {
    my ($item, $items) = (shift, shift);
    my $minute = sprintf(":%02d", int( 60 * $item / $items));
    $minute .= sprintf(":%02d", int((60 * $item % $items) / $items * 60))
        unless (60 % $items == 0);
    return $minute
}

sub round_1000 { return round(shift() / 1000); }
sub kWh     { return sprintf("%5.2f kWh", shift() / 1000) if $test;
              return sprintf("%5d kWh", round_1000(shift()  )); }
sub W       { return sprintf("%5d W"  , round(shift()       )); }
sub percent { return sprintf("%2d"    , round(shift() *  100)); }

# all hours according to local time without switching for daylight saving
use constant NIGHT_START   =>  0; # at night (with just basic load)
use constant NIGHT_END     =>  6;
use constant MORNING_START =>  6; # in early morning
use constant MORNING_END   =>  9;
use constant BRIGHT_START  =>  9; # bright sunshine time
use constant BRIGHT_END    => 15;
use constant LAFTERN_START => 15; # in late afternoon
use constant LAFTERN_END   => 18;
use constant EVENING_START => 18; # after sunset
use constant EVENING_END   => 24;
use constant WINTER_START  => 10; # dark season
use constant WINTER_END    => 04;

my $sum_items = 0;
my $items = 0; # number of load measure points in current hour
my $load_max = 0;
my $load_max_time;
my ($load_sum, $night_sum, $morning_sum, $bright_sum,
    $earleve_sum, $evening_sum, $winter_sum) = (0, 0, 0, 0, 0, 0, 0);

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

my @items_by_hour;
my @load_by_item;
my @load_by_hour;
my @load_by_weekday;
my $hour_per_year = 0;
sub get_profile {
    my $file = shift;
    open(my $IN, '<', $file) or die "Could not open profile file $file: $!\n"
        unless $test;

    my $warned_just_before = 0;
    my $weekday = 4; # HTW load profiles start on Friday (of 2010), Monday == 0
    while (my $line =
           !$test ? <$IN>
           : "date".(",".($hour_per_year < TEST_START ||
                          ($hour >= TEST_DROP_START && $hour <  TEST_DROP_END)
                          ? 0 : $test_load)) x $test ."\n") {
        chomp $line;
        next if $line =~ m/^\s*#/; # skip comment line
        my @sources = split "," , $line;
        my $n = $#sources;
        shift @sources; # ignore date & time info
        if ($items > 0 && $items != $n) {
            print "Warning: unequal number of items per line: $n vs. $items in "
                ."$file\n";
            $items = -1;
        }
        $items = $n if $items >= 0;
        $sum_items += ($items_by_hour[$month][$day][$hour] = $n);
        my $hload = 0;
        if (defined $load_const && $weekday < $load_days &&
            ($load_from>$load_to ? ($load_from <= $hour || $hour < $load_to)
                                 : ($load_from <= $hour && $hour < $load_to))) {
            $items_by_hour[$month][$day][$hour] = 1;
            $load_by_item[$month][$day][$hour][0] = $hload = $load_const;
            if ($hload > $load_max) {
                $load_max = $hload;
                $load_max_time = time_string($month, $day, $hour, 0);
            }
        } else {
            for (my $item = 0; $item < $n; $item++) {
                my $load = $sources[$item];
                die "Error parsing load item: '$load' in $file line $."
                    unless $load =~ m/^\s*-?[\.\d]+\s*$/;
                if ($load > $load_max) {
                    $load_max = $load;
                    $load_max_time =
                        time_string($month, $day, $hour, 60 * $item / $n);
                }
                $hload += $load;
                $load_by_item[$month][$day][$hour][$item] = $load;
                if ($load <= 0) {
                    my $lang = $en;
                    $en = 1;
                    print "Warning: load on YYYY-".
                        date_string($month, $day, $hour)
                        .sprintf("h, item %4d", $item)." = $load\n"
                        unless $test || $warned_just_before;
                    $en = $lang;
                    $warned_just_before = 1;
                } else {
                    $warned_just_before = 0;
                }
            }
            $hload /= $n;
        }
        $load_by_hour[$month][$day][$hour] += $hload;
        $load_by_weekday[$weekday] += $hload;
        $load_sum    += $hload;
        $night_sum   += $hload if   NIGHT_START <= $hour && $hour <   NIGHT_END;
        $morning_sum += $hload if MORNING_START <= $hour && $hour < MORNING_END;
        $bright_sum  += $hload if  BRIGHT_START <= $hour && $hour <  BRIGHT_END;
        $earleve_sum += $hload if LAFTERN_START <= $hour && $hour < LAFTERN_END;
        $evening_sum += $hload if EVENING_START <= $hour && $hour < EVENING_END;
        $winter_sum  += $hload if WINTER_START <= $month || $month < WINTER_END;

        if (++$hour == 24) {
            $hour = 0;
            $weekday = 0 if ++$weekday == 7;
        }
        adjust_day_month();
        $hour_per_year++;
        last if $test && $hour_per_year == TEST_END;
    }
    close $IN unless $test;
    $month--;
    check_consistency($month, 12, "months", $file);
    check_consistency($hour_per_year, YearHours, "hours", $file);
}

get_profile($load_profile);

my $profile_txt = $en ? "load profile file" : "Lastprofil-Datei";
my $pv_data_txt = $en ? "PV data file(s)"   : "PV-Daten-Datei(en)";
my $p_txt = $en ? "load data points per hour  " : "Last-Datenpunkte pro Stunde";
my $t_txt = $en ? "total cons. acc. to profile" : "Verbrauch gemäß Lastprofil ";
my $W_txt = $en ? "portion per weekday (Mo-Su)" :"Anteil pro Wochentag (Mo-So)";
my $n_txt = $en ? "portion 12 AM -  6 AM      " : "Anteil  0 -  6 Uhr MEZ     ";
my $m_txt = $en ? "portion  6 AM -  9 PM      " : "Anteil  6 -  9 Uhr MEZ     ";
my $s_txt = $en ? "portion  9 AM -  3 PM      " : "Anteil  9 - 15 Uhr MEZ     ";
my $a_txt = $en ? "portion  3 AM -  6 PM      " : "Anteil 15 - 18 Uhr MEZ     ";
my $e_txt = $en ? "portion  6 PM - 12 PM      " : "Anteil 18 - 24 Uhr MEZ     ";
my $w_txt = $en ? "portion October-March      " : "Anteil Oktober - März      ";
my $b_txt = $en ? "basic load                 " : "Grundlast                  ";
my $M_txt = $en ? "maximal load               " : "Maximallast                ";
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
print "$profile_txt$de1$s10 : $load_profile\n" unless $test;
print "$p_txt = ".sprintf("%4d", $sum_items / $hour_per_year)."\n";
print "$t_txt =".kWh($load_sum)."\n";
if ($load_sum != 0) {
    if (!$test) {
        print "$W_txt=   ";
        for (my $weekday = 0; $weekday < 7; $weekday++) {
            print "".percent($load_by_weekday[$weekday] / $load_sum)." %";
            print ", " unless $weekday == 6;
        }
    }
    print "\n";
    print "$n_txt =   ".percent($night_sum   / $load_sum)." %\n";
    print "$m_txt =   ".percent($morning_sum / $load_sum)." %\n";
    print "$s_txt =   ".percent($bright_sum  / $load_sum)." %\n";
    print "$a_txt =   ".percent($earleve_sum / $load_sum)." %\n";
    print "$e_txt =   ".percent($evening_sum / $load_sum)." %\n";
    print "$w_txt =   ".percent($winter_sum  / $load_sum)." %\n" unless $test;
}
print "$b_txt =".W($night_sum / ($test ? TEST_LENGTH / 24 : 365)
                   / (NIGHT_END - NIGHT_START))."\n";
print "$M_txt =".W($load_max)." $on $load_max_time\n";
print "\n";

################################################################################
# read PV production data

my @PV_gross_out;
my ($start_year, $years);
sub get_power {
    my ($file, $nominal_power) = (shift, shift);
    open(my $IN, '<', $file) or die "Could not open PV data file $file: $!\n"
        unless $test;
    print "".($en ? "PV data file  " : "PV-Daten-Datei")."$s13 : $file"
        unless $test;

    # my $sum_needed = 0;
    my ($slope, $azimuth);
    my $current_years = 0;
    my $nominal_power_deflt; # PVGIS default: 1 kWp
    my     $pvsys_eff_deflt; # PVGIS default: 0.86
    my      $power_provided; # PVGIS default: none
    my $power_rate;
    my ($months, $hours) = (0, 0);
    while ($_ = !$test ? <$IN>
           : "2000010".(1 + int($hours / 24)).":"
           .sprintf("%02d", $hours % 24)."00, "
           .(TEST_PV_START <= $hours && $hours < TEST_PV_END ?
             $nominal_power : 0)."\n") {
        chomp;
        if (m/^Latitude \(decimal degrees\):\s*([\d\.]+)/) {
            check_consistency($1, $lat, "latitude", $file) if $lat;
            $lat = $1;
        }
        if (m/^Longitude \(decimal degrees\):\s*([\d\.]+)/) {
            check_consistency($1, $lon, "longitude", $file) if $lon;
            $lon = $1;
        }
        $slope   = $1 if (!$slope   && m/^Slope:\s*(.+)$/);
        $azimuth = $1 if (!$azimuth && m/^Azimuth:\s*(.+)$/);
        if (!$nominal_power_deflt && m/Nominal power.*? \(kWp\):\s*([\d\.]+)/) {
            $nominal_power_deflt = $1 * 1000;
            $nominal_power = $nominal_power_deflt unless $nominal_power;
            $nominal_power_sum += $nominal_power;
        }
        if (!$pvsys_eff_deflt && m/System losses \(%\):\s*([\d\.]+)/) {
            my $seff = 1 - $1/100;
            my $eff = percent($seff);
            $pvsys_eff_deflt = $seff / $inverter_eff_never_0;
            print ", ".($en ?
                        "contained system efficiency $eff% was overridden" :
                        "enthaltene System-Effizienz $eff% wurde übersteuert")
                if !$test && $pvsys_eff && $pvsys_eff != $pvsys_eff_deflt;
            $pvsys_eff = $pvsys_eff_deflt unless defined $pvsys_eff;
            if ($pvsys_eff > 1) {
                print "\n";
                die "unreasonable PV system efficiency ".round($pvsys_eff * 100)
                    ."% - have -peff and -ieff been used properly?";
            }
            $power_rate = $nominal_power / $nominal_power_deflt
                / ($pvsys_eff_deflt * $inverter_eff);
        }
        $power_provided = 1 if m/^time,P,/;

        next unless m/^20\d\d\d\d\d\d:\d\d\d\d,/;
        unless ($test) {
            die "Missing latitude in $file"  unless $lat;
            die "Missing longitude in $file" unless $lon;
            die "Missing slope in $file"     unless $slope;
            die "Missing azimuth in $file"   unless $azimuth;
            die "Missing nominal power in $file" unless $nominal_power_deflt;
            die "Missing system efficiency in $file" unless $pvsys_eff_deflt;
            die "Missing PV power output data in $file" unless $power_provided;
        }

        next if m/^20\d\d0229:/; # skip data of Feb 29th (in leap year)
        $start_year = $1 if (!$start_year && m/^(\d\d\d\d)/);
        if ($tmy) {
            # typical metereological year
            my $selected_month = 0;
            $selected_month = $1 if m/^2016(01)/;
            $selected_month = $1 if m/^2016(02)/;
            $selected_month = $1 if m/^2012(03)/;
            $selected_month = $1 if m/^2008(04)/;
            $selected_month = $1 if m/^2011(05)/;
            $selected_month = $1 if m/^2010(06)/;
            $selected_month = $1 if m/^2012(07)/;
            $selected_month = $1 if m/^2014(08)/;
            $selected_month = $1 if m/^2015(09)/;
            $selected_month = $1 if m/^2017(10)/;
            $selected_month = $1 if m/^2013(11)/;
            $selected_month = $1 if m/^2018(12)/;
            next unless $selected_month;
        }
        $current_years++ if m/^20..0101:00/;
        $months++ if m/^20....01:00/;

        die "Missing power data in $file line $_"
            unless m/^(\d\d\d\d)(\d\d)(\d\d):(\d\d)(\d\d),\s?([\d\.]+)/;
        my ($year, $month, $day, $hour, $minute_unused, $power) =
            ($tmy ? $start_year : $1, $2, $3, $4, $5, $6);
        $power *= $power_rate unless $test;
        $PV_gross_out[$year - $start_year][$month][$day][$hour] += $power;
        $hours++;
        last if $test && $hours == TEST_END;
    }
    close $IN unless $test;
    check_consistency($months, 12 * $current_years, "months", $file);
    check_consistency($hours, YearHours * $current_years, "hours", $file);
    check_consistency($years, $current_years, "years", $file) if $years;
    $years = $current_years;
    if ($test) {
        $years = 1;
        return $nominal_power;
    }

    $slope   = " $slope"   unless $slope   =~ m/^-/;
    $azimuth = " $azimuth" unless $azimuth =~ m/^-/;
    $slope   = " $slope"   unless $slope   =~ m/\d\d/;
    $azimuth = " $azimuth" unless $azimuth =~ m/\d\d/;
    $slope   =~ s/ deg\./°/;
    $azimuth =~ s/ deg\./°/;
    $slope   =~ s/optimum/opt./;
    $azimuth =~ s/optimum/opt./;
    print "\n".($en ? "slope         " : "Neigungswinkel")."$s13 =  $slope\n";
    print "".($en ? "azimuth       " : "Azimut        ")."$s13 =  $azimuth\n";
    return $nominal_power;
}

for (my $i = 0; $i <= $#PV_files; $i++) {
    $PV_peaks[$i] = get_power($PV_files[$i], $PV_peaks[$i]);
}
my $PV_peaks = join("+", @PV_peaks);

################################################################################
# PV usage simulation

my $PV_gross_out_sum = 0;
my $PV_gross_max = 0;
my $PV_gross_max_tm;
my @PV_net_out;
my $PV_net_out_sum = 0;
my $PV_net_bright_sum = 0;

my $load_scale = defined $consumption && $load_sum != 0
    ? 1000 * $consumption / $load_sum : 1;
my $load_scale_never_0 = $load_scale != 0 ? $load_scale : 1;
my @PV_net_loss;
my $PV_net_losses = 0;
my $PV_net_loss_hours = 0;
my @PV_usage_loss_by_item;
my @PV_usage_loss;
my $PV_usage_losses = 0;
my $PV_usage_loss_hours = 0;
my @PV_used_by_item;
my @PV_used;
my $PV_used_sum = 0;
my $PV_used_via_storage = 0;
my $charge        = 0 if defined $capacity;
my $charge_sum    = 0 if defined $capacity;
my $charging_loss = 0 if defined $capacity;
my $spill_loss    = 0 if defined $capacity;
my $AC_coupling_losses = 0 if defined $capacity;
my $grid_feed_in  = 0;

sub simulate()
{
    my $year = 0;
    ($month, $day, $hour) = (1, 1, 0);
    my $minute = 0; # currently fixed

    # factor out $load_scale for optimizing the inner loop
    $capacity /= $load_scale_never_0 if $capacity;
    $bypass /= $load_scale_never_0 if defined $bypass;
    $max_feed /= ($load_scale_never_0 * $storage_eff_never_0)
        if defined $max_feed;

    while ($year < $years) {
        my $power = $PV_gross_out[$year][$month][$day][$hour];
        if (!defined $power) {
            $en = 1;
            die "No power data at ".($start_year + $year)."-".
                time_string($month, $day, $hour, $minute);
        }
        $PV_gross_out_sum += $power;
        if ($power > $PV_gross_max) {
            $PV_gross_max = $power;
            $PV_gross_max_tm = ($tmy ? "TMY" : $start_year + $year)."-"
                .time_string($month, $day, $hour, $minute);
        }
        my $net_pv_power = $power * $pvsys_eff * $inverter_eff;

        $PV_net_loss[$month][$day][$hour] = 0;
        my $PV_loss = 0;
        if ($curb && $net_pv_power > $curb) { # TODO adapt to storage use
            $PV_loss = $net_pv_power - $curb;
            $PV_net_loss[$month][$day][$hour] += $PV_loss;
            $PV_net_losses += $PV_loss;
            $PV_net_loss_hours++;
            $net_pv_power = $curb;
            # print "$year-".time_string($month, $day, $hour, $minute).
            #"\tPV=".round($net_pv_power)."\tcurb=".round($curb).
            #"\tloss=".round($PV_net_losses)."\t$_\n";
        }
        $PV_net_out[$month][$day][$hour] += $net_pv_power;
        $PV_net_out_sum += $net_pv_power;
        $PV_net_bright_sum += $net_pv_power
            if BRIGHT_START <= $hour && $hour < BRIGHT_END;

        # factor out $load_scale for optimizing the inner loop
        $net_pv_power /= $load_scale_never_0;
        $PV_loss /= $load_scale_never_0;
        # my $needed = 0;
        my $usages = 0;
        my $curb_losses_this_hour = 0;
        my $AC_coupling_losses_this_hour = 0;
        $items = $items_by_hour[$month][$day][$hour];

        # factor out $items for optimizing the inner loop
        if (defined $capacity && $items != 1) {
            $capacity *= $items;
            $charge *= $items;
            $charge_sum *= $items;
            # $charging_loss *= $items;
            $spill_loss *= $items;
            $PV_used_via_storage *= $items;
            $grid_feed_in *= $items;
        }

        my $test_started = $test && ($day - 1) * 24 + $hour >= TEST_START;
        # my $feed_sum = 0 if defined $max_feed;
        for (my $item = 0; $item < $items; $item++) {
            my $loss = 0;
            my $power_needed = $load_by_item[$month][$day][$hour][$item];
            die "load_by_item[$month][$day][$hour][$item] is undefined"
                unless defined $power_needed;
            printf("%02d".minute_string($item, $items)." load=%4d PV net=%4d ",
                   $hour, $power_needed, $net_pv_power) if $test_started;
            # $needed += $power_needed;
            # load will be reduced by constant $bypass or $bypass_spill

            # $pv_used locally accumulates PV own consumption
            # $grid_feed_in accumulates feed to grid

            # feed by constant bypass or just as much as used (optimal charge)
            my $pv_used = defined $bypass ? $bypass * $inverter_eff
                : $power_needed; # preliminary
            my $avail_power = $net_pv_power - $pv_used;
            if ($avail_power < 0) {
                $pv_used = $net_pv_power;
                # == min($net_pv_power,
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

            if (defined $capacity) { # storage available
                my $capacity_to_fill = $capacity - $charge;
                $capacity_to_fill = 0 if $capacity_to_fill < 0;
                my ($charge_input, $charge_delta) = (0, 0);
                if ($avail_power > 0) {
                    # when charging is DC-coupled, no loss through inverter
                    $avail_power /= $inverter_eff_never_0 unless $AC_coupled;
                    # $avail_power is the power available for charging
                    my $need_for_fill = $capacity_to_fill / $charge_eff_never_0;
                    # optimal charge: exactly as much as unused and fits in
                    $charge_input = $avail_power;
                        # will become min($avail_power, $need_for_fill);
                    my $surplus = $avail_power - $need_for_fill;
                    printf("[avail=%4d,surplus=%4d] ",
                           $avail_power, max($surplus, 0) +.5) if $test_started;
                    if ($surplus > 0) {
                        $charge_input = $need_for_fill; # TODO check AC-coupled
                        my $surplus_net = $surplus;
                        $surplus_net *= $inverter_eff unless $AC_coupled;
                        if (!defined $bypass) {
                            $grid_feed_in += $surplus_net; # on optimal charge
                            printf(" (grid feed=%4d)", $surplus_net)
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
                        printf("              "); # no surplus
                    }

                    # The following adds reduced charging due to curb
                    # to the usage losses, which is not entirely correct
                    # in case of constant feed (non-optimal discharge).
                    $loss += min($PV_loss, $capacity_to_fill
                                 * $storage_eff * $inverter_eff)
                        if $AC_coupled && $PV_loss != 0;

                    $charge_delta = $charge_input * $charge_eff;
                    $charge += $charge_delta;
                    $charge_sum += $charge_delta;
                    # $charging_loss += $charge_input - $charge_delta;
                } elsif ($test_started) {
                    printf("          "); # no $avail_power
                    printf("              "); # no surplus
                }
                my $print_charge = $test_started &&
                    ($avail_power > 0 || $charge > 0);
                printf("chrg loss=%4d dischrg needed=%4d [charge %4d + %4d ",
                       ($charge_input - $charge_delta) *
                       ($AC_coupled ? 1 : $inverter_eff) + .5,$power_needed +.5,
                       $charge - $charge_delta, $charge_delta) if $print_charge;
                my $AC_loss = 0;
                if ($charge > 0) { # storage not empty
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
                    $discharge = $charge if $discharge > $charge;
                    printf("- lost=%4d - %4d ", $discharge * (1 - $storage_eff)
                           + .5, $discharge*$storage_eff + .5) if $test_started;
                    if ($discharge != 0) {
                        $charge -= $discharge; # includes storage loss
                        $discharge *= $storage_eff;
                        my $discharge_net = $discharge * $inverter_eff;
                        if ($AC_coupled) {
                            $AC_loss = $discharge - $discharge_net;
                            $AC_coupling_losses_this_hour += $AC_loss;
                        }
                        if (defined $max_feed && $const_feed) {
                            # $feed_sum += $discharge;
                            my $dis_feed_in = $discharge_net - $power_needed;
                            print "####$dis_feed_in = $discharge_net - $power_needed\n";
                            if ($dis_feed_in > 0) {
                                $grid_feed_in += $dis_feed_in;
                                $discharge_net -= $dis_feed_in;
                            }
                        }
                        $pv_used += $discharge_net;
                        $PV_used_via_storage += $discharge_net;
                    }
                }
                printf("= %4d] ", $charge) if $print_charge;
                printf("AC coupling loss=%4d ", $AC_loss + .5)
                    if $print_charge && $AC_coupled;
            }
            printf("used=%4d\n", $pv_used + .5) if $test_started;
            $usages += $pv_used;
            $PV_used_by_item[$month][$day][$hour][$item] =
                round($pv_used * $load_scale) if $max;

            if ($PV_loss != 0 && $power_needed > 0) {
                $loss += min($PV_loss, $power_needed);
            }
            if ($loss != 0) {
                $curb_losses_this_hour += $loss;
                $PV_usage_loss_by_item[$month][$day][$hour][$item] =
                    round($loss * $load_scale) if $max;
                $PV_usage_loss_hours++; # will be normalized by $items
            } elsif ($max) {
                $PV_usage_loss_by_item[$month][$day][$hour][$item] = 0;
            }
        }
        # $spill_loss += ($net_pv_power - $feed_sum / $items) * $load_scale
        #     if defined $bypass;

        $curb_losses_this_hour *= $load_scale / $items;
        $PV_usage_loss[$month][$day][$hour] += $curb_losses_this_hour;
        $PV_usage_losses += $curb_losses_this_hour;
        $AC_coupling_losses_this_hour *= $load_scale / $items;
        $AC_coupling_losses += $AC_coupling_losses_this_hour;
        # $sum_needed += $needed * $load_scale / $items; # per hour
        # print "$year-".time_string($month, $day, $hour, $minute).
        # "\tPV=".round($net_pv_power)."\tPN=".round($needed)."\tPU=".round($usages).
        # "\t$_\n" if $net_pv_power != 0 && m/^20160214:1010/; # m/^20....02:12/;
        $usages *= $load_scale / $items;
        $PV_used[$month][$day][$hour] += $usages;
        $PV_used_sum += $usages;

        # revert factoring out $items for optimizing the inner loop
        if (defined $capacity && $items != 1) {
            $capacity /= $items;
            $charge /= $items;
            $charge_sum /= $items;
            # $charging_loss /= $items;
            $spill_loss /= $items;
            $PV_used_via_storage /= $items;
            $grid_feed_in /= $items;
        }

        $hour = 0 if ++$hour == 24;
        adjust_day_month();
        ($year, $month, $day) = ($year + 1, 1, 1) if $month > 12;
        last if $test && ($day - 1) * 24 + $hour == TEST_END;
    }
    print "\n" if $test;
    $load_max *= $load_scale;
    $load_sum *= $load_scale;
    $PV_gross_out_sum /= $years;
    $PV_net_losses /= $years;
    $PV_net_loss_hours /= $years;
    $PV_net_out_sum /= $years;
    $PV_net_bright_sum /= $years;
    $PV_used_sum /= $years;
    $PV_usage_losses *= $load_scale / $years;
    $PV_usage_loss_hours /= ($sum_items / YearHours * $years);
    # die "Inconsistent load calculation: sum = $sum vs. needed = $sum_needed"
    #     if round($sum) != round($sum_needed);

    if (defined $capacity) {
        $capacity *= $load_scale_never_0;
        $bypass *= $load_scale_never_0 if defined $bypass;
        $max_feed *= $load_scale_never_0 * $storage_eff_never_0
            if defined $max_feed;
        $charge *= $load_scale / $years;
        $charge_sum *= $load_scale / $years;
        # $charging_loss *= $load_scale / $years;
        $spill_loss *= $load_scale / $years;
        $PV_used_via_storage *= $load_scale / $years;
        $grid_feed_in *= $load_scale / $years;
    } else {
        $grid_feed_in = $PV_net_out_sum - $PV_used_sum;
    }
}

simulate();
die "PV power data all zero" unless $PV_gross_max_tm;

################################################################################
# statistics output

my $nominal_txt      = $en ? "nominal PV power"     : "PV-Nominalleistung";
my $max_gross_txt    = $en ? "max gross PV power"   : "Max. PV-Bruttoleistung";
my $curb_txt         = $en ? "power curb by inverter"
                           : "Leistungsbegrenzung (Drosselung)";
my $system_eff_txt   = $en ? "PV system eff."       : "PV-System-Eff.";
my $own_txt          = $en ? "PV own use"           : "PV-Eigenverbrauch";
my $own_ratio_txt    = $own_txt . ($en ? " ratio"  : "santeil");
my $own_storage_txt  = $en ?"PV own use via storage":"PV-Nutzung über Speicher";
my $load_cover_txt   = $en ? "use coverage ratio"  : "Eigendeckungsanteil";
my $PV_gross_txt     = $en ? "PV gross yield"       : "PV-Bruttoertrag";
my $PV_net_txt       = $en ? "PV net yield"         : "PV-Nettoertrag";
my $PV_net_curb_txt  = $en ? "inverter output ".($curb ? "" : "not ")."curbed"
                  : "WR-Ausgangsleistung ".($curb ? "" : "nicht ")."gedrosselt";
my $PV_loss_txt      = $en ? "PV yield net loss"    : "PV-Netto-Ertragsverlust";
my $consumpt_txt     = $en ? "consumption by household"
                           : "Verbrauch durch Haushalt";
my $load_const_txt   = $en ? "constant load"        : "Konstante Last";
my $load_during_txt  =($en ? "on the first $load_days days a week"
                           : "an den ersten $load_days Tagen der Woche").
                      ($feed_from == 0 && $feed_to == 24) ? "" :
                      ($en ? " from $load_from to $load_to h"
                           : " von $load_from bis $load_to Uhr")
                      if defined $load_const;
my $usage_loss_txt   = $en ? "$own_txt net loss"    :  $own_txt."sverlust";
my $net_de           = $en ? ""                     : " netto";
my $by_curb          = $en ? "by curbing"           : "durch Drosselung";
my $without_curb     = $en ? "without curb"         : "ohne Drosselung";
my $with_curb        = $en ? "with curb"            : "mit Drosselung";
my $opt_with_curb    = ($curb ? " $with_curb" : "");
my $grid_feed_txt    = $en ? "grid feed-in"         : "Netzeinspeisung";
my $each             = $en ? "each"                 : "alle";
my $yearly_txt       = $en ? "yearly"               : "jährlich";
my $capacity_txt     = $en ? "storage capacity"     : "Speicherkapazität";
my $coupled_txt = ($AC_coupled ? "AC" : "DC").($en ? " coupled": "-gekoppelt");
my $optimal_charge   = $en ? "optimal charging strategy (power not consumed)"
                           :"Optimale Ladestrategie (nicht gebrauchte Energie)";
my $bypass_txt       = $en ? "storage bypass"       : "Speicher-Umgehung";
my $spill_txt        = !$bypass_spill ? "" :
                      ($en ?  " and for surplus"    : " und für Überschuss");
my $optimal_discharge= $en ? "optimal discharging strategy (as much as needed)"
                           :"Optimale Entladestrategie (so viel wie gebraucht)";
my $feed_txt        = ($const_feed
                    ? ($en ? "constant "            : "Konstant")
                    : ($en ?  "maximal "            : "Maximal"))
                     .($en ? "feed-in"              : "einspeisung");
my $feed_during_txt  =!$const_feed || ($feed_from == 0 && $feed_to == 24) ? "" :
                       $en ? " from $feed_from to $feed_to h"
                           : " von $feed_from bis $feed_to Uhr";
my $ceff_txt         = $en ? "charging eff."        : "Lade-Eff.";
my $seff_txt         = $en ? "storage eff."         : "Speicher-Eff.";
my $ieff_txt         = $en ?"inverter eff."         : "Wechselrichter-Eff.";
my $stored_txt       = $en ? "buffered energy"      : "Zwischenspeicherung";
my $spill_loss_txt   = $en ? "loss by spill"        : "Verlust durch Überlauf";
my $AC_coupl_loss_txt= $en ? "loss by AC coupling" :"Verlust durch AC-Kopplung";
my $charging_loss_txt= $en ? "charging loss"        : "Ladeverlust";
my $storage_loss_txt = $en ? "storage loss"         : "Speicherverlust";
my $cycles_txt       = $en ? "full cycles per year" : "Vollzyklen pro Jahr";

my $own_usage =
    round($PV_net_out_sum ? 100 * $PV_used_sum / $PV_net_out_sum : 0);
my $load_coverage = round($load_sum ? 100 * $PV_used_sum / $load_sum : 0);
my $yield_daytime =
    percent($PV_net_out_sum ? $PV_net_bright_sum / $PV_net_out_sum : 0);
my $storage_loss;
my $cycles = 0;
if (defined $capacity) { # also for future loss when discharging the rest:
    $charging_loss = ($charge_eff ? $charge_sum * (1 / $charge_eff - 1) : 0);
    $charging_loss *= $inverter_eff unless $AC_coupled;
    $storage_loss = $charge_sum * (1 - $storage_eff);
    $storage_loss *= $inverter_eff unless $AC_coupled;
    $cycles = round($charge_sum / $capacity) if $capacity != 0;
    # future losses:
    $charge *= $storage_eff;
    $AC_coupling_losses += $charge * (1 - $inverter_eff) if $AC_coupled;
    $charge *= $inverter_eff;
}

$load_const = round($load_const * $load_scale) if defined $load_const;
$pvsys_eff = round($pvsys_eff * 100);
$charge_eff *= 100 if $charge_eff;
$storage_eff = round($storage_eff * 100) if $storage_eff;
$inverter_eff *= 100;

sub save_statistics {
    my $file = shift;
    return unless $file;
    my $res_txt = shift;
    my $max     = shift;
    my $hourly  = shift;
    my $daily   = shift;
    my $weekly  = shift;
    my $monthly = shift;
    open(my $OU, '>',$file) or die "Could not open statistics file $file: $!\n";

    my $nominal_sum = $#PV_peaks == 0 ? $PV_peaks[0] : "=$PV_peaks";
    print $OU "$consumpt_txt in kWh, ";
    print $OU "$load_const_txt in W $load_during_txt, " if defined $load_const;
    print $OU "$profile_txt, $pv_data_txt\n";
    print $OU "".round_1000($load_sum).", ";
    print $OU "$load_const, " if defined $load_const;
    print $OU "$load_profile, ".join(", ", @PV_files)."\n";
    print $OU " $nominal_txt in Wp, $max_gross_txt in W, "
        ."$curb_txt in W, $system_eff_txt in %, $ieff_txt in %, "
        ."$own_ratio_txt in %, $load_cover_txt in %\n";
    print $OU "$nominal_sum, ".round($PV_gross_max).", "
        .($curb ? $curb : $en ? "none" : "keine").
        ", $pvsys_eff, $inverter_eff, $own_usage, $load_coverage\n";
    print $OU "$capacity_txt in Wh, "
        .(defined $bypass ? "$bypass_txt in W$spill_txt" : $optimal_charge).", "
        .(defined $max_feed ? "$feed_txt in W$feed_during_txt, "
                            : $optimal_discharge)
        ."$ceff_txt in %, $seff_txt in %, "
        .($AC_coupled ? "$AC_coupl_loss_txt in kWh": $coupled_txt)
        .", $spill_loss_txt in kWh, "
        ."$charging_loss_txt in kWh, $storage_loss_txt in kWh, "
        ."$own_storage_txt in kWh, $stored_txt in kWh, $cycles_txt\n"
        if defined $capacity;
    print $OU "$capacity, ". (defined $bypass ? $bypass : "")
        .", $max_feed, $charge_eff, $storage_eff, "
        .($AC_coupled ? round_1000($AC_coupling_losses) : "").", "
        .round_1000($spill_loss).", "
        .round_1000($charging_loss).", ".round_1000($storage_loss).", "
        .round_1000($PV_used_via_storage).", ".round_1000($charge_sum)
        .", $cycles\n" if defined $capacity;
    print $OU "$yearly_txt, $PV_gross_txt, $PV_net_txt, "
        .($curb?"$PV_loss_txt $by_curb, $usage_loss_txt $net_de $by_curb, ": "")
        ."$own_txt$opt_with_curb, ".
        "$grid_feed_txt, $each in kWh\n";
    print $OU "".($en ? "sums" : "Summen"  ).", ".
        round_1000($PV_gross_out_sum ).", ".
        round_1000($PV_net_out_sum   ).", ".
        ($curb ? round_1000($PV_net_losses  ).", ".
                 round_1000($PV_usage_losses).", " : "").
        round_1000($PV_used_sum      ).", ".
        round_1000($grid_feed_in     )."\n";
    print $OU "$res_txt, $PV_gross_txt, $PV_net_txt, "
        .($curb ? "$PV_net_curb_txt, "        : "")
        .($curb ? "$own_txt $without_curb, " : "")."$own_txt$opt_with_curb, "
        ."$consumpt_txt, "."$each in Wh\n";
    ($month, my $week, my $days, $day, $hour) = (1, 1, 0, 1, 0);
    my ($gross, $PV_loss, $net, $hload, $loss, $used) = (0, 0, 0, 0, 0, 0);
    while ($month <= 12) {
        my $tim;
        if ($weekly) {
            $tim = sprintf("%02d", $week);
        } else {
            $tim = sprintf("%02d", $month);
            $tim = $tim."-".sprintf("%02d", $day ) if $daily || $hourly || $max;
            $tim = $tim." ".sprintf("%02d", $hour) if $hourly || $max;
        }
        my $gross_across = 0;
        for (my $year = 0; $year < $years; $year++) {
            $gross_across += $PV_gross_out[$year][$month][$day][$hour];
        }
        $gross   += round($gross_across / $years);
        $PV_loss += round($PV_net_loss  [$month][$day][$hour] / $years);
        $net     += round($PV_net_out   [$month][$day][$hour] / $years);
        $hload   += round($load_by_hour [$month][$day][$hour] * $load_scale);
        $loss    += round($PV_usage_loss[$month][$day][$hour] / $years);
        $used    += round($PV_used      [$month][$day][$hour] / $years);
        my ($m, $d, $h) = ($month, $day, $hour);
        if (++$hour == 24) {
            $hour = 0;
            $days++;
            $week++ if $days % 7 == 0
                && $days < 365; # count Dec-31 as part of week 52
        }
        adjust_day_month();
        if ($max || $hourly || ($hour == 0 &&
            ($daily || $weekly && $days % 7 == 0 || $monthly && $day == 1))) {
            for (my $i = 0; $i < ($max ? $items : 1); $i++) {
                my $minute = $hourly ? ":00" : "";
                if ($max) {
                    $minute = minute_string($i, $items);
                    $hload = round($load_by_item[$m][$d][$h][$i] * $load_scale);
                    $loss = $PV_usage_loss_by_item[$m][$d][$h][$i];
                    $used = $PV_used_by_item      [$m][$d][$h][$i];
                }
                print $OU "$tim$minute, $gross, ".($net + $PV_loss).", "
                    .($curb ?             "$net, " : "")
                    .($curb ? ($used + $loss).", " : "")."$used ,$hload\n";
            }
            ($gross, $PV_loss, $net, $hload, $loss, $used) = (0, 0, 0, 0, 0, 0);
        }
        last if $test && ($day - 1) * 24 + $hour == TEST_END;
    }
    close $OU;
}

my $max_txt   = $items == 60 ? ($en ? "minute" : "Minute")
                             : ($en ? "point in time" : "Zeitpunkt");
my $hour_txt  = $en ? "hour"   : "Stunde";
my $date_txt  = $en ? "date"   : "Datum";
my $week_txt  = $en ? "week"   : "Woche";
my $month_txt = $en ? "month"  : "Monat";
save_statistics($max    , $max_txt  , 1, 0, 0, 0, 0);
save_statistics($hourly , $hour_txt , 0, 1, 0, 0, 0);
save_statistics($daily  , $date_txt , 0, 0, 1, 0, 0);
save_statistics($weekly , $week_txt , 0, 0, 0, 1, 0);
save_statistics($monthly, $month_txt, 0, 0, 0, 0, 1);

my $at             = $en ? "with"                 : "bei";
my $and            = $en ? "and"                  : "und";
my $due_to         = $en ? "due to"               : "durch";
my $only           = $en ? "only"                 : "nur";
my $during         = $en ? "during"               : "während";
my $by_curb_at     = $en ? "$by_curb at"          : "$by_curb auf";
my $yield          = $en ? "yield portion"        : "Ertragsanteil";
my $daytime        = $en ? "9 AM - 3 PM "         : "9-15 Uhr MEZ";
my $of_yield       = $en ? "of net yield"   : "des Nettoertrags (Nutzungsgrad)";
my $of_consumption = $en ? "of consumption" : "des Verbrauchs (Autarkiegrad)";
# PV-Abregelungsverlust"
if (!$test) {
    $lat = " $lat" unless $lat =~ m/^-?\d\d/;
    $lon = " $lon" unless $lon =~ m/^-?\d\d/;
    print "".($en ? "latitude      " : "Breitengrad   ")."$s13 =   $lat\n";
    print "".($en ? "longitude     " : "Längengrad    ")."$s13 =   $lon\n";
    print "\n";
}
my $nominal_sum = $#PV_peaks == 0 ? "" : " = $PV_peaks Wp";
print "$nominal_txt $en2         =" .W($nominal_power_sum)."p$nominal_sum\n";
print "$max_gross_txt $en4     =".W($PV_gross_max)." $on $PV_gross_max_tm\n";
print "$PV_gross_txt $en1            =".kWh($PV_gross_out_sum)."\n";
print "$PV_net_txt $en2             =" .kWh($PV_net_out_sum).
    " $at $system_eff_txt $pvsys_eff%, $ieff_txt $inverter_eff%\n";
print "$PV_loss_txt $en2 $en2 $en2  ="       .kWh($PV_net_losses).
    " $during ".round($PV_net_loss_hours)." h $by_curb_at $curb W\n" if $curb;
print "$yield $daytime  =   $yield_daytime %\n";

print "\n";
print "$consumpt_txt    =" .kWh($load_sum)."\n";
print "$load_const_txt $en1             =".W($load_const)." $load_during_txt\n"
    if defined $load_const;
if (defined $capacity) {
    print "\n".
        "$capacity_txt $en1          =" .W($capacity)."h, $coupled_txt\n";
    print "$optimal_charge\n" unless defined $bypass;
    print "$bypass_txt $en3          =".W($bypass)
        .($bypass_spill ? "  " : "")."$spill_txt\n" if defined $bypass;
    print "$optimal_discharge\n" unless defined $max_feed;
    print "$feed_txt $en3        ".($const_feed ? "" : " ")
        ."=".W($max_feed)."$feed_during_txt\n" if defined $max_feed;
    print "$AC_coupl_loss_txt $en3$en3  =" .kWh($AC_coupling_losses)."\n"
        if $AC_coupled;
    print "$spill_loss_txt $en3 $en3 $en3   =".kWh($spill_loss)."\n";
    print "$charging_loss_txt $de2              =".kWh($charging_loss)
        ." $due_to $ceff_txt $charge_eff%\n";
    print "$storage_loss_txt $en3            =".kWh($storage_loss)
        ." $due_to $seff_txt $storage_eff%\n";
    print "$own_storage_txt $en2   =".kWh($PV_used_via_storage)."\n";
    print "$stored_txt $en2$en2        =" .kWh($charge_sum)." ($at $system_eff_txt $and $ceff_txt)\n";
    # Vollzyklen, Kapazitätsdurchgänge pro Jahr Kapazitätsdurchsatz:
    printf "$cycles_txt $de1       =  %3d\n", $cycles;

    # $psys_eff   /= 100;
    # $charge_eff /= 100;
    $storage_eff  /= 100;
    $inverter_eff /= 100;
    my $grid_feed_in_alt = $PV_net_out_sum - $PV_used_sum - $AC_coupling_losses
        - $spill_loss - $charging_loss - $storage_loss - $charge;
    my $discrepancy = $grid_feed_in - $grid_feed_in_alt;
    die "Internal error: grid feed-in calculation discrepancy $discrepancy: ".
        "grid feed-in $grid_feed_in vs. ".int($grid_feed_in_alt + .5)." =\n".
        "PV net sum $PV_net_out_sum - PV used $PV_used_sum".
        " - AC coupling loss $AC_coupling_losses - loss by spill $spill_loss".
        " - charging loss $charging_loss - storage loss $storage_loss".
        " - charge $charge" if abs($discrepancy) > 0.001; # 1 mWh
    print "\n";
}

print "$own_txt $en4 $en3         =" .kWh($PV_used_sum)."$opt_with_curb\n";
print "$usage_loss_txt $en3$en3  =" .kWh($PV_usage_losses)." $net_de $during "
    .round($PV_usage_loss_hours)." h $by_curb_at $curb W\n" if $curb;
print "$grid_feed_txt $en3            =" .kWh($grid_feed_in)."\n";
print "$own_ratio_txt $en4 $en4  =  ".sprintf("%3d", $own_usage)." % $of_yield\n";
my $load_coverage_str = sprintf("%3d", $load_coverage);
print "$load_cover_txt $en1        =  $load_coverage_str % $of_consumption\n";
