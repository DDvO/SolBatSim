#!/usr/bin/perl
################################################################################
# Eigenverbrauchs-Simulation mit stündlichen PV-Daten und einem mindestens
# stündlichen Lastprofil, optional mit Ausgangs-Drosselung des Wechselrichters
# und optional mit Stromspeicher (Batterie o.ä.)
#
# Nutzung: Solar.pl <Lastprofil-Datei> [<Jahresverbrauch in kWh>]
#            (<PV-Daten-Datei> [<Nennleistung in Wp>])+
#            [-load <konstante Last in W> [<Zahl der Tage pro Woche, sonst 5>:
#                   <von Uhrzeit, sonst 8 Uhr>..<bis Uhrzeit, sonst 16 Uhr>]]
#            [-peff <PV-System-Wirkungsgrad in %, ansonsten von PV-Daten-Datei>]
#            [-capacity <Speicherkapazität Wh, ansonsten 0 (kein Batterie)>]
#            [-ac] (bedeutet AC-gekoppelte Ladung, nach Wechselrichtung)
#            [-pass [spill] <Speicher-Umgehung in W, opt. auch bei Überlauf>]
#            [-feed [max] <begrenzte/konstante Einspeisung aus Speicher in W>]
#            [-ceff <Lade-Wirkungsgrad in %, ansonsten 94]
#            [-seff <Speicher-Wirkungsgrad in %, ansonsten 95]
#            [-ieff <Wechselrichter-Wirkungsgrad in %, ansonsten 94]
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
# Simulation of actual own usage of photovoltaic power output according to load
# profiles with a resolution of at least one hour, typically per minute.
# Optionally takes into account power output cropping by solar inverter.
# Optionally with energy storage (using a battery or the like).
#
# Usage: Solar.pl <load profile file> [<consumption per year in kWh>]
#          (<PV data file> [<nominal power in Wp>])+
#          [-load <constant load in W> [<count of days per week, default 5>:
#                 <from hour, default 8 o'clock>..<to hour, default 16>]]
#          [-peff <PV system efficiency in %, default from PV data file(s)>]
#          [-capacity <storage capacity in Wh, default 0 (no battery)>]
#          [-ac] (means AC-coupled charging, after inverter)
#          [-pass [spill] <storage bypass in W, optionally also on surplus>]
#          [-feed [max] <limited/constant feed-in from storage in W>]
#          [-ceff <charging efficiency in %, default 94]
#          [-seff <storage efficiency in %, default 95]
#          [-ieff <inverter efficiency in %, default 94]
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

die "Missing load profile file" if $#ARGV < 0;
my $profile             = shift @ARGV; # load profile file
my $simulated_load      = shift @ARGV # default by solar data file
    if $#ARGV >= 0 && $ARGV[0] =~ m/^[\d\.]+$/;
my $load_const;     # constant load in W, during certain times:
my $load_days = 5;  # count of days per week with constant load
my $load_from = 8;  # hour of constant load begin
my $load_to  = 16;  # hour of constant load en

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
my $AC_coupled;     # by default, charging is DC-coupled (w/o inverter)
my $charge_eff;     # charge efficiency; default see below
my $storage_eff;    # storage efficiency; default see below
# storage efficiency and discharge efficiency could have been combined
my $nominal_power_sum = 0;

while ($#ARGV >= 0 && $ARGV[0] =~ m/^\s*[^-]/) {
    push @PV_files, shift @ARGV; # PV data file
    push @PV_peaks, $#ARGV >= 0 && $ARGV[0] =~ m/^[\d\.]+$/ ? shift @ARGV : 0;
}
die "Missing PV data file" if $#PV_files < 0;

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

$inverter_eff  = .94;
my ($en, $tmy, $curb, $max, $hourly, $daily, $weekly,  $monthly);
while ($#ARGV >= 0) {
    if      ($ARGV[0] eq "-en"      ) { $en           =  no_arg();
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
    } elsif ($ARGV[0] eq "-ac"      ) { $AC_coupled   =  no_arg();
    } elsif ($ARGV[0] eq "-pass"    ) { $bypass_spill = 1 if $#ARGV >= 1 &&
                                            $ARGV[1] eq "spill" && shift @ARGV;
                                        $bypass       = num_arg();
    } elsif ($ARGV[0] eq "-feed"    ) { $const_feed   = 0 if $#ARGV >= 1 &&
                                            $ARGV[1] eq "max" && shift @ARGV;
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
if (defined $load_const) {
    die "days count for -load option must be in range 0..7"
        if  $load_days > 7;
    die "begin hour for -load option must be in range 0..24"
        if  $load_from > 24;
    die "end hour for -load option must be in range 0..24"
        if  $load_to > 24;
}
if (defined $capacity) {
    $charge_eff  = .94 unless defined $charge_eff;
    $storage_eff = .95 unless defined $storage_eff;
} else {
    die   "-ac option requires -capacity option" if defined $AC_coupled;
    die "-pass option requires -capacity option" if defined $bypass;
    die "-feed option requires -capacity option" if defined $max_feed;
    die "-ceff option requires -capacity option" if defined $charge_eff;
    die "-seff option requires -capacity option" if defined $storage_eff;
}
my $inverter_eff_never_0 = $inverter_eff != 0 ? $inverter_eff : 1;
my   $charge_eff_never_0 = $charge_eff   != 0 ?   $charge_eff : 1
    if defined $charge_eff;

# deliberately not using any extra packages like Math
sub min { return $_[0] < $_[1] ? $_[0] : $_[1]; }
sub max { return $_[0] > $_[1] ? $_[0] : $_[1]; }
sub round { return int(.5 + shift); }
sub check_consistency {
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
sub index_string {
    return date_string(shift, shift, shift).sprintf("h, index %4d", shift);
}

sub round_1000 { return round(shift() / 1000); }
sub kWh     { return sprintf("%5d kWh", round_1000(shift()  )); }
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

use constant YearHours => 24 * 365;

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
    open(my $IN, '<', $file) or die "Could not open profile file $file: $!\n";

    my $warned_just_before = 0;
    my $weekday = 4; # HTW load profiles start on Friday (of 2010), Monday == 0
    while (my $line = <$IN>) {
        chomp $line;
        next if $line =~ m/^\s*#/; # skip comment line
        my @sources = split "," , $line;
        my $n = $#sources;
        shift @sources; # ignore date & time info
        if ($items > 0 && $items != $n) {
            print "Unequal number of items per line: $n vs. $items in $file\n";
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
                    print "Load on YYYY-".index_string($month, $day, $hour, $item).
                        " = ".sprintf("%4d", $load)."\n"
                        unless $warned_just_before;
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
    }
    close $IN;
    $month--;
    check_consistency($month, 12, "months", $file);
    check_consistency($hour_per_year, YearHours, "hours", $file);
}

get_profile($profile);

my $profile_txt = $en ? "load profile file" : "Lastprofil-Datei";
my $pv_data_txt = $en ? "PV data file(s)"   : "PV-Daten-Datei(en)";
my $p_txt = $en ? "load data points per hour  " : "Last-Datenpunkte pro Stunde";
my $W_txt = $en ? "load per weekday (Mon-Sun) " : "Last pro Wochentag (Mo-So) ";
my $n_txt = $en ? "load portion 12 AM -  6 AM " : "Lastanteil  0 -  6 Uhr MEZ ";
my $m_txt = $en ? "load portion  6 AM -  9 PM " : "Lastanteil  6 -  9 Uhr MEZ ";
my $s_txt = $en ? "load portion  9 AM -  3 PM " : "Lastanteil  9 - 15 Uhr MEZ ";
my $a_txt = $en ? "load portion  3 AM -  6 PM " : "Lastanteil 15 - 18 Uhr MEZ ";
my $e_txt = $en ? "load portion  6 PM - 12 PM " : "Lastanteil 18 - 24 Uhr MEZ ";
my $w_txt = $en ? "load portion October-March " : "Lastanteil Oktober - März  ";
my $t_txt = $en ? "total load acc. to profile " : "Verbrauch gemäß Lastprofil ";
my $b_txt = $en ? "basic load                 " : "Grundlast                  ";
my $M_txt = $en ? "maximal load               " : "Maximallast                ";
my $on    = $en ? "on" : "am";
my $en1 = $en ? " "   : "";
my $en2 = $en ? "  "  : "";
my $en3 = $en ? "   " : "";
my $de1 = $en ? ""    : " ";
my $de2 = $en ? ""    : "  ";
my $de3 = $en ? ""    : "   ";
my $s10   = "          "; 
print "$profile_txt$de1$s10 : $profile\n";
print "$p_txt = ".sprintf("%4d", $sum_items / $hour_per_year)."\n";
if ($load_sum != 0) {
    print "$W_txt =   ";
    for (my $weekday = 0; $weekday < 7; $weekday++) {
        print "".percent($load_by_weekday[$weekday] / $load_sum)." %";
        print ", " unless $weekday == 6;
    }
    print "\n";
    print "$n_txt =   ".percent($night_sum   / $load_sum)." %\n";
    print "$m_txt =   ".percent($morning_sum / $load_sum)." %\n";
    print "$s_txt =   ".percent($bright_sum  / $load_sum)." %\n";
    print "$a_txt =   ".percent($earleve_sum / $load_sum)." %\n";
    print "$e_txt =   ".percent($evening_sum / $load_sum)." %\n";
    print "$w_txt =   ".percent($winter_sum  / $load_sum)." %\n";
}
print "$t_txt =".kWh($load_sum)."\n";
print "$b_txt =".W($night_sum / 365 / (NIGHT_END - NIGHT_START))."\n";
print "$M_txt =".W($load_max)." $on $load_max_time\n";
print "\n";

################################################################################
# read PV production data

my @PV_gross_out;
my ($start_year, $years);
sub get_power {
    my ($file, $nominal_power) = (shift, shift);
    open(my $IN, '<', $file) or die "Could not open PV data file $file: $!\n";
    my $s13 = "             ";
    print "".($en ? "PV data file  " : "PV-Daten-Datei")."$s13 : $file";

    # my $sum_needed = 0;
    my ($slope, $azimuth);
    my $current_years = 0;
    my $nominal_power_deflt; # PVGIS default: 1 kWp
    my $pvsys_eff_deflt; # PVGIS default: 0.86
    my $power_rate;
    my $power_provided = 0; # PVGIS default: none
    my ($months, $hours) = (0, 0);
    while (<$IN>) {
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
                if $pvsys_eff && $pvsys_eff != $pvsys_eff_deflt;
            $pvsys_eff = $pvsys_eff_deflt unless defined $pvsys_eff;
            if ($pvsys_eff > 1) {
                print "\n";
                die "unreasonable PV system efficiency ".round($pvsys_eff * 100)
                    ."% - have -peff and -ieff been used properly?";
            }
        }
        $power_provided = 1 if m/^time,P,/;

        next unless m/^20\d\d\d\d\d\d:\d\d\d\d,/;
        die "Missing latitude in $file"  unless $lat;
        die "Missing longitude in $file" unless $lon;
        die "Missing slope in $file"     unless $slope;
        die "Missing azimuth in $file"   unless $azimuth;
        die "Missing nominal power in $file"     unless $nominal_power_deflt;
        die "Missing system efficiency in $file" unless $pvsys_eff_deflt;
        die "Missing PV power output data in $file" unless $power_provided;
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
            unless m/^(\d\d\d\d)(\d\d)(\d\d):(\d\d)(\d\d),([\d\.]+)/;
        my ($year, $month, $day, $hour, $minute_unused, $power) =
            ($tmy ? $start_year : $1, $2, $3, $4, $5, $6 * $nominal_power
             / $nominal_power_deflt / ($pvsys_eff_deflt * $inverter_eff));
        $PV_gross_out[$year - $start_year][$month][$day][$hour] += $power;
        $hours++;
    }
    close $IN;
    check_consistency($months, 12 * $current_years, "months", $file);
    check_consistency($hours, YearHours * $current_years, "hours", $file);
    check_consistency($years, $current_years, "years", $file) if $years;
    $years = $current_years;

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

my $load_scale = defined $simulated_load && $load_sum != 0
    ? 1000 * $simulated_load / $load_sum : 1;
my $load_scale_never_0 = $load_scale != 0 ? $load_scale : 1;
my @PV_net_loss;
my $PV_net_loss_sum = 0;
my $PV_net_loss_hours = 0;
my @PV_usage_loss_by_item;
my @PV_usage_loss;
my $PV_usage_loss_sum = 0;
my $PV_usage_loss_hours = 0;
my @PV_used_by_item;
my @PV_used;
my $PV_used_sum = 0;
my $PV_used_via_storage = 0;
my $charge        = 0 if defined $capacity;
my $charge_sum    = 0 if defined $capacity;
my $charging_loss = 0 if defined $capacity;
my $surplus_loss  = 0 if defined $capacity;
my $grid_feed_in  = 0;

sub simulate()
{
    my $year = 0;
    ($month, $day, $hour) = (1, 1, 0);
    my $minute = 0; # currently fixed

    # factor out $load_scale for optimizing the inner loop
    $capacity /= $load_scale_never_0 if $capacity;
    $bypass /= $load_scale_never_0 if defined $bypass;
    $max_feed /= $load_scale_never_0 if defined $max_feed;

    while ($year < $years) {
        my $power = $PV_gross_out[$year][$month][$day][$hour];
        die "No power data at ".($start_year + $year)."-".
            time_string($month, $day, $hour, $minute) unless defined($power);
        $PV_gross_out_sum += $power;
        if ($power > $PV_gross_max) {
            $PV_gross_max = $power;
            $PV_gross_max_tm = ($tmy ? "TMY" : $start_year + $year)."-".
                time_string($month, $day, $hour, $minute);
        }
        my $net_pv_power = $power * $pvsys_eff * $inverter_eff;

        $PV_net_loss[$month][$day][$hour] = 0;
        my $PV_loss = 0;
        if ($curb && $net_pv_power > $curb) {
            $PV_loss = $net_pv_power - $curb;
            $PV_net_loss[$month][$day][$hour] += $PV_loss;
            $PV_net_loss_sum += $PV_loss;
            $PV_net_loss_hours++;
            $net_pv_power = $curb;
            # print "$year-".time_string($month, $day, $hour, $minute).
            #"\tPV=".round($net_pv_power)."\tcurb=".round($curb).
            #"\tloss=".round($PV_net_loss_sum)."\t$_\n";
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
        my $losses = 0;
        $items = $items_by_hour[$month][$day][$hour];

        # factor out $items for optimizing the inner loop
        if (defined $capacity && $items != 1) {
            $capacity *= $items;
            $charge *= $items;
            $charge_sum *= $items;
            # $charging_loss *= $items;
            $surplus_loss *= $items;
            $PV_used_via_storage *= $items;
            $grid_feed_in *= $items;
        }

        # my $feed_sum = 0 if defined $max_feed;
        for (my $item = 0; $item < $items; $item++) {
            my $loss = 0;
            my $power_needed = $load_by_item[$month][$day][$hour][$item];
            die "load_by_item[$month][$day][$hour][$item] is undefined"
                unless defined $power_needed;
            # $needed += $power_needed;
            # load will be reduced by constant $bypass or $bypass_spill

            # $pv_used locally accumulates PV own consumption
            # $grid_feed_in accumulates feed to grid

            # feed by constant bypass or just as much as used (optimal charge)
            my $pv_used = defined $bypass ? $bypass : $power_needed; # preliminary
            my $extra_power = $net_pv_power - $pv_used;
            if ($extra_power < 0) {
                $pv_used = $net_pv_power;
                # == min($net_pv_power, defined $bypass ? $bypass : $power_needed);
            }
            if (defined $bypass) {
                my $unused_bypass = $pv_used - $power_needed;
                if ($unused_bypass > 0) {
                    $pv_used -= $unused_bypass;
                    $grid_feed_in += $unused_bypass;
                }
            }
            $power_needed -= $pv_used;

            if (defined $capacity) { # storage available
                my $capacity_to_fill = $capacity - $charge;
                my $storage_full = $capacity_to_fill <= 0;
                if ($extra_power > 0) {
                    # when charging is DC-coupled, no loss through inverter
                    $extra_power /= $inverter_eff_never_0 unless $AC_coupled;
                    my $max_to_charge = $storage_full ? 0 :
                        $capacity_to_fill / $charge_eff_never_0;
                    # optimal charge: exactly as much as unused and fits in
                    my $charge_input = $extra_power;
                        # will become min($extra_power, $max_to_charge);
                    my $surplus = $extra_power - $max_to_charge;
                    if ($surplus > 0) {
                        $charge_input = $max_to_charge;
                        $surplus *= $inverter_eff;
                        if (!defined $bypass) {
                            $grid_feed_in += $surplus; # on optimal charge
                        } elsif ($bypass_spill) {
                            my $remaining_load = $power_needed - $bypass;
                            if ($remaining_load > 0) {
                                if ($remaining_load > $surplus) {
                                    $pv_used += $surplus;
                                    $power_needed -= $surplus;
                                } else {
                                    $pv_used += $remaining_load;
                                    $power_needed -= $remaining_load;
                                    $grid_feed_in += $surplus - $remaining_load;
                                }
                            } else {
                                $grid_feed_in += $surplus;
                            }
                        } else { # defined $bypass && !$bypass_spill
                            $surplus_loss += $surplus;
                        }
                    }

                    # This leads to adding reduced charging due to curb
                    # to the usage losses, which is not entirely correct
                    # in case of constant feed (non-optimal discharge).
                    $loss += min($PV_loss, $max_to_charge * $charge_eff
                                 * $storage_eff * $inverter_eff)
                        if $AC_coupled && $PV_loss != 0;

                    my $charge_delta = $charge_input * $charge_eff;
                    # print "#> $net_pv_power $power_needed ".
                    # "$charge += $charge_delta\n"
                    # if $net_pv_power * $load_scale > 595;
                    $charge += $charge_delta;
                    $charge_sum += $charge_delta;
                    # $charging_loss += $charge_input - $charge_delta;
                }
                if ($charge > 0) { # storage not empty
                    my $discharge = 0;
                    if (!defined $max_feed && $power_needed > 0) {
                        # optimal discharge: exactly as much as currently needed
                        $discharge = min($power_needed, $charge);
                    }
                    if (defined $max_feed) {
                        $discharge = $max_feed;
                        $discharge = $power_needed # optimal but limited feed
                            if !$const_feed && $power_needed < $max_feed;
                        $discharge = min($discharge, $charge);
                        # print "## $net_pv_power $power_needed ".
                        # "$charge -= $discharge max_feed $max_feed\n"
                        # if $net_pv_power * $load_scale > 595;
                    }
                    if ($discharge != 0) {
                        $charge -= $discharge;
                        $discharge *= $storage_eff * $inverter_eff;
                        if (defined $max_feed && $const_feed) {
                            # $feed_sum += $discharge;
                            my $discharge_feed_in = $discharge - $power_needed;
                            if ($discharge_feed_in > 0) {
                                $grid_feed_in += $discharge_feed_in;
                                $discharge -= $discharge_feed_in;
                            }
                        }
                        $pv_used += $discharge;
                        $PV_used_via_storage += $discharge;
                    }
                }
            }
            $usages += $pv_used;
            $PV_used_by_item[$month][$day][$hour][$item] =
                round($pv_used * $load_scale) if $max;

            if ($PV_loss != 0 && $power_needed > 0) {
                $loss += min($PV_loss, $power_needed);
            }
            if ($loss != 0) {
                $losses += $loss;
                $PV_usage_loss_by_item[$month][$day][$hour][$item] =
                    round($loss * $load_scale) if $max;
                $PV_usage_loss_hours++; # will be normalized by $items
            } elsif ($max) {
                $PV_usage_loss_by_item[$month][$day][$hour][$item] = 0;
            }
        }
        # $surplus_loss += ($net_pv_power - $feed_sum / $items) * $load_scale
        #     if defined $bypass;

        $losses *= $load_scale / $items;
        $PV_usage_loss[$month][$day][$hour] += $losses;
        $PV_usage_loss_sum += $losses;
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
            $surplus_loss /= $items;
            $PV_used_via_storage /= $items;
            $grid_feed_in /= $items;
        }

        $hour = 0 if ++$hour == 24;
        adjust_day_month();
        ($year, $month, $day) = ($year + 1, 1, 1) if $month > 12;
    }
    $load_max *= $load_scale;
    $load_sum *= $load_scale;
    $PV_gross_out_sum /= $years;
    $PV_net_loss_sum /= $years;
    $PV_net_loss_hours /= $years;
    $PV_net_out_sum /= $years;
    $PV_net_bright_sum /= $years;
    $PV_used_sum /= $years;
    $PV_usage_loss_sum *= $load_scale / $years;
    $PV_usage_loss_hours /= ($sum_items / YearHours * $years);
    # die "Inconsistent load calculation: sum = $sum vs. needed = $sum_needed"
    #     if round($sum) != round($sum_needed);

    if (defined $capacity) {
        $charge *= $load_scale / $years;
        $charge_sum *= $load_scale / $years;
        # $charging_loss *= $load_scale / $years;
        $surplus_loss *= $load_scale / $years;
        $capacity *= $load_scale_never_0 ;
        $max_feed *= $load_scale_never_0 if defined $max_feed;
        $PV_used_via_storage *= $load_scale / $years;
        $grid_feed_in *= $load_scale / $years;
    } else {
        $grid_feed_in = $PV_net_out_sum - $PV_used_sum;
    }
}

simulate();

################################################################################
# statistics output

my $nominal_txt      = $en ? "nominal PV power"     : "PV-Nominalleistung";
my $max_gross_txt    = $en ? "max gross PV power"   : "Bruttoleistung max.";
my $curb_txt         = $en ? "power curb by inverter": "Leistungsbegrenzung";
my $system_eff_txt   = $en ? "PV system eff."       : "PV-System-Eff.";
my $own_txt          = $en ? "own consumption "     : "Eigenverbrauch";
my $own_ratio_txt    = $own_txt . ($en ? "ratio"    : "santeil");
my $via_storage_txt  = $en ? "via storage"          : "über Speicher";
my $own_storage_txt  = "$own_txt$de1$via_storage_txt";
my $load_cover_txt   = $en ? "load coverage ratio"  : "Eigendeckungsanteil";
my $PV_gross_txt     = $en ? "PV gross yield"       : "PV-Bruttoertrag";
my $PV_net_txt       = $en ? "PV net yield"         : "PV-Nettoertrag";
my $PV_net_curb_txt  = $en ? "inverter output curbed"
                           : "WR-Ausgangsleistung gedrosselt";
my $PV_loss_txt      = $en ? "PV yield lost"        : "PV-Ertragsverlust";
my $load_txt         = $en ? "load by household"    : "Last durch Haushalt";
my $load_const_txt   = $en ? "constant load"        : "Konstante Last";
my $load_during_txt  =($en ? "on the first $load_days days a week from ".
                       "$load_from to $load_to h"  : "an den ersten $load_days".
                       " Tagen der Woche von $load_from bis $load_to Uhr")
    if defined $load_const;
my $usage_loss_txt   = $en ? "PV $own_txt"."loss"   : $own_txt."sverlust";
my $by_curb          = $en ? "by curbing"           : "durch Drosselung";
my $use_wo_curb_txt  = $en ? "$own_txt"."without curb": "$own_txt ohne Drossel";
my $use_w_curb_txt   = $en ? "own consumption".($curb ? " with curb"   : "")
                           : "Eigenverbrauch" .($curb ? " mit Drossel" : "");
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
my $ceff_txt         = $en ? "charging eff."        : "Lade-Eff.";
my $seff_txt         = $en ? "storage eff."         : "Speicher-Eff.";
my $ieff_txt         = $en ?"inverter eff."         : "Wechselrichter-Eff.";
my $stored_txt       = $en ? "buffered energy"      : "Zwischenspeicherung";
my $storage_loss_txt = $en ? "charging and storage losses"
                           : "Lade- und Speicherverluste";
my $surplus_loss_txt = $en ? "loss by surplus"      :"Verlust durch Überschuss";
my $cycles_txt       = $en ? "full cycles per year" : "Vollzyklen pro Jahr";

my $own_usage     =
    percent($PV_net_out_sum ? $PV_used_sum / $PV_net_out_sum : 0);
my $load_coverage =
    percent($load_sum ? $PV_used_sum / $load_sum : 0);
my $yield_daytime =
    percent($PV_net_out_sum ? $PV_net_bright_sum / $PV_net_out_sum : 0);
my $storage_loss;
my $cycles = 0;
if (defined $capacity) {
    $charging_loss = $charge_eff ? $charge_sum * (1 / $charge_eff - 1) : 0;
    $storage_loss = $charging_loss + $charge_sum * (1 - $storage_eff);
    $cycles = round($charge_sum / $capacity) if $capacity != 0;
}

$load_const = round($load_const * $load_scale) if defined $load_const;
$pvsys_eff = round($pvsys_eff * 100);
$charge_eff   *= 100 if $charge_eff;
$storage_eff  *= 100 if $storage_eff;
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
    print $OU "$load_txt in kWh, ";
    print $OU "$load_const_txt in W $load_during_txt, " if defined $load_const;
    print $OU "$profile_txt, $pv_data_txt\n";
    print $OU "".round_1000($load_sum).", ";
    print $OU "$load_const, " if defined $load_const;
    print $OU "$profile, ".join(", ", @PV_files)."\n";
    print $OU " $nominal_txt in Wp, $max_gross_txt in W, "
        ."$curb_txt in W, $system_eff_txt in %, $ieff_txt in %, "
        ."$own_ratio_txt in %, $load_cover_txt in %\n";
    print $OU "$nominal_sum, ".round($PV_gross_max).", "
        .($curb ? $curb : "none").
        ", $pvsys_eff, $inverter_eff, $own_usage, $load_coverage\n";
    print $OU "$capacity_txt in Wh $coupled_txt, "
        .(defined $bypass ? "$bypass_txt in W$spill_txt, " : $optimal_charge)
        .(defined $max_feed ? "$feed_txt in W, " : $optimal_discharge)
        ."$ceff_txt in %, $seff_txt in %, "
        ."$surplus_loss_txt in kWh, $storage_loss_txt in kWh, "
        ."$own_storage_txt in kWh, $stored_txt in kWh, $cycles_txt\n"
        if defined $capacity;
    print $OU "$capacity, $bypass, $max_feed, $charge_eff, $storage_eff, ".
        round_1000($surplus_loss).", ".round_1000($storage_loss).", ".
        round_1000($PV_used_via_storage).", ".
        round_1000($charge_sum).", $cycles\n"
        if defined $capacity;
    print $OU "$yearly_txt, $PV_gross_txt, $PV_net_txt, $PV_loss_txt $by_curb, ".
        "$usage_loss_txt $by_curb, $use_w_curb_txt, ".
        "$grid_feed_txt, $each in kWh\n";
    print $OU "".($en ? "sums" : "Summen"  ).", ".
        round_1000($PV_gross_out_sum ).", ".
        round_1000($PV_net_out_sum   ).", ".
        round_1000($PV_net_loss_sum  ).", ".
        round_1000($PV_usage_loss_sum).", ".
        round_1000($PV_used_sum      ).", ".
        round_1000($grid_feed_in     )."\n";
    print $OU "$res_txt, $PV_gross_txt, $PV_net_txt, $PV_net_curb_txt, "
        ."$load_txt, $use_wo_curb_txt, $use_w_curb_txt, $each in Wh\n";
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
                    $minute  = sprintf(":%02d", int( 60 * $i / $items));
                    $minute .= sprintf(":%02d", int((60 * $i % $items)
                                                    / $items * 60))
                        unless (60 % $items == 0);
                    $hload = round($load_by_item[$m][$d][$h][$i] * $load_scale);
                    $loss = $PV_usage_loss_by_item[$m][$d][$h][$i];
                    $used = $PV_used_by_item      [$m][$d][$h][$i];
                }
                print $OU "$tim$minute, $gross, ".($net + $PV_loss).
                      ", $net, $hload, ".($used + $loss).", $used\n";
            }
            ($gross, $PV_loss, $net, $hload, $loss, $used) = (0, 0, 0, 0, 0, 0);
        }
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

my $s13 = "             ";
my $at             = $en ? "with"                 : "bei";
my $and            = $en ? "and"                  : "und";
my $due_to         = $en ? "due to"               : "durch";
my $only           = $en ? "only"                 : "nur";
my $during         = $en ? "during"               : "während";
my $by_curb_at     = $en ? "$by_curb at"          : "$by_curb auf";
my $yield          = $en ? "yield portion"        : "Ertragsanteil";
my $daytime        = $en ? "9 AM - 3 PM "         : "9-15 Uhr MEZ";
my $of_yield       = $en ? "of net yield"         : "des Nettoertrags (Nutzungsgrad)";
my $of_consumption = $en ? "of consumption"       : "des Verbrauchs (Autarkiegrad)";
# PV-Abregelungsverlust"
   $use_w_curb_txt  = "$use_w_curb_txt $de1          " unless $curb;
my $nominal_sum = $#PV_peaks == 0 ? "" : " = $PV_peaks Wp";
$lat = " $lat" unless $lat =~ m/^-?\d\d/;
$lon = " $lon" unless $lon =~ m/^-?\d\d/;
print "".($en ? "latitude      " : "Breitengrad   ")."$s13 =   $lat\n";
print "".($en ? "longitude     " : "Längengrad    ")."$s13 =   $lon\n";
print "\n";
print "$nominal_txt $en2         =" .W($nominal_power_sum)."p$nominal_sum\n";
print "$max_gross_txt $en1        =".W($PV_gross_max)." $on $PV_gross_max_tm\n";
print "$PV_gross_txt $en1            =".kWh($PV_gross_out_sum)."\n";
print "$PV_net_txt $en2             =" .kWh($PV_net_out_sum).
    " $at $system_eff_txt $pvsys_eff%, $ieff_txt $inverter_eff%\n";
print "$PV_loss_txt $en2 $en2         ="       .kWh($PV_net_loss_sum).
    " $during ".round($PV_net_loss_hours)." h $by_curb_at $curb W\n" if $curb;
print "$yield $daytime  =   $yield_daytime %\n";

print "\n";
print "$load_txt $en2        =" .kWh($load_sum)."\n";
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
        ."=".W($max_feed)."\n" if defined $max_feed;
    print "$surplus_loss_txt $en3 $en3 $en3 =".kWh($surplus_loss)."\n"
        if defined $bypass;
    print "$storage_loss_txt$de1 =".kWh($storage_loss)." $due_to ".
        "$ceff_txt $charge_eff%, $seff_txt $storage_eff%\n";
    print "$own_storage_txt$en1=".kWh($PV_used_via_storage)."\n";
    print "$stored_txt $en2$en2        =" .kWh($charge_sum)." ($at $system_eff_txt $and $ceff_txt)\n";
    # Vollzyklen, Kapazitätsdurchgänge pro Jahr Kapazitätsdurchsatz:
    printf "$cycles_txt $de1       =  %3d\n", $cycles;

    # $psys_eff   /= 100;
    # $charge_eff /= 100;
    $storage_eff  /= 100;
    $inverter_eff /= 100;
    $charge *= $storage_eff;
    my $grid_feed_in_alt = $PV_net_out_sum - $PV_used_sum
        - $surplus_loss - ($storage_loss + $charge) * $inverter_eff;
    #print "$grid_feed_in_alt = $PV_net_out_sum - $PV_used_sum - $surplus_loss"
    #    ." - ".($storage_loss * $inverter_eff)
    #    ." - ".($charge * $inverter_eff)."\n";
    my $discrepancy = abs($grid_feed_in_alt - $grid_feed_in);
    die "grid feed-in calc discrepancy: $grid_feed_in_alt vs. $grid_feed_in"
        if $discrepancy > 0.001; # 1 mWh
    print "\n";
}

print "$use_w_curb_txt  ".($curb ? $en1 : "")."=" .kWh($PV_used_sum)."\n";
print "$usage_loss_txt $de1    =" .kWh($PV_usage_loss_sum)." $during "
    .round($PV_usage_loss_hours)." h $by_curb_at $curb W\n" if $curb;
print "$grid_feed_txt $en3            =" .kWh($grid_feed_in)."\n";
print "$own_ratio_txt       =   $own_usage % $of_yield\n";
print "$load_cover_txt         =   $load_coverage % $of_consumption\n";


# ./Solar.pl Lastprofil_4673_kWh.csv 3000 Timeseries_48.215_11.727_SA2_1kWp_crystSi_0_38deg_0deg_2005_2020.csv 600 -peff 100 -ieff 100 -tmy -capacity 100 -ceff 100 -seff 100 -pass 0 -feed 600
# storage must have no effect,
# if large enough (>=7 Wh) or with spill for any capacity
# is the case also with feed 500 apparently because
# small delay in feeding is negligible if capacity is >= 290

# ./Solar.pl Lastprofil_4673_kWh.csv 3000 Timeseries_48.215_11.727_SA2_1kWp_crystSi_0_38deg_0deg_2005_2020.csv 600 -peff 92 -tmy -capacity 1 -ceff 100 -seff 100 -pass spill 0 -feed max 600
# apparently even tiny capacity helps a lot with spill enabled and adaptive feed
