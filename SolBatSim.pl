#!/usr/bin/perl
################################################################################
# Eigenverbrauchs-Simulation mit stündlichen PV-Daten und einem mindestens
# stündlichen Lastprofil, optional mit Ausgangs-Limitierung des Wechselrichters
# und optional mit Stromspeicher (Batterie o.ä.)
#
# Nutzung: Solar.pl <Lastprofil-Datei> [<Jahresverbrauch in kWh>]
#            (<PV-Daten-Datei> [<Nennleistung in Wp>])+
#            [-eff <System-Wirkungsgrad in %, ansonsten von PV-Daten-Datei(en)>]
#            [-capacity <Speicherkapazität Wh, ansonsten 0 (kein Batterie)>
#            [-ceff <Lade-Wirkungsgrad in %, ansonsten 94]
#            [-seff <Speicher-Wirkungsgrad in %, ansonsten 95]
#            [-deff <Entlade-Wirkungsgrad in %, ansonsten 94]
#            [-en] [-tmy] [-lim <Wechselrichter-Ausgangsleistungs-Limit in W>]
#            [-hour <Datei>] [-day <Datei>] [-week <Datei>] [-month <Datei>]
# Mit "-en" erfolgen die Textausgaben auf Englisch.
# Wenn PV-Daten für mehrere Jahre gegeben sind, wird der Durchschnitt berechnet
# oder mit Option "-tmy" Monate für ein typisches meteorologisches Jahr gewählt.
# Mit den Optionen "-hour"/"-day"/"-week"/"-month" wird jeweils eine CSV-Datei
# mit gegebenen Namen mit Statistik-Daten pro Stunde/Tag/Woche/Monat erzeugt.
#
# Beispiel:
# Solar.pl Lastprofil.csv 3000 Solardaten_1215_kWh.csv 1000 -lim 600 -tmy
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
#          [-eff <system efficiency in %, default from PV data file(s)>]
#          [-capacity <storage capacity in Wh, default 0 (no battery)>
#          [-ceff <charging efficiency in %, default 94]
#          [-seff <storage efficiency in %, default 95]
#          [-deff <discharging efficiency in %, default 94]
#          [-en] [-tmy] [-lim <inverter output power limit in W>]
#          [-hour <file>] [-day <file>] [-week <file>] [-month <file>]
# Use "-en" for text output in English.
# When PV data for more than one year is given, the average is computed, while
# with the option "-tmy" months for a typical meteorological year are selected.
# With each the options "-hour"/"-day"/"-week"/"-month" a CSV file is produced
# with the given name containing with statistical data per hour/day/week/month.
#
# Example:
# Solar.pl loadprofile.csv 3000 solardata_1215_kWh.csv 1000 -lim 600 -tmy -en
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

my @PV_files;
my @PV_peaks;       # nominal/maximal PV output(s), default from PV data file(s)
my ($lat, $lon);    # from PV data file(s)
my $sys_efficiency; # system efficiency, default from PV data file(s)
my $capacity;            # usable storage capacity in Wh on average degradation
my $charge_eff    = .94; # charge efficiency
my $storage_eff   = .95; # storage efficiency
my $discharge_eff = .94; # discharge efficiency
my $nominal_power_sum = 0;

while ($#ARGV >= 0 && $ARGV[0] =~ m/^\s*[^-]/) {
    push @PV_files, shift @ARGV; # PV data file
    push @PV_peaks, $#ARGV >= 0 && $ARGV[0] =~ m/^[\d\.]+$/ ? shift @ARGV : 0;
}
die "Missing PV data file" if $#PV_files < 0;

sub num_arg {
    die "Missing number arg for -eff/-lim option"
        unless $ARGV[0] =~ m/^[\d\.]+$/;
    return shift @ARGV;
}
sub str_arg {
    die "Missing arg for -max/-hour/-day/-week/-month option" if $#ARGV < 0;
    return shift @ARGV;
}
my ($en, $tmy, $PV_limit, $max, $hourly, $daily, $weekly,  $monthly);
while ($#ARGV >= 0) {
    if      ($ARGV[0] eq "-en"    && shift @ARGV) { $en       = 1;
    } elsif ($ARGV[0] eq "-eff"   && shift @ARGV) { $sys_efficiency =
                                                        num_arg() / 100;
    } elsif ($ARGV[0] eq "-tmy"   && shift @ARGV) { $tmy      = 1;
    } elsif ($ARGV[0] eq "-lim"   && shift @ARGV) { $PV_limit = num_arg();
    } elsif ($ARGV[0] eq "-max"   && shift @ARGV) { $max      = str_arg();
    } elsif ($ARGV[0] eq "-capacity"&&shift@ARGV) { $capacity = str_arg();
    } elsif ($ARGV[0] eq "-ceff"  && shift @ARGV) { $charge_eff =
                                                        num_arg() / 100;
    } elsif ($ARGV[0] eq "-seff"  && shift @ARGV) { $storage_eff =
                                                        num_arg() / 100;
    } elsif ($ARGV[0] eq "-deff"  && shift @ARGV) { $discharge_eff =
                                                        num_arg() / 100;
    } elsif ($ARGV[0] eq "-hour"  && shift @ARGV) { $hourly   = str_arg();
    } elsif ($ARGV[0] eq "-day"   && shift @ARGV) { $daily    = str_arg();
    } elsif ($ARGV[0] eq "-week"  && shift @ARGV) { $weekly   = str_arg();
    } elsif ($ARGV[0] eq "-month" && shift @ARGV) { $monthly  = str_arg();
    } else { die "Invalid option: $ARGV[0]";
    }
}

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

sub kWh     { return sprintf("%5d kWh", round(shift() / 1000)); }
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
my $hour_per_year = 0;
sub get_profile {
    my $file = shift;
    open(my $IN, '<', $file) or die "Could not open profile file $file: $!\n";

    my $warned_just_before = 0;
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
        my $load = 0;
        for (my $item = 0; $item < $n; $item++) {
            my $point = $sources[$item];
            die "Error parsing load item: '$point' in $file line $."
                unless $point =~ m/^\s*-?[\.\d]+\s*$/;
            if ($point > $load_max) {
                $load_max = $point;
                $load_max_time =
                    time_string($month, $day, $hour, 60 * $item / $n);
            }
            $load += $point;
            $load_by_item[$month][$day][$hour][$item] = $point;
            if ($point <= 0) {
                print "Load on YYYY-".index_string($month, $day, $hour, $item).
                    " = ".sprintf("%4d", $point)."\n"
                    unless $warned_just_before;
                $warned_just_before = 1;
            } else {
                $warned_just_before = 0;
            }
        }
        $load /= $n;
        $load_by_hour[$month][$day][$hour] += $load;
        $load_sum    += $load;
        $night_sum   += $load if   NIGHT_START <= $hour && $hour <   NIGHT_END;
        $morning_sum += $load if MORNING_START <= $hour && $hour < MORNING_END;
        $bright_sum  += $load if  BRIGHT_START <= $hour && $hour <  BRIGHT_END;
        $earleve_sum += $load if LAFTERN_START <= $hour && $hour < LAFTERN_END;
        $evening_sum += $load if EVENING_START <= $hour && $hour < EVENING_END;
        $winter_sum  += $load if WINTER_START <= $month || $month < WINTER_END;

        $hour = 0 if ++$hour == 24;
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
print "$n_txt =   ".percent($night_sum   / $load_sum)." %\n";
print "$m_txt =   ".percent($morning_sum / $load_sum)." %\n";
print "$s_txt =   ".percent($bright_sum  / $load_sum)." %\n";
print "$a_txt =   ".percent($earleve_sum / $load_sum)." %\n";
print "$e_txt =   ".percent($evening_sum / $load_sum)." %\n";
print "$w_txt =   ".percent($winter_sum  / $load_sum)." %\n";
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

    # my $sum_needed = 0;
    my ($slope, $azimuth);
    my $current_years = 0;
    my $nominal_power_deflt; # PVGIS default: 1 kWp
    my $sys_efficiency_deflt; # PVGIS default: 0.86
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
        if (!$sys_efficiency_deflt && m/System losses \(%\):\s*([\d\.]+)/) {
            $sys_efficiency_deflt = 1 - $1/100;
            print "Ignoring efficiency $sys_efficiency_deflt provided by PVGIS\n"
                if 0 && $sys_efficiency
                && $sys_efficiency != $sys_efficiency_deflt;
            $sys_efficiency = $sys_efficiency_deflt
                unless defined $sys_efficiency;
        }
        $power_provided = 1 if m/^time,P,/;

        next unless m/^20\d\d\d\d\d\d:\d\d\d\d,/;
        die "Missing latitude in $file"  unless $lat;
        die "Missing longitude in $file" unless $lon;
        die "Missing slope in $file"     unless $slope;
        die "Missing azimuth in $file"   unless $azimuth;
        die "Missing nominal power in $file"     unless $nominal_power_deflt;
        die "Missing system efficiency in $file" unless $sys_efficiency_deflt;
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
             / $nominal_power_deflt / $sys_efficiency_deflt);
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
    my $s13 = "             ";
    print "".($en ? "PV data file  " : "PV-Daten-Datei")."$s13 : $file\n";
    print "".($en ? "slope         " : "Neigungswinkel")."$s13 =  $slope\n";
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

my $load_scale =
    defined $simulated_load ? 1000 * $simulated_load / $load_sum : 1;
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
my $charge_sum = 0;
my $grid_feed_in = 0;

sub simulate()
{
    my $year = 0;
    ($month, $day, $hour) = (1, 1, 0);
    my $minute = 0; # currently fixed

    my $charge = 0;
    # factor out $load_scale for optimizing the inner loop
    $capacity /= ($load_scale != 0 ? $load_scale : 1) if $capacity;

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
        $power *= $sys_efficiency;

        $PV_net_loss[$month][$day][$hour] = 0;
        my $PV_loss = 0;
        if ($PV_limit && $power > $PV_limit) {
            $PV_loss = ($power - $PV_limit); # * $sys_efficiency;
            $PV_net_loss[$month][$day][$hour] += $PV_loss;
            $PV_net_loss_sum += $PV_loss;
            $PV_net_loss_hours++;
            $power = $PV_limit;
            # print "$year-".time_string($month, $day, $hour, $minute).
            #"\tPV=".round($power)."\tlimit=".round($PV_limit).
            #"\tloss=".round($PV_net_loss_sum)."\t$_\n";
        }
        # $power *= $sys_efficiency;
        $PV_net_out[$month][$day][$hour] += $power;
        $PV_net_out_sum += $power;
        $PV_net_bright_sum += $power
            if BRIGHT_START <= $hour && $hour < BRIGHT_END;

        # factor out $load_scale for optimizing the inner loop
        my $effective_PV_power = $power / ($load_scale != 0 ? $load_scale : 1);
        $PV_loss /= $load_scale if $load_scale != 0;
        # my $needed = 0;
        my $usages = 0;
        my $losses = 0;
        $items = $items_by_hour[$month][$day][$hour];

        # factor out $items for optimizing the inner loop
        if ($capacity) {
            $capacity *= $items;
            $charge *= $items;
            $charge_sum *= $items;
            $PV_used_via_storage *= $items;
            $grid_feed_in *= $items;
        }

        for (my $item = 0; $item < $items; $item++) {
            my $loss = 0;
            my $point = $load_by_item[$month][$day][$hour][$item];
            # $needed += $point;
            my $power_diff = $effective_PV_power - $point;
            my $pv_used = $effective_PV_power; # will be PV own consumption
            if ($power_diff > 0) {
                $pv_used = $point; # == min($effective_PV_power, $point);
                $grid_feed_in += $power_diff if $capacity;
            }

            if ($capacity) { # storage available
                if ($capacity > $charge # storage not full
                    && $effective_PV_power > $point) {
                    # optimal charge: exactly as much as currently unused
                    my $charge_delta = min($power_diff, $capacity - $charge);
                    $grid_feed_in -= $charge_delta;
                    $charge_delta *= $charge_eff;
                    $charge += $charge_delta;
                    $charge_sum += $charge_delta;
                    $loss += min(($capacity - $charge) * $charge_eff
                                 * $storage_eff * $discharge_eff, $PV_loss)
                        if $PV_loss != 0;
                } elsif ($charge > 0 # storage no empty
                         && $point > $effective_PV_power) {
                    # optimal discharge: exactly as much as currently needed
                    my $discharge = min($point - $effective_PV_power, $charge);
                    $charge -= $discharge;
                    $discharge *= $storage_eff * $discharge_eff;
                    $pv_used += $discharge;
                    $PV_used_via_storage += $discharge;
                }
            }
            $usages += $pv_used;
            $PV_used_by_item[$month][$day][$hour][$item] =
                round($pv_used * $load_scale) if $max;

            if ($PV_loss != 0 && $point > $effective_PV_power) {
                $loss += min($point - $effective_PV_power, $PV_loss);
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
        $PV_usage_loss[$month][$day][$hour] += $losses * $load_scale / $items;
        $PV_usage_loss_sum += $losses / $items;
        # $sum_needed += $needed * $load_scale / $items; # per hour
        # print "$year-".time_string($month, $day, $hour, $minute).
        # "\tPV=".round($power)."\tPN=".round($needed)."\tPU=".round($usages).
        # "\t$_\n" if $power != 0 && m/^20160214:1010/; # m/^20....02:12/;
        $usages *= $load_scale / $items;
        $PV_used[$month][$day][$hour] += $usages;
        $PV_used_sum += $usages;

        # revert factoring out $items for optimizing the inner loop
        if ($capacity) {
            $capacity /= $items;
            $charge /= $items;
            $charge_sum /= $items;
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
    if ($capacity) {
        $grid_feed_in *= $load_scale / $years;
    } else {
        $grid_feed_in = $PV_net_out_sum - $PV_used_sum;
    }

    $charge_sum *= $load_scale / $years;
    $capacity *= ($load_scale != 0 ? $load_scale : 1) if $capacity;
}

simulate();

################################################################################
# statistics output

my $nominal_txt      = $en ? "nominal PV power"     : "PV-Nominalleistung";
my $max_gross_txt    = $en ? "max gross PV power"   : "Bruttoleistung max.";
my $PV_limit_txt     = $en ? "cropping by inverter" : "Leistungsbegrenzung";
my $system_eff_txt   = $en ? "system efficiency"    : "System-Wirkungsgrad";
my $own_txt          = $en ? "own consumption "     : "Eigenverbrauch";
my $own_ratio_txt    = $own_txt . ($en ? "ratio"    : "santeil");
my $own_storage_txt  = $en ? $own_txt."via storage":"$own_txt via Speicher";
my $load_cover_txt   = $en ? "load coverage ratio"  : "Eigendeckungsanteil";
my $PV_gross_txt     = $en ? "PV gross yield"       : "PV-Bruttoertrag";
my $PV_net_txt       = $en ? "PV net yield"         : "PV-Nettoertrag";
my $PV_net_lim_txt   = $en ? "inverter output cropped"
    : "WR-Ausgangsleistung limitiert";
my $PV_loss_lim_txt  = $en ? "output cropping loss" : "WR-Abregelungsverlust";
my $load_txt         = $en ? "load by household"    : "Last durch Haushalt";
my $use_loss_txt     = $en ? "PV $own_txt"."loss"   : $own_txt."sverlust";
my $by_limit         = $en ? "by cropping"          : "durch Limit";
my $use_wo_lim_txt   = $en ? "$own_txt"."w/o cropping" : "$own_txt ohne Limit";
my $use_with_lim_txt = $en ? "own consumption".($PV_limit ? " w/ cropping" : "")
                           : "Eigenverbrauch" .($PV_limit ? " mit Limit"   :"");
my $grid_feed_txt    = $en ? "grid feed-in"         : "Netzeinspeisung";
my $each             = $en ? "each"   : "alle";
my $yearly_txt       = $en ? "yearly" : "jährlich";
my $capacity_txt     = $en ? "storage capacity"     : "Speicherkapazität";
my $ceff_txt         = $en ? "charging efficiency"  : "Lade-Wirkungsgrad";
my $seff_txt         = $en ? "storage efficiency" : "Speicher-Wirkungsgrad";
my $deff_txt         = $en ? "discharging efficiency" : "Entlade-Wirkungsgrad";
my $stored_txt       = $en ? "buffered energy"      : "Zwischenspeicherung";
my $cycles_txt       = $en ? "full cycles per year" : "Vollzyklen pro Jahr";

my $own_usage     =
    percent($PV_net_out_sum ? $PV_used_sum / $PV_net_out_sum : 0);
my $load_coverage =
    percent($load_sum ? $PV_used_sum / $load_sum : 0);
my $yield_daytime =
    percent($PV_net_out_sum ? $PV_net_bright_sum / $PV_net_out_sum : 0);
my $cycles = round($charge_sum / $capacity) if $capacity;

$sys_efficiency *= 100;
$charge_eff *= 100;
$storage_eff *= 100;
$discharge_eff *= 100;

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
    print $OU
        "$load_txt in kWh, $profile_txt,$pv_data_txt\n";
    print $OU "".round($load_sum / 1000).
        ", $profile, ".join(", ", @PV_files)."\n";
    print $OU " $nominal_txt in Wp, $max_gross_txt in W, "
        ."$PV_limit_txt in W, $system_eff_txt in %, "
        ."$own_ratio_txt in %, $load_cover_txt in %\n";
    print $OU "$nominal_sum, ".round($PV_gross_max).", "
        .($PV_limit ? $PV_limit : "none").
        ", $sys_efficiency, $own_usage, $load_coverage\n";
    print $OU "$capacity_txt in Wh, $ceff_txt in %, $seff_txt in %, $deff_txt".
        "in %, $own_storage_txt in kWh, $stored_txt in kWh, $cycles_txt\n"
        if $capacity;
    print $OU "$capacity, $charge_eff, $storage_eff, $discharge_eff, ".
        round($PV_used_via_storage / 1000).", ".
        round($charge_sum / 1000).", $cycles\n" if $capacity;
    print $OU "$yearly_txt, $PV_gross_txt, $PV_net_txt, $PV_loss_lim_txt, ".
        "$use_loss_txt $by_limit, $use_with_lim_txt, ".
        "$grid_feed_txt, $each in kWh\n";
    print $OU "".($en ? "sums" : "Summen"  ).", ".
        round($PV_gross_out_sum  / 1000).", ".
        round($PV_net_out_sum    / 1000).", ".
        round($PV_net_loss_sum   / 1000).", ".
        round($PV_usage_loss_sum / 1000).", ".
        round($PV_used_sum       / 1000).", ".
        round($grid_feed_in      / 1000)."\n";
    print $OU "$res_txt, $PV_gross_txt, $PV_net_txt, $PV_net_lim_txt, "
        ."$load_txt, $use_wo_lim_txt, $use_with_lim_txt, $each in Wh\n";
    ($month, my $week, my $days, $day, $hour) = (1, 1, 0, 1, 0);
    my ($gross, $PV_loss, $net, $load, $loss, $used) = (0, 0, 0, 0, 0, 0);
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
        $load    += round($load_by_hour [$month][$day][$hour] * $load_scale);
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
                    $load =    round($load_by_item[$m][$d][$h][$i]
                                     * $load_scale);
                    $loss = $PV_usage_loss_by_item[$m][$d][$h][$i];
                    $used = $PV_used_by_item      [$m][$d][$h][$i];
                }
                print $OU "$tim$minute, $gross, ".($net + $PV_loss).
                      ", $net, $load, ".($used + $loss).", $used\n";
            }
            ($gross, $PV_loss, $net, $load, $loss, $used) = (0, 0, 0, 0, 0, 0);
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
my $with           = $en ? "with"                 : "mit";
my $during         = $en ? "during"               : "während";
my $lim            = $en ? "by cropping at"       : "durch Limitierung auf";
my $yield          = $en ? "yield portion"        : "Ertragsanteil";
my $daytime        = $en ? "9 AM - 3 PM "         : "9-15 Uhr MEZ";
my $of_yield       = $en ? "of net yield"         : "des Nettoertrags (Nutzungsgrad)";
my $of_consumption = $en ? "of consumption"       : "des Verbrauchs (Autarkiegrad)";
my $PV_loss_txt    = $en ? "PV loss             " : "PV-Abregelungsverlust";
   $use_with_lim_txt  = "$use_with_lim_txt $en2         " unless $PV_limit;
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
              " $at $system_eff_txt $sys_efficiency%\n";
print "$PV_loss_txt $en1      ="       .kWh($PV_net_loss_sum)." $during "
    .round($PV_net_loss_hours)." h $lim $PV_limit W\n" if $PV_limit;
print "$yield $daytime  =   $yield_daytime %\n";

print "\n";
print "$load_txt $en2        =" .kWh($load_sum)."\n";
print "$use_with_lim_txt $de3=" .kWh($PV_used_sum)."\n";
print "$use_loss_txt $de1    =" .kWh($PV_usage_loss_sum)." $during "
    .round($PV_usage_loss_hours)." h $lim $PV_limit W\n" if $PV_limit;
print "$grid_feed_txt $en3            =" .kWh($grid_feed_in)."\n";
print "$own_ratio_txt       =   $own_usage % $of_yield\n";
print "$load_cover_txt         =   $load_coverage % $of_consumption\n";
if ($capacity) {
    print "\n".
        "$capacity_txt $en1          =" .W($capacity)."h $with ".
        "$ceff_txt $charge_eff %,\n";
    printf "$seff_txt $en3         %3d %%,".
        " $en1 $deff_txt $discharge_eff %%\n", $storage_eff;
    print "$own_storage_txt =" .kWh($PV_used_via_storage)."\n";
    print "$stored_txt $en2$en2        =" .kWh($charge_sum)."\n";
    # Vollzyklen, Kapazitätsdurchgänge pro Jahr Kapazitätsdurchsatz:
    printf "$cycles_txt $de1       =  %3d\n", $cycles;
}
