#!/usr/bin/perl
#
################################################################################
# Eigenverbrauchs-Simulation mit stündlichen PV-Daten und einem Lastprofil
# Nutzung: Solarertrag.pl <Lastprofil.csv> <Solardaten.csv> [<Nennleistung in Wp>
#          [<Wirkungsgrad in %> [<Jahresverbrauch in kW> [<Limitierung in W>]]]]
# Beispiel:
# Solarertrag.pl Lastprofil_4673_kWh.csv Solardaten_1215_kWh.csv 1000 88 3000 600
#
# Solardaten zu beziehen von https://re.jrc.ec.europa.eu/pvg_tools/de/#api_5.2
# "Daten pro Stunde" wählen, optional PV-Leistung und Systemverlust anpassen
# Startjahr 2008 (oder früher), Endjahr 2018 (oder später)
# Häkchen bei PV-Leistung setzen! Dann Download-Knopf "csv" drücken.
# (c) 2022 David von Oheimb - License: MIT - Version 2.3
################################################################################

use strict;
use warnings;

sub min { return $_[0] < $_[1] ? $_[0] : $_[1]; }
sub max { return $_[0] > $_[1] ? $_[0] : $_[1]; }
sub round { return int(.5 + $_[0]); }

sub time_string {
    return
        sprintf("%02d", $_[0])."-".
        sprintf("%02d", $_[1])." um ".
        sprintf("%02d", $_[2]).":".
        sprintf("%02d", $_[3]);
}

sub kWh { return sprintf("%5d kWh", round(shift()/1000)); }
sub W   { return sprintf("%5d W"  , round(shift())); }

use constant TMY => 1; # Typisches meteorologisches Jahr

# Alle Uhrzeiten ohne Sommerzeit-Umschaltung
use constant BRIGHT_START => 9; # Start heller Sonnenschein
use constant BRIGHT_END   => 15; # Ende heller Sonnenschein
use constant NIGHT_START  => 18; # Start Abend
use constant NIGHT_END    =>  6; # Ende Morgen

my $profile = $ARGV[0]; # Profildatei
my $items = 0; # Zahl der verwendeten Messpunkte pro Minute
my $load_max = 0;
my $load_max_time;
my $load_sum = 0;
my $load_bright_sum = 0;
my $load_night_sum = 0;

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
# Lastprofil einlesen

my @load_by_item;
sub get_profile {
    my $file = shift;
    open(my $IN, '<', $file) or die "Could not open profile file $file $!\n";

    for (my $hour_per_year = 0; my $line = <$IN>; $hour_per_year++) {
        chomp $line;
        my @sources = split "," , $line;
        shift @sources; # Ignoriere Datums- und Uhrzeit-Information
        die "Inconsistent number of items per line in $file"
            if $items != 0 && $items != $#sources + 1;
        $items = $#sources + 1;
        my $load = 0;
        for (my $item = 0; $item < $items; $item++) {
            my $point = $sources[$item];
            if ($point > $load_max) {
                $load_max = $point;
                $load_max_time =
                    time_string($month, $day, $hour, $item / $items * 60);
            }
            $load += $point;
            $load_by_item[$month][$day][$hour][$item] = $point;
        }
        $load_sum += $load;
        $load_bright_sum += $load
            if BRIGHT_START <= $hour && $hour < BRIGHT_END;
        $load_night_sum += $load
            if $hour < NIGHT_END || NIGHT_START <= $hour;
        $hour = 0 if ++$hour == 24;
        adjust_day_month();
    }
    close $IN;
    $month--;
    die "Got $month months rather than 12" if $month != 12;

    $load_sum /= $items;
    $load_bright_sum /= $items;
    $load_night_sum /= $items;
}

get_profile($profile);

my $load_bright_part = $load_bright_sum / $load_sum;
my $load_night_part = $load_night_sum / $load_sum;
print "Last-Datenpunkte pro Stunde = $items\n";
print "Lastanteil 9-15 Uhr MEZ     = ".round($load_bright_part * 100)."%\n";
print "Lastanteil 18-6 Uhr MEZ     = ".round($load_night_part * 100)."%\n";
print "Verbrauch gemäß Lastprofil  =".kWh($load_sum)."\n";
print "Maximallast                 =".W($load_max)." am $load_max_time\n";

################################################################################
# Simulation

my $nominal_power_deflt = 0; # PVGIS default: 1 kWp
my $nominal_power = $ARGV[2];
my $PV_gross_max = 0;
my $PV_gross_max_time;
my $PV_gross_out_sum = 0;

my $sys_efficiency_deflt = 0; # PVGIS default: 0.86
my $sys_efficiency = $ARGV[3];
my @PV_net_out;
my $PV_net_out_sum = 0;
my $PV_net_bright_sum = 0;

my $simulated_load = $ARGV[4]; # default by load profile
my $load_scale = $simulated_load ? 1000 * $simulated_load / $load_sum : 1;
my $power_limit = $ARGV[5]; # default none
my $PV_net_loss = 0;
my $PV_net_loss_hours = 0;
my $PV_usage_loss = 0;
my $PV_usage_loss_hours = 0;
my @PV_used;
my $PV_used_sum = 0;

sub timeseries {
    my $file = shift;
    open(my $IN, '<', $file) or die "Could not open PV data file $file $!\n";

    # my $sum_needed = 0;
    my $power_rate;
    my $power_provided = 0; # PVGIS default: none
    my ($years, $months, $hours) = (0, 0, 0);
    while (<$IN>) {
        chomp;
        if (!$nominal_power_deflt && m/Nominal power.*? \(kWp\):\s*([\d\.]+)/) {
            $nominal_power_deflt = $1 * 1000;
            $nominal_power = $nominal_power_deflt unless $nominal_power;
        }
        if (!$sys_efficiency_deflt && m/System losses \(%\):\s*([\d\.]+)/) {
            $sys_efficiency_deflt = 1 - $1/100;
            $sys_efficiency = $sys_efficiency
                ? $sys_efficiency / 100 : $sys_efficiency_deflt;
        }
        $power_provided = 1 if m/^time,P,/;

        next unless m/^20\d\d\d\d\d\d:\d\d\d\d,/;
        die "Missing nominal power in $file" unless $nominal_power_deflt;
        die "Missing system efficiency in $file" unless $sys_efficiency_deflt;
        die "Missing PV power output data in $file" unless $power_provided;
        next if m/^20\d\d0229:/; # Schaltjahr

        if (TMY) {
            # Typisches metereologisches Jahr
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
        $years++ if m/^20..0101:00/;
        $months++ if m/^20....01:00/;

        die "Missing power data in $file line $_"
            unless m/^(\d\d\d\d)(\d\d)(\d\d):(\d\d)(\d\d),([\d\.]+)/;
        my ($year, $month, $day, $hour, $minute, $power) =
            ($1, $2, $3, $4, $5, $6 * $nominal_power
             / $nominal_power_deflt / $sys_efficiency_deflt);
        if ($power > $PV_gross_max) {
            $PV_gross_max = $power;
            $PV_gross_max_time =
                "$year-".time_string($month, $day, $hour, $minute);
        }
        $PV_gross_out_sum += $power;

        my $PV_loss = 0;
        if ($power_limit && $power > $power_limit) {
            $PV_loss = ($power - $power_limit) * $sys_efficiency;
            $PV_net_loss += $PV_loss;
            $PV_net_loss_hours++;
            $power = $power_limit;
            # print "$year-".time_string($month, $day, $hour, $minute).
            #"\tPV=".round($power)."\tlimit=".round($power_limit).
            #"\tloss=".round($PV_net_loss)."\t$_\n";
        }
        $power *= $sys_efficiency;
        $PV_net_out[$month][$day][$hour] += $power;
        $PV_net_out_sum += $power;
        $PV_net_bright_sum += $power
            if BRIGHT_START <= $hour && $hour < BRIGHT_END;

        # Faktor $load_scale herausziehen zur Optimierung der inneren Schleife
        my $effective_power = $power / $load_scale;
        $PV_loss /= $load_scale;
        my $needed = 0;
        my $used = 0;
        for (my $item = 0; $item < $items; $item++) {
            my $point = $load_by_item[$month][$day][$hour][$item];
            $needed += $point;
            my $power_used = min($effective_power, $point); # PV power self use
            $used += $power_used;
            if ($PV_loss != 0 && $point > $effective_power) {
                $PV_usage_loss += min($point - $effective_power, $PV_loss);
                $PV_usage_loss_hours++; # will be normalized by $items
            }
        }
        # $sum_needed += $needed * $load_scale; # pro Stunde
        $used = $used * $load_scale / $items; # mit Berücksichtigung Ausdünnung
        # print "$year-".time_string($month, $day, $hour, $minute).
        # "\tPV=".round($power)."\tPN=".round($needed)."\tPU=".round($used).
        # "\t$_\n" if $power != 0 && m/^20160214:1010/; # m/^20....02:12/;
        $PV_used[$month][$day][$hour] += $used;
        $PV_used_sum += $used;

        $hours++;
    }
    close $IN;
    die "Got $months months rather than ".(12 * $years)
        if $months != $years * 12;
    die "Got $hours hours rather than ".(8760 * $years)
        if $hours != $years * 8760;

    $load_max *= $load_scale;
    $load_sum *= $load_scale;
    $PV_gross_out_sum /= $years;
    $PV_net_loss /= $years;
    $PV_net_loss_hours /= $years;
    $PV_net_out_sum /= $years;
    $PV_net_bright_sum /= $years;
    $PV_used_sum /= $years;
    $PV_usage_loss *= $load_scale / ($items + $years);
    $PV_usage_loss_hours /= ($items * $years);
    # $sum_needed /= $items;
    # die "Inconsistent load calculation: sum = $sum vs. needed = $sum_needed"
    #     if round($sum) != round($sum_needed);
}

timeseries($ARGV[1]);

my $lim =" durch Limitierung auf $power_limit W";
my $PV_net_bright_part = $PV_net_bright_sum / $PV_net_out_sum;
print "\n";
print "PV-Nomialleistung           =".W($nominal_power)."p\n";
print "PV-Bruttoleistung Maximum   =".W($PV_gross_max).
    " am $PV_gross_max_time\n";
print "PV-Bruttoertrag             =".kWh($PV_gross_out_sum)."\n";
print "Ertrag nach Wechselrichtung =".kWh($PV_net_out_sum).
    " bei Systemwirkungsgrad ".($sys_efficiency * 100)."%\n";
print "PV-Netto-Abregelungsverlust =".kWh($PV_net_loss).
    " während ".round($PV_net_loss_hours)." h". "$lim\n" if $power_limit;
print "Ertragsanteil 9-15 Uhr MEZ  =   ".round($PV_net_bright_part * 100)."%\n";

print "\n";
print "Haushalts-Stromverbrauch    =".kWh($load_sum)."\n";
print "Eigenverbrauch              =".kWh($PV_used_sum)."\n";
print "Eigenverbrauchsverlust      =".kWh($PV_usage_loss).
    " während ".round($PV_usage_loss_hours)." h"."$lim\n" if $power_limit;
print "Eigenverbrauchsanteil       =   ".
    round($PV_used_sum / $PV_net_out_sum * 100)."% des Ertrags\n";
print "Eigendeckungsanteil         =   ".
    round($PV_used_sum / $load_sum * 100)."% des Verbrauchs\n";
