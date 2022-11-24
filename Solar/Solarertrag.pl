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

sub kWh { return sprintf("%5d kWh", int(shift()/1000)); }
sub W   { return sprintf("%5d W"  , int(shift())); }

use constant TMY => 1; # Typisches meteorologisches Jahr

my $profile = $ARGV[0]; # Profildatei
my $items = 0; # Zahl der verwendeten Messpunkte pro Minute
my $PV_bright_sum = 0;
my $bright_sum = 0;
my $night_sum = 0;
my $load_sum = 0;

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

my @load_by_date;
sub get_profile {
    my $file = shift;
    open(my $IN, '<', $file) or die "Could not open profile file $file $!\n";

    $hour = 0;
    for (my $hour_per_year = 0; my $line = <$IN>; $hour_per_year++) {
        chomp $line;
        my @sources = split "," , $line;
        shift @sources; # Ignoriere Datums- und Uhrzeit-Information
        die "Inconsistent number of items per line in $file"
            if $items != 0 && $items != $#sources + 1;
        $items = $#sources + 1;
        for (my $item = 0; $item < $items; $item++) {
            my $point = $sources[$item];
            $load_by_date[$month][$day][$hour][$item] = $point;
            $load_sum += $point;
            $bright_sum += $point if 9 <= $hour && $hour < 15;
            $night_sum += $point if $hour < 6 || 18 <= $hour;
        }
        $hour = 0 if ++$hour == 24;
        adjust_day_month();
    }
    close $IN;
    $month--;
    die "Got $month months rather than 12" if $month != 12;
}

get_profile($profile);
$load_sum /= $items;
$bright_sum /= $items;
$night_sum /= $items;
print "Last-Datenpunkte pro Stunde = $items\n";
print "Lastanteil von 9-15 Uhr MEZ = ".int($bright_sum / $load_sum * 100)."%\n";
print "Lastanteil von 18-6 Uhr MEZ = ".int($night_sum / $load_sum * 100)."%\n";
print "Verbrauch gemäß Lastprofil  =".kWh($load_sum)."\n";

################################################################################
# Simulation

my $PV_max = 0;
my $PV_max_time;
my $years = 0;
my $PV_gross = 0;
my @PV_out;
my $PV_out_sum = 0;
my $provided_nominal_power = 0; # PVGIS default: 1 kWp
my $nominal_power = $ARGV[2];
my $provided_efficiency = 0; # PVGIS default: 0.86
my $system_efficiency = $ARGV[3];
my $simulated_load = $ARGV[4]; # default by load profile
my $load_max = 0;
my $load_max_time;
my $load_scale = $simulated_load ? 1000 * $simulated_load / $load_sum : 1;
my $power_limit = $ARGV[5]; # default none
my $limitation_PVout_loss = 0;
my $limitation_usage_loss = 0;
my $hours_limit_loss = 0;
my $hours_usage_loss = 0;
my @PV_used;
my $PV_used_sum = 0;

sub timeseries {
    my $file = shift;
    open(my $IN, '<', $file) or die "Could not open PV data file $file $!\n";

    # my $sum_needed = 0;
    my $power_rate;
    my $provided_power_output = 0; # PVGIS default: none
    my $months = 0;
    my $hours = 0;
    while (<$IN>) {
        chomp;
        if (!$provided_nominal_power && m/^Nominal power.*? \(kWp\):\s*([\d\.]+)/) {
            $provided_nominal_power = $1 * 1000;
            $nominal_power = $provided_nominal_power unless $nominal_power;
        }
        if (!$provided_efficiency && m/^System losses \(%\):\s*([\d\.]+)/) {
            $provided_efficiency = 1 - $1/100;
            $system_efficiency = $system_efficiency
                ? $system_efficiency / 100 : $provided_efficiency;
        }
        $provided_power_output = 1 if m/^time,P,/;

        next unless m/^20\d\d\d\d\d\d:\d\d\d\d,/;
        die "Missing nominal power in $file" unless $provided_nominal_power;
        die "Missing system efficiency in $file" unless $provided_efficiency;
        die "Missing PV power output data in $file" unless $provided_power_output;
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
            ($1, $2, $3, $4, $5, $6
             * $nominal_power / $provided_nominal_power / $provided_efficiency);
        if ($power > $PV_max) {
            $PV_max = $power;
            $PV_max_time = "$year-$month-$day um $hour:$minute";
        }
        $PV_gross += $power;

        my $power_loss = 0;
        if ($power_limit && $power > $power_limit) {
            $power_loss = ($power - $power_limit) * $system_efficiency;
            $limitation_PVout_loss += $power_loss;
            $hours_limit_loss++;
            $power = $power_limit;
            # print "$year-$month-$day:$hour:$minute\tPV=".int($power)."\tlimit=".int($power_limit)."\tloss=".int($limitation_PVout_loss)."\t$_\n";
        }
        $power *= $system_efficiency;
        $PV_out[$month][$day][$hour] += $power;
        $PV_out_sum += $power;
        $PV_bright_sum += $power if 9 <= $hour && $hour < 15;

        # Faktor $load_scale herausziehen zur Optimierung der inneren Schleife
        my $effective_power = $power / $load_scale;
        $power_loss /= $load_scale;
        my $needed = 0;
        my $used = 0;
        for (my $item = 0; $item < $items; $item++) {
            my $point = $load_by_date[$month][$day][$hour][$item];
            if ($point > $load_max) {
                $load_max = $point;
                $load_max_time = "$year-$month-$day um $hour:".
                    sprintf("%02d",$item / $items * 60);
            }
            $needed += $point;
            my $power_used = min($effective_power, $point); # PV power self use
            $used += $power_used;
            if ($power_loss && $point > $effective_power) {
                $limitation_usage_loss +=
                    min($point - $effective_power, $power_loss) / $items;
                $hours_usage_loss++; # will be normalized by $items
            }
        }
        # $sum_needed += $needed * $load_scale; # pro Stunde
        $used = $used * $load_scale / $items; # mit Berücksichtigung Ausdünnung
        # print "$year-$month-$day:$hour:$minute\tPV=".int($power)."\tPN=".int($needed)."\tPU=".int($used)."\t$_\n" if $power != 0 && m/^20160214:1010/; # m/^20....02:12/;
        $PV_used[$month][$day][$hour] += $used;
        $PV_used_sum += $used;

        $hours++;
    }
    close $IN;
    die "Got $months months rather than ".(12 * $years)
        if $months != $years * 12;
    die "Got $hours hours rather than ".(8760 * $years)
        if $hours != $years * 8760;
    $load_sum *= $load_scale;
    $load_max *= $load_scale;
    $PV_gross /= $years;
    $limitation_PVout_loss /= $years;
    $PV_out_sum /= $years;
    $PV_bright_sum /= $years;
    $PV_used_sum /= $years;
    $limitation_usage_loss *= $load_scale / $years;
    $hours_limit_loss /= $years;
    $hours_usage_loss /= ($items * $years);
    # $sum_needed /= $items;
    # die "Inconsistent load calculation: sum = $sum vs. needed = $sum_needed"
    #     if int($sum + 0.5) != int($sum_needed + 0.5);
}

timeseries($ARGV[1]);

print "Haushalts-Stromverbrauch    =".kWh($load_sum)."\n";
#print "Maximallast                 =".W($load_max)." am $load_max_time\n";
print "PV-Bruttoleistung Maximum   =".W($PV_max)." am $PV_max_time\n";
print "PV-Nomialleistung           =".W($nominal_power)."p\n";
print "PV-Bruttoertrag             =".kWh($PV_gross)."\n";
print "Ertrag nach Wechselrichtung =".kWh($PV_out_sum).
    " bei Systemwirkungsgrad ".($system_efficiency * 100)."%\n";
print "Ertragsanteil 9-15 Uhr MEZ  =   ".
    int($PV_bright_sum / $PV_out_sum * 100)."%\n";
print "PV-Netto-Abregelungsverlust =".kWh($limitation_PVout_loss).
    " während ".int($hours_limit_loss)." h".
    " durch Limitierung auf $power_limit W\n" if $power_limit;
print "Eigenverbrauchsverlust      =".kWh($limitation_usage_loss).
    " während ".int($hours_usage_loss)." h".
    " durch die Limitierung\n" if $power_limit;
print "Eigenverbrauch              =".kWh($PV_used_sum)."\n";
print "Eigenverbrauchsanteil       =   ".
    int($PV_used_sum / $PV_out_sum * 100)."% vom Ertrag\n";
print "Eigendeckungsanteil         =   ".
    int($PV_used_sum / $load_sum * 100)."% vom Verbrauch\n";
