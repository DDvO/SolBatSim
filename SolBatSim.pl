#!/usr/bin/perl
################################################################################
# Eigenverbrauchs-Simulation mit stündlichen PV-Daten und einem mindestens
# stündlichen Lastprofil, optional mit PV-Strom-Limitierung durch Wechselrichter
#
# Nutzung: Solar.pl <Lastprofil-Datei> <Solardaten-Datei> <Nennleistung in Wp>
#            <Wirkungsgrad in %> <Jahresverbrauch in kW>
#            [-en] [-tmy] [-lim <PV-Limitierung durch Wechselrichter in W>]
#            [-hour <Datei>] [-day <Datei>] [-week <Datei>] [-month <Datei>]
# Mit "-en" erfolgen die Textausgaben auf Englisch.
# Wenn PV-Daten für mehrere Jahre gegeben sind, wird der Durchschnitt berechnet
# oder mit Option "-tmy" Monate für ein typisches meteorologisches Jahr gewählt.
# Mit den Optionen "-hour"/"-day"/"-week"/"-month" wird jeweils eine CSV-Datei
# mit gegebenen Namen mit Statistik-Daten pro Stunde/Tag/Woche/Monat erzeugt.
#
# Beispiel:
# Solar.pl Lastprofil.csv Solardaten_1215_kWh.csv 1000 88 3000 -lim 600 -tmy
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
# Optionally takes into account cropping of PV input by solar inverter.
#
# Usage: Solar.pl <load profile file> <solar data file> <nominal power in Wp>
#          <system efficiency in %> <consumption per year in kW>
#          [-en] [-tmy] [-lim <PV input limit in W>]
#          [-hour <file>] [-day <file>] [-week <file>] [-month <file>]
# Use "-en" for text output in English.
# When PV data for more than one year is given, the average is computed, while
# with the option "-tmy" months for a typical meteorological year are selected.
# With each the options "-hour"/"-day"/"-week"/"-month" a CSV file is produced
# with the given name containing with statistical data per hour/day/week/month.
#
# Example:
# Solar.pl loadprofile.csv solardata_1215_kWh.csv 1000 88 3000 -lim 600 -tmy -en
#
# Take solar data from https://re.jrc.ec.europa.eu/pvg_tools/
# Select location, click "HOURLY DATA", and set the check mark at "PV power".
# Optionally may adapt "Installed peak PV power" and "System loss" already here.
# For using TMY data, choose Start year 2008 or earlier, End year 2018 or later.
# Then press the download button marked "csv".
#
# (c) 2022 David von Oheimb - License: MIT - Version 2.3
################################################################################

use strict;
use warnings;

my $profile             = shift @ARGV; # load profile file
my $PV_data             = shift @ARGV; # solar data file
my $nominal_power       = shift @ARGV; # maximal PV output
my $sys_efficiency      = shift @ARGV; # default by solar data file
my $simulated_load      = shift @ARGV; # default by solar data file
my ($en, $tmy, $PV_limit, $max, $hourly, $daily, $weekly,  $monthly);
while ($#ARGV >= 0) {
    if      ($ARGV[0] eq "-en"                  ) { $en        = shift @ARGV;
    } elsif ($ARGV[0] eq "-tmy"                 ) { $tmy       = shift @ARGV;
    } elsif ($ARGV[0] eq "-lim"    && shift @ARGV) { $PV_limit = shift @ARGV;
    } elsif ($ARGV[0] eq "-max"    && shift @ARGV) { $max      = shift @ARGV;
    } elsif ($ARGV[0] eq "-hour"   && shift @ARGV) { $hourly   = shift @ARGV;
    } elsif ($ARGV[0] eq "-day"    && shift @ARGV) { $daily    = shift @ARGV;
    } elsif ($ARGV[0] eq "-week"   && shift @ARGV) { $weekly   = shift @ARGV;
    } elsif ($ARGV[0] eq "-month"  && shift @ARGV) { $monthly  = shift @ARGV;
    } else { die "Invalid option: $ARGV[0]";
    }
}

# deliberately not using any extra packages like Math
sub min { return $_[0] < $_[1] ? $_[0] : $_[1]; }
sub max { return $_[0] > $_[1] ? $_[0] : $_[1]; }
sub round { return int(.5 + shift); }

my $at = $en ? "at" : "um";
sub time_string {
    return
        sprintf("%02d", shift)."-".
        sprintf("%02d", shift)." $at ".
        sprintf("%02d", shift).":".
        sprintf("%02d", shift);
}

sub kWh     { return sprintf("%5d kWh", round(shift() / 1000)); }
sub W       { return sprintf("%5d W"  , round(shift()       )); }
sub percent { return sprintf("%2d"    , round(shift() *  100)); }

# all hours according to local time without switching for daylight saving
use constant BRIGHT_START =>  9; # start bright sunshine
use constant BRIGHT_END   => 15; # end bright sunshine
use constant NIGHT_START  => 18; # start no sunshine (yearly average)
use constant NIGHT_END    =>  6; # end no shunshine (yearly average)

use constant YearHours => 24 * 365;

my $sum_items = 0;
my $items = 0; # number of load measure points in current hour
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
# read load profile data

my @items_by_hour;
my @load_by_item;
my @load_by_hour;
my $hour_per_year = 0;
sub get_profile {
    my $file = shift;
    open(my $IN, '<', $file) or die "Could not open profile file $file $!\n";

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
            if ($point > $load_max) {
                $load_max = $point;
                $load_max_time =
                    time_string($month, $day, $hour, int(60 * $item / $n));
            }
            $load += $point;
            $load_by_item[$month][$day][$hour][$item] = $point;
        }
        $load /= $n;
        $load_sum += $load;
        $load_by_hour[$month][$day][$hour] += $load;
        $load_bright_sum += $load
            if BRIGHT_START <= $hour && $hour < BRIGHT_END;
        $load_night_sum += $load
            if $hour < NIGHT_END || NIGHT_START <= $hour;
        $hour = 0 if ++$hour == 24;
        adjust_day_month();
        $hour_per_year++;
    }
    close $IN;
    $month--;
    die "Got $month months rather than 12" if $month != 12;
}

get_profile($profile);

my $p_txt = $en ? "load data points per hour  " : "Last-Datenpunkte pro Stunde";
my $b_txt = $en ? "load portion 9 AM - 3 PM   " : "Lastanteil 9 - 15 Uhr MEZ  ";
my $n_txt = $en ? "load portion 6 PM - 6 AM   " : "Lastanteil 18 - 6 Uhr MEZ  ";
my $t_txt = $en ? "total load acc. to profile " : "Verbrauch gemäß Lastprofil ";
my $m_txt = $en ? "maximal load               " : "Maximallast                ";
my $on    = $en ? "on" : "am";
print "$p_txt =   ".round($sum_items / $hour_per_year)."\n";
print "$b_txt =   ".percent($load_bright_sum / $load_sum)." %\n";
print "$n_txt =   ".percent($load_night_sum  / $load_sum)." %\n";
print "$t_txt =".kWh($load_sum)."\n";
print "$m_txt =".W($load_max)." $on $load_max_time\n";

################################################################################
# PV usage simulation

my $nominal_power_deflt = 0; # PVGIS default: 1 kWp
my $PV_gross_max = 0;
my $PV_gross_max_tm;
my @PV_gross_out;
my $PV_gross_out_sum = 0;

my $sys_efficiency_deflt = 0; # PVGIS default: 0.86
my @PV_net_out;
my $PV_net_out_sum = 0;
my $PV_net_bright_sum = 0;

my $load_scale = $simulated_load ? 1000 * $simulated_load / $load_sum : 1;
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

my $years = 0;
sub timeseries {
    my $file = shift;
    open(my $IN, '<', $file) or die "Could not open PV data file $file $!\n";

    # my $sum_needed = 0;
    my $power_rate;
    my $power_provided = 0; # PVGIS default: none
    my ($months, $hours) = (0, 0, 0);
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
        next if m/^20\d\d0229:/; # skip data of Feb 29th (in leap year)

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
        $years++ if m/^20..0101:00/;
        $months++ if m/^20....01:00/;

        die "Missing power data in $file line $_"
            unless m/^(\d\d\d\d)(\d\d)(\d\d):(\d\d)(\d\d),([\d\.]+)/;
        my ($year, $month, $day, $hour, $minute, $power) =
            ($1, $2, $3, $4, $5, $6 * $nominal_power
             / $nominal_power_deflt / $sys_efficiency_deflt);
        if ($power > $PV_gross_max) {
            $PV_gross_max = $power;
            $PV_gross_max_tm="$year-".time_string($month, $day, $hour, $minute);
        }
        $PV_gross_out[$month][$day][$hour] += $power;
        $PV_gross_out_sum += $power;

        $PV_net_loss[$month][$day][$hour] = 0;
        my $PV_loss = 0;
        if ($PV_limit && $power > $PV_limit) {
            $PV_loss = ($power - $PV_limit) * $sys_efficiency;
            $PV_net_loss[$month][$day][$hour] += $PV_loss;
            $PV_net_loss_sum += $PV_loss;
            $PV_net_loss_hours++;
            $power = $PV_limit;
            # print "$year-".time_string($month, $day, $hour, $minute).
            #"\tPV=".round($power)."\tlimit=".round($PV_limit).
            #"\tloss=".round($PV_net_loss_sum)."\t$_\n";
        }
        $power *= $sys_efficiency;
        $PV_net_out[$month][$day][$hour] += $power;
        $PV_net_out_sum += $power;
        $PV_net_bright_sum += $power
            if BRIGHT_START <= $hour && $hour < BRIGHT_END;

        # factor out $load_scale for optimizing the inner loop
        my $effective_PV_power = $power / $load_scale;
        $PV_loss /= $load_scale;
        # my $needed = 0;
        my $usages = 0;
        my $losses = 0;
        $items = $items_by_hour[$month][$day][$hour];
        for (my $item = 0; $item < $items; $item++) {
            my $point = $load_by_item[$month][$day][$hour][$item];
            # print "load[$month][$day][$hour][".sprintf("%4d", $item).
            #     "] = ".sprintf("%4d", $point)."\n" if $point <= 0;
            # $needed += $point;
            my $pv_used = min($effective_PV_power, $point); # PV own consumption
            $usages += $pv_used;
            if ($PV_loss != 0 && $point > $effective_PV_power) {
                my $loss = min($point - $effective_PV_power, $PV_loss);
                $losses += $loss;
                $PV_usage_loss_by_item[$month][$day][$hour][$item] =
                    round($loss * $load_scale) if $max;
                $PV_usage_loss_hours++; # will be normalized by $items
            } elsif ($max) {
                $PV_usage_loss_by_item[$month][$day][$hour][$item] = 0;
            }
            $PV_used_by_item[$month][$day][$hour][$item] =
                round($pv_used * $load_scale) if $max;
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

        $hours++;
    }
    close $IN;
    die "Got $months months rather than ".(12 * $years)
        if $months != $years * 12;
    die "Got $hours hours rather than ".(YearHours * $years)
        if $hours != $years * YearHours;

    $load_max *= $load_scale;
    $load_sum *= $load_scale;
    $PV_gross_out_sum /= $years;
    $PV_net_loss_sum /= $years;
    $PV_net_loss_hours /= $years;
    $PV_net_out_sum /= $years;
    $PV_net_bright_sum /= $years;
    $PV_used_sum /= $years;
    $PV_usage_loss_sum *= $load_scale / $years;
    $PV_usage_loss_hours /= ($sum_items / $hours * $years);
    # die "Inconsistent load calculation: sum = $sum vs. needed = $sum_needed"
    #     if round($sum) != round($sum_needed);
}

timeseries($PV_data);

################################################################################
# statistics output

my $nominal_txt      = $en ? "nominal PV power"     : "PV-Nominalleistung";
my $max_gross_txt    = $en ? "max gross PV power"   : "Bruttoleistung max.";
my $PV_limit_txt     = $en ? "cropping by inverter" : "Leistungsbegrenzung";
my $system_eff_txt   = $en ? "system efficiency"    : "System-Wirkungsgrad";
my $own_consumpt_txt = $en ? "own consumption ratio": "Eigenverbrauchsanteil";
my $load_cover_txt   = $en ? "load coverage ratio"  : "Eigendeckungsanteil";
my $PV_gross_txt     = $en ? "PV gross yield"       : "PV-Bruttoertrag";
my $PV_net_txt       = $en ? "PV net yield"         : "PV-Nettoertrag";
my $PV_net_lim_txt   = $en ? "PV net cropped"       :"PV-Nettoertrag limitiert";
my $PV_loss_lim_txt  = $en ? "PV loss by cropping"  : "PV-Abregelungsverlust";
my $load_txt         = $en ? "load by household"    : "Last durch Haushalt";
my $use_loss_txt     = $en ?"PV own consumption loss":"Eigenverbrauchsverlust";
my $by_limit         = $en ? "by cropping"          : "durch Limit";
my $use_wo_lim_txt   = $en ? "own consumption w/o cropping"
    : "Eigenverbrauch ohne Limit";
my $use_with_lim_txt = $en ? "own consumption".($PV_limit ? " w/ cropping" : "")
                           : "Eigenverbrauch" .($PV_limit ? " mit Limit"   :"");
my $each             = $en ? "each"   : "alle";
my $yearly_txt       = $en ? "yearly" : "jährlich";

my $own_usage     = percent($PV_used_sum       / $PV_net_out_sum);
my $load_coverage = percent($PV_used_sum       /       $load_sum);
my $yield_daytime = percent($PV_net_bright_sum / $PV_net_out_sum);

sub save_statistics {
    my $file = shift;
    return unless $file;
    my $res_txt = shift;
    my $max     = shift;
    my $hourly  = shift;
    my $daily   = shift;
    my $weekly  = shift;
    my $monthly = shift;
    open(my $OU, '>', $file) or die "Could not open statistics file $file $!\n";

    print $OU ", $nominal_txt in Wp, $max_gross_txt in W, "
        ."$PV_limit_txt in W, $system_eff_txt in %, "
        ."$own_consumpt_txt in %, $load_cover_txt in %, $PV_data\n";
    print $OU ", $nominal_power, ".round($PV_gross_max).", "
        .($PV_limit ? $PV_limit : 0).", ".percent($sys_efficiency)
        .", $own_usage, $load_coverage, $profile\n";
    print $OU "$yearly_txt, $PV_gross_txt, $PV_net_txt, $PV_loss_lim_txt, ".
        "$load_txt, $use_loss_txt $by_limit, $use_with_lim_txt, $each in kWh\n";
    print $OU ($en ? "sums" : "Summen"  ).", "
        .round($PV_gross_out_sum  / 1000).", "
        .round($PV_net_loss_sum   / 1000).", "
        .round($PV_net_out_sum    / 1000).", "
        .round($load_sum          / 1000).", "
        .round($PV_usage_loss_sum / 1000).", "
        .round($PV_used_sum       / 1000)."\n";

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
        $gross   += round($PV_gross_out [$month][$day][$hour] / $years);
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

my $en1 = $en ? " "  : "";
my $en2 = $en ? "  " : "";
my $de1 = $en ? ""   : " ";
my $de2 = $en ? ""   : "  ";
my $de3 = $en ? ""   : "   ";
my $with           = $en ? "with"                 : "bei";
my $during         = $en ? "during"               : "während";
my $lim            = $en ? "by cropping at"       : "durch Limitierung auf";
my $yield          = $en ? "yield portion"        : "Ertragsanteil";
my $daytime        = $en ? "9 AM - 3 PM "         : "9-15 Uhr MEZ";
my $of_yield       = $en ? "of yield"             : "des Ertrags";
my $of_consumption = $en ? "of consumption"       : "des Verbrauchs";
my $PV_loss_txt    = $en ? "PV loss             " : "PV-Abregelungsverlust";
$use_with_lim_txt  = "$use_with_lim_txt $en2         " unless $PV_limit;
print "\n";
print "$nominal_txt $en2         =" .W($nominal_power)."p\n";
print "$max_gross_txt $en1        =".W($PV_gross_max)." $on $PV_gross_max_tm\n";
print "$PV_gross_txt $en1            =".kWh($PV_gross_out_sum)."\n";
print "$PV_net_txt $en2             =" .kWh($PV_net_out_sum).
              " $with $system_eff_txt ".($sys_efficiency * 100)."%\n";
print "$PV_loss_txt $en1      ="       .kWh($PV_net_loss_sum)." $during "
    .round($PV_net_loss_hours)." h $lim $PV_limit W\n" if $PV_limit;
print "$yield $daytime  =   $yield_daytime %\n";

print "\n";
print "$load_txt $en2        =" .kWh($load_sum)."\n";
print "$use_with_lim_txt $de3=" .kWh($PV_used_sum)."\n";
print "$use_loss_txt $de1    =" .kWh($PV_usage_loss_sum)." $during "
    .round($PV_usage_loss_hours)." h $lim $PV_limit W\n" if $PV_limit;
print "$own_consumpt_txt         = $own_usage % $of_yield\n";
print "$load_cover_txt           = $load_coverage % $of_consumption\n";
