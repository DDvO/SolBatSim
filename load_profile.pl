#!/usr/bin/perl
#
################################################################################
# Lastprofil-Synthetisierung auf Minuntenbasis
# Nutzung: Lastprofil.pl Lastprofil-Ausgabe-Datei.csv
# (c) 2022 David von Oheimb - License: MIT - Version 2.3
################################################################################

use strict;
use warnings;

# https://solar.htw-berlin.de/elektrische-lastprofile-fuer-wohngebaeude/
# 74 Lastprofile mit folgenden Jahressummen in kWh:
# 3238  4500  6619  2663  3196  1398  2937  5175  8001  3369
# 3887  4507  4892  3259  3402  3196  5924  5738  5489  4946
# 6903  4044  7497  1847  5149  4211  4554  4360  3683  4695
# 5010  5318  3429  3081  6689  4147  3466  5556  5169  3774
# 5520  6041  2296  2629  4948  8635  4227  4641  3786  3962
# 2391  4363  3390  6089  4784  5560  4287  4980  6535  4494
# 4569  5220  6944  5953  6761  3615  4958  5229  6944  6224
# 4323  4837  4342  4266
# Durchschnitt: 4044

use constant L1  =>  6; # 54% 1398 kleinster Jahresverbrauch
use constant L2  => 24; # 57% 1847 nahe 2000
use constant L30 =>  7; # 73% 2937 nahe 3000
use constant L31 => 34; # 66% 3081 nahe 3000
use constant L32 =>  5; # 66% 3196 nahe 3000, gleich L33
use constant L33 => 16; # 72% 3196 nahe 3000, gleich L32
use constant L34 =>  1; # 72% 3238 nahe 3000, ähnlich L32
use constant L4  => 22; # 66% 4044 nahe 4000
use constant L45 =>  2; # 77% 4500, fast gleich L46
use constant L46 => 12; # 70% 4507, fast gleich L45
use constant L5  => 31; # 71% 5010 nahe 5000, repräsentativ bzgl. Jahreszeit
use constant L0  => 17; # 72% 5924 nahe 6000, repräsentativ bzgl. Tageszeit
use constant L6  => 42; # 72% 6041 nahe 6000
use constant L8  =>  9; # 65% 8001 nahe 8000
use constant L9  => 46; # 63% 8635 größter Jahresverbrauch

use constant L_days    =>   17; # 5924, repräsentativ bzgl. Tageszeit
use constant L_subst   =>   64; # 5953, ähnlicher Jahresverbrauch
#use constant L_subst   =>   31;
use constant Subst_beg => 6691; # Startstunde Urlaubsloch von L_days: 10-06:19
use constant Subst_end => 7066; # Endstunde   Urlaubsloch von L_days: 10-22:10

use constant ITEMS => 60; # 74; # Ziel-Anzahl der Datensätze pro Stunde
use constant DELTA => 74/2; # Verschiebung der Datensätze innerhalb einer Stunde

use constant N => 24 * 365; # Nur zum Debugging
use constant YearHours => 24 * 365;

# all hours according to local time without switching for daylight saving
use constant BRIGHT_START  =>  9; # start bright sunshine time
use constant BRIGHT_END    => 15; # end   bright sunshine time
use constant EVENING_START => 18; # start evening
use constant EVENING_END   => 24; # end   evening
use constant NIGHT_START   =>  0; # start night (basic load)
use constant NIGHT_END     =>  6; # end   night (basic load)

use constant HOUR_AVERAGE => 0;
use constant TIMEZONE_DELTA => +1; # for local timezone MEZ

use constant Normierung => 0;
use constant Datenlog => 1;

my $profile = $ARGV[0]; # Profildatei

# deliberately not using any extra packages like Math
sub min { return $_[0] < $_[1] ? $_[0] : $_[1]; }
sub max { return $_[0] > $_[1] ? $_[0] : $_[1]; }
sub round { return int(.5 + shift); }

my $en = 1;
sub date_string {
    return sprintf("%02d", shift)."-".sprintf("%02d", shift).
        shift.sprintf("%02d", shift);
}
sub time_string {
    return date_string(shift, shift, "T", shift).sprintf(":%02d", int(shift));
}
sub index_string {
    return date_string(shift, shift, " at ", shift).
        sprintf("h, index %4d", shift);
}

sub percent { return sprintf("%2d", round(shift() *  100)); }

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

my $items = 0; # total number of load items
my @items_by_hour;
#my @load_by_elem;
my @load_by_hour;
my @load_by_day;
my @avg_load_by_day;

sub parse_PV {
    my $PV = shift;
    my $PV_file = shift;
    my $line;
    while (($line = <$PV>) =~ m/^#/) {}
    die "Cannot parse '$line' in $PV_file"
        unless $line =~ m/^,,0,\d+-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d)Z,(\d+)/;
    my ($month, $day, $hour, $min, $sec) = ($1, $2 - 1, $3, $4, $5);
    $day += 31 if --$month > 0;
    $day += 28 if --$month > 0;
    $day += 31 if --$month > 0;
    $day += 30 if --$month > 0;
    $day += 31 if --$month > 0;
    $day += 30 if --$month > 0;
    $day += 31 if --$month > 0;
    $day += 31 if --$month > 0;
    $day += 30 if --$month > 0;
    $day += 31 if --$month > 0;
    $day += 30 if --$month > 0;
    $day += 31 if --$month > 0;
    return ((($day * 24 + $hour + TIMEZONE_DELTA) * 60 + $min) * 60 + $sec, $6);
}

sub after_hours {
    my $hour = shift;
    return "after ".(($hour - TIMEZONE_DELTA) % YearHours)." hours";
}

sub synthesize_profile {
    my $file = shift;
    open(my $LOAD, '<', $file) or die "Could not open lood file $file $!\n";
    my $PV_file = "PV_Power.csv";
    my $PV;
    open($PV, '<', $PV_file) or die "Could not open lood file $PV_file $!\n"
        if Datenlog && $file eq "PL2.csv";
    my ($PV_time, $PV_power) = parse_PV($PV, $PV_file) if $PV;

    my ($day_per_year, $hour_per_year, $minute, $item, $k) =
        (0, TIMEZONE_DELTA, 0, 0, 0);
    my ($year_before, $month_before, $day_before, $hour_before, $minute_before,
        $second_before) = (0, 1, 1, 0, 0, 0, 0) if Datenlog;
    my $warned_just_before = 0;
    # https://perlmaven.com/how-to-read-a-csv-file-using-perl
    while (my $line = <$LOAD>) {
        if (Datenlog) {
            chomp $line;
            next if $line =~ m/^#/;
            my @sources = split "," , $line;
            if ($#sources < 4) { # EOF
                $items += ($items_by_hour[$hour_per_year++] = $item);
                $hour_per_year = 0 if $hour_per_year == YearHours;

                print "EOF ".after_hours($hour_per_year)." in $file\n"
                    if $hour_per_year != TIMEZONE_DELTA;
                last;
            }

            my $date = $sources[3];
            die "Cannot parse '$line' ".after_hours($hour_per_year)." in $file"
                unless $date =~
                m/^(202\d)-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d)(.\d+)?Z$/;
            my ($year, $month, $day, $hour, $minute, $second) =
                ($1, $2, $3, $4, $5, $6);
            my $value = $sources[4];
            # print "Processing $year-$month-$day ..\n" if $day != $day_before;

            if ($PV) {
                my $time = (($hour_per_year * 60) + $minute) * 60 + $second;
                while ($time > $PV_time) {
                    ($PV_time, $PV_power) = parse_PV($PV, $PV_file);
                }
                $value += $PV_power;
            }
            if ($value <= 0 && 0) {
                print "In $file on $year-$month-$day"."T$hour:$minute:$second".
                    ", load = ".sprintf("%4d", $value)."\n"
                    unless $warned_just_before;
                $warned_just_before = 1;
            } else {
                $warned_just_before = 0;
            };
            if ($day != $day_before || $hour != $hour_before) {
                $items += ($items_by_hour[$hour_per_year++] = $item);
                $hour_per_year = 0 if $hour_per_year == YearHours;
                $item = 0;
                my $hours_before = $day_before * 24 + $hour_before;
                my $hours = $day * 24 + $hour;
                if ($hours - $hours_before > 1) {
                    print "Gap in $file from ".
                        "$year_before-$month_before-$day_before".
                        "T$hour_before:$minute_before:$second_before to $year-".
                        "$month-$day"."T$hour:$minute:$second, load $value\n";
                    die "Cannot handle gap over month limit"
                        if $year_before > 0 && $year != $year_before
                        || $month_before > 0 && $month != $month_before;
                    while (++$hours_before != $hours) {
                        $items += ($items_by_hour[$hour_per_year] = 1);
                        $load_by_hour[$hour_per_year++][$item] +=
                            $value; # reuse next value
                        $hour_per_year = 0 if $hour_per_year == YearHours;
                    }
                }
            }
            $load_by_hour[$hour_per_year][$item++] += $value;
            ($year_before, $month_before, $day_before, $hour_before,
             $minute_before, $second_before) = ($year, $month, $day,
                                                $hour, $minute, $second);
            next;
        }
        if (
            # Normierung auf 1500 kWh/a durch Ausdünnung:
            # nimm 17 von 60 minütlichen Datenpunkten
            0 && ($minute == 1 || $minute == 3 || $minute == 5 ||
                  $minute == 11 || $minute == 16 || $minute == 19 ||
                  $minute == 24 || $minute == 27 || $minute == 32 ||
                  $minute == 35 || $minute == 40 || $minute == 43 ||
                  $minute == 48 || $minute == 51 || $minute == 53 ||
                  $minute == 57 || $minute == 59)
            ||
            # Normierung auf 3003 kWh/a durch Ausdünnung:
            # nimm 33 von 60 minütlichen Datenpunkten
            0 && ($minute % 2 == 0 || $minute ==  1 ||
                  $minute == 3 || $minute == 59)
            ||
            1
            ) {
            chomp $line;
            my @sources = split "," , $line;
            my $n = $#sources;
            shift @sources; # ignore date & time info

            # my $point = $sources[L34-1];
            # Bei Normierung auf 1500 und 3003:
            # my $point = $minute % 2 == 0 ? $sources[L0-1] : $sources[L5-1];
            # Bei Normierung auf 3000:
            # my $point = $k++ % 2 == 0 ? $sources[L0-1] : $sources[L5-1];
            # Bei Mischung der beiden 3196:
            # my $point = $k++ % 2 == 0 ? $sources[L32-1] : $sources[L33-1];
            # Bei Mischung der beiden 3196 mit 5010 und 5924:
            # my $point = ++$k % 4 == 3 ? $sources[L32-1] :
            #     $k % 4 == 2 ? $sources[L33-1] :
            #     $k % 4 == 1 ? $sources[L5-1] : $sources[L0-1];
            # my $point = $sources[($k++ + DELTA) % ITEMS];
            my $point = $sources[L_days - 1];
            $point = $sources[L_subst - 1] if L_subst
                && Subst_beg <= $hour_per_year && $hour_per_year < Subst_end;
            $load_by_hour[$hour_per_year][$item] = $point;

            my $sum = 0;
            for (my $i = 0; $i < $n; $i++) {
            #   $load_by_elem[$hour_per_year][$item][$i] += $sources[$i];
                $sum += $sources[$i];
            }
                $load_by_day[$day_per_year] += $point;
            $avg_load_by_day[$day_per_year] += $sum / $n;

            # print "$file " if $hour_per_year == N && $item == 0;
            # print "$item:$point," if $hour_per_year == N;
            $item++;

            # Fülle ggf. den Rest auf mit übrigen Daten aus letzter Minute
            while ($item >= 60 && ITEMS > 60 && $item < ITEMS) {
                $load_by_hour[$hour_per_year][$item++] =
                    $sources[($k++ + DELTA)% ITEMS];
            }
            $items += ($items_by_hour[$hour_per_year] = $item);
        }
        $minute++;
        if ($minute == 60) {
            $minute = 0;
            $item = 0;
            # print "\n" if $hour_per_year == N;
            # last if $hour_per_year == N + 2;
            $hour_per_year++;
            $day_per_year++ if $hour_per_year % 24 == 0;
        }
    }
    close $LOAD;
    close $PV if $PV;
    $items /= YearHours;
}

my ($sum, $bright_sum, $evening_sum, $night_sum) = (0, 0, 0, 0);
my $sum_negative = 0;

sub save_profile {
    my $file = shift;
    open(my $OUT, '>', $file) or die "Could not open profile file $file $!\n";
    if (Normierung) {
        print $OUT "# Lastprofil ".L_days;
        print $OUT ", aber im Oktober zwischen Stunde ".Subst_beg.
            " und ".Subst_end." Lastprofil ".L_subst if L_subst;
        print $OUT ", tageweise normiert ".
            "auf durchschnittliche tgl. Last aller Haushalte\n";
        print $OUT "# Lastprofil ".L_days."\n" if L_subst;
    }

    $hour = 0;
    my $day_per_year = 0;
    my $warned_just_before = 0;
    for (my $hour_per_year = 0; $hour_per_year < YearHours; $hour_per_year++) {
        print $OUT "# Lastprofil ".L_days."\n"
            if Normierung && L_subst && $hour_per_year == Subst_beg;
        print $OUT "# Lastprofil ".L_subst."\n"
            if Normierung && L_subst && $hour_per_year == Subst_end;
        print $OUT date_string($month, $day, ":", $hour).",";
        my $n = $items_by_hour[$hour_per_year];

        # Vorbereitung Normierung
        my $factor = 1 if Normierung;
        if (Normierung) {
            # Normierung auf Durchschnitt aller Haushalte pro Tag:
            $factor = $avg_load_by_day[$day_per_year]
                / $load_by_day[$day_per_year]
                if $load_by_day[$day_per_year] != 0;
            # auf 3000 kWh/a:
            # $factor *= 3000000 / 4044002;
            # $factor *= 3000000 / 4685054;

            # Nach grober Normierung auf 3000 kWh/a mit abwechselnd L1 und L5
            # mit geraden Minuten + Minute 1 + Minute 31:
            # Feine Normierung auf 3000 kWh/a: geringfügige Streckung
            # $factor = 3000000000 / 2916807575;
        }

        # try eliminating points <= 0 without changing sum in current hour
        my ($before, $todo) = (0, 0);
        my ($load, $negative_load) = (0, 0);
        if (!HOUR_AVERAGE) {
            # eliminate as far as possible using point before and future ones
            for (my $item = 0; $item < $n; $item++) {
                my $point = $load_by_hour[$hour_per_year][$item];
                $point *= $factor if Normierung;
                $point = round($point);
                if ($point <= 0) {
                    print "In $file on ".
                        index_string($month, $day, $hour, $item).
                        ", load = ".sprintf("%4d", $point)."\n"
                        unless $warned_just_before;
                    $negative_load -= $point;
                    $warned_just_before = 1;
                } else {
                    $warned_just_before = 0;
                }
                $load += $point;
                if ($todo + $point <= 0 && $before > 1) {
                    my $delta = min($before - 1, -($todo + $point));
                    $load_by_hour[$hour_per_year][$item - 1] -= $delta;
                    $todo += $delta;
                    ($todo, $point) = (0, $point + $todo) if $todo > 0;
                }
                if ($point <= 0 && $item < $n - 1) {
                    $todo += $point - 1;
                    $load_by_hour[$hour_per_year][$item] = $point = 1;
                }
                if ($todo < 0 && $item == $n - 1) {
                    $load_by_hour[$hour_per_year][$item] = ($point += $todo);
                }
                $before = $point;
            }
            $todo = $before;
        }
        if (!HOUR_AVERAGE && $todo <= 0 && $load + $todo >= 1) {
            # eliminate last point <= 0 using points from beginning of hour
            $todo--;
            $load_by_hour[$hour_per_year][$n - 1] = 1; # same as -= $todo;
            for (my $item = 0; $todo < 0 && $item < $n - 1; $item++) {
                my $point = $load_by_hour[$hour_per_year][$item];
                if ($point > 1) {
                    my $delta = min($point - 1, -$todo);
                    $load_by_hour[$hour_per_year][$item] -= $delta;
                    $todo += $delta;
                }
            }
        }
        $load = 0 if !HOUR_AVERAGE;
        for (my $item = 0; $item < $n; $item++) {
            # print index_string($month, $day, $hour, $item).",";
            # for (my $i = 0; $i < 74; $i++) {
            #     print "".round($load_by_elem[$hour_per_year][$item][$i]);
            #     print "".($i < 73 ? "," : "\n");
            # }
            my $point = $load_by_hour[$hour_per_year][$item];

            $point *= $factor if Normierung && !HOUR_AVERAGE;
            if ($point <= 0) {
                print "In $file on ".index_string($month, $day, $hour, $item).
                    ", load = ".sprintf("%4d", $point)."\n"
            };

            if (!HOUR_AVERAGE) {
                $point = round($point);
                print $OUT "$point,";
            }
            $load += $point if !HOUR_AVERAGE;
        }
        $sum_negative += ($negative_load / $n);
        $load /= $n;
        print $OUT round($load) if HOUR_AVERAGE;
        print $OUT "\n";
        $sum         += $load;
        $bright_sum  += $load if  BRIGHT_START <= $hour && $hour <  BRIGHT_END;
        $evening_sum += $load if EVENING_START <= $hour && $hour < EVENING_END;
        $night_sum   += $load if   NIGHT_START <= $hour && $hour <   NIGHT_END;

        $hour = 0 if ++$hour == 24;
        adjust_day_month();
        $day_per_year++ if $hour == 0;
    }
    close $OUT;
}

#synthesize_profile("PV_Power.csv");
synthesize_profile("PL1.csv");
synthesize_profile("PL2.csv");
synthesize_profile("PL3.csv");
#synthesize_profile("PL.csv");

save_profile($profile);

print "Last-Datenpunkte pro Stunde = ".round($items)."\n";
print "Lastanteil  9 - 15 Uhr MEZ  = ".percent($bright_sum / $sum)." %\n";
print "Lastanteil 18 - 24 Uhr MEZ  = ".percent($evening_sum / $sum)." %\n";
print "Lastanteil  0 -  6 Uhr MEZ  = ".percent($night_sum / $sum)." %\n";
print "Summe negativer Verbrauch   = ".($sum_negative / 1000)." kWh\n";
print "Verbrauch gemäß Lastprofil  = ".($sum / 1000)." kWh\n";
