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
use constant Normierung => 1;
use constant Datenlog => 0;

my $profile = $ARGV[0]; # Profildatei

sub round { return int(.5 + shift); }
sub percent { return sprintf("%2d"    , round(shift() *  100)); }

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

sub synthesize_profile {
    my $file = shift;
    open(my $LOAD, '<', $file) or die "Could not open lood file $file $!\n";
    my ($day_per_year, $hour_per_year, $minute, $item, $k) = (0, 0, 0, 0, 0);
    my ($year_before, $month_before, $day_before, $hour_before, $minute_before,
        $second_before, $index) = (0, 0, 0, 0, 0, 0, (365-32) * 24) if Datenlog;
    # https://perlmaven.com/how-to-read-a-csv-file-using-perl
    while (my $line = <$LOAD>) {
        if (Datenlog) {
            chomp $line;
            my @sources = split "," , $line;
            my $eof =  $#sources < 4;
            my $date = $eof ? "2022-11-30T00:00:00" : $sources[3];
            next unless $date =~ m/^202/;
            die "Cannot parse entry after $hour_per_year hours in $file"
                unless $date =~ m/^(202\d)-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d)/;
            my $value = $sources[4];
            my ($year, $month, $day, $hour, $minute, $second) =
                ($1, $2, $3, $4, $5, $6);
            # print "Processing $year-$month-$day ..\n" if $day != $day_before;

            print "Zero load in $file at $year-$month-$day".
                "T$hour:$minute:$second\n" unless $value || $eof;
            if ($hour != $hour_before || $eof) {
                $items += ($items_by_hour[$index] = $item);
                $item = 0;
                $hour_per_year++;
                $index++;
                last if $hour_per_year == YearHours;
                die "EOF after $hour_per_year hours in $file\n" if $eof;
                if (($hour_before + 1) % 24 != $hour) {
                    print "Gap in $file from ".
                        "$year_before-$month_before-$day_before".
                        "T$hour_before:$minute_before:$second_before to $year-".
                        "$month-$day"."T$hour:$minute:$second, load $value\n";
                    while (++$hour_before % 24 != $hour) {
                        $items += ($items_by_hour[$index] = 1);
                        $load_by_hour[$index++][$item] += $value; # reuse next value
                        $hour_per_year++;
                    }
                }
            }
            $index = 0 if $index == YearHours;
            $load_by_hour[$index][$item++] += $value;
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
            #print "\n" if $hour_per_year == N;
            #last if $hour_per_year == N + 2;
            $hour_per_year++;
            $day_per_year++ if $hour_per_year % 24 == 0;
        }
    }
    close $LOAD;
    $items /= $hour_per_year;
}

my $sum = 0;
my $solar_time_sum = 0;
my $night_time_sum = 0;
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
    for (my $hour_per_year = 0; $hour_per_year < YearHours; $hour_per_year++) {
        print $OUT "# Lastprofil ".L_days."\n"
            if Normierung && L_subst && $hour_per_year == Subst_beg;
        print $OUT "# Lastprofil ".L_subst."\n"
            if Normierung && L_subst && $hour_per_year == Subst_end;
        print $OUT sprintf("%02d", $month)."-".sprintf("%02d", $day).":"
            .sprintf("%02d", $hour).",";

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

        my $load = 0;
        my $n =  $items_by_hour[$hour_per_year];
        for (my $item = 0; $item < $n; $item++) {
            #print sprintf("%02d", $month)."-".sprintf("%02d", $day).":"
            #    .sprintf("%02d", $hour).":".sprintf("%02d", $item).",";
            #for (my $i = 0; $i < 74; $i++) {
            #    print "".round($load_by_elem[$hour_per_year][$item][$i]);
            #    print "".($i < 73 ? "," : "\n");
            #}
            my $point = $load_by_hour[$hour_per_year][$item];

            $point = round($point * $factor) if Normierung;
            print "Zero load at hour $hour_per_year item $item\n" unless $point;

            print $OUT "$point,";
            $load += $point;
        }
        print $OUT "\n";
        $load /= $n;
        $sum            += $load;
        $solar_time_sum += $load if 9 <= $hour && $hour < 15;
        $night_time_sum += $load if $hour < 6 || 18 <= $hour;

        $hour = 0 if ++$hour == 24;
        adjust_day_month();
        $day_per_year++ if $hour == 0;
    }
    close $OUT;
}

#synthesize_profile("PL1.csv");
#synthesize_profile("PL2.csv");
#synthesize_profile("PL3.csv");
synthesize_profile("PL.csv");

save_profile($profile);

print "Last-Datenpunkte pro Stunde = $items\n";
print "Lastanteil von 9-15 Uhr MEZ = ".percent($solar_time_sum / $sum)." %\n";
print "Lastanteil von 18-6 Uhr MEZ = ".percent($night_time_sum / $sum)." %\n";
print "Verbrauch gemäß Lastprofil  = ".($sum / 1000)." kWh\n";
