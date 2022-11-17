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

use constant N => 24 * 365; # Nur zum Debugging
use constant ITEMS => 74;   # Maximalzahl der Datensätze pro Stunde
use constant DELTA => 74/2; # Verschiebung der Datensätze innerhalb einer Stunde

my $profile = $ARGV[0]; # Profildatei
my $items = 0; # Zahl der verwendeten Messpunkte pro Minute
my $solar_time_sum = 0;
my $night_time_sum = 0;
my $sum = 0;

sub kWh { return sprintf("%5d kWh", int(shift()/1000)); }
sub W   { return sprintf("%5d W"  , int(shift())); }

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

my @load_by_hour;

sub synthesize_profile {
    my $file = shift;
    open(my $LOAD, '<', $file) or die "Could not open lood file $file $!\n";
    my ($hour_per_year, $minute, $item, $k) = (0, 0, 0, 0);
    # https://perlmaven.com/how-to-read-a-csv-file-using-perl
    while (my $line = <$LOAD>) {
        if (
            # Normierung auf 1500 kWh/a durch Ausdünnung:
            # nimm 17 von 60 minütlichen Datenpunkten
            0 && ($minute == 1 || $minute == 3 || $minute == 5 ||
                  $minute == 11 ||  $minute == 16 || $minute == 19 ||
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
            my $point = $sources[($k++ + DELTA) % ITEMS];
            $load_by_hour[$hour_per_year][$item] += $point;
            # print "$file " if $hour_per_year == N && $item == 0;
            # print "$item:$point," if $hour_per_year == N;
            $item++;

            # Fülle ggf. den Rest auf mit übrigen Daten aus letzer Minute
            while ($item >= 60 && ITEMS > 60 && $item < ITEMS) {
                $load_by_hour[$hour_per_year][$item++] +=
                    $sources[($k++ + DELTA)% ITEMS];
            }
        }
        $minute++;
        if ($minute == 60) {
            $minute = 0;
            $items = $item;
            $item = 0;
            #print "\n" if $hour_per_year == N;
            #last if $hour_per_year == N + 2;
            $hour_per_year++;
        }
    }
    close $LOAD;
}

sub save_profile {
    my $file = shift;
    open(my $OUT, '>', $file) or die "Could not open profile file $file $!\n";

    $hour = 0;
    for (my $hour_per_year = 0; $hour_per_year < 24 * 365; $hour_per_year++) {
        print $OUT sprintf("%02d", $month)."-".sprintf("%02d", $day).":"
            .sprintf("%02d", $hour).",";
        for (my $item = 0; $item < $items; $item++) {
            my $point = $load_by_hour[$hour_per_year][$item];
            # Nach grober Normierung auf 3000 kWh/a mit abwechselnd L1 und L5
            # mit geraden Minuten + Minute 1 + Minute 31:
            # Feine Normierung auf 3000 kWh/a: geringfügige Streckung
            # $point = int($point * 3000000000 / 2916807575);
            print $OUT "$point,";
            $sum += $point;
            $solar_time_sum += $point if 9 <= $hour && $hour < 15;
            $night_time_sum += $point if $hour < 6 || 18 <= $hour;
        }
        print $OUT "\n";

        $hour = 0 if ++$hour == 24;
        adjust_day_month();
    }
    close $OUT;
}

synthesize_profile("PL1.csv");
synthesize_profile("PL2.csv");
synthesize_profile("PL3.csv");

save_profile($profile);

$sum /= $items;
$solar_time_sum /= $items;
$night_time_sum /= $items;
print "Last-Datenpunkte pro Stunde = $items\n";
print "Lastanteil von 9-15 Uhr MEZ = ".int($solar_time_sum / $sum * 100)."%\n";
print "Lastanteil von 18-6 Uhr MEZ = ".int($night_time_sum / $sum * 100)."%\n";
print "Verbrauch gemäß Lastprofil  =".kWh($sum)."\n";
