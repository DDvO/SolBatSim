#!/usr/bin/perl

# Collect status data reported each second by Shelly (Pro) 3EM energy meter.
# Uses the http://<addr>/status endpoint (yet so far not the MQTT interface).
# Outputs data in the following files, each of which is optional:
# * <base_name><energy_name>_<year>.csv energy imported and exported per hour
# * <base_name><load_min>_<year>.csv    total load per minute, one line per hour
# * <base_name><load_sec>_<date>.csv    total load per second, one line per hour
# * <base_name><status_name>_<date>.csv status of the three phases per second
# * <base_name><log_name>_<year>.txt    info on the data collection per event
#
# CLI options, each of which may be a value or "" indicating none/default:
# <base_name> <energy_name> <load_min> <load_sec> <status_name> <log_name>
# <time_zone> <3em_addr> <3em_username> <3em_password>
#
# Alternatively to providing options at the CLI, they may also be given
# via environment variables, which is advisable for sensitive passwords.
#
# By default, the time zone is CET (Central European Time)
# without the typical adaptations to DST (Daylight Saving Time) twice a year.
# This prevents misalignment and confusion on the interpretation of timed data
# and makes sure that the length of the daily output is the same for all days.
#
# (c) 2023 David von Oheimb - License: MIT

use strict;
use warnings;
#use IO::Null; # using IO::Null->new, print gives: print() on unopened filehandle GLOB
use IO::Handle; # for flush

my $i = 0;
my $out_prefix   = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_BASENAME}; # e.g., ~/3EM_
my $out_energy   = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_ENERGY};   # per hour
my $out_load_min = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_LOAD_MIN}; # per minute
my $out_load_sec = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_LOAD_SEC}; # per second
my $out_stat     = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_STATUS};   # per second
my $out_log      = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_LOG};      # per event

# time zone for output e.g., "local"
my $tz = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_TZ} || "CET";
my $date_format        = "%Y-%m-%d";
my $date_format_out    = $ENV{Shelly_3EM_OUT_DATE_FORMAT} || $date_format;
my $date_time_sep      = " "; # "T";
my $date_time_sep_out  = $ENV{Shelly_3EM_OUT_DATE_TIME_SEP} || $date_time_sep;
my $time_format        = "%H:%M:%S";
my $time_format_out    = $ENV{Shelly_3EM_OUT_TIME_FORMAT} || $time_format;

my $addr = $ARGV[$i++] || $ENV{Shelly_3EM_ADDR} || "3em"; # e.g. 192.168.178.123
my $user = $ARGV[$i++] || $ENV{Shelly_3EM_USER}; # HTTP user name, if needed
my $pass = $ARGV[$i++] || $ENV{Shelly_3EM_PASS}; # HTTP password, if needed
my $url  = "http://$addr/status";

my $debug = $ENV{Shelly_3EM_DEBUG} // 0;

die "missing CLI argument or env. variable 'Shelly_3EM_ADDR'"   unless $addr;
die "missing CLI argument or env. variable 'Shelly_3EM_OUT_TZ'" unless $tz;

sub round { return int(($_[0] < 0 ? -.5 : .5) + $_[0]); }

# local system time
# https://stackoverflow.com/questions/60107110/perl-strftime-localtime-minus-12-hours
use DateTime;
sub date_time {
    return ($_[0]->strftime($date_format),
            $_[0]->strftime($time_format));
}
sub date_time_out {
    return ($_[0]->strftime($date_format_out),
            $_[0]->strftime($time_format_out));
}
my $start = DateTime->now(time_zone => $tz);
my ($date, $time) = date_time($start);
# $start->add(hours => 1);
# my $time_format_hour = "%H:00:00";
# my $end_time = $start->strftime($time_format_hour);
# $start->subtract(hours => 1);
# $end_time = "24:00:00" if $end_time eq "00:00:00";

sub time_epoch { return DateTime->from_epoch(epoch => $_[0], time_zone => $tz);}
use constant SECONDS_PER_MINUTE => 60;
use constant SECONDS_PER_HOUR => 60 * SECONDS_PER_MINUTE;
# max seconds per day even for days with daylight saving time (DST) adaptation:
# use constant MAX_SECONDS => (24 + 1) * SECONDS_PER_HOUR;
my ($count_seconds, $count_gaps) = (0, 0);

sub out_name {
    my ($name, $period, $ext) = @_;
    # https://stackoverflow.com/questions/1376607/how-can-i-suppress-stdout-temporarily-in-a-perl-program
    my $none = File::Spec->devnull(); # sink for no-op output
    return $name ? $out_prefix.$name."_".$period.$ext : $none;
}

my ($energy, $EO);
my ($load_min, $LM);
my ($load_sec, $LS);
my ($status, $SO);
my ($log, $LOG);
# preliminary log:
sub log_name { return out_name($out_log, $_[0], ".txt"); }
$log = log_name($start->strftime("%Y"));
open($LOG,'>>', $log) || die "cannot open '$log' for appending: $!";

sub log_msg {
    my $msg = "$date $time: $_[0]\n";
    print $msg;
    print $LOG $msg if defined $LOG;
}
sub log_warn { log_msg("WARNING: $_[0]"); }

sub do_after_day {
    print $LS "\n" if defined $LS;
    close $LS if defined $LS;
    close $SO if defined $SO;
}

sub do_after_year {
    print $LM "\n" if defined $LM;
    close $LM  if defined $LM;
    close $EO  if defined $EO;
    close $LOG if defined $LOG;
}

sub cleanup() {
    do_after_day();
    log_msg("end after $count_seconds seconds, $count_gaps gaps");
    do_after_year();
}

# https://stackoverflow.com/questions/77302036/atexit3-in-perl-end-or-sig-die
$SIG{'__DIE__'} = sub {
    my $msg = $_[0];
    log_msg("aborting on fatal error: $msg");
    cleanup();
    die $msg; # actually die
};
$SIG{'INT'} = $SIG{'TERM'} = sub {
    my $sig = $_[0];
    log_msg("aborting on signal $sig");
    cleanup();
    $SIG{$sig} = 'DEFAULT';
    kill($sig, $$); # execute default action
};

# https://stackoverflow.com/questions/19842400/perl-http-post-authentication
sub http_get {
    use HTTP::Request::Common;
    require LWP::UserAgent;
    my ($url, $user, $pass) = @_;
    my $ua = new LWP::UserAgent;

    $ua->timeout(1);
    my $request = GET $url;
    $request->authorization_basic($user, $pass) if $user;
    my $response = $ua->request($request);
    return $response->content;
}

my $last_valid_unixtime = 0;
sub get_line {
    my ($check_time) = @_;

  retry:
    my $status_json = http_get($url, $user, $pass);
    if ($status_json =~ m/(Network is unreachable|No route to host|Can't connect|Server closed connection|Connection reset by peer|(Connection|Operation) timed out|read timeout)/) {
        log_warn($1);
        # 'timed out' can be misleading: also occurs on router down without having waited
        sleep(5) unless $1 =~ m/Connection reset by peer/;
        goto retry;
    }
    unless ($status_json =~ /\"time\":\"([\d:]*)\",\"unixtime\":(\d+),.*\"emeters\":\[\{\"power\":([\-\d\.]+),\"pf\":([\-\d\.]+),\"current\":([\-\d\.]+),\"voltage\":([\-\d\.]+),\"is_valid\":true,\"total\":([\d\.]+),\"total_returned\":([\d\.]+)}\,\{\"power\":([\-\d\.]+),\"pf\":([\-\d\.]+),\"current\":([\-\d\.]+),\"voltage\":([\-\d\.]+),\"is_valid\":true,\"total\":([\d\.]+),\"total_returned\":([\d\.]+)\},\{\"power\":([\-\d\.]+),\"pf\":([\-\d\.]+),\"current\":([\-\d\.]+),\"voltage\":([\-\d\.]+),\"is_valid\":true,\"total\":([\d\.]+),\"total_returned\":([\d\.]+)\}\],\"total_power\":([\-\d\.]+),.*,\"uptime\":(\d+)/) {
        if ($status_json =~ /ERROR:\s?([\s0-9A-Za-z]*)/i) {
            # e.g.: The requested URL could not be retrieved
            log_warn("skipping error response: $1"); # e.g., by Squid
        } else {
            my $shown = substr($status_json, 0, 1100); # typically ~1020 chars
            log_warn("error parsing 3EM status response '$shown'");
        }
        sleep(1);
        goto retry;
    }

    my ($hour, $unixtime,
        $powerA, $pfA, $currentA, $voltageA, $totalA, $total_returnedA,
        $powerB, $pfB, $currentB, $voltageB, $totalB, $total_returnedB,
        $powerC, $pfC, $currentC, $voltageC, $totalC, $total_returnedC,
        $total_power, $uptime)
        = ($1, $2 + 0,
           $3, $4, $5, $6, $7, $8,
           $9, $10, $11, $12, $13, $14,
           $15, $16, $17, $18, $19, $20,
           $21, $22);
    my $dataA = "$powerA,$pfA,$currentA,$voltageA,$totalA,$total_returnedA";
    my $dataB = "$powerB,$pfB,$currentB,$voltageB,$totalB,$total_returnedB";
    my $dataC = "$powerC,$pfC,$currentC,$voltageC,$totalC,$total_returnedC";
    if ($unixtime) {
        $last_valid_unixtime = $unixtime;
    } else {
        if ($last_valid_unixtime) {
            log_warn("approximating missing 3EM status unixtime from last valid"
                     ." one $last_valid_unixtime + uptime $uptime");
            $unixtime = $last_valid_unixtime + $uptime;
        } else {
            log_warn("missing 3EM status unixtime, discarding '$status_json'");
            sleep(1);
            goto retry;
        }
    }
    my ($date_3em, $time_3em) = date_time(time_epoch($unixtime));
    my $data = "$dataA,$dataB,$dataC";
    print "($time, $hour, $unixtime, $time_3em, $data)\n" if $debug;

    if ($check_time) {
        my $time_hour = substr($time, 0, 5);
        log_warn("3EM status time '$hour' does not equal '$time_hour'")
            unless $hour eq $time_hour;
        log_warn("3EM status unixtime '$date_3em"."$date_time_sep$time_3em' ".
                 "does not closely match '$date"."$date_time_sep$time'")
            unless abs($unixtime - $start->epoch) <= 1
            # 3 seconds diff can happen easily
    }
    my $power = $powerA + $powerB + $powerC;
    log_warn("inconsistent total_power = $total_power ".
             "vs. $powerA + $powerB + $powerC")
        unless abs($power - $total_power) <= 0.1;
    return ($unixtime, $power, $data);
}

my ($energy_imported_this_hour, $energy_exported_this_hour) = (0, 0);
my $power_sum_minute = 0;
my ($prev_power, $prev_timestamp) = (0, -1);
# my $prev = "";

sub do_before_year {
    my ($date_3em, $first) = @_;
    die "error matching time" unless $date_3em =~ m/^(\d+)-(\d\d-\d\d)$/;
    my $year_3em = $1;
    return unless $first || $2 eq "01-01";

    $load_min = out_name($out_load_min, $year_3em, ".csv");
    $energy = out_name($out_energy, $year_3em, ".csv");
    $log    = log_name($year_3em);
    open($LOG, '>>', $log  ) || die "cannot open '$log' for appending: $!";
    open($EO, '>>', $energy) || die "cannot open '$energy' for appending: $!";
    open($LM,'>>',$load_min) || die "cannot open '$load_min' for appending: $!";

    $LOG->autoflush; # immediately show each line reporting an event
    $EO ->autoflush; # immediately show each line reporting energy per hour
    $LM ->autoflush; # immediately show each load per minute
    log_msg("start - will connect to $url") if $first;
    # on empty energy output CSV file, add header:
    print $EO "time [$tz],imported [Wh],exported [Wh]\n" if -z $energy;
    # no header for load output CSV file
}

sub do_before_day {
    my ($date_3em, $time_3em, $date_3em_out, $time_3em_out, $first) = @_;
    return unless $first || $time_3em eq "00:00:00";

    do_after_day();
    do_before_year($date_3em, $first);
    $load_sec = out_name($out_load_sec, $date_3em_out, ".csv");
    $status = out_name($out_stat, $date_3em_out, ".csv");
    open($LS,'>>',$load_sec) || die "cannot open '$load_sec' for appending: $!";
    open($SO, '>>', $status) || die "cannot open '$status' for appending: $!";

    print $SO "time [$tz],total_power,".
   "powerA [W],pfA,currentA [A],voltageA [V],totalA [Wh],total_returnedA [Wh],".
   "powerB [W],pfB,currentB [A],voltageB [V],totalB [Wh],total_returnedB [Wh],".
   "powerC [W],pfC,currentC [A],voltageC [V],totalC [Wh],total_returnedC [Wh]\n"
        if -z $status; # on empty status output CSV file, add header
    # no header for load output CSV file
}

sub do_before_hour {
    my ($date_3em, $time_3em, $date_3em_out, $time_3em_out, $first) = @_;
    die "error matching time" unless $time_3em =~ m/^(\d\d):(\d\d:\d\d)$/;
    my ($hour_3em, $min_sec_3em) = ($1, $2);
    return unless $first || $min_sec_3em eq "00:00";

    do_before_day($date_3em, $time_3em, $date_3em_out, $time_3em_out, $first);
    print $LM "\n" unless ($first || -z $load_min);
    print $LS "\n" unless ($first || -z $load_sec);
    my $date_time_out = $date_3em_out.$date_time_sep_out.$time_3em_out; # $hour_3em
    print $LM $date_time_out;
    print $LS $date_time_out;
}

sub do_each_second {
    my ($timestamp, $power, $data) = @_;
    my $time = time_epoch($timestamp);
    my ($date_3em    , $time_3em    ) = date_time($time);
    my ($date_3em_out, $time_3em_out) = date_time_out($time);
    my $first = $count_seconds == 0;

    do_before_hour($date_3em, $time_3em, $date_3em_out, $time_3em_out, $first);

    ++$count_seconds;
    $power_sum_minute += $power;
# https://www.promotic.eu/en/pmdoc/Subsystems/Comm/PmDrivers/IEC62056_OBIS.htm
# https://de.wikipedia.org/wiki/Stromz%C3%A4hler Zweirichtungszähler für
# Verbrauch (OBIS-Kennzahl 1.8.0) und Einspeisung (OBIS-Kennzahl 2.8.0)
    $energy_imported_this_hour += $power
        if $power > 0; # Positive active energy, energy meter register 1.8.0
    $energy_exported_this_hour -= $power
        if $power < 0; # Negative active energy, energy meter register 2.8.0
    my $power_ = round($power);
    print $LS ",$power_";
    print $SO "$time_3em_out,$power_$data\n";

    if (!$first && $time_3em =~/:59$/) { # end of each minute
        print $LM ",".round($power_sum_minute / SECONDS_PER_MINUTE);
        $power_sum_minute = 0;
        # make load and status output visible
        $LM->flush();
        $LS->flush();
        $SO->flush();

        # at end of each hour, calculate total imported/exported energy
        if ($time_3em =~/59:59$/) {
            printf $EO "$date_3em_out$date_time_sep_out$time_3em_out,%d,%d\n",
                round($energy_imported_this_hour / SECONDS_PER_HOUR),
                round($energy_exported_this_hour / SECONDS_PER_HOUR);
            ($energy_imported_this_hour, $energy_exported_this_hour) = (0, 0);
        }
    }
}

use Time::HiRes qw(usleep);

do {
    my $first = $count_seconds == 0;
    my ($timestamp, $power, $data) = get_line($first);
    my $diff_seconds = $first ? 1 : $timestamp - $prev_timestamp;
    if ($diff_seconds == 0) {
        print "$time: $timestamp (skipping same result)\n" if $debug;
    } else {
        print "$time: $timestamp,$power,$data\n" if $debug;
        if ($diff_seconds > 1) {
            log_warn("time gap ".++$count_gaps.": $diff_seconds seconds");
            # linear interpolation of missing time and power
            my $power_step = ($power - $prev_power) / $diff_seconds;
            while (--$diff_seconds) {
                $prev_power += $power_step;
                do_each_second(++$prev_timestamp, $prev_power, "");
            }
        } elsif ($diff_seconds < 0) {
            log_warn("negative 3EM status unixtime difference: $diff_seconds");
        }
        do_each_second($timestamp, $power, ",$data");
    }

    $prev_power = $power;
    $prev_timestamp = $timestamp;
    # $prev = $time;
    usleep(500000); # 0.5 secs; each iteration otherwise takes about .2 seconds
    ($date, $time) = date_time(DateTime->now(time_zone => $tz));
} while(1);
# while ($count_seconds < MAX_SECONDS); # stop after 1 day at the latest
# while $time ge $prev # not yet wrap around at 24:00:00
#     && $time lt $end_time;
# cleanup();
