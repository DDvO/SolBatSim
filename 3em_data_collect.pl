#!/usr/bin/perl

# Collect status data reported each second by a Shelly (Pro) 3EM energy meter,
# using the http://<3em_addr>/status endpoint (yet not the MQTT interface) and
# optionally collecting PV status data reported each second by Shelly Plus 1PM
# using the http://<1pm_addr>/rpc/Shelly.GetStatus endpoint but not MQTT.
# In this case, the total load reported by the 3EM energy meter is corrected
# by adding the absolute value of the PV power reported by the 1PM power meter.
# Optionally collect also the storage battery charger status data reported each
# second by another Shelly Plus 1PM using http://<chg_addr>/rpc/Shelly.GetStatus
# In this case, the total load is further corrected by subtracting charge power.
# Optionally collect also the storage battery inverter status data reported
# each second by an OpenDTU using http://<dtu_addr>/api/livedata/status and/or
# the status data of any Shelly Plus 1PM attached to the battery inverter.
# In this case, the total load is further corrected adding the discharge power.
#
# Alternatively, take as input per-second load and (optional) PV power data
# obtained, e.g., using Home Assistant. In this case, <3em_addr> must be '-'.
#
# Outputs data in the following files, each of which is optional:
# * <base_name><power_name>_<date>.csv  total load, PV input, charge, discharge,
#                                       and the three phase powers per second
# * <base_name><energy_name>_<year>.csv energy consumed, produced, own use (self
#                                     consumption), balance, imported, exported,
#                                     charged, and discharged in total per hour,
#                                     and any battery voltage at the end of hour
# * <base_name><load_min>_<year>.csv  average load per minute, one line per hour
# * <base_name><load_sec>_<date>.csv    load per second, one line per hour
# * <base_name><status_name>_<date>.csv status of the three phases per second,
#                             preceded by PV input, charge, and discharge power
# * <base_name><pvstat_name>_<date>.csv status of optional PV input per second
# * <base_name><chgstat_name>_<date>.csv status of optional charger per second
# * <base_name><disstat_name>_<date>.csv status of optional discharge per second
# * <base_name><log_name>_<year>.txt    info on the data collection per event
# The script is robust w.r.t. intermittently missing power data by interpolating
# the data over the range of seconds where no power measurement is available.
# In order to cope with inadvertent abortion of script execution (e.g., due to
# system reboot), the script should be started automatically when not currently
# running, for instance using a Linux cron job that is triggered each minute.
# It can recover the per-minute and per-hour data accumulation for the current
# day if the file with the load per second is available. For correct recovery
# including PV production, the file with PV status data per second is needed.
# With a charger being used, also the file with charger status data is needed.
# With battery discharge, also the file with discharge status data is needed.

# day if the file containing the load values per second is available.
#
# CLI options, each of which may be a value or "" indicating none/default:
# <base_name> <power_name> <energy_name>
# <load_min> <load_sec> <status_name> <pvstat_name> <log_name> <time_zone>
# - |(<3em_addr> <1pm_addr> <chg_addr> <dis_addr> <dtu_addr> <dtu_serial>
#     <3em_username> <3em_password> <1pm_username> <1pm_password>
#     <chg_username> <chg_password> <dis_username> <dis_password>
#     <dtu_username> <dtu_password>)
# where '-' means that data shall be read from subsequent file(s) or STDIN.
#
# Alternatively to providing options at the CLI, they may also be given
# via environment variables, which is advisable for sensitive passwords.
#
# By default, the time zone is CET (Central European Time)
# without the typical adaptations to DST (Daylight Saving Time) twice a year.
# This prevents misalignment and confusion on the interpretation of timed data
# and makes sure that the length of the daily output is the same for all days.
#
# (c) 2023-2024 David von Oheimb - License: MIT

use strict;
use warnings;
#use IO::Null; # using IO::Null->new, print gives: print() on unopened filehandle GLOB
use IO::Handle; # for flush
use Time::HiRes qw(usleep);

my $i = 0;
my $out_prefix   = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_BASENAME}; # e.g., ~/3EM_
my $out_power    = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_POWER};    # per second
my $out_energy   = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_ENERGY};   # per hour
my $out_load_min = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_LOAD_MIN}; # per minute
my $out_load_sec = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_LOAD_SEC}; # per second
my $out_stat     = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_STATUS};   # per second
my $out_pvstat   = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_PV};       # per second
my $out_chgstat  = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_CHG};      # per second
my $out_disstat  = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_DIS};      # per second
my $out_log      = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_LOG};      # per event

# time zone for output e.g., "local"
my $tz = $ARGV[$i++] || $ENV{Shelly_3EM_OUT_TZ} || "CET";
my $date_format        = "%Y-%m-%d";
my $date_format_out    = $ENV{Shelly_3EM_OUT_DATE_FORMAT} || $date_format;
my $date_time_sep      = " "; # "T";
my $date_time_sep_out  = $ENV{Shelly_3EM_OUT_DATE_TIME_SEP} || $date_time_sep;
my $time_format        = "%H:%M:%S";
my $time_format_out    = $ENV{Shelly_3EM_OUT_TIME_FORMAT} || $time_format;
my $format_out = $date_format_out.$date_time_sep_out.$time_format_out;

my (@times, @loads, @ppowers, @phases);
my $item = -1;
sub may_fill_last_sec {
    my $n = $#times;
    return if $n < 0;
    $times[$n] =~ m/^([^\s]+) (23:59:59)?/;
    return if defined $2;
    push @times  , "$1 23:59:59";
    push @loads  , $loads  [$n];
    push @ppowers, $ppowers[$n];
    push @phases , $phases [$n];

}
my $addr = $ARGV[$i++] || $ENV{Shelly_3EM_ADDR}; # e.g. 192.168.178.123
my ($url, $user, $pass);
my ($addr_1pm, $url_1pm, $user_1pm, $pass_1pm);
my ($addr_chg, $url_chg, $user_chg, $pass_chg);
my ($addr_dis, $url_dis, $user_dis, $pass_dis);
my ($addr_dtu, $url_dtu, $user_dtu, $pass_dtu, $serial_dtu);
if ($addr eq "-") {
    # read from power.csv file produced by Home Assistant configuration.yaml
    splice(@ARGV,0,$i);
    while(<>) {
        chomp;
        my @elems = (split ",", $_);
        if ($#elems == 5) {
            my ($time, $load, $pv, $phA, $phB, $phC) = @elems;
            may_fill_last_sec() if $time =~ m/^([^\s])+ 00:00:00$/;
            push @times  , $time;
            push @loads  , $load;
            push @ppowers, $pv;
            push @phases , "$phA,$phB,$phC";
            ++$item;
            my $sum = $pv + $phA + $phB + $phC;
            print("at $time, load $load is not consistent with ".
                  sprintf("%7.2f", $sum)." = sum of phases ".
                  "$phA + $phB + $phC and PV production $pv")
                unless abs($load - $sum) <= 0.01;
        } else {
            print "ignoring input line: $_\n";
        }
    }
    may_fill_last_sec() if $#times >= 0;
    $item = -1;
} else {
    $addr_1pm = $ARGV[$i++] || $ENV{Shelly_1PM_ADDR}; # e.g. 192.168.178.124
    $addr_1pm = 0 if $addr_1pm eq "";
    $addr_chg = $ARGV[$i++] || $ENV{Shelly_CHG_ADDR}; # e.g. 192.168.178.125
    $addr_chg = 0 if $addr_chg eq "";
    $addr_dis = $ARGV[$i++] || $ENV{Shelly_DIS_ADDR}; # e.g. 192.168.178.126
    $addr_dis = 0 if $addr_dis eq "";
    $addr_dtu = $ARGV[$i++] || $ENV{Shelly_DTU_ADDR}; # e.g. 192.168.178.127
    $addr_dtu = 0 if $addr_dtu eq "";
    $serial_dtu = $ARGV[$i++] || $ENV{Shelly_DTU_SERIAL}; # e.g. 112183822756
    $url      = "http://$addr/status";
    $url_1pm  = "http://$addr_1pm/rpc/Shelly.GetStatus" if $addr_1pm;
    $url_chg  = "http://$addr_chg/rpc/Shelly.GetStatus" if $addr_chg;
    $url_dis  = "http://$addr_dis/rpc/Shelly.GetStatus" if $addr_dis;
    $url_dtu  = "http://$addr_dtu/api/livedata/status"  if $addr_dtu; # OpenDTU

    $user     = $ARGV[$i++] || $ENV{Shelly_3EM_USER}; # HTTP username, if needed
    $pass     = $ARGV[$i++] || $ENV{Shelly_3EM_PASS}; # HTTP password, if needed
    $user_1pm = $ARGV[$i++] || $ENV{Shelly_1PM_USER} || $user;
    $pass_1pm = $ARGV[$i++] || $ENV{Shelly_1PM_PASS} || $pass;
    $user_chg = $ARGV[$i++] || $ENV{Shelly_CHG_USER} || $user;
    $pass_chg = $ARGV[$i++] || $ENV{Shelly_CHG_PASS} || $pass;
    $user_dis = $ARGV[$i++] || $ENV{Shelly_DIS_USER} || $user;
    $pass_dis = $ARGV[$i++] || $ENV{Shelly_DIS_PASS} || $pass;
    $user_dtu = $ARGV[$i++] || $ENV{Shelly_DTU_USER} || $user;
    $pass_dtu = $ARGV[$i++] || $ENV{Shelly_DTU_PASS} || $pass;
}

my $debug = $ENV{Shelly_3EM_DEBUG} // 0;
my $test_extra_power = $ENV{Shelly_3EM_EXTRA} // 0;
my $test_extra_pv_power = $ENV{Shelly_1PM_EXTRA} // 0;

die "missing CLI argument or env. variable 'Shelly_3EM_ADDR'"   unless $addr;
die "missing CLI argument or env. variable 'Shelly_3EM_OUT_TZ'" unless $tz;

# deliberately not using any extra packages like Math
sub min { return $_[0] < $_[1] ? $_[0] : $_[1]; }
sub max { return $_[0] > $_[1] ? $_[0] : $_[1]; }
sub round { return int(($_[0] < 0 ? -.5 : .5) + $_[0]); }

# local system time
# https://stackoverflow.com/questions/60107110/perl-strftime-localtime-minus-12-hours
use DateTime;
sub date_time {
    return ($_[0]->strftime($date_format),
            $_[0]->strftime($time_format));
}
sub date_time_now { return date_time(DateTime->now(time_zone => $tz)); }
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

# https://stackoverflow.com/questions/7486470/how-to-parse-a-string-into-a-datetime-object-in-perl
my $parse_err;
# moving var into sub would lead to: Variable "$parse_err" will not stay shared
sub parse_datetime {
    use DateTime::Format::Strptime;
    my $in = "$_[0] $tz";
    $parse_err = "cannot parse date+time '$in' w.r.t. pattern '$format_out %Z'";
    sub die_err { die($parse_err); }
    my $time_parser = DateTime::Format::Strptime->new(
        pattern => "$format_out %Z",
        on_error => \&die_err,
    );
    return $time_parser->parse_datetime($in);
}

sub time_epoch { return DateTime->from_epoch(epoch => $_[0], time_zone => $tz);}
use constant SECONDS_PER_MINUTE => 60;
use constant SECONDS_PER_HOUR => 60 * SECONDS_PER_MINUTE;
# max seconds per day even for days with daylight saving time (DST) adaptation:
# use constant MAX_SECONDS => (24 + 1) * SECONDS_PER_HOUR;
my ($count_seconds, $count_gaps, $count_1pm_miss) = (0, 0, 0);
my ($count_chg_miss, $count_dis_miss) = (0, 0);

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
my ($pvstat, $PO);
my ($chgstat, $CH);
my ($disstat, $DS);
my ($powers, $PW);
my ($log, $LOG);
# preliminary log:
sub log_name { return out_name($out_log, $_[0], ".txt"); }
$log = log_name($start->strftime("%Y"));
open($LOG,'>>', $log) || die "cannot open '$log' for appending: $!";

sub log_msg {
    my $tim = $addr eq "-" && defined $times[$item] ? $times[$item]
        : "$date $time";
    my $msg = "$tim: $_[0]\n";
    print $msg;
    print $LOG $msg if defined $LOG;
}
sub log_warn { log_msg("WARNING: $_[0]"); }

sub do_after_day {
    close $LS if defined $LS;
    close $SO if defined $SO;
    close $PO if defined $PO;
    close $CH if defined $CH;
    close $DS if defined $DS;
    close $PW if defined $PW;
}

sub do_after_year {
    close $LM  if defined $LM;
    close $EO  if defined $EO;
    close $LOG if defined $LOG;
}

my ($energy_consumed_this_hour, $energy_produced_this_hour) = (0, 0);
my ($energy_charged_this_hour,$energy_discharged_this_hour) = (0, 0);
my ($energy_own_used_this_hour, $energy_balanced_this_hour) = (0, 0);
my ($energy_imported_this_hour, $energy_exported_this_hour) = (0, 0);

sub cleanup() {
    do_after_day();
    log_msg "energy sums in Wh so far last hour: "
        ."consumed ".round($energy_consumed_this_hour / SECONDS_PER_HOUR).", "
        ."produced ".round($energy_produced_this_hour / SECONDS_PER_HOUR).", "
        ."charged " .round( $energy_charged_this_hour / SECONDS_PER_HOUR).", "
        ."discharged ".round($energy_discharged_this_hour/SECONDS_PER_HOUR).", "
        ."own use " .round($energy_own_used_this_hour / SECONDS_PER_HOUR).", "
        ."balance " .round($energy_balanced_this_hour / SECONDS_PER_HOUR).", "
        ."imported ".round($energy_imported_this_hour / SECONDS_PER_HOUR).", "
        ."exported ".round($energy_exported_this_hour / SECONDS_PER_HOUR);
    log_msg("end after $count_seconds seconds, ".
            "$count_gaps gaps, $count_1pm_miss PV data misses, ".
            "$count_chg_miss charger data misses, ".
            "$count_dis_miss discharge data misses");
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

# https://www.xmodulo.com/how-to-send-http-get-or-post-request-in-perl.html
# https://stackoverflow.com/questions/19842400/perl-http-post-authentication
sub http_get {
    my ($url, $user, $pass) = @_;

    use LWP::UserAgent;
    my $ua = LWP::UserAgent->new;
    $ua->timeout(1);
    my $req = HTTP::Request->new(GET => $url);
    $req->authorization_basic($user, $pass) if $user || $pass;
    my $resp = $ua->request($req);
    if ($resp->is_success) {
        # print "HTTP GET response: ".$resp->decoded_content."\n";
        return $resp->content;
    }
    # print "HTTP GET status: ", $resp->code, ", err: ", $resp->message, "\n";
    return $resp->message;
}

my $last_valid_unixtime = 0;
my $last_3em_status_serial = 0;
sub get_3em {
    if ($addr eq "-") {
        print "$times[$item],$loads[$item],$phases[$item]\n" if $debug;
        my $dt = parse_datetime($times[$item]);
        return ($dt->epoch, $loads[$item], $phases[$item]);
    }

    my ($check_time) = @_;

  retry:
    ($date, $time) = date_time_now();
    my $status_json = http_get($url, $user, $pass);
    if ($status_json =~ m/(Network is unreachable|No route to host|Server closed connection|Connection reset by peer|(Connection|Operation) timed out|read timeout)/) {
        # e.g.: "Can't connect to 3em:80 (Operation timed out)"
        log_warn("$1 for 3EM");
        sleep(1) unless $1 =~ m/timed out|read timeout/;
        goto retry;
    }
    unless ($status_json =~ /"time":"([\d:]*)","unixtime":(\d+),"serial":(\d+),.*"emeters":\[\{"power":([\-\d\.]+),"pf":([\-\d\.]+),"current":([\-\d\.]+),"voltage":([\-\d\.]+),"is_valid":true,"total":([\d\.]+),"total_returned":([\d\.]+)}\,\{"power":([\-\d\.]+),"pf":([\-\d\.]+),"current":([\-\d\.]+),"voltage":([\-\d\.]+),"is_valid":true,"total":([\d\.]+),"total_returned":([\d\.]+)\},\{"power":([\-\d\.]+),"pf":([\-\d\.]+),"current":([\-\d\.]+),"voltage":([\-\d\.]+),"is_valid":true,"total":([\d\.]+),"total_returned":([\d\.]+)\}\],"total_power":([\-\d\.]+),.*,"uptime":(\d+)/) {
        if ($status_json =~ /ERROR:\s?([\s0-9A-Za-z]*)/i) {
            # e.g.: The requested URL could not be retrieved
            log_warn("skipping error response: $1 for 3EM"); # e.g., by Squid
        } else {
            my $shown = substr($status_json, 0, 1100); # typically ~1020 chars
            log_warn("error parsing 3EM status response '$shown'");
        }
        sleep(1);
        goto retry;
    }

    my ($hour, $unixtime, $status_serial,
        $powerA, $pfA, $currentA, $voltageA, $totalA, $total_returnedA,
        $powerB, $pfB, $currentB, $voltageB, $totalB, $total_returnedB,
        $powerC, $pfC, $currentC, $voltageC, $totalC, $total_returnedC,
        $total_power, $uptime)
        = ($1, $2 + 0, $3 + 0,
           $4, $5, $6, $7, $8, $9,
           $10, $11, $12, $13, $14, $15,
           $16, $17, $18, $19, $20, $21,
           $22, $23);

    # not checking the 'serial' field because it would lead to many gaps - why?
    if (0 && $status_serial == $last_3em_status_serial) {
        usleep(100000);
        goto retry;
    }
    $last_3em_status_serial = $status_serial;

    if ($unixtime) {
        $last_valid_unixtime = $unixtime;
    } elsif ($last_valid_unixtime) {
        log_warn("approximating missing 3EM status unixtime from last valid"
                 ." one $last_valid_unixtime + uptime $uptime");
        $unixtime = $last_valid_unixtime + $uptime;
    } else {
        $unixtime = time();
        unless ($unixtime) {
            log_warn("missing 3EM status unixtime and system time, discarding '$status_json'");
            sleep(1);
            goto retry;
        }
        log_warn("missing 3EM status unixtime, taking host system time");
    }

    my ($date_3em, $time_3em) = date_time(time_epoch($unixtime));
    if ($check_time) {
        my $time_hour = substr($time, 0, 5);
        log_warn("3EM status time '$hour' does not equal '$time_hour'")
            unless $hour eq $time_hour;
        log_warn("3EM status unixtime '$date_3em"."$date_time_sep$time_3em'".
                 " does not very closely match ".
                 "host system time '$date"."$date_time_sep$time'")
            unless abs($unixtime - $start->epoch) <= 1
            # 3 seconds diff can happen easily
    }
    my $power = $powerA + $powerB + $powerC;
    # may include PV power, charge, and discharge
    my ($pA, $pB, $pC) = (sprintf("%6.2f", $powerA),
                          sprintf("%6.2f", $powerB),
                          sprintf("%6.2f", $powerC));
    my $dataA = "$pA,$pfA,$currentA,$voltageA,$totalA,$total_returnedA";
    my $dataB = "$pB,$pfB,$currentB,$voltageB,$totalB,$total_returnedB";
    my $dataC = "$pC,$pfC,$currentC,$voltageC,$totalC,$total_returnedC";
    my $data = "$dataA,$dataB,$dataC";
    print "($time, $hour, $unixtime, 3EM $time_3em, $power, $data)\n" if $debug;

    log_warn("inconsistent 3EM total_power = $total_power ".
             "vs. $powerA + $powerB + $powerC")
        unless abs($power - $total_power) <= 0.1;
    return ($unixtime, $power, $data);
}

sub get_1pm {
    if ($addr eq "-") {
        my $dt = parse_datetime($times[$item]);
        return ($dt->epoch, $ppowers[$item], "");
    }

    my ($name, $timestamp, $url, $user, $pass) = @_;
    my ($unixtime, $power, $data) = (0, 0, ""); # default: no current data

    my $status_json = http_get($url, $user, $pass);
    if ($status_json =~ m/(Network is unreachable|No route to host|Server closed connection|Connection reset by peer|(Connection|Operation) timed out|read timeout)/) {
        # e.g.: "Can't connect to pm1:80 (Operation timed out)"
        log_warn("$1 for $name");
        goto end;
    }
    unless ($status_json =~ /"switch:0":\{"id":0, "source":"\w+", "output":\w+, "apower":([\-\d\.]+), "voltage":([\-\d\.]+), "current":([\-\d\.]+), "aenergy":\{"total":([\-\d\.]+),"by_minute":\[([\-\d\.]+),([\-\d\.]+),([\-\d\.]+)\],"minute_ts":(\d+)\},"temperature":\{"tC":([\-\d\.]+), "tF":([\-\d\.]+)\}\}/) {
        if ($status_json =~ /ERROR:\s?([\s0-9A-Za-z]*)/i) {
            # e.g.: The requested URL could not be retrieved
            log_warn("skipping error response: $1 for $name"); # e.g., by Squid
        } else {
            my $shown = substr($status_json, 0, 800); # typically ~710 chars
            log_warn("error parsing $name status response '$shown'");
        }
        goto end;
    }

    my ($apower, $voltage, $current, $total, $min1, $min2, $min3, $ts, $tC, $tF)
        = ($1, $2, $3, $4, $5, $6, $7, $8 + 0, $9, $10);
    if ($ts) {
        $unixtime = $ts;
    } elsif ($timestamp) {
        log_warn("substituting missing $name status minute_ts from 3EM timestamp $timestamp");
        $unixtime = $timestamp;
    } else {
        log_warn("missing $name status minute_ts, discarding '$status_json'");
        goto end;
    }
    $current = "0    " if abs($current) < 0.0005;
    # PV power might be reported negative, but usually is reported >= 0
    ($power, $data) = (abs($apower), "$voltage,$current,$total,$tC");

    my $dt = time_epoch($unixtime);
    my $hour = sprintf("%02d:%02d", $dt->hour, $dt->minute);
    my ($date_1pm, $time_1pm) = date_time($dt);
    print "($time, $hour, $unixtime, $name $time_1pm, $power, $data)\n"
        if $debug;

    log_warn("$name status minute_ts '$date_1pm"."$date_time_sep$time_1pm' ".
             "does not very closely match 3EM timestamp '$date"."$date_time_sep$time'")
        unless abs($unixtime - $timestamp) <= 3; # 3 seconds diff can happen easily
  end:
    return ($unixtime, $power, $data);
}

# https://tbnobody.github.io/OpenDTU-docs/firmware/web_api/#get-current-livedata
sub get_dtu {
    my ($name, $timestamp, $url, $user, $pass) = @_;

    my $data = "";
    my $status_json = http_get($url, $user, $pass);
    if ($status_json =~ m/(Network is unreachable|No route to host|Server closed connection|Connection reset by peer|(Connection|Operation) timed out|read timeout)/) {
        # e.g.: "Can't connect to dtu:80 (Operation timed out)"
        log_warn("$1 for $name");
        $timestamp = 0;
        goto end;
    }

    # \{"inverters":\[
    unless ($status_json =~ /\{"serial":"$serial_dtu","name":"[\w\-]+","order":\d+,"data_age":\d+,"poll_enabled":\w+,"reachable":\w+,"producing":\w+,"limit_relative":([\-\d]+),"limit_absolute":([\-\d]+),"AC":\{"0":\{"Power":\{"v":([\-\d\.]+),"u":"W","d":\d+\},"Voltage":\{"v":([\-\d\.]+),"u":"V","d":\d+\},"Current":\{"v":([\-\d\.]+),"u":"A","d":\d+\},"Power DC":\{"v":[\-\d\.]+,"u":"W","d":\d+\},"YieldDay":\{"v":[\-\d\.]+,"u":"Wh","d":\d+\},"YieldTotal":\{"v":[\-\d\.]+,"u":"kWh","d":\d+\},"Frequency":\{"v":([\-\d\.]+),"u":"Hz","d":\d+\},"PowerFactor":\{"v":([\-\d\.]+),"u":"","d":\d+\},"ReactivePower":\{"v":([\-\d\.]+),"u":"var","d":\d+\},"Efficiency":\{"v":([\-\d\.]+),"u":"%","d":\d+\}\}\},"DC":\{(("\d+":\{"name":\{"u":"\w*"\},"Power":\{"v":[\-\d\.]+,"u":"W","d":\d+\},"Voltage":\{"v":[\-\d\.]+,"u":"V","d":\d+\},"Current":\{"v":[\-\d\.]+,"u":"A","d":\d+\},"YieldDay":\{"v":[\-\d\.]+,"u":"Wh","d":\d+\},"YieldTotal":\{"v":[\-\d\.]+,"u":"kWh","d":\d+\}\},?)+)\},"INV":\{"0":\{"Temperature":\{"v":([\-\d\.]+),"u":"°C","d":\d+\}\}\},"events":\d+\}/) {
    # \],"total":\{"Power":\{"v":[\-\d\.]+,"u":"W","d":\d\},"YieldDay":\{"v":[\-\d\.]+,"u":"Wh","d":\d+\},"YieldTotal":\{"v":[\-\d\.]+,"u":"kWh","d":\d+\}\},"hints":\{"time_sync":\w+,"radio_problem":\w+,"default_password":\w+\}\}
        if ($status_json =~ /ERROR:\s?([\s0-9A-Za-z]*)/i) {
            # e.g.: The requested URL could not be retrieved
            log_warn("skipping error response: $1 for $name"); # e.g., by Squid
        } else {
            my $shown = substr($status_json, 0, 1800); # may be ~1750 chars
            log_warn("error parsing $name status response '$shown'");
        }
        $timestamp = 0;
        goto end;
    }

    my ($limit_relative, $limit_absolute, $power, $voltage, $current, $frequency,
        $power_factor, $reactive_power, $efficiency, $DC, $temperature)
        = ($1, $2, $3 + 0, $4, $5, $6, $7, $8, $9, $10, $12);
    my $AC = sprintf("%3d,%2d,%.1f,%.2f,%.1f,%.2f,%.1f,%.1f,%.1f",
                     $limit_absolute, $limit_relative,
                     $voltage, $current, $frequency, $power_factor,
                     $reactive_power, $efficiency, $temperature);
    $DC =~ s/,"d":\d+//g;
    $DC =~ s/\{"u":""\},/ /g; # separator between input string
    $DC =~ s/"\w*"://g;
    $DC =~ s/(\d\.\d)\d*/$1/g;
    $DC =~ s/,"(W|V|A|Wh|kWh)"//g;
    $DC =~ s/[\{\}]//g;
    $data = "$AC,$DC";
    print "($time, $name, $power, $data)\n" if $debug;

  end:
    return ($timestamp, $power, $data);
}

my $load_sum_minute = 0;
my ($prev_power, $prev_timestamp) = (0, 0);
# may include PV power, charge, and discharge
# my $prev = "";

# previous plausible (i.e., non-negative) values:
my ($prev_pv_power, $prev_chg_power, $prev_dis_power) = (0, 0, 0);

log_msg("start - will connect to $url"
        .($url_1pm ? " and $url_1pm" : "")
        .($url_chg ? " and $url_chg" : "")
        .($url_dis ? " and $url_dis" : "")
        .($url_dtu ? " and $url_dtu" : ""))
    unless $addr eq "-";

# try recover data from any previous run
# TODO maybe add recovery also from $out_power, $out_load_min or $out_stat
my $cannot_recover = "cannot recover earlier data for the current hour";
if ($addr ne "-" && $out_load_sec) {
    my $load_sec = out_name($out_load_sec, $date, ".csv");
    if (open(my $LS, '<' ,$load_sec)) {
        my $prev_second = 0;
        my $line;
        while (<$LS>) {
            $line = $_;
        }
        close $LS;
        die "empty high-resolution load file '$load_sec'" unless defined $line;
        chomp $line;
        my @elems = (split ",", $line);
        my $n = $#elems;
        die "cannot parse last line '$line' of high-resolution load file '$load_sec'"
            if $n < 0;
        my $date_time = $elems[0];
        my $dt = parse_datetime($date_time);
        die "cannot parse date+time '$date_time' in last line of high-resolution load file '$load_sec'"
            unless $dt;
        $prev_timestamp = $dt->epoch;
        my ($minute, $second) = ($dt->minute, $dt->second);

        my $count = -1;
        my @pv_power = ("");
        if ($addr_1pm) { # TODO also for addr_chg and (addr_dis or addr_dtu)   
            if ($out_pvstat) {
                my $pvstat = out_name($out_pvstat, $date, ".csv");
                if (open(my $PO, '<' ,$pvstat)) {
                    my $time = $dt->strftime($time_format_out);
                    while (<$PO>) {
                        next if $count++ == -1;
                        $count--;
                        chomp;
                        my @values = (split ",", $_);
                        die "cannot parse line '$_' of PV status file '$pvstat'"
                            if $#values < 1;
                        next if $count == 0 && $values[0] lt $time;
                        $count++;
                        my $pv = $values[1];
                        die "cannot parse PV power value '$pv' in line '$_' of PV status file '$pvstat'"
                            unless $pv =~ m/^\s*-?[\d\.]+$/;
                        push @pv_power, $pv;
                    }
                    close $PO;
                    log_warn("for the time frame according to the last line of the high-resolution load file '$load_sec', the count of PV power values in PV status file '$pvstat': $count does not match number of load values: $n") if $count != $n;
                } else {
                    log_warn("no previouly produced PV status file '$pvstat' found, so $cannot_recover");
                }
            } else {
                log_warn("as no PV status file <pvstat_name> is defined, $cannot_recover");
            }
        }

        for (my $i = 1; $i <= $n; $i++) {
            my $load = $elems[$i];
            die "cannot parse power value '$load' in last line '$line' of high-resolution load file '$load_sec'"
                unless $load =~ m/^\s*-?\d+$/;
            $prev_pv_power = $addr_1pm && $i <= $count ? $pv_power[$i] + 0 : 0;
            my $pv_used = min($load, $prev_pv_power);
            $prev_power = $load - $prev_pv_power; # TODO 
            $load_sum_minute += $load;
            $energy_consumed_this_hour += $load;
            $energy_produced_this_hour += $prev_pv_power;
             $energy_charged_this_hour += $prev_chg_power;
          $energy_discharged_this_hour += $prev_dis_power;
            $energy_own_used_this_hour += $pv_used;
            $energy_balanced_this_hour += $prev_power;
            if ($prev_power > 0) {
                $energy_imported_this_hour += $prev_power;
            } else {
                $energy_exported_this_hour -= $prev_power;
            }
            print "$time $i<=$n, load $load - pv $prev_pv_power = $prev_power, "
                ."energy sums: consumed $energy_consumed_this_hour, "
                ."produced $energy_produced_this_hour, "
                . "charged $energy_charged_this_hour, "
           ."discharged $energy_discharged_this_hour, "
                ."own use $energy_own_used_this_hour, "
                ."balance $energy_balanced_this_hour, "
                ."imported $energy_imported_this_hour, "
                ."exported $energy_exported_this_hour\n"
                if $debug && $addr_1pm && $i > $n - 10; # TODO adapt
            if (++$second >= 60) {
                $second = 0;
                $minute++;
                die "too many power values in last line '$line' of high-resolution load file '$load_sec'"
                    if $minute >= 60;
                $load_sum_minute = 0;
            }
        }
        $prev_timestamp += $n;
    } else {
        log_warn("no previouly produced high-resolution load file '$load_sec' found, so $cannot_recover");
    }
} elsif ($addr ne "-") {
    log_warn("as no high-resolution load file <load_sec> is defined, $cannot_recover");
}

sub do_before_year {
    my ($date_3em, $first) = @_;
    die "error matching time" unless $date_3em =~ m/^(\d+)-(\d\d-\d\d)$/;
    my ($year_3em, $month_day_3em) = ($1, $2);
    return unless $first || $month_day_3em eq "01-01";

    print $LM "\n" if defined $LM && !($first && $prev_timestamp);
    do_after_year();
    $load_min = out_name($out_load_min, $year_3em, ".csv");
    $energy = out_name($out_energy, $year_3em, ".csv");
    $log    = log_name($year_3em);
    open($LOG, '>>', $log  ) || die "cannot open '$log' for appending: $!";
    open($EO, '>>', $energy) || die "cannot open '$energy' for appending: $!";
    open($LM,'>>',$load_min) || die "cannot open '$load_min' for appending: $!";

    $LOG->autoflush; # immediately show each line reporting an event
    $EO ->autoflush; # immediately show each line reporting energy per hour
    $LM ->autoflush; # immediately show each load per minute
    # on empty energy output CSV file, add header:
    print $EO "time [$tz],consumed [Wh],produced [Wh],own use [Wh],".
        "balance [Wh],imported [Wh],exported [Wh],".
        "charged [Wh],discharged [Wh],battery [V]\n" if -z $energy;
    # no header for load output CSV file
}

sub do_before_day {
    my ($date_3em, $time_3em, $date_3em_out, $time_3em_out, $first) = @_;
    return unless $first || $time_3em eq "00:00:00";

    print $LS "\n" if defined $LS && !($first && $prev_timestamp);;
    do_after_day();
    do_before_year($date_3em, $first);
    $load_sec = out_name($out_load_sec, $date_3em_out, ".csv");
    $status = out_name($out_stat, $date_3em_out, ".csv");
    $pvstat = out_name($out_pvstat, $date_3em_out, ".csv");
    $chgstat = out_name($out_chgstat, $date_3em_out, ".csv");
    $disstat = out_name($out_disstat, $date_3em_out, ".csv");
    $powers = out_name($out_power , $date_3em_out, ".csv");
    open($LS,'>>',$load_sec) || die "cannot open '$load_sec' for appending: $!";
    open($SO, '>>', $status) || die "cannot open '$status' for appending: $!";
    open($PO, '>>', $pvstat) || die "cannot open '$pvstat' for appending: $!";
    open($CH, '>>', $chgstat)|| die "cannot open '$chgstat' for appending: $!";
    open($DS, '>>', $disstat)|| die "cannot open '$disstat' for appending: $!";
    open($PW, '>>', $powers) || die "cannot open '$powers' for appending: $!";

    # no header for load output CSV file
    print $SO "time [$tz],PV power [W],charge power [W],discharge power [W],".
        "total_power [W],".
   "powerA [W],pfA,currentA [A],voltageA [V],totalA [Wh],total_returnedA [Wh],".
   "powerB [W],pfB,currentB [A],voltageB [V],totalB [Wh],total_returnedB [Wh],".
   "powerC [W],pfC,currentC [A],voltageC [V],totalC [Wh],total_returnedC [Wh]\n"
        if -z $status; # on empty status output CSV file, add header
    my $pm_data_header = "voltage [V],current [A],total [Wh],temperature [°C]";
    print $PO "time [$tz],PV power [W],$pm_data_header\n"
        if -z $pvstat; # on empty PV output CSV file, add header
    print $CH "time [$tz],charge power [W],$pm_data_header\n"
        if -z $chgstat; # on empty charger output CSV file, add header
    print $DS "time [$tz],discharge power [W],".
        (!$addr_dtu ? "$pm_data_header" : "limit [W],limit [%],".
         "voltage [V],current [A],frequency [Hz],power factor,".
         "reactive power [var],efficiency [%], temperature [°C],".
         "string 0 power [W],string 0 voltage [V],string 0 current [A],".
         "string 0 yield day [Wh],string 0 yield total [kWh],".
         "string 1 power [W],string 1 voltage [V],string 1 current [A],".
         "string 1 yield day [Wh],string 1 yield total [kWh]")."\n"
        if -z $disstat; # on empty discharge output CSV file, add header
    print $PW "time [$tz],load [W],PV power [W],charge power [W],".
        "discharge power [W],powerA [W],powerB [W],powerC [W]\n"
        if -z $powers; # on empty power output CSV file, add header
}

sub do_before_hour {
    my ($date_3em, $time_3em, $date_3em_out, $time_3em_out, $first) = @_;
    die "error matching time" unless $time_3em =~ m/^(\d\d):(\d\d:\d\d)$/;
    my ($hour_3em, $min_sec_3em) = ($1, $2);
    return unless $first || $min_sec_3em eq "00:00";

    do_before_day($date_3em, $time_3em, $date_3em_out, $time_3em_out, $first);

    return if $first && $prev_timestamp;
    print $LM "\n" unless (-z $load_min);
    print $LS "\n" unless (-z $load_sec);
    my $date_time_out = $date_3em_out.$date_time_sep_out.$time_3em_out; # $hour_3em
    print $LM $date_time_out;
    print $LS $date_time_out;
}

my $prev_load =  0;
sub do_each_second {
    my ($timestamp, $powers_ok, $power, $data, $pv_power, $pv_data,
        $chg_power, $chg_data, $dis_power, $dis_data) = @_;
    my $pvpower  =  $pv_power ? sprintf("%5.1f",  $pv_power) : "    0";
    my $chgpower = $chg_power ? sprintf("%5.1f", $chg_power) : "    0";
    my $dispower = $dis_power ? sprintf("%5.1f", $dis_power) : "    0";

    my $load = $power + $pv_power - $chg_power + $dis_power;
    $powers_ok &&=
        $load > 0 && $pv_power >= 0 && $chg_power >= 0 && $dis_power >= 0;
    if ($powers_ok) {
        $prev_load = $load;
    } else {
        my $r = $load > 0 ? "not from current plausible data" : "not positive";
        my $l = sprintf("%.2f", $load);
        log_warn("load is $r: $l; substituting previous value $prev_load");
        $load = $prev_load;
    }
    if ($pv_power >= 0) {
        $prev_pv_power = $pv_power;
    } else {
        log_warn("PV power is negative: $pv_power; ".
                 "substituting for energy the previous value $prev_pv_power");
        $pv_power = $prev_pv_power;
    }
    my $pv_own_used = min($load, $pv_power);
    if ($chg_power >= 0) {
        $prev_chg_power = $chg_power;
    } else {
        log_warn("charge power is negative: $chg_power; ".
                 "substituting for energy the previous value $prev_chg_power");
        $chg_power = $prev_chg_power;
    }
    if ($dis_power >= 0) {
        $prev_dis_power = $dis_power;
    } else {
        log_warn("discharge power is negative: $dis_power; ".
                 "substituting for energy the previous value $prev_dis_power");
        $dis_power = $prev_dis_power;
    }
    my $time = time_epoch($timestamp);
    my ($date_3em    , $time_3em    ) = date_time($time);
    my ($date_3em_out, $time_3em_out) = date_time_out($time);
    my $first = $count_seconds == 0;

    do_before_hour($date_3em, $time_3em, $date_3em_out, $time_3em_out, $first);

    ++$count_seconds;
    $load_sum_minute += $load;
    $energy_consumed_this_hour += $load;
    $energy_produced_this_hour += $pv_power;
     $energy_charged_this_hour += $chg_power;
  $energy_discharged_this_hour += $dis_power;
    $energy_own_used_this_hour += $pv_own_used; # self-consumption
    $energy_balanced_this_hour += $power;
# https://www.promotic.eu/en/pmdoc/Subsystems/Comm/PmDrivers/IEC62056_OBIS.htm
# https://de.wikipedia.org/wiki/Stromz%C3%A4hler Zweirichtungszähler für
# Verbrauch (OBIS-Kennzahl 1.8.0) und Einspeisung (OBIS-Kennzahl 2.8.0)
    $energy_imported_this_hour += $power
        if $power > 0; # Positive active energy, energy meter register 1.8.0
    $energy_exported_this_hour -= $power
        if $power < 0; # Negative active energy, energy meter register 2.8.0
    print $LS ",".round($load) if $powers_ok;  # suppress unclear load values
    print $SO "$time_3em_out,$pvpower,$chgpower,$dispower,"
        .sprintf("%+6.2f", $power)."$data\n";
    print $PO "$time_3em_out,$pvpower$pv_data\n";
    print $CH "$time_3em_out,$chgpower$chg_data\n";
    print $DS "$time_3em_out,$dispower$dis_data\n";
    if ($data ne "") {
        my @dat = (split ",", $data);
        my $inc = $#dat == 3 ? 1 : 6;
        $data =
            ",".sprintf("%6.2f", $dat[1]).
            ",".sprintf("%6.2f", $dat[1 + $inc]).
            ",".sprintf("%6.2f", $dat[1 + $inc * 2]);
    }
    my $date_time_out = $date_3em_out.$date_time_sep_out.$time_3em_out;
    print $PW "$date_time_out,".sprintf("%7.2f", $load)
        .",$pvpower,$chgpower,$dispower$data\n"
        if $powers_ok;  # suppress unclear power values

    if (!$first && $time_3em =~/:59$/) { # end of each minute
        print $LM ",".round($load_sum_minute / SECONDS_PER_MINUTE);
        $load_sum_minute = 0;
        # make load and status output visible
        $LM->flush();
        $LS->flush();
        $SO->flush();
        $PO->flush();
        $CH->flush();
        $DS->flush();
        $PW->flush();

        if ($time_3em =~/59:59$/) { # at end of each hour
            my $consumed = round($energy_consumed_this_hour / SECONDS_PER_HOUR);
            my $produced = round($energy_produced_this_hour / SECONDS_PER_HOUR);
            my $charged  = round( $energy_charged_this_hour / SECONDS_PER_HOUR);
            my $discharged=round($energy_discharged_this_hour/SECONDS_PER_HOUR);
            my $own_used = round($energy_own_used_this_hour / SECONDS_PER_HOUR);
            my $balanced = round($energy_balanced_this_hour / SECONDS_PER_HOUR);
            my $imported = round($energy_imported_this_hour / SECONDS_PER_HOUR);
            my $exported = round($energy_exported_this_hour / SECONDS_PER_HOUR);
            my $voltage = " 0";
            $voltage = sprintf("%2.1f", $1) if $dis_data =~
                m/,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,[^,]*,([\-\d\.]+),[^,]*,[^,]*,[^,]*/;
            printf $EO
                "$date_time_out,%4d,%4d,%4d,%4d,%4d,%4d,%4d,%4d,%s\n",
                $consumed, $produced, $own_used, $balanced,
                $imported, $exported, $charged, $discharged, $voltage;

            my $diff1 = round(($energy_consumed_this_hour
                               +$energy_charged_this_hour
                               -$energy_discharged_this_hour
                               -$energy_produced_this_hour) / SECONDS_PER_HOUR);
            log_warn("energy balance = $balanced vs. $diff1 = ".
                     "energy consumed $consumed + charged $charged - ".
                     "energy discharged $discharged - produced $produced")
                if abs($balanced - $diff1) > 1;
            my $diff2 = round(($energy_imported_this_hour -
                               $energy_exported_this_hour) / SECONDS_PER_HOUR);
            log_warn("energy balance = $balanced vs. $diff2 = ".
                     "energy imported $imported - exported $exported")
                if abs($balanced - $diff2) > 1;

            ($energy_consumed_this_hour, $energy_produced_this_hour) = (0, 0);
            ($energy_charged_this_hour,$energy_discharged_this_hour) = (0, 0);
            ($energy_own_used_this_hour, $energy_balanced_this_hour) = (0, 0);
            ($energy_imported_this_hour, $energy_exported_this_hour) = (0, 0);
        }
    }
}

# re-calculate due to potenital delays recovering data from any previous run:
$start = DateTime->now(time_zone => $tz);
my ($nseconds, $rounds) = (0, 0) if $debug;
do {
    $rounds++ if $debug;
    goto end if $addr eq "-" && ++$item > $#times;
    my $first = $count_seconds == 0;
    my ($timestamp, $power, $data) = get_3em($first);
    # may include PV power, charge, and discharge
    my ( $pv_timestamp,  $pv_power,  $pv_data) = (0, 0, "");
    my ($chg_timestamp, $chg_power, $chg_data) = (0, 0, "");
    my ($dis_timestamp, $dis_power, $dis_data) = (0, 0, "");
    my $diff_seconds = $prev_timestamp ? $timestamp - $prev_timestamp : 1;
    if ($diff_seconds == 0) {
        print "$time: $timestamp (skipping result for same time)\n" if $debug;
    } else {
        my $powers_ok = 1;
        $power += $test_extra_power;
        $nseconds += $diff_seconds unless $first;
        if ($addr_1pm && $diff_seconds >= 1) {
            ($pv_timestamp, $pv_power, $pv_data) =
                get_1pm("PV", $timestamp, $url_1pm, $user_1pm, $pass_1pm);
            if ($pv_timestamp) {
                $pv_power = 0 if 0 < $pv_power
                    && $pv_power < 0.9;  # inverter drags ~0.7 W on standby
                $pv_power += $test_extra_pv_power;
                $power -= $test_extra_pv_power;
            } else {
                $powers_ok = 0;
                log_warn("taking previous PV power value $prev_pv_power "
                         ."as no current PV status data available");
                ++$count_1pm_miss;
                $pv_power = $prev_pv_power;
            }
        }
        if ($addr_chg && $diff_seconds >= 1) {
            ($chg_timestamp, my $chg, $chg_data) =
                get_1pm("charger", $timestamp, $url_chg, $user_chg, $pass_chg);
            if ($chg_timestamp) {
                $chg_power = max(0, $chg - 2.5) # HLG-600H drags ~2.4 W on standby
                    if $chg_power > 0;
            } else {
                $powers_ok = 0;
                log_warn("taking previous charge power value $prev_chg_power "
                         ."as no current status data available from charger");
                ++$count_chg_miss;
                $chg_power = $prev_chg_power;
            }
        }
        if (($addr_dis || $addr_dtu) && $diff_seconds >= 1) {
            my $dtu_timestamp = 0;
            ($dis_timestamp, $dis_power, $dis_data) =
                get_1pm("discharge data from 1PM", $timestamp, $url_dis,
                        $user_dis, $pass_dis) if ($addr_dis);
            ($dtu_timestamp, my $dtu_power, my $dtu_data) =
                get_dtu("discharge data from DTU", $timestamp, $url_dtu,
                        $user_dtu, $pass_dtu) if ($addr_dtu);
            if ($dis_timestamp || $dtu_timestamp) {
                if ($dtu_timestamp) {
                    if ($dis_timestamp) {
                        $dis_data = $dtu_data;
                        # timestamp and power preferred from more correct 1PM
                    } else {
                        ($dis_timestamp, $dis_power, $dis_data) =
                            ($dtu_timestamp, $dtu_power, $dtu_data);
                    }
                }
                $dis_power = 0 if 0 < $dis_power
                    && $dis_power < 0.5;  # suppress flicker < 0.5 W on standby
            } else {
                $powers_ok = 0;
                log_warn("taking previous discharge power value $prev_dis_power"
                         ." as no current status data available");
                ++$count_dis_miss;
                $dis_power = $prev_dis_power;
            }
        }
        print "$time: $timestamp ($diff_seconds seconds)\n" if $debug;
        if ($diff_seconds > 1) {
            log_warn("time gap ".++$count_gaps.": $diff_seconds seconds");
            # linear interpolation of missing time and power
            my $power_step = ($power - $prev_power) / $diff_seconds;
            my  $pv_power_step = ( $pv_power -  $prev_pv_power) / $diff_seconds;
            my $chg_power_step = ($chg_power - $prev_chg_power) / $diff_seconds;
            my $dis_power_step = ($dis_power - $prev_dis_power) / $diff_seconds;
            while (--$diff_seconds) {
                $prev_power += $power_step;
                $prev_pv_power += $pv_power_step;
                $prev_chg_power += $chg_power_step;
                $prev_dis_power += $dis_power_step;
                do_each_second(++$prev_timestamp, $powers_ok,
                               $prev_power, "", $prev_pv_power, "",
                               $prev_chg_power, "", $prev_dis_power, "");
            }
        }
        if ($diff_seconds < 0) {
            log_warn("skipping status entry due to negative 3EM unixtime difference: $diff_seconds");
            $timestamp = $prev_timestamp;
        } else {
            do_each_second($timestamp, $powers_ok,
                           $power, ",$data", $pv_power, ",$pv_data",
                           $chg_power, ",$chg_data", $dis_power, ",$dis_data");
        }
    }

    # handled by do_each_second:
    # $prev_pv_power = $pv_power;
    # $prev_chg_power = $chg_power;
    # $prev_dis_power = $dis_power;
    $prev_power = $power;
    $prev_timestamp = $timestamp;
    # $prev = $time;
    usleep(600000) # 0.6 secs; each iteration otherwise takes about .4 seconds
        unless $addr eq "-";
    ($date, $time) = date_time_now();
    printf("%.2f seconds/round\n", $nseconds / $rounds) if $debug;
} while(1);
# while ($count_seconds < MAX_SECONDS); # stop after 1 day at the latest
# while $time ge $prev # not yet wrap around at 24:00:00
#     && $time lt $end_time;

 end:
--$item;
cleanup();

# Local IspellDict: american
# LocalWords: addr pm chg dtu em pvstat chgstat disstat dis
