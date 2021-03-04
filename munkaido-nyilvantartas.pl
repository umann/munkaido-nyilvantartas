#!/usr/bin/perl -w

## no critic(ProhibitEnumeratedClasses)
## no critic(ProhibitPostfixControls)
## no critic(ProhibitCStyleForLoops)
## no critic(ProhibitMagicNumbers)

use strict;
use warnings;

our $VERSION = '0.0.40';

use 5.010;

use utf8;

use Carp qw(cluck confess croak longmess);
use Cwd qw(abs_path);
use Data::Dump qw(dump);
use Digest::MD5 qw(md5_hex);
use File::Basename qw(dirname);
use File::Temp qw(tempfile);
use File::Slurp qw(read_file write_file);
use File::Spec::Functions qw(catfile rel2abs);
use FindBin qw($RealBin);
use Getopt::Long;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;
use Spreadsheet::WriteExcel;
use Spreadsheet::ParseExcel::Utility qw(int2col);
use Time::Local qw(timelocal);
use YAML qw(Dump);

sub terminate;
sub stderr;

local $SIG{__DIE__}  = \&confess;
local $SIG{__WARN__} = \&cluck;

binmode STDERR, ':encoding(UTF-8)';
binmode STDOUT, ':encoding(UTF-8)';
system 'chcp 65001 >nul';

my $D = q{.};  # date delimiter
{
    no warnings 'redefine'; ## no critic(ProhibitNoWarnings)

    sub Spreadsheet::ParseExcel::_convert_col_width {
        my $self        = shift;
        my $excel_width = shift;
        my $pixels      = $excel_width * 7 / 256;
        return $pixels <= 12 ? $pixels / 12 : ($pixels - 5) / 7;
    }
}

my %config_token = (
    dir                 => 'mappa',
    calendar_filesheet  => 'munkanaptár munkalapja',
    employee_filesheet  => 'dolgozók munkalapja',
    workday_filesheet   => 'hétköznapi munkaidő munkalapja',
    saturday_filesheet  => 'szombati munkaidő munkalapja',
    leave_register_file => 'szabadság-nyilvántartó kartonok munkafüzete',
    template_file       => 'sablon munkafüzete',
    output_file         => 'eredmény munkafüzete',
);

my %config_default = (
    dir                 => $RealBin,
    calendar_filesheet  => 'munkaido-nyilvantartas.ADATOK.xls[munkanaptár]',
    employee_filesheet  => 'munkaido-nyilvantartas.ADATOK.xls[dolgozók]',
    workday_filesheet   => 'munkaido-nyilvantartas.ADATOK.xls[hétköznap]',
    saturday_filesheet  => 'munkaido-nyilvantartas.ADATOK.xls[szombat]',
    leave_register_file => 'Szabadsag_kartonok_{ÉÉÉÉ}.xls',
    template_file       => 'munkaido-nyilvantartas.SABLON.xls',
    output_file         => 'munkaido-nyilvantartas.{ÉÉÉÉ}-{HH}.xls ',
);

my %opt;
GetOptions(
    \%opt,
    qw(overwrite start build version config_file=s editor=s nocritic norebuild),
    map { "$_=s" } sort keys %config_token
);
version() if $opt{version};
build()   if $opt{build};

my $NAME         = qr/\S+(?:\x20\S+)*/msx;
my $ANY_DATE_RE0 = qr/\d+\D\d+\D\d+/msx;
my $YM_RE        = qr/(\d{4}\D\d{2})/msx;

my $SECS_PER_DAY = 24 * 60 * 60;

my $_CONFIG = 'konfiguráció';
my $_VALUE  = 'érték';

my $FIRST_COL_NUMBER = 0;
my $FIRST_ROW_NUMBER = 0;

my @_WDAYNAME = qw(
    vasárnap
    hétfő
    kedd
    szerda
    csütörtök
    péntek
    szombat
    vasárnap
);
my @_MONTHNAME = qw(
    THE_ONLY_PLACE_WHERE_IT_IS_A_GOOD_IDEA_TO_START_MONTH_NUMBERING_AT_0
    január
    február
    március
    április
    május
    június
    július
    augusztus
    szeptember
    október
    november
    december
);

my $_MONTH                 = 'hónap';
my $_EXT                   = 'fájlkiterjesztés';
my $_EMPLOYEE_NAME         = 'dolgozónév';
my $_INVALID               = 'érvénytelen';
my $_NO_SUCH               = 'nincs ilyen';
my $_WORKSHEET             = 'munkalap';
my $_ENV_VAR               = 'környezeti változó';
my $_ALREADY_EXISTS        = 'már létezik';
my $_NO                    = 'nincs';
my $_MISSING               = 'hiányzik';
my $_NOT                   = 'nem';
my $_BUT                   = 'hanem';
my $_ONLY_HERE             = 'csak itt';
my $_DATE                  = 'dátum';
my $_DAY_OF_WEEK           = 'a hét napja';
my $_DAYTYPE               = 'nap típusa';
my $_WEEKDAY               = 'hétköznap';
my $_WORKDAY               = 'munkanap';
my $_WORKHOURS             = 'm.órák';
my $_WORKTIME              = 'munkaidő';
my $_SATURDAY              = 'szombat';
my $_SUNDAY                = 'vasárnap';
my $_WEEKEND               = 'hétvége';
my $_VACATION              = 'szabadság';
my $_SICKLEAVE             = 'betegszabadság';
my $_VACATION_ABBR         = 'sz';
my $_SICKLEAVE_ABBR        = 'tp';
my $_START                 = 'kezd';
my $_END                   = 'vége';
my $_VALIDITY              = 'érv.';
my $_ERROR                 = 'hiba';
my $_OFFICIAL_WORK_HOURS   = 'hivatalos munkaórák';
my $_MANDATORY_CONFIG_KEYS = 'kötelező konfig paraméterek';
my $_ENTERED_CONFIG_KEYS   = 'megadott konfig paraméterek';
my $_WORKED_DAYS           = 'hivatalos / ledolgozott munkanapok';
my $_WORKED_HOURS          = 'ledolgozott órák';
my $_OVER_HOURS            = 'túlóra';
my $_VACATION_DAYS         = 'szabadság napok';
my $_SICKLEAVE_DAYS        = 'betegség napok';
my $_PAID_VACATION_HOURS   = 'szabadság után fizetett órák';
my $_PAID_SICKLEAVE_HOURS  = 'betegség után fizetett órák';
my $_VACATION_DAYS_LEFT    = 'megmaradt szabadság napok';
my $_DAYS_AT_START         = 'nyitó szabadság napok';
my $_NICKNAMES             = 'becenevek';
my $_YYYY                  = 'ÉÉÉÉ';
my $_MM                    = 'HH';
my $_TOTAL                 = 'Összesen';
my $_SYS                   = 'RENDSZER';

my $_USR_ERR_BOTH_OFFFICAL_AND_NICKNAME = 'egyszerre hivatalos és becenév';
my $_USR_ERR_SAT_HOLIDAY_SAT_WORK = 'szombati ünnepnap munkabejegyzéssel';
my $_USR_ERR_SAT_WORKDAY_SAT_WORK =
    'szombatra áthelyezett munkanap, szombati munkabejegyzéssel';
my $_USR_ERR_SUN_VACATION = 'szabadság vasárnap';
my $_USR_ERR_SAT_VACATION = 'szabadság szombaton';
my $_USR_ERR_SAT_VACATION_SAT_WORK =
    'szabadság szombaton (+szombati munkabejegyzés)';
my $_USR_ERR_SUN_VACATION_HOLIDAY = 'szabadság ünnepnap (+vasárnap)';
my $_USR_ERR_VACATION_HOLIDAY     = 'szabadság ünnepnap';
my $_USR_ERR_SAT_VACATION_HOLIDAY = 'szabadság ünnepnap (+szombat)';
my $_USR_ERR_SAT_VACATION_HOLIDAY_SAT_WORK =
    'szabadság ünnepnap (+szombat +szombati munkabejegyzés)';
my $_USR_ERR_SAT_WORKDAY_VACATION_SAT_WORK =
'szombatra áthelyezett munkanap, szombati munkabejegyzéssel (+szabadság)';
my $_SYS_ERR_SUN_MOVED_WORKDAY = "$_SYS: vasárnapra áthelyezett munkanap";
my $_SYS_ERR_WEEKDAY_MOVED_WORKDAY =
    "$_SYS: hétköznapra áthelyezett munkanap";
my $_SYS_USR_ERR_SUN_MOVED_WORKDAY_HOLIDAY =
    "$_SYS: vasárnapra áthelyezett munkanap (+szabadság)";
my $_SYS_USR_ERR_WEEKDAY_MOVED_WORKDAY_HOLIDAY =
    "$_SYS: hétköznapra áthelyezett munkanap (+szabadság)";
my $_ERR_CANNOT_WRITE = 'nem lehet írni';
my $_ERR_CANNOT_READ  = 'nem lehet olvasni';

my @_TO_BE_SUMMED = (
    $_WORKED_HOURS, $_PAID_SICKLEAVE_HOURS, $_PAID_VACATION_HOURS,
    $_OVER_HOURS, $_WORKED_DAYS, $_SICKLEAVE_DAYS, $_VACATION_DAYS,
);

my %_DAYTYPE = (
    PUB_HOLIDAY       => 'munkaszüneti nap',
    MOVED_WORKDAY     => 'áthelyezett munkanap',
    MOVED_PUB_HOLIDAY => 'áthelyezett pihenőnap',
);

my $TIME      = time;
my $target_ym = input_target_ym(shift);

my ($target_yyyy, $target_mm) = split /\D/msx, $target_ym;
my $first_ymd_of_target_month = first_ymd_of_month($target_ym);
my $last_ymd_of_target_month  = last_ymd_of_month($target_ym);

my %config         = get_config();
my $unnick         = get_unnick();
my $calendar_hoh   = get_calendar_hoh();
my $vacation_hoh   = get_vacation_hoh();
my $work_hours_hoh = get_work_hours_hoh();

diff_keys(
    $_VACATION   => $vacation_hoh,
    $_WORKTIME   => $work_hours_hoh,
    'asymmetric' => 1
);

write_result();

sub write_result {
    my @employees = sort keys %{$work_hours_hoh};
    my $out_lol;

    my @days =
        days_between($first_ymd_of_target_month, $last_ymd_of_target_month,
        'MONTH');
    (my $ym_name = ym($days[0])) =~ s/^(....)$D(..).*/$1. $_MONTHNAME[$2]/msx;
    push @{$out_lol}, [];  #TITLE
    push @{$out_lol}, [];
    push @{$out_lol}, [ $ym_name, q{}, q{}, map { (q{}, $_, q{}) } @employees ];
    push @{$out_lol},
        [
        $_DATE, $_DAY_OF_WEEK, $_DAYTYPE,
        map { ("$_WORKTIME $_START", "$_WORKTIME $_END", $_WORKHOURS) }
            @employees
        ];

    for my $ymd (@days) {
        push @{$out_lol},
            [
            $ymd, wdayname_of_ymd($ymd), daytype_of_ymd($ymd),
            map { ymd_employee_data3($ymd, $_) } @employees
            ];
    }
    push @{$out_lol},
        map { [] }
        scalar(@days) + 1 .. 31;  # add missing lines for 28 and 30 days months

    my %total = ymd_employee_work_data($target_ym, @employees);
    my %total_by_empl =
        map { $_ => { ymd_employee_work_data($target_ym, $_) } } @employees;
    my $n_workdays = grep { is_workday($_) } @days;
    $total{$_WORKED_DAYS} = "($n_workdays)";
    for my $key ($_WORKED_HOURS, $_OVER_HOURS, $_PAID_SICKLEAVE_HOURS,
        $_PAID_VACATION_HOURS, $_WORKED_DAYS, $_SICKLEAVE_DAYS, $_VACATION_DAYS)
    {
        push @{$out_lol},
            [
            $_TOTAL, $key,
            $total{$key} || 0,
            map { (q{}, q{}, $total_by_empl{$_}{$key} // q{}) } @employees
            ];
    }
    push @{$out_lol}, [
        q{},
        $_VACATION_DAYS_LEFT,
        q{},
        map {
            (   q{}, q{},
                ($vacation_hoh->{$_}{$_DAYS_AT_START} || 0) -
                    ($total_by_empl{$_}{$_VACATION_DAYS} || 0)
                )
        } @employees
    ];
    push @{$out_lol}, [q{}];  # dolgozo alairasa sor

    #print lol_to_txt($out_lol);

    write_xls_via_template(
        {
            lol      => $out_lol,
            template => get_fullpath($config{template_file}),
            outfile  => get_fullpath($config{output_file}),
            col_map  => sub {
                my $fix_cols    = 3;
                my $repeat_cols = 3;
                my $col         = shift;
                $col < $fix_cols
                    ? $col
                    : $fix_cols + ($col - $fix_cols) % $repeat_cols;
            },
            keep_cells => [
                qw(A1 A4 B4 C4 D4 E4 F4 A36 B36 A37 B37 A38 B38 A39 B39 A40 B40 A41 B41 A42 B42 A43 B43 A44 B44)
            ],
        }
    );
    return;
}

sub ymd_employee_data3 {
    my ($ymd, $employee_name) = @_;
    my %ymd_employee_work_data = ymd_employee_work_data($ymd, $employee_name);

    return q{}, q{}, $_VACATION_ABBR  if $ymd_employee_work_data{$_VACATION};
    return q{}, q{}, $_SICKLEAVE_ABBR if $ymd_employee_work_data{$_SICKLEAVE};
    return map { /X/smx ? q{} : $_ }
        map { $_ || q{} }
        @ymd_employee_work_data{ ($_START, $_END, $_WORKED_HOURS) };
}

sub sum_work_data {
    my ($ymds_ar, $employee_names_ar) = @_;
    my %sum;
    for my $employee_name (@{$employee_names_ar}) {
        for my $ymd (@{$ymds_ar}) {
            my %r = ymd_employee_work_data($ymd, $employee_name);
            for (@_TO_BE_SUMMED) {
                $sum{$_} += $r{$_} || 0
                    if +($r{$_} // q{}) =~ /^[1-9][0-9]*$/smx;
            }

            #for (@_TO_BE_KEPT) {
            #    $sum{$_} = $r{$_};
            #}
        }
    }
    return %sum;
}

sub unnick {
    my $x      = shift;
    my $debug  = shift;
    my $return = $unnick->{$x} // $x;
    terminate("$debug: $_INVALID $_EMPLOYEE_NAME: $return")
        if $return !~ /\S\s\S/smx;
    return $return;
}

sub subst_tmpl {
    my $x = shift;
    $x =~ s/[{](?:$_YYYY|YYYY)[}]/$target_yyyy/gmsx;
    $x =~ s/[{](?:$_MM|MM)[}]/$target_mm/gmsx;
    $x =~ s{[%](\w+)[%]}{$ENV{$1} // terminate("$_NO_SUCH $_ENV_VAR: $1")}egmsx;
    return $x;
}

sub ymd_employee_work_data {
    my ($ymd, @employee_names) = @_;

    # alkalmazottak osszesitett adatai:
    if (1 != @employee_names) {
        return sum_work_data([$ymd], \@employee_names);
    }

    my $employee_name = shift @employee_names;

    # egesz honap osszesitett adatai:
    if ($ymd =~ /^....$D..$/msx) {
        return sum_work_data(
            [ days_between(first_ymd_of_month($ymd), last_ymd_of_month($ymd)) ],
            [$employee_name]
        );
    }
    return ymd_employee_work_data_1($ymd, $employee_name);
}

sub ymd_employee_work_data_1 {
    my ($ymd, $employee_name) = @_;

    state $subcase_hr = get_subcase_hr();

    my ($start, $end) =
        @{ $work_hours_hoh->{$employee_name}{$ymd} }{ $_START, $_END };
    my $wdayname = wdayname_of_ymd($ymd);

    my $subcase = (
          $vacation_hoh->{$employee_name}{$_VACATION}{$ymd}  ? 'S'
        : $vacation_hoh->{$employee_name}{$_SICKLEAVE}{$ymd} ? 'B'
        :                                                      '0'
        )
        . (
        is_in($calendar_hoh->{$ymd}{$_DAYTYPE}, $_DAYTYPE{MOVED_WORKDAY})
        ? 'X'
        : is_in(
            $calendar_hoh->{$ymd}{$_DAYTYPE}, $_DAYTYPE{PUB_HOLIDAY},
            $_DAYTYPE{MOVED_PUB_HOLIDAY}
            )
        ? '1'
        : '0'
        )
        . (
          $wdayname eq $_SUNDAY ? '00'
        : $wdayname eq $_SATURDAY ? ($start ? '11' : '10')
        :                           '01'
        );

    my %return;

    @return{
        (   $_ERROR,                $_OFFICIAL_WORK_HOURS,
            $_WORKED_HOURS,         $_PAID_VACATION_HOURS,
            $_PAID_SICKLEAVE_HOURS, $_OVER_HOURS
        )
    } = @{ $subcase_hr->{$subcase} // terminate("$_SYS: $subcase") };
    terminate(Dump $return{$_ERROR},
        $ymd, $employee_name, $vacation_hoh->{$employee_name})
        if delete $return{$_ERROR};

    my $workhours = workhours_minus_lunchbreak($start, $end);
    for (values %return) {
        $_ = $workhours ? $_ * $workhours : undef;
    }

    $return{$_SICKLEAVE_DAYS} = $return{$_PAID_SICKLEAVE_HOURS} ? 1 : 0;
    $return{$_SICKLEAVE}      = $return{$_PAID_SICKLEAVE_HOURS} ? 1 : 0;
    $return{$_VACATION_DAYS}  = $return{$_PAID_VACATION_HOURS}  ? 1 : 0;
    $return{$_VACATION}       = $return{$_PAID_VACATION_HOURS}  ? 1 : 0;
    $return{$_WORKED_DAYS} =
        is_workday($ymd) && $return{$_WORKED_HOURS} ? 1 : 0;
    @return{ ($_START, $_END) } = ($start, $end) if $return{$_WORKED_HOURS};
    return %return;
}

sub get_subcase_hr {

    # Esetek:
    # ..00 vasárnap
    # ..01 hétfőtől péntekig
    # ..10 szombat, munkabejegyzés nélkül
    # ..11 szombat, munkabejegyzéssel
    # .0. naptár szerint: nincs külön bejegyzése
    # .1. naptár szerint: ünnepnap vagy áthelyezett pihenőnap
    # .a. naptár szerint: áthelyezett munkanap
    # 0.. nem szabadság
    # S.. szabadság
    # B.. betegszabadság
    #
    # Aleset => [
    #    hiba,
    #    hivatalos munkaóra szorzó,
    #    ledolgozott munkaóra szorzó,
    #    kifizetett szabadság munkaóra szorzó,
    #    kifizetett betegszabadság munkaóra szorzó,
    #    túlóra szorzó
    # ]
    return {
        '0000' => [ 0, 0, 0, 0, 0, 0 ],  # normál vasárnap
        '0001' => [ 0, 1, 1, 0, 0, 0 ],  # normál hétfőtől péntekig
        '0010' => [ 0, 0, 0, 0, 0, 0 ]
        ,  # normál szombat, munkabejegyzés nélkül
        '0011' => [ 0, 0, 1, 0, 0, 1 ],  # normál szombat, munkabejegyzéssel
        '0100' => [ 0, 0, 0, 0, 0, 0 ],  # ünnepnap vasárnap
        '0101' => [ 0, 0, 0, 0, 0, 0 ],  # ünnepnap hétfőtől péntekig
        '0110' => [ 0, 0, 0, 0, 0, 0 ]
        ,  # ünnepnap szombat, munkabejegyzés nélkül
        '0a10' => [ 0, 0, 0, 0, 0, 0 ]
        ,  # áthelyezett munkanap, szombati munkabejegyzés nélkül
        'S001' => [ 0, 1, 0, 1, 0, 0 ]
        ,  # szabadság hétfőtől péntekig valamelyik napon
        'Sa10' => [ 0, 1, 0, 1, 0, 0 ]
        , # szabadság szombatra áthelyezett munkanapon szombati munkabejegyzés nélkül
        'B001' => [ 0, 1, 0, 0, 1, 0 ]
        ,  # betegszabadság hétfőtől péntekig valamelyik napon
        'Ba10' => [ 0, 1, 0, 0, 1, 0 ]
        , # betegszabadság szombatra áthelyezett munkanapon szombati munkabejegyzés nélkül
        '0a00' => [$_SYS_ERR_SUN_MOVED_WORKDAY]
        ,  # áthelyezett munkanap vasárnap
        '0a01' => [$_SYS_ERR_WEEKDAY_MOVED_WORKDAY]
        ,  # áthelyezett munkanap hétfőtől péntekig
        '0111' => [$_USR_ERR_SAT_HOLIDAY_SAT_WORK]
        ,  # ünnepnap szombat, munkabejegyzéssel
        '0a11' => [$_USR_ERR_SAT_WORKDAY_SAT_WORK]
        ,  # áthelyezett munkanap szombat, munkabejegyzéssel
        map {
            (   "${_}000" => [ $_USR_ERR_SUN_VACATION . " ($_)" ]
                ,  # (beteg)szabadság vasárnap
                "${_}010" => [ $_USR_ERR_SAT_VACATION . " ($_)" ]
                ,  # (beteg)szabadság szombat, munkabejegyzés nélkül
                "${_}011" => [ $_USR_ERR_SAT_VACATION_SAT_WORK . " ($_)" ]
                ,  # (beteg)szabadság szombat, munkabejegyzéssel
                "${_}100" => [ $_USR_ERR_SUN_VACATION_HOLIDAY . " ($_)" ]
                ,  # (beteg)szabadság ünnepnap vasárnap
                "${_}101" => [ $_USR_ERR_VACATION_HOLIDAY . " ($_)" ]
                ,  # (beteg)szabadság ünnepnap hétfőtől péntekig
                "${_}110" => [ $_USR_ERR_SAT_VACATION_HOLIDAY . " ($_)" ]
                , # (beteg)szabadság ünnepnap szombat, munkabejegyzés nélkül
                "${_}111" =>
                    [ $_USR_ERR_SAT_VACATION_HOLIDAY_SAT_WORK . " ($_)" ]
                ,  # (beteg)szabadság ünnepnap szombat, munkabejegyzéssel
                "${_}a11" =>
                    [ $_USR_ERR_SAT_WORKDAY_VACATION_SAT_WORK . " ($_)" ]
                , # (beteg)szabadság szombatra áthelyezett munkanap szombat, munkabejegyzéssel
                "${_}a00" =>
                    [ $_SYS_USR_ERR_SUN_MOVED_WORKDAY_HOLIDAY . " ($_)" ]
                ,  # (beteg)szabadság szombatra áthelyezett munkanap vasárnap
                "${_}a01" =>
                    [ $_SYS_USR_ERR_WEEKDAY_MOVED_WORKDAY_HOLIDAY . " ($_)" ]
                , # (beteg)szabadság szombatra áthelyezett munkanap hétfőtől péntekig
                )
        } qw(S B)
    };
}

sub get_fullpath {
    my $file = shift // return;
    my $dir  = shift // $config{dir};
    my $worksheet = $file =~ s/(\[.*)$//msx ? $1 : q{};
    return +(
          $file =~ m{^(?://|\\\\|[a-zA-Z]:[/\\])}msx
        ? $file
        : File::Spec->catfile($dir, $file)
    ) . $worksheet;
}

### PROPRIETARY EXCEL STUFF {

sub get_config {
    my $config_file = $opt{config_file} // do {
        (my $base = $0) =~ s/.(?:pl|exe)$//msx || terminate($0);
        "$base.KONFIG.xls";
    };  # do
    my %return;
    if (-f $config_file) {
        my $filesheet = $config_file . "[$_CONFIG]";
        my $hoh       = filesheet_to_hoh($filesheet);
        diff_keys(
            $_MANDATORY_CONFIG_KEYS => [ values %config_token ],
            $filesheet              => $hoh,
            do                      => \&terminate
        );
        my $dir = dirname abs_path $config_file;
        for (keys %{$hoh}) {
            my $key = { reverse %config_token }->{$_};
            my $val = $hoh->{$_}{$_VALUE};
            $return{$key} = $val;
        }
        $return{dir} =
            get_fullpath($return{dir}, dirname abs_path $config_file);
    }

    elsif ($opt{config_file}) {
        terminate "$_NO $opt{config_file}";
    }
    for (sort keys %config_token) {
        $return{$_} = $opt{$_} // $return{$_} // $config_default{$_};
    }
    for (values %return) {
        $_ = subst_tmpl($_);
    }

    # diff_keys(
    # $_MANDATORY_CONFIG_KEYS => [ values %config_token ],
    # $_ENTERED_CONFIG_KEYS   => \%return,
    # do                 => \&terminate,
    # );
    return %return;
}

sub get_unnick {
    my $hr;  #return value
    my $file = get_fullpath($config{employee_filesheet});
    my $hoh  = filesheet_to_hoh($file);
    for my $employee_name (keys %{ $hoh // {} }) {
        for my $nick (split /\s*,\s*/msx,
            $hoh->{$employee_name}{$_NICKNAMES} // q{})
        {
            terminate(
                "$file: $_NICKNAMES: $nick: $hr->{$nick} <> $employee_name")
                if $hr->{$nick};
            $hr->{$nick} = $employee_name;
        }
    }
    return $hr;
}

sub get_calendar_hoh {

    my $filesheet = get_fullpath($config{calendar_filesheet});
    my $hoh       = filesheet_to_hoh($filesheet);              # return value
    transform_keys_in_place(\&format_date, $hoh, $filesheet);
    diff_keys(
        $_DAYTYPE  => [ values %_DAYTYPE ],
        $filesheet => [ map { $_->{$_DAYTYPE} } values %{$hoh} ],
        do         => \&terminate
    );

    return $hoh;
}

sub get_vacation_hoh {

    my $hoh;  # return value
    my $file        = get_fullpath($config{leave_register_file});
    my $workbookhoh = file_to_workbookhoh($file);

    for my $sheet_name (sort keys %{$workbookhoh}) {
        my $filesheet = "$file\[$sheet_name]";
        next if $sheet_name =~ /alapadatok|munkanapok/msx;
        my $employee_name = transform(\&unnick, $sheet_name, $filesheet);
        my $sheethoh = $workbookhoh->{$sheet_name};

## no critic(ProhibitComplexRegexes)
        my $txt = sheethoh_to_txt($sheethoh);
        terminate("$filesheet syntax\n" . $txt) if $txt !~ m{\A
            $NAME\n
            [^\n]*\n
            (?:[^\n\t]*\t){3}\d+\n
            (?:$ANY_DATE_RE0\t$ANY_DATE_RE0\t[^\n]*\n)*
            (?:\t\t\t\d+(?:\t[^\n]*)?\n)*
         \Z}msx;

## use critic(ProhibitComplexRegexes)

        my $employee_name2 =
            transform(\&unnick, get_cell($sheethoh, 0, 0), $filesheet);
        if ($employee_name ne $employee_name2) {
            terminate("$filesheet: $employee_name <> $employee_name2");
        }
        if ($employee_name !~ /\S\s+\S/msx) {
            terminate("$filesheet: $employee_name");
        }
        my $dat = get_cell($sheethoh, 2, 3);
        terminate(
"$filesheet: $sheet_name: $_DAYS_AT_START $_MISSING/$_INVALID (`$dat`)"
        ) if $dat !~ /^\d+$/msx;
        $hoh->{$employee_name}{$_DAYS_AT_START} = $dat;
    ROW: for my $row (sort { $a <=> $b } keys %{$sheethoh}) {
            next if $row < 3 + $FIRST_ROW_NUMBER;
            my %leave;
            my $leave_type =
                get_cell($sheethoh, $row, 4) ? $_SICKLEAVE : $_VACATION;

            my $col = 0;
            for ($_START, $_END) {
                $leave{$_} = format_date(get_cell($sheethoh, $row, $col))
                    // last ROW;
                terminate("$employee_name: "
                        . "$leave_type $_ $leave{$_} $_NOT $_WORKDAY, $_BUT "
                        . daytype_of_ymd($leave{$_}))
                    if !is_workday($leave{$_});
                $col++;
            }
            for my $ymd (grep { is_workday($_) }
                days_between($leave{$_START}, $leave{$_END}, $filesheet))
            {
                $hoh->{$employee_name}{$leave_type}{$ymd} = 1;
            }
        }
    }
    return $hoh;
}

sub get_work_hours_hoh {
    my $hoh;  # return value

    my $workday_filesheet = get_fullpath($config{workday_filesheet});
    my $workday_sheethoh  = filesheet_to_hoh($workday_filesheet);
    for my $employee_name (sort keys %{$workday_sheethoh}) {
        my $rec = $workday_sheethoh->{$employee_name};
        for (my $cnt = 0 ; $rec->{"$_START.$cnt"} ; $cnt++) {
            my $start =
                maxs_nonempty(format_date($rec->{"$_VALIDITY $_START.$cnt"}),
                $first_ymd_of_target_month);
            my $end =
                mins_nonempty(format_date($rec->{"$_VALIDITY $_END.$cnt"}),
                $last_ymd_of_target_month);
            for my $ymd (grep { is_workday($_) }
                days_between($start, $end, $workday_filesheet))
            {
                $hoh->{$employee_name}{$ymd} =
                    { map { $_ => $rec->{"$_.$cnt"} } $_START, $_END };
            }
        }
    }
    my $saturday_filesheet = get_fullpath($config{saturday_filesheet});
    my $saturday_sheethoh  = filesheet_to_sheethoh($saturday_filesheet);

    my @sat_employee_names;
    for (
        my $col = 3 + $FIRST_COL_NUMBER ;
        ($saturday_sheethoh->{$FIRST_ROW_NUMBER}{$col} // q{}) =~ /\S/msx ;
        $col += 2
        )
    {
        push @sat_employee_names, $saturday_sheethoh->{$FIRST_ROW_NUMBER}{$col};
    }
    my $sat = sheethoh_to_hoh($saturday_sheethoh, 1, $saturday_filesheet);
    transform_keys_in_place(\&format_date, $sat);

    for my $key (sort keys %{$sat}) {
        my $ymd      = format_date($key);
        my $rec      = $sat->{$ymd};
        my $wdayname = wdayname_of_ymd($ymd);
        terminate("$saturday_sheethoh: $ymd $wdayname <> $_SATURDAY")
            if $wdayname ne $_SATURDAY;
        for (my $cnt = 0 ; exists $rec->{"$_START.$cnt"} ; $cnt++) {
            my $employee_name = $sat_employee_names[$cnt] // terminate(
                "$saturday_filesheet: $ymd: #$cnt: $_NO $_EMPLOYEE_NAME");
            next if !$rec->{"$_START.$cnt"};
            $hoh->{$employee_name}{$ymd} =
                { map { $_ => $rec->{"$_.$cnt"} } $_START, $_END };
        }
    }
    transform_keys_in_place($unnick, $hoh,
        "$workday_filesheet / $saturday_filesheet");

    return $hoh;
}

### PROPRIETARY EXCEL STUFF }

### PROPRIETARY DATE/TIME STUFF {

sub is_workday {
    my $ymd = shift;
    return is_in(daytype_of_ymd($ymd), $_DAYTYPE{MOVED_WORKDAY}, $_WEEKDAY);
}

sub daytype_of_ymd {
    my $ymd = shift;
    state $cache;
    return $cache->{$ymd} //= $calendar_hoh->{$ymd}{$_DAYTYPE} // (
        is_in(wdayname_of_ymd($ymd), $_SATURDAY, $_SUNDAY)
        ? $_WEEKEND
        : $_WEEKDAY
    );
}

sub workhours_minus_lunchbreak {
    my ($start_hm, $end_hm) = @_;
    return if !$start_hm;
    if ($start_hm eq 'X' && $end_hm =~ /X[+](.+)/msx) {
        return hm2h($1);
    }
    my $return = hm2h($end_hm) - hm2h($start_hm);
    terminate("$start_hm, $end_hm") if $return < 1;
    $return -= 0.5 if $return > 6;  # lunch break
    return sprintf '%.2f', $return;

}

sub input_target_ym {
    my $input = shift;
    my %ym;
    $ym{curr} = ym($TIME);
    $ym{prev} = ym($TIME - 31 * $SECS_PER_DAY);
    my $days_left_from_curr_month =
        diff_days(last_ymd_of_month($TIME), ymd($TIME));
    $ym{default} = $days_left_from_curr_month < 3 ? $ym{curr} : $ym{prev};
    if (!$input) {
        stderr(ucfirst "$_MONTH ($ym{default}): ");
        ($input = <>) =~ s/^\s*$/default/msx;
    }
    $input = $ym{$input} // $input;
    my ($yy, $mm) = $input =~ /^\D*(?:20)?(\d{2})\D*(\d{1,2})\D*$/msx
        or terminate("$_INVALID: $input");
    return sprintf "%04d$D%02d", 2000 + $yy, $mm;
}

### PROPRIETARY DATE/TIME STUFF }

### GENERAL DATE/TIME STUFF {

sub hm2h {
    my $hm = shift;
    my ($h, $m) = $hm =~ /^(\d{1,2}):(\d{2})$/msx or terminate($hm);
    return sprintf '%.2f', $h + $m / 60;
}

sub ym {
    my $ts_or_ymd_or_ym = shift;
    my $ymd =
          $ts_or_ymd_or_ym =~ /($ANY_DATE_RE0)/msx ? format_date($1)
        : $ts_or_ymd_or_ym =~ /($YM_RE)/msx        ? format_date("$1${D}01")
        :                                            ts2ymd($ts_or_ymd_or_ym);
    return substr $ymd, 0, 7;
}

sub ymd {
    my $ts_or_ymd = shift;
    return $ts_or_ymd =~ /($ANY_DATE_RE0)/msx
        ? format_date($1)
        : ts2ymd($ts_or_ymd);
}

sub first_ymd_of_month {
    my $ts_or_ymd_or_ym = shift;
    return ym($ts_or_ymd_or_ym) . "${D}01";
}

sub last_ymd_of_month {
    my $ts_or_ymd_or_ym         = shift;
    my $first_ymd_of_month      = first_ymd_of_month($ts_or_ymd_or_ym);
    my $a_ymd_in_next_month     = add_days($first_ymd_of_month, 31);
    my $first_ymd_of_next_month = first_ymd_of_month($a_ymd_in_next_month);
    my $last_ymd_of_month       = add_days($first_ymd_of_next_month, -1);
    return $last_ymd_of_month;

#return add_days(first_ymd_of_month(add_days(first_ymd_of_month(ymd($ts_or_ymd_or_ym)), 31)), -1)
}

sub days_between {
    my ($ymd1, $ymd2, $debug) = @_;

    return if $ymd1 gt $ymd2;

    my @return;
    my $max  = 40;
    my $prev = q{};
    for (my $ymd = $ymd1 ; $ymd le $ymd2 ; $ymd = add_days($ymd, 1)) {
        push @return, $ymd;
        terminate(join q{ }, @return) if !$max-- || $prev eq $ymd;
        $prev = $ymd;
    }
    return @return;
}

sub add_days {
    my ($ymd, $days) = @_;

    return ts2ymd(ymd2ts($ymd) + $days * $SECS_PER_DAY);
}

sub diff_days {
    my ($ymd1, $ymd2) = @_;

    return +(ymd2ts($ymd1) - ymd2ts($ymd2)) / $SECS_PER_DAY;
}

sub ts2ymd {
    my $ts = shift // terminate("$_NO timestamp");
    terminate("$_INVALID timestamp $ts") if $ts !~ /^(?:0|[1-9][0-9]{0,9})$/msx;
    my ($mday, $mon0, $year19) = (localtime $ts)[ 3 .. 5 ];
    return sprintf_y_m_d($year19 + 1900, $mon0 + 1, $mday);
}

sub sprintf_y_m_d {
    my ($y, $m, $d) = @_;
    return sprintf "%04d$D%02d$D%02d", $y, $m, $d;
}

sub ymd2ts {
    my $ymd = shift;
    my ($year, $mon1, $mday) = split /\D/msx, $ymd;

    # fun fact: if you set hour to 0 instead of 12,
    # add_days('2021-10-25', 1) will return the same '2021-10-25'
    return timelocal(0, 0, 12, $mday, $mon1 - 1, $year);
}

sub wdayname_of_ymd {
    my $ymd = shift;
    return $_WDAYNAME[ day_of_week($ymd) ];
}

sub day_of_week {
    my $ymd = shift;
    return +(localtime ymd2ts($ymd))[6];
}

sub format_date {
    my $x = shift // return;
    my $debug = shift;
    return if $x !~ /\S/msx;
    return $x
        if $x =~
s/^D*(\d{1,2})\D+(\d{1,2})\D+(\d\d\d\d)\D*/sprintf_y_m_d($3, $1, $2)/emsx
        ;  # xls # 1/4/2021
    return $x
        if $x =~
        s/^\D*(\d\d\d\d)\D*(\d\d)\D*(\d\d)\D*/sprintf_y_m_d($1, $2, $3)/emsx
        ;  # HU
     #return $x if $x =~ s/^\D*(\d\d)\D+(\d\d)\D+(\d\d\d\d)\D*/sprintf_y_m_d($3, $2, $1)/emsx;                                        # SK
    terminate("$debug: $_INVALID $_DATE: $x");
    return;  # for perlcritic
}

### GENERAL DATE/TIME STUFF }

### GENERAL SPREADSHEET/TABLE STUFF {

sub sheethoh_to_hoh {
    my $sheethoh         = shift;
    my $skip_rows_on_top = shift // 0;
    my $debug            = shift;
    my $hoh;
    my $header;
    my %colname_occurences;

ROW: for my $row (sort { $a <=> $b } keys %{$sheethoh}) {
        my %current_colname_count;
        next if $skip_rows_on_top-- > 0;
        my $row_hr = $sheethoh->{$row};

        # first (unskipped) row is header (column names)
        if (!$header) {
            $header = { map { $_ => $row_hr->{$_} } keys %{$row_hr} };
            for my $colname (grep { /\S/msx } values %{$header}) {
                $colname_occurences{$colname}++;
            }
            next;
        }
        my $key;  # of row (first column)
    COL: for my $col (sort { $a <=> $b } keys %{$row_hr}) {
            my $val = $row_hr->{$col};
            my $cellname = row_col_to_cell($row, $col);  # for debug

            # first column is key of row
            if (!defined $key) {
                $key = $val // q{};
                next ROW
                    if $key !~ /\S/msx;  # skip rows with missing first column
                terminate("$debug!$cellname ($header->{$col}): $key 2x")
                    if $hoh->{$key};     # duplicated first column value
                next;
            }
            my $colname = $header->{$col} // q{};

            # check missing column name:
            if ($colname !~ /\S/msx) {
                next COL if $val !~ /\S/msx;  # value is empty, too, ok
                terminate(
                    "$debug!$cellname: value is non-empty (`$val`),",
                    "column name should not be empty (`$colname`)"
                );
            }
            if ($colname_occurences{$colname} > 1) {
                $current_colname_count{$colname} //= 0;
                if (   exists $hoh->{$key}
                    && exists $hoh->{$key}
                    { $colname . q{.} . $current_colname_count{$colname} })
                {
                    $current_colname_count{$colname}++;
                }
                $colname = $colname . q{.} . $current_colname_count{$colname};
            }
            $hoh->{$key}{$colname} = $val;

        }
    }
    return $hoh;
}

sub get_cell {
    my ($sheet, $row, $col) = @_;
    return $sheet->{ $FIRST_ROW_NUMBER + $row }{ $FIRST_COL_NUMBER + $col };
}

sub file_sheet_names_to_sheethohs {
    my ($file, @sheet_names) = @_;
    my $workbookhoh = file_to_workbookhoh($file);
    my @return      = map {
        $workbookhoh->{$_} // terminate("$_NO_SUCH $_WORKSHEET: $file" . "[$_]")
    } @sheet_names;
    return 1 == @return ? $return[0] : @return;
}

sub filesheet_to_hoh {
    my $file_sheet = shift;
    return sheethoh_to_hoh(
        filesheet_to_sheethoh($file_sheet, undef, $file_sheet));
}

sub filesheet_to_sheethoh {
    my $filesheet = shift;
    my ($file, $sheet_name) = $filesheet =~ /^(.+)\[(.*)\]$/msx
        or terminate("$_INVALID $_WORKSHEET: `$filesheet`");
    my $workbookhoh = file_to_workbookhoh($file);
    return $workbookhoh->{$sheet_name}
        // terminate("$_NO_SUCH $_WORKSHEET: $file" . "[$sheet_name]");
}

sub file_to_workbook {
    my $file = shift // terminate();

    $file =~ /[.]xls$/imsx or terminate("$_INVALID $_EXT: $file");
    return eval {
        my $parser = Spreadsheet::ParseExcel->new();
        $parser->parse($file) // die $parser->error() . "\n";
    } // terminate("$_ERR_CANNOT_READ: $file\n$@");
}

sub file_to_workbookhoh {
    my $file = shift // terminate();

    my $workbook = file_to_workbook($file);

    my $workbookhoh;  # return value

    for my $worksheet ($workbook->worksheets()) {
        next if $worksheet->{SheetHidden};
        my $sheet_name = $worksheet->get_name();
        my ($row_min, $row_max) = $worksheet->row_range();
        my ($col_min, $col_max) = $worksheet->col_range();
        terminate(
            "$file: col_min = $col_min < $FIRST_COL_NUMBER = FIRST_COL_NUMBER")
            if $col_min < $FIRST_COL_NUMBER;
        terminate(
            "$file: row_min = $row_min < $FIRST_ROW_NUMBER = FIRST_ROW_NUMBER")
            if $row_min < $FIRST_ROW_NUMBER;
        for my $row ($row_min .. $row_max) {
            for my $col ($col_min .. $col_max) {
                my $cell = $worksheet->get_cell($row, $col) // next;
                my $val = $cell->value() // next;
                $val =~ s/^\h*(.*?)\h*$/$1/gmsx;  # lrtrim
                $workbookhoh->{$sheet_name}{$row}{$col} = $val;
            }
        }
    }
    return $workbookhoh;
}

sub write_xls_via_template {
    my $p = shift;

    if (-e $p->{outfile} && !$opt{overwrite}) {
        terminate("$_ALREADY_EXISTS: $p->{outfile}");
    }

    my %keep_cell = map { $_ => 1 } @{ $p->{keep_cells} // [] };
    my $lol = $p->{lol} // terminate();

    # Open the template with SaveParser
    my $parser   = Spreadsheet::ParseExcel::SaveParser->new;
    my $template = $parser->Parse($p->{template})
        // terminate("$_ERR_CANNOT_READ: $p->{template}");
    my $sheet      = 0;
    my $format0    = $template->{Worksheet}[$sheet]{Cells}[1][0]{FormatNo};
    my $new_maxcol = 0;
    for my $row (0 .. $#{$lol}) {
        $new_maxcol = $#{ $lol->[$row] } if $new_maxcol <= $#{ $lol->[$row] };
    }

    for my $row (0 .. max($#{$lol}, $template->{Worksheet}[$sheet]{MaxRow})) {
        for my $col (
            0 .. max($#{ $lol->[$row] }, $template->{Worksheet}[$sheet]{MaxCol})
            )
        {
            my $tmpl_col = transform($p->{col_map}, $col);
            my $tmpl_row = transform($p->{row_map}, $row);
            my $val = $lol->[$row][$col] // q{};
            my $tmpl_cell_obj =
                $template->{Worksheet}[$sheet]{Cells}[$tmpl_row][$tmpl_col];
            my $format = $tmpl_cell_obj->{FormatNo};
            if ($col <= $new_maxcol) {
                if ($keep_cell{ row_col_to_cell($tmpl_row, $tmpl_col) }) {
                    $val = $tmpl_cell_obj->{_Value};
                }
                $template->AddCell(0, $row, $col, $val, $format);
            }
            else {
                $format = 0;
            }
        }
    }

    my $workbook = eval { $template->SaveAs($p->{outfile}) }
        // terminate("$_ERR_CANNOT_WRITE: $p->{outfile}");
    $workbook->close();
    system "start $p->{outfile}" if $opt{start};
    return;
}

sub row_col_to_cell {
    my ($row, $col) = @_;
    return int2col($col - $FIRST_COL_NUMBER) . ($row - $FIRST_ROW_NUMBER + 1);
}

sub sheethoh_to_txt {
    my $sheethoh = shift;
    return lol_to_txt(sheethoh_to_lol($sheethoh));
}

sub sheethoh_to_lol {
    my $sheethoh = shift;
    my $lol;
    for my $row (sort { $a <=> $b } keys %{$sheethoh}) {
        my $row_hr = $sheethoh->{$row};
        for my $col (sort { $a <=> $b } keys %{$row_hr}) {
            $lol->[ $row - $FIRST_ROW_NUMBER ][ $col - $FIRST_COL_NUMBER ] =
                $row_hr->{$col};
        }
    }
    return $lol;
}

sub lol_to_txt {
    my $lol = shift;
    my $txt = join q{}, map {
        join("\t", map { $_ // q{} } @{$_}) . "\n"
    } @{$lol};
    $txt =~ s/\t+$//gmsx;       # cut trailing empty cells of lines
    $txt =~ s/\n+\z//msx;       # cut trailing empty rows
    $txt =~ s/(?<=.)\z/\n/msx;  # add trailing NL only if any content
    return $txt;
}

### GENERAL SPREADSHEET/TABLE STUFF }

### GENERAL STUFF {

sub max {
    my ($x, $y) = @_;
    return $x > $y ? $x : $y;
}

sub stderr {
    my @msg = @_;
    return out(*STDERR, @msg);
}

sub stdout {
    my @msg = @_;
    return out(*STDOUT, @msg);
}

sub out {
    my ($fh, @msg) = @_;
    print {$fh} @msg or confess $!;
    return;
}

sub is_in {
    my ($elem, @in) = @_;
    return grep { ($elem // q{}) eq $_ } @in;
}

sub diff_keys {
    my ($name1, $hr1, $name2, $hr2, %p) = @_;
    for ($hr1, $hr2) {
        $_ = { map { $_ => 1 } @{$_} } if 'ARRAY' eq ref;
    }
    my $msg = diff_keys_one_side($name1, $hr1, $name2, $hr2);
    $msg .= diff_keys_one_side($name2, $hr2, $name1, $hr1) if !$p{asymmetric};
    ($p{do} // \&CORE::warn)->($msg) if $msg;
    return;
}

sub diff_keys_one_side {
    my ($name1, $hr1, $name2, $hr2) = @_;
    if (my @miss2 = grep { !defined $hr2->{$_} } sort keys %{$hr1}) {
        return
              "$name2: $_NO: "
            . join(q{, }, @miss2)
            . " ($_ONLY_HERE: $name1)\n";
    }
    return q{};
}

sub transform_keys_in_place {
    my ($transformator, $hr, $debug) = @_;
    for my $key (sort keys %{$hr}) {
        my $processed_key = transform($transformator, $key) // next;
        next if $processed_key eq $key;
        if (exists $hr->{$processed_key}) {

            #terminate("$debug: $processed_key <> $key " . dump $hr);
            if (ref $hr->{$processed_key} eq 'HASH') {
                $hr->{$key} = { %{ $hr->{$key} }, %{ $hr->{$processed_key} } };
            }
            elsif (ref $hr->{$processed_key} eq 'ARRAY') {
                $hr->{$key} = [ @{ $hr->{$key} }, @{ $hr->{$processed_key} } ];
            }
            else {
                terminate("$debug: ", dump $hr->{$processed_key}, $hr->{$key});
            }
        }
        $hr->{$processed_key} = delete $hr->{$key};
    }
    return;  # in-place, no return value
}

sub transform {
    my ($transformator, @values) = @_;
    return if 0 == grep { defined } @values;
    return wantarray ? @values : $values[0] if !defined $transformator;
    my $ref = ref $transformator;
    return $transformator->(@values)          if 'CODE' eq $ref;
    return $transformator->{ $values[0] }     if 'HASH' eq $ref;
    return $values[0] =~ /$transformator/gmsx if 'Regexp' eq $ref;
    if ('ARRAY' eq ref $transformator) {
        my ($sub, @args) = @{$transformator};
        return $sub->(@args) if 'CODE' eq ref $sub;
    }

    terminate("$_INVALID transformator:" . dump $ref, @values);
    return;  # fir perlcritic
}

sub terminate {
    my @msg      = @_;
    my $longmess = undump(longmess);
    $longmess =~
        s{[^\n]*(PAR[.]pm|PAR::|line\s0|at\s-e\s)[^\n]*\n?}{}gmsx;  # PAR
    $longmess =~ s{script/}{}gmsx;                                  # PAR

    (my $base = $0) =~ s{.*?([^/\\]+)[.]\w+$}{$1}msx;
    $longmess =~ s/\sat\s\S*$base[.]\w+//gmsx;
    $longmess =~ s/\bmain:://gmsx;
    $longmess =~ s/\bcalled\s//gmsx;
    $longmess =~ s/(?:ARRAY|HASH)[(]0x[0-9a-f]+[)]/.../gmsx;
    $longmess =~ s/^\h+//gmsx;
    stderr(
        dump {
            stacktrace => [ split /\n/msx, $longmess ],
            options    => \%opt,
            config     => \%config
        }
    );
    stderr("\n*** HIBA: ", @msg, "\n\nNyomj ENTER-t\n");
    my $dummy = <>;
    exit -1;
}

sub undump {
    my $x = shift;
    $x =~ s/\\\\/\\/gmsx;
    $x =~ s/\\x([0-9a-f]{2})/sprintf '%s', hex $1/egmsx;
    $x =~ s/\\x[{]([0-9a-f]+)[}]/sprintf '%c', hex $1/egmsx;
    return $x;
}

sub maxs_nonempty {
    my ($x, $y) = @_;
    return $x if is_empty($y);
    return $y if is_empty($x);
    return $x gt $y ? $x : $y;
}

sub mins_nonempty {
    my ($x, $y) = @_;
    return $x if is_empty($y);
    return $y if is_empty($x);
    return $x lt $y ? $x : $y;
}

sub is_empty {
    my $x = shift;
    return +($x // q{}) !~ /\S/msx;
}

sub _qx {
    my $cmd = shift;
## no critic(ProhibitPunctuationVars ProhibitBacktickOperators)
    my $retval = qx($cmd);
    my $exit   = $? >> 8;
## use critic(ProhibitPunctuationVars ProhibitBacktickOperators)
    confess "$cmd\nexit $exit\n$retval" if $exit;
    return $retval;
}

### GENERAL STUFF }

### BUILD {

sub version {
    print "$VERSION\n" or confess $!;
    exit;
}

sub build {
    if (my $editor = $opt{editor} // $ENV{EDITOR} // 'notepad++.exe') {
        terminate "$editor running" if 0 <= index _qx('tasklist'), $editor;
    }

    my $pl = $0;
    (my $exe = $pl) =~ s/.pl$/.exe/msx || confess "Cannot build from $pl";

    if (!$opt{nocritic}) {
        _qx("perlcritic --brutal --verbose 8 $pl");
    }

    my $content = read_file $pl, { binmode => ':raw' };
    my $curr_version = $VERSION . q{-} . md5_hex($content);

    if (-e $exe) {
        (my $v = _qx("$exe --version")) =~ s/\s//gmsx;
        if ($v eq $curr_version) {
            stderr("Same version $curr_version for $pl and $exe");
            exit;
        }
        unlink $exe or confess "unlink $exe: $!";
    }

    $content =~ s{^(our.\$VERSION.=.'(?:[0-9]+[.]){2})([0-9]+)(')}
      {$1 . ($2 + 1) . $3}emsx
        or croak;
    if (!$opt{norebuild}) {
        write_file $pl, { binmode => ':raw' }, $content;
    }

    my $new_md5 = md5_hex($content);
    $content =~
        s/^(our.\$VERSION.=.'(?:[0-9]+[.]){2})([0-9]+)(')/$1$2-$new_md5$3/msx
        or croak;
    my ($fh, $tmp) = tempfile();
    binmode $fh;
    out($fh, $content);

    my $p = $ENV{Strawberry} // 'C:/Strawberry';
    my $s = $ENV{SystemRoot};
    (local $ENV{PATH} = "$s/system32;$p/bin;$p/perl/site/bin;$p/perl/bin") =~
        tr{/}{\\};
    _qx("spp -o $exe $tmp");

    stdout(_qx("$exe --version"));

    exit;
}

### BUILD }

__END__
