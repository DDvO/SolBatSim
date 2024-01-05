# SolBatSim

### Abstract

Simulate own consumption of a household having a PV solar system,
optionally with a storage battery for buffering currently unused energy.

This simulator is implemented as a [Perl script](SolBatSim.pl) and thus can
run on an system supporting Perl. It uses hourly solar yield data in CSV format,
e.g., from [PVGIS](https://re.jrc.ec.europa.eu/pvg_tools/en/)
and load profiles with a resolution of at least one hour (better: minute-based).
It optionally produces statistical data per hour/day/week/month in CSV format.
It optionally takes into account input limit and output crop of solar inverters.

Various efficiency/loss parameters can be set freely, with reasonable defaults.
Storage may be AC or DC coupled, with adjustable charge/discharge limits.
Various charge/discharge strategies are supported, including suboptimal ones.

Due to its openness, flexibility, and (depending on input) realistic results,
this simulator may be used as a reference for comparing PV simulations.

Load profiles may be taken from public sources such as [HTW Berlin](
https://solar.htw-berlin.de/elektrische-lastprofile-fuer-wohngebaeude/)
or obtained locally using a suitable local digital metering device,
such as a digital 'smart meter' or a Shelly (Pro) 3EM energy meter.
The home automation software [Home Assistant](https://www.home-assistant.io/),
which may continuously run on a micro controller or a home server.
With the definitions in [this YAML configuration file](configuration.yaml),
it can produce both a per-minute load and PV power profile and a per-second
protocol, which may be post-processed with the following Perl script.

[This Perl script](3em_data_collect.pl) gathers the status data
reported each second by a Shelly (Pro) 3EM energy meter.\
It can also collect the PV status data reported each second by a Shelly Plus 1PM.
In this case, the total load reported by the 3PM energy meter is corrected by
adding the absolute value of the PV power input reported by the 1PM power meter.\
Alternatively,
it takes its input from per-second load and (optional) PV power data obtained,
e.g., using Home Assistant with [configuration.yaml](configuration.yaml).
* The script can save the per-second status data in a file per day, including a
total load value obtained by summing up the load values of the three phases.
* It can also produce another file per day with load and PV power per second,
* a further file per day with one load value per second in a row per hour,
* a file per year with the average load per minute in a row per hour,
  which is well suited as SolBatSim input file,
* and another file per year with
a record per hour of the energy consumption, production (if any), and balance,
as well as the imported energy and exported energy, obtained by accumulating
the positive/negative per-second total power values over the given hour.

The script is robust against intermittently missing power data, by interpolating
the data over the range of seconds where no power measurement is available.\
In order to cope with inadvertent abortion of script execution (e.g., due to
system reboot), the script should be started automatically when not currently
running, for instance using a Linux cron job that is triggered each minute.
It can recover the per-minute and per-hour data accumulation for the current day
if the file with the load values per second is available. For correct recovery
including PV production, also the file with PV status data per second is needed.

## Kurzbeschreibung

PV-Solar-Eigenverbrauchs-Simulation, optional mit Batterie-Stromspeicher.
[Dieses Perl-Skript](SolBatSim.pl) verwendet stündliche PV-Daten im CSV-Format
z.B. von [PVGIS](https://re.jrc.ec.europa.eu/pvg_tools/de/)
und Lastprofile mit mindestens stündlicher (typisch minütlicher) Auflösung.
Erzeugt optional statistische Daten pro Stunde/Tag/Woche/Monat im CSV-Format.

Optional mit Drosselung des Wechselrichters auf Ein- und/oder Ausgangsseite.
Diverse Wirkungsgrad-Parameter können frei gewählt werden,
ansonsten wirden sinnvolle Standard-Werte verwendet.
Speicheranbindung AC- oder DC-gekoppelt mit anpassbaren Lade-/Entladelimits.
Unterstützung diverser, auch suboptimaler, Lade- und Entladestrategien.

Aufgrund seiner Offenheit, großen Flexibilität
und (abhängig von den Eingabedaten) sehr realistischen Ergebnisse
ist er auch als Referenz für andere PV-Simulationen verwendbar.

[Hier](https://ddvo.github.io/Solar/#Eigenverbrauch) mehr zum Thema
Eigenverbrauch mit häuslichen PV-Anlagen und seiner Simulation.

Lastprofile kann man aus öffentlichen Quellen beziehen wie [HTW Berlin](
https://solar.htw-berlin.de/elektrische-lastprofile-fuer-wohngebaeude/)
oder mit Hilfe eines geeigneten Messgeräts bzw. Stromzählers selbst gewinnen.

Lastprofil-Dateien sollten folgendes Format haben:
* CSV (Textdatei mit Daten, die durch Kommata getrennt sind)
* Zeilen mit `#` am Anfang werden als Kommentarzeilen ignoriert, ebenso Zeilen,
  die mindestens ein Wort mit vier Buchstaben enthalten, z.B. die Kopfzeile
  `"Datetime";"Power"` von Dateien zum "Import individueller stündlicher
  Verbauch" des [PVTool-Rechners](https://www.akkudoktor.net/pvtool-rechner/).
* Entweder genau 365 Datenzeilen mit jeweils Last-Daten pro Tag\
  oder mindestens 24 und maximal 24 * 365 Daten-Zeilen mit Last-Daten pro Stunde
  &mdash; bei weniger als 8760 Datenzeilen wird für den Rest das Jahres zyklisch
  mit den vorhandenen Daten aufgefüllt, so dass man z.B. auch mit Daten
  von nur einem Tag (24 Stunden) oder einer Woche etwas anfangen kann.
* In jeder Datenzeile steht in der ersten Spalte meist Datum und Uhrzeit,
  aber diese Spalte wird ignoriert.\
  Dahinter stehen Datenpunkte mit Last/Verbrauchs-Werten, die als gleichmäßig
  über den jeweiligen Zeitraum (Tag oder Stunde) verteilt angesehen werden,
  Also bei Last-Daten pro Stunde z.B. nur eine Zahl bei Stundenauflösung,
  12 Zahlen bei Daten im 5-Minuten-Abstand, 60 Zahlen bei Minutenauflösung, usw.
* Die Einheit der Last-Datenwerte (z.B. W oder kW oder kWh) ist unerheblich
  &mdash; es wird ohnehin auf den Jahresverbrauch skaliert,
  den man getrennt angeben muss.

Lastprofile des eigenen Haushalts können mit geeigneten digitalen Stromzählern
erzeugt werden oder mit einem Energiemessgeräts wie dem Shelly (Pro) 3EM und der
Hausautomatisierungs-Software [Home Assistant](https://www.home-assistant.io/),
welche z.B. auf einem Mikrocontroller oder einem Heimserver ständig laufen kann.
Mit Hilfe der Definitionen in [dieser YAML-Konfiguration](configuration.yaml)
kann sowohl ein minutenweises Ertrags- und Lastprofil erzeugt werden als auch
ein sekundenweises Protokoll der Lastdaten und ggf. der PV-Erzeugungsleistung,
welches dann mit dem folgenden Skript weiter verarbeitbar ist.

[Dieses Perl-Skript](3em_data_collect.pl​) liest die sekündlichen Status-Daten
eines Energiemessgeräts Shelly (Pro) 3EM kontinuierlich aus. Es kann sie mit den
Status-Daten eines Shelly Plus 1PM verknüpfen, der die Leistung einer kleinen
PV-Anlage misst, indem es als Haushalts-Last die Summe aus der saldierten
Leistung am Energiemessgerät und dem Absolutbetrag der PV-Leistung bildet.\
Alternativ verarbeitet es als Eingabe
sekündliche Last- und ggf. PV-Leistungsdaten, die z.B. vom Home Assistant
mit [configuration.yaml](configuration.yaml) erzeugt wurden.\
Es kann folgende Daten abspeichern:
* In einer Datei pro Tag alle Statusdaten je Sekunde aus dem Energiemessgerät
  für die drei Phasen und die saldierte Leistung,
  also die Summe der Leistungswerte über alle drei Phasen.
* In einer Datei pro Tag alle Statusdaten je Sekunde aus dem PV-Daten-Messgerät.
* In einer Datei pro Tag die Gesamt-Last, PV-Leistung (wenn vorhanden, sonst 0)
  und die einzelnen Leistungswerte der drei Phasen je Sekunde.
* In einer Datei pro Tag ein Lastprofil mit der saldierten Leistung je Sekunde.
* In einer Datei pro Jahr ein Lastprofil mit einem Eintrag je Minute
  mit dem Durchschnitt der saldierten Leistungswerte über die Minute.
  Diese Datei ist sehr gut als Eingabe-Datei für den SolBatSim geeignet.
* In einer Datei pro Jahr je Stunde die verbrauchte und erzeugte Energie, die
  Energiebilanz (Riemann-Summe über die saldierte Leistung am Energiemessgerät),
  sowie die bezogene (importierte) und die eingespeiste (exportierte) Energie,
  welche sich durch Akkumulation der positiven
  bzw. negativen saldierten Leistungswerte je Sekunde über die Stunde ergibt.

Das Skript ist robust gegen zeitweise fehlende Leistungs-Daten, z.B. wegen
hängender Verbindung zum Energiemessgerät oder zeitweisem Ausfall in seiner
Ausführung, indem es die Messdaten über die ausgefallenen Sekunden interpoliert.
Damit nach Abbrüchen der Skript-Ausführung z.B. durch System-Neustarts das
Skript wieder automatisch weiter laufen kann, sollte es z.B. über einen Linux
cron job o.ä. jede Minute neu gestartet werden, sofern es aktuell nicht läuft.
Wenn für den aktuellen Tag die Lastprofil-Datei mit der saldierten Leistung je
Sekunde verfügbar ist, kann es basierend auf dessen Inhalt die Akkumulation
minütlicher und stündlicher Durchschnittsleistungs- und Energiedaten fortsetzen.
Bei vorhandener PV-Produktion wird dabei für eine korrekte Wiederaufnahme auch
die Datei mit den sekündlichen bisherigen PV-Statusdaten des Tages benötigt.

### Details

Das [Simulator-Skript](SolBatSim.pl) benötigt als Eingabe eine Lastprofil-Datei.
Für die Simulation allgemeiner Haushaltssituationen
können zum Beispiel die 74 von der Forschungsgruppe Solarspeichersysteme
der HTW Berlin [veröffentlichen Lastprofile](
https://solar.htw-berlin.de/elektrische-lastprofile-fuer-wohngebaeude/)
mit 1-Minuten-Auflösung (oder gar 1-Sekunden-Auflösung) verwendet werden.
Mit dem [Lastprofil-Skript](load_profile.pl) kann man aus diesen Rohdaten
Lastprofil-Dateien wie [diese](load_4673_kWh.csv) synthetisieren.

Die zweite wichtige Eingabe sind die PV-Ertragsdaten, welche meistens
in Stunden-Auflösung erhältlich sind, wie etwa die PV-Ertragsdaten
von [PVGIS](https://re.jrc.ec.europa.eu/pvg_tools/de/). Von dort kann man
für einen gegebenen Standort und eine gegebene PV-Modul-Ausrichtung
(wahlweise für einen Abschnitt von Jahren zwischen 2005 und 2020
oder für ein [*typisches meteorologisches Jahr*](
https://help.valentin-software.com/pvsol/de/berechnungsgrundlagen/einstrahlung/klimadaten/))
Solardaten wie [diese](yield_1274_kWh.csv) herunterladen.
Der Simulator kann eine oder mehrere solcher Dateien als Eingabe verwenden,
womit sich auch eine Linearkombination von PV-Modulsträngen unterschiedlicher
Ausrichtung, Verschattung und Leistungsparameter abbilden lässt. Jede Datei
stellt dabei einen Strang von Modulen mit gleicher Einstrahlung dar.

Die Simulation läuft normalerweise über alle Jahre mit vorhandenen PV-Daten
und mittelt in der Ausgabe die Energie-Werte über die betrachteten Jahre.
Sie kann aber auch beschränkt werden auf ein typisches meteorologisches Jahr
oder auf eine bestimmte Jahresspanne, für die PV-Daten vorhanden sind.
Außerdem kann man weiter einschränken auf eine bestimmte Monatspanne,
Tagesspanne und/oder Stundenspanne.

Für die Simulation kann das Lastprofil in einem wählbaren täglichen
Zeitabschnitt durch eine konstante oder minimale (Grund-)Last adaptiert werden,
ebenso der Gesamt-Jahresverbrauch aus dem Lastprofil,
die Nennleistung jeder PV-Modulgruppe
und weitere Parameter wie der System-Wirkungsgrad der PV-Anlage
(resultierend aus Verlusten z.B. in den Leitungen und durch
Verschmutzung, Eigenverschattung und Alterung der Module)
und der Wirkungsgrad des Wechselrichters, welche als konstant angenommen werden.
Auch eine Limitierung der Leistung einzelner Modulstränge (an MPPT-Eingängen)
und der Wechselrichter-Gesamt-Ausgangsleistung
(auf [z.B. 600 W](https://Solar.DavOh.de/#Kappungsverlust)) wird unterstützt.

Außerdem kann die Verwendung eines
[Stromspeichers](https://Solar.DavOh.de/#Batteriepuffer) simuliert werden,
dessen Ladung DC- oder AC-seitig gekoppelt sein kann.
Für jeden Strang von PV-Modulen lässt sich angeben, ob er mit dem Speicher
gekoppelt ist oder direkt (über den Wechselrichter) ins Hausnetz einspeist.
Parameter sind die Brutto-Kapazität, die maximale Lade- und Entladetiefe,
die maximale Lade- und Entladerate (Leistung als Vielfaches der Kapazität/h),
die angenommenen Wirkungsgrade der Ladung und Speicherung, sowie optional
der Wirkungsgrad des für die Entladung verwendeten Wechselrichters.

Zudem kann aus folgenden [hier](https://Solar.DavOh.de/#Regelungsstrategien)
näher behandelten Lade- und Entladestrategien gewählt werden:
- Ladestrategie (solange die definierte Maximalladung nicht erreicht ist):
  - Lastvorrang, auch *Überschussladung* genannt (optimal):
    Speicherung der aktuell nicht anderweitig gebrauchten PV-Energie
  - vorrangige Speicherung (ohne Berücksichtigung der Last), wobei wahlweise
    Strom auch teils am Speicher vorbei geleitet werden kann (Bypass):
    - für Überschuss, der nicht mehr in den Speicher passt, und/oder
    - für eine konstante PV-Nettoleistung
- Entladestrategie (solange die definierte Minimalladung nicht erreicht ist):
  - lastgeregelte Einspeisung (optimal):
    Entnahme so viel wie zusätzlich zum PV-Ertrag gebraucht wird
  - lastgeregelte Einspeisung, aber mit Limitierung der abgegebenen Leistung
    (wobei die Limitierung auf ein Uhrzeit-Intervall eingeschränkt werden kann)
  - Speicherentladung kompensiert PV-Leistung maximal auf Mindestlast-Zielwert
    (wobei die Einspeisung auf ein Uhrzeit-Intervall eingeschränkt werden kann)
  - Umschaltung auf Konstanteinspeisung mit Mindestlast-Zielwert, wenn die
    PV-Leistung unterhalb eines Schwellwerts (z.B. 100 W) liegt und zudem
    der Abstand zwischen Zielwert und PV-Leistung über dem Schwellwert liegt -
    bei Einspeisung geht aber die PV-Leistung verloren (wie beim Anker Solix),
    und die Einspeisung kann auf ein Uhrzeit-Intervall eingeschränkt werden
  - Konstanteinspeisung: Entnahme einer definierten Leistung aus dem Speicher,
    optional auf ein Uhrzeit-Intervall eingeschränkt (z.B. für Nachteinspeisung)

Die Ausgabe aller Parameter und Ergebnisse erfolgt textuell im Terminal.
Die Ergebnisse, wie z.B. die PV-Erträge und der Eigenverbrauch, sowie ggf. der
Speicherdurchsatz usw., werden über alle simulierten Jahre gemittelt ausgegeben.
Optional wird auch die Verteilung der PV-Leistung, Last, Netzeinspeiseleistung,
Lade- und Entladeleistung über die 24 Stunden der Tage ausgegeben,
und zwar gemittelt über alle Tage und als Maximalwerte für die jeweilige Stunde.
Optional kann die Ausgabe zusätzlich in CSV-Dateien geschehen. Dann erfolgt
zusätzlich eine tabellarische Ausgabe der wichtigsten variablen Größen:
PV-Brutto- und Netto-Ertrag, Verbrauch, Eigenverbrauch und Netzeinspeisung,
sowie bei Verwendung eines Speichers Ladung, Entladung und Ladezustand.
Diese werden wahlweise in voller Auflösung (also mit je einer Zeile pro Wert
im Lastprofil) oder über Stunden, Tage, Wochen oder Monate gemittelt ausgegeben.

Der Simulator hat auch einen Testmodus für Debugging- und Demonstrationszwecke
und kann bei Bedarf detaillierte Daten für jeden Simulationsschritt anzeigen.

## Usage

```
SolBatSim.pl <load profile file> [<consumption per year in kWh>]
  (<PV file> [direct] [<nominal power in Wp> [<inverter input limit in W]])+
  [-only <*|year[..year]>[-<*|mon[..mon]>[-<*|day[..day]>[:<*|hour[..hour]]]]]
  [-dist <relative load distribution over each day, per hour 0,..,23>
  [-bend <load distort factors for hour 0,..,23 each day, default 1>
  [-load [min] <constant load, at same scale as in PV data file>
         [<count of days per week starting Monday, default 5>:<from hour,
          default 8 o'clock>..<to hour, default 16, also across midnight>]]
  [-avg_hour] [-verbose]
  [-peff <PV system efficiency in %, default from PV data file(s)>]
  [-ac | -dc] [-capacity <storage capacity in Wh, default 0 (no battery)>]
  [-pass [spill] <constant storage bypass in W in addition to 'direct PV',
                  with 'spill' or AC coupling also when storage is full>]
  [-feed ( lim <limit of optimal feed (according to load) from storage in W >
        | comp <target rate in W up to which storage compensates PV output>
        | excl <threshold in W for discharging storage instead of using PV>
        | <constant feed-in from storage in W> )
        [<from hour, default 0>..<to hour, default 24, also over midnight>]]
  [-max_charge <SoC in %, default 90> [<max charge rate, default 1 C>]]
  [-max_discharge <DoD in %, default 90> [<max rate, default 1 C>]]
  [-ceff <charging efficiency in %, default 94>]
  [-seff <storage efficiency in %, default 95>]
  [-ieff <inverter efficiency in %, default 94>]
  [-ieff2 <efficiency of discharge inverter in %, default from -ieff>]
  [-debug] [-test <load points per hour, for using test data over 24 hours>]
  [-en] [-tmy] [-curb <inverter output power limitation in W>]
  [-hour <statistics file>] [-day <stat file>] [-week <stat file>]
  [-month <stat file>] [-season <file>] [-max <stat file>]

Example:
SolBatSim.pl load_4673_kWh.csv 3000 yield_1274_kWh.csv 1000 -curb 600 -tmy -en
```

All times (hours) are in local time without daylight saving (CET, GMT/UTC +1).

Use `-en` for text output in English. Error messages are all in English.

When PV data for more than one year is given, the average is computed, while
with the option `-tmy` months for a typical meteorological year are selected.

With storage, AC-coupled charging is the default. It has extra inverter loss,
but no spill loss. DC-coupled charging bypasses first inverter and its limits.

With each the options `-hour`/`-day`/`-week`/`-month` a CSV file is produced
with the given name containing with statistical data per hour/day/week/month.


PV data can be obtained from https://re.jrc.ec.europa.eu/pvg_tools/ \
Select location, click "HOURLY DATA", and set the check mark at "PV power".\
Optionally may adapt "Installed peak PV power" and "System loss" already here.\
For using TMY data, choose Start year 2008 or earlier, End year 2020 or later.\
Then press the download button marked "csv".

## Nutzung

```
SolBatSim.pl <Lastprofil-Datei> [<Jahresverbrauch in kWh>]
  (<PV-Datei> [direct] [<Nennleistung Wp> [<WR-Eingangsbegrenzung in W>]])+
  [-only <*|Jahr[..Jahr]>[-<*|Mon[..Mon]>[-<*|Tag[..Tag]>[:<*|Std[..Std]]]]]
  [-dist <relative Lastverteilung über den Tag pro Stunde 0,..,23>
  [-bend <Lastverzerrungsfaktoren tgl. pro Std. 0,..,23, sonst 1>
  [-load [min] <konstante Last, auf gleicher Skala wie in der PV-Daten-Datei>
         [<Zahl der Tage pro Woche ab Montag, sonst 5>:<von Uhrzeit,
          sonst 8 Uhr>..<bis Uhrzeit, sonst 16 Uhr, auch über Mitternacht>]]
  [-avg_hour] [-verbose]
  [-peff <PV-System-Wirkungsgrad in %, ansonsten von PV-Daten-Datei>]
  [-ac | -dc] [-capacity <Speicherkapazität Wh, ansonsten 0 (kein Batterie)>]
  [-pass [spill] <konstante Speicher-Umgehung in W zusätzlich zu 'direct' PV,
                  mit 'spill' oder AC-Kopplung auch bei vollem Speicher>]
  [-feed ( lim <Begrenzung der lastoptimierten Entladung aus Speicher in W>
        | comp <Einspeiseziel in W bis zu dem der Speicher die PV kompensiert>
        | excl <Grenzwert in W zur Speicherentladung statt PV-Nutzung>
        | <konstante Entladung aus Speicher in W> )
        [<von Uhrzeit, sonst 0 Uhr>..
         <bis Uhrzeit, sonst 24 Uhr, auch über Mitternacht>]]
  [-max_charge <Ladehöhe in %, sonst 90> [<max Laderate, sonst 1 C>]]
  [-max_discharge <Entladetiefe in %, sonst 90> [<Rate, sonst 1 C>]]
  [-ceff <Lade-Wirkungsgrad in %, ansonsten 94>]
  [-seff <Speicher-Wirkungsgrad in %, ansonsten 95>]
  [-ieff <Wechselrichter-Wirkungsgrad in %, ansonsten 94>]
  [-ieff2 <Wirkungsgrad des Entlade-Wechselrichters in %, Standard von -ieff>]
  [-debug] [-test <Lastpunkte pro Stunde, für Prüfberechnung über 24 Stunden>]
  [-en] [-tmy] [-curb <Wechselrichter-Ausgangs-Drosselung in W>]
  [-hour <Statistik-Datei>] [-day <Stat.Datei>] [-week <Stat.Datei>]
  [-month <Stat.Datei>] [-season <Stat.Datei>] [-max <Stat.Datei>]

Beispiel:
SolBatSim.pl load_4673_kWh.csv 3000 yield_1274_kWh.csv 1000 -curb 600 -tmy
```

Alle Uhrzeiten sind in lokaler Winterzeit (MEZ, GMT/UTC + 1 ohne Sommerzeit).

Mit `-en` erfolgen die Textausgaben auf Englisch. Fehlertexte sind englisch.

Wenn PV-Daten für mehrere Jahre gegeben sind, wird der Durchschnitt berechnet
oder mit Option `-tmy` Monate für ein typisches meteorologisches Jahr gewählt.

Beim Speicher ist AC-Kopplung Standard. Dabei Verluste durch zweimal WR, aber
kein Überlauf. DC-gekoppelte Ladung umgeht den ersten WR und seine Limits.

Mit den Optionen `-hour`/`-day`/`-week`/`-month` wird jeweils eine CSV-Datei
mit gegebenen Namen mit Statistik-Daten pro Stunde/Tag/Woche/Monat erzeugt.

PV-Daten können bezogen werden von https://re.jrc.ec.europa.eu/pvg_tools/de/ \
Wähle den Standort und "DATEN PRO STUNDE", setze Häkchen bei "PV-Leistung".\
Optional "Installierte maximale PV-Leistung" und "Systemverlust" anpassen.\
Bei Nutzung von "-tmy" Startjahr 2008 oder früher, Endjahr 2020 oder später.\
Dann den Download-Knopf "csv" drücken.

<!--
Local IspellDict: german8
LocalWords: pl load csv yield direct only Mon dist ac debug em mon stat Datetime
LocalWords: bend avg hour verbose peff dc capacity spill feed min en lim comp
LocalWords: excl charge discharge ceff seff ieff test tmy data collect curb day
LocalWords: week month season cron job mdash configuration yaml
LocalWords:
-->
