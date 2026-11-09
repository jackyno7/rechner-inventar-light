#Requires -Version 5.1
<#
.SYNOPSIS
    LIGHT-VERSION (3 von 12 Prüfpunkten, kostenlos). Rechner-Inventar: listet alle Computerkonten eines Active Directory mit Betriebssystem, Version, Support-Ende,
    letzter Anmeldung und LAPS-Status und erzeugt einen HTML-Bericht mit Auffälligkeiten.

.DESCRIPTION
    Das Skript arbeitet ausschließlich LESEND. Es ändert nichts im Active Directory, fragt die Rechner selbst
    nicht ab und überträgt keine Daten. Es schreibt genau eine HTML-Datei (und auf Wunsch CSV-Dateien)
    in den angegebenen Ausgabeordner.

    Voraussetzungen: Windows PowerShell 5.1, das Modul "ActiveDirectory" und ein Domänenkonto mit Leserechten.
    Administratorrechte sind nicht erforderlich. Grundlage sind allein die Angaben, die jeder Rechner selbst im
    Verzeichnis einträgt (Betriebssystem, Version, letzte Anmeldung).

    Die Support-Enddaten stehen in der Tabelle $SupportTabelle weiter unten (Stand 10.2026, Quelle: Microsoft).
    Neue Windows-Versionen lassen sich dort mit einer Zeile ergänzen.

.PARAMETER Ausgabe
    Pfad der HTML-Datei. Standard: Rechner-Inventar_<Domäne>_<Datum>.html im aktuellen Ordner.

.PARAMETER InaktivTage
    Ab wie vielen Tagen ohne Anmeldung ein Computerkonto als inaktiv gilt (Standard 90).

.PARAMETER VorwarnTage
    Meldet Betriebssysteme, deren Support innerhalb dieser Frist endet (Standard 180).

.PARAMETER Suchbasis
    Nur Computer unterhalb dieser Organisationseinheit auswerten (DistinguishedName), z. B. "OU=Standort A,DC=firma,DC=local".

.PARAMETER Server
    Bestimmter Domänencontroller oder Domäne (optional).

.PARAMETER Anmeldung
    Anmeldedaten eines Domänenkontos (Get-Credential), falls das Skript nicht als Domänenbenutzer läuft.
    Dann zusätzlich -Server angeben.

.PARAMETER Csv
    Schreibt das vollständige Inventar und je Prüfpunkt mit Treffern eine CSV-Datei neben den Bericht.

.PARAMETER NichtOeffnen
    Öffnet den fertigen Bericht nicht automatisch im Browser.

.PARAMETER PassThru
    Gibt die Befunde zusätzlich als Objekte zurück (für eigene Auswertungen).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Rechner-Inventar.ps1

.EXAMPLE
    .\Rechner-Inventar.ps1 -InaktivTage 180 -Csv -Ausgabe C:\Berichte\rechner.html

.NOTES
    IT Admin Tools by MM - Version 1.0 (10.2026)
    Kein Ersatz für eine Inventarisierungssoftware oder Sicherheitsprüfung. Nutzung auf eigene Verantwortung.
#>
[CmdletBinding()]
param(
    [string]$Ausgabe,
    [ValidateRange(1, 3650)][int]$InaktivTage = 90,
    [ValidateRange(0, 3650)][int]$VorwarnTage = 180,
    [string]$Suchbasis,
    [string]$Server,
    [System.Management.Automation.PSCredential]$Anmeldung,
    [switch]$Csv,
    [switch]$NichtOeffnen,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$Version = '1.0'

# ------------------------------------------------------------------ Support-Enddaten (Stand 10.2026, Quelle: Microsoft)
# Spalten: Muster für den Betriebssystemnamen | Build | Anzeigename | Ende für Home/Pro | Ende für Enterprise/Education
# Server und LTSC haben nur ein Datum; IoT-Ausgaben stehen vor den gleichnamigen LTSC-Zeilen, weil sie länger gelten. Leeres Build = gilt für alle Builds dieses Namens. Die erste passende Zeile gewinnt.
$SupportTabelle = @(
    @('Windows 11.*IoT.*LTSC', 26100, 'Windows 11 IoT Enterprise LTSC 2024', '2034-10-10', '2034-10-10'),
    @('Windows 10.*IoT.*LTS[CB]', 19044, 'Windows 10 IoT Enterprise LTSC 2021', '2032-01-13', '2032-01-13'),
    @('Windows 11.*LTSC',     26100, 'Windows 11 LTSC 2024',  '2029-10-09', '2029-10-09'),
    @('Windows 11',           26300, 'Windows 11 26H2',       '2028-10-10', '2029-10-09'),
    @('Windows 11',           28000, 'Windows 11 26H1',       '2028-03-14', '2029-03-13'),
    @('Windows 11',           26200, 'Windows 11 25H2',       '2027-10-12', '2028-10-10'),
    @('Windows 11',           26100, 'Windows 11 24H2',       '2026-10-13', '2027-10-12'),
    @('Windows 11',           22631, 'Windows 11 23H2',       '2025-11-11', '2026-11-10'),
    @('Windows 11',           22621, 'Windows 11 22H2',       '2024-10-08', '2025-10-14'),
    @('Windows 11',           22000, 'Windows 11 21H2',       '2023-10-10', '2024-10-08'),
    @('Windows 10.*LTS[CB]',  19044, 'Windows 10 LTSC 2021',  '2027-01-12', '2027-01-12'),
    @('Windows 10.*LTS[CB]',  17763, 'Windows 10 LTSC 2019',  '2029-01-09', '2029-01-09'),
    @('Windows 10.*LTS[CB]',  14393, 'Windows 10 LTSB 2016',  '2026-10-13', '2026-10-13'),
    @('Windows 10.*LTS[CB]',  10240, 'Windows 10 LTSB 2015',  '2025-10-14', '2025-10-14'),
    @('Windows 10',           $null, 'Windows 10',            '2025-10-14', '2025-10-14'),
    @('Windows (8|7|Vista|XP|2000|NT)', $null, 'Windows 8.1 und älter', '2023-01-10', '2023-01-10'),
    @('Server.? 2025',        $null, 'Windows Server 2025',   '2034-11-14', '2034-11-14'),
    @('Server.? 2022',        $null, 'Windows Server 2022',   '2031-10-14', '2031-10-14'),
    @('Server.? 2019',        $null, 'Windows Server 2019',   '2029-01-09', '2029-01-09'),
    @('Server.? 2016',        $null, 'Windows Server 2016',   '2027-01-12', '2027-01-12'),
    @('Server.? 2012',        $null, 'Windows Server 2012 / 2012 R2', '2023-10-10', '2023-10-10'),
    @('Server.? (2008|2003|2000)', $null, 'Windows Server 2008 R2 und älter', '2020-01-14', '2020-01-14')
)

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host 'Das PowerShell-Modul "ActiveDirectory" fehlt.' -ForegroundColor Red
    Write-Host 'Windows 10/11: Einstellungen > Optionale Features > "RSAT: Active Directory Domain Services" hinzufügen.'
    Write-Host 'Windows Server:  Install-WindowsFeature RSAT-AD-PowerShell'
    exit 2
}
Import-Module ActiveDirectory

$ad = @{}
if ($Server) { $ad.Server = $Server }
if ($Anmeldung) { $ad.Credential = $Anmeldung }

# ------------------------------------------------------------------ Hilfsfunktionen
function Datum($d, $leer = 'nie') { if ($d) { ([datetime]$d).ToString('dd.MM.yyyy') } else { $leer } }
function JaNein($b) { if ($b) { 'Ja' } else { 'Nein' } }
function Ort($dn) { ($dn -replace '^(CN|OU)=.+?(?<!\\),', '') -replace ',DC=.*$', '' }
function Html($t) { [System.Net.WebUtility]::HtmlEncode([string]$t) }
function Kurz($t, $n = 220) { $t = ([string]$t -replace '\s+', ' ').Trim(); if ($t.Length -gt $n) { $t.Substring(0, $n) + ' ...' } else { $t } }

# CSV-Dateien öffnen sich in Excel: Werte, die mit = + - @ beginnen, würden dort als Formel ausgeführt (Namen von Dateien, Ordnern, Konten oder Computern
# stammen von Benutzern). Diese Funktion ersetzt für dieses Skript den Befehl Export-Csv und setzt vor solche Werte ein Hochkomma.
function Export-Csv {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline = $true)]$InputObject, [Parameter(Mandatory)][string]$Path, [switch]$NoTypeInformation, [string]$Delimiter = ';', [string]$Encoding = 'UTF8')
    begin { $alle = New-Object System.Collections.Generic.List[object] }
    process { if ($null -ne $InputObject) { $alle.Add($InputObject) } }
    end {
        $sicher = foreach ($o in $alle) {
            $n = [ordered]@{}
            foreach ($p in $o.PSObject.Properties) { $v = $p.Value; if (($v -is [string]) -and ($v -match '^[=+@\t\r]|^-.')) { $v = "'" + $v }; $n[$p.Name] = $v }
            [pscustomobject]$n
        }
        $sicher | Microsoft.PowerShell.Utility\Export-Csv -Path $Path -NoTypeInformation:$NoTypeInformation -Delimiter $Delimiter -Encoding $Encoding
    }
}

# Tabellen: ConvertTo-Html und Export-Csv nehmen die Spalten des ersten Objekts. Haben die Zeilen einer Liste verschiedene
# Eigenschaften (z. B. Fehlzeile neben normalen Treffern), gingen Spalten verloren: alle Namen der Reihe nach vereinigen.
function Einheitlich($liste) {
    $liste = @($liste | Where-Object { $null -ne $_ })
    if ($liste.Count -lt 2) { return $liste }
    $namen = New-Object System.Collections.Generic.List[string]
    foreach ($o in $liste) { foreach ($p in $o.PSObject.Properties.Name) { if (-not $namen.Contains($p)) { $namen.Add($p) } } }
    if ($namen.Count -eq @($liste[0].PSObject.Properties.Name).Count) { return $liste }
    @($liste | Select-Object -Property $namen)
}
# Sortierschluessel fuer Nummern wie K2, K10: erst Buchstaben, dann die Zahl als Zahl
function IdZahl($id) { $m = [regex]::Match([string]$id, '\d+'); if ($m.Success) { [int]$m.Value } else { 0 } }
function IdText($id) { [regex]::Replace([string]$id, '\d+', '') }

$Befunde = New-Object System.Collections.Generic.List[object]

# Platzhalterzeilen ("nicht geprüft") erkennt das Skript nur an den vollständigen, festen Texten, die es selbst erzeugt - keine Teilstring-Suche
# in Konto-, Objekt- oder Beschreibungstexten, die ein Angreifer beeinflussen könnte
$PlatzhalterTexte = @('Abfrage fehlgeschlagen - nicht geprüft', 'Berechtigungen nicht lesbar - nicht geprüft', 'SYSVOL nicht lesbar - nicht geprüft', 'WMI-Filter nicht lesbar - nicht geprüft', 'Kein Domänencontroller ermittelbar - SYSVOL nicht geprüft', 'Inhalt des Ordners nicht lesbar - Richtlinie nicht geprüft')
function IstPlatzhalter($o) {
    # Ein Skript kann $script:PlatzhalterSpalten setzen: dann zählen nur diese Spalten (Namensspalten sind von Benutzern beeinflussbar)
    foreach ($pr in $o.PSObject.Properties) {
        if ($script:PlatzhalterSpalten -and ($pr.Name -notin $script:PlatzhalterSpalten)) { continue }
        if (($pr.Value -is [string]) -and ($PlatzhalterTexte -ccontains $pr.Value)) { return $true }
    }
    ("$($o.Feststellung)" -clike 'Nicht geprüft - *')
}

# Besteht ein Prüfpunkt nur aus Platzhalterzeilen (Abfrage fehlgeschlagen), ist das kein Mangel dieses Prüfpunkts: als Info führen statt
# mit seiner Stufe, damit es in den Kacheln nicht als hoher Befund zählt. Nur für Prüfpunkte, deren Ergebnis von einer Abfrage abhängt.
function InfoWennNurNichtGeprueft([string]$IdMuster) {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster -and $_.Anzahl -gt 0 -and $_.Stufe -ne 'Info' })) {
        if (-not @($x.Objekte | Where-Object { -not (IstPlatzhalter $_) }).Count) { $x.Stufe = 'Info' }
    }
}

# Prüfpunkte, die für einzelne Server nicht ausgewertet werden konnten (Abfrage fehlgeschlagen), dürfen nicht als "ohne Befund" erscheinen:
# je Server eine Zeile "nicht geprüft" anhängen; hatte der Prüfpunkt sonst keinen Treffer, steht er als Info statt als OK.
function NichtGeprueft([string]$IdMuster, [string[]]$Namen, [string]$Verweis, [string]$Spalte = 'Server', [string]$Grund = 'Abfrage nicht möglich') {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster })) {
        $war = $x.Anzahl
        $zeilen = @($Namen | ForEach-Object { [pscustomobject][ordered]@{ $Spalte = $_; Feststellung = "Nicht geprüft - $Grund (siehe $Verweis)"; Details = '' } })
        $x.Objekte = @($x.Objekte) + $zeilen
        $x.Anzahl = $x.Objekte.Count
        if (-not $war) { $x.Stufe = 'Info' }
    }
}
function Befund {
    param([string]$Id, [string]$Bereich, [string]$Titel, [ValidateSet('Hoch', 'Mittel', 'Niedrig', 'Info')][string]$Stufe,
        $Objekte, [string]$Erklaerung, [string]$Empfehlung)
    $liste = @($Objekte | Where-Object { $null -ne $_ })
    $status = $Stufe
    if ($liste.Count -eq 0) { $status = 'OK' }
    $Befunde.Add([pscustomobject]@{
            Id = $Id; Bereich = $Bereich; Titel = $Titel; Stufe = $status; Anzahl = $liste.Count
            Objekte = $liste; Erklaerung = $Erklaerung; Empfehlung = $Empfehlung
        })
}

# Schreibt den HTML-Bericht (und auf Wunsch CSV-Dateien). $Abschnitte: Überschrift -> Objekte oder fertiges HTML,
# erscheinen vor den Befunden. $Anhang: Überschrift -> Text, erscheint am Ende unverändert.
function BerichtSchreiben {
    param([string]$Name, [string]$Datei, [string]$Bezug, [string]$BezugName = 'Domäne', [System.Collections.IDictionary]$Umgebung,
        [System.Collections.IDictionary]$Abschnitte = @{}, [System.Collections.IDictionary]$Anhang = @{}, [string]$Fuss)
    Write-Host '  Bericht schreiben ...'
    $jetzt = Get-Date
    $ziel = $script:Ausgabe
    if (-not $ziel) { $ziel = Join-Path (Get-Location).Path ("{0}_{1}_{2:yyyy-MM-dd_HHmm}.html" -f $Datei, $Bezug, $jetzt) }
    $ziel = [System.IO.Path]::GetFullPath($ziel)
    $ordner = Split-Path $ziel -Parent
    if (-not (Test-Path $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }

    $rang = @{ Hoch = 0; Mittel = 1; Niedrig = 2; Info = 3; OK = 4 }
    $zahl = @{}
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') { $zahl[$s] = @($Befunde | Where-Object Stufe -eq $s).Count }
    $css = @'
body{font-family:Segoe UI,Arial,sans-serif;color:#2b2f36;margin:0;background:#f4f6fa}
.seite{max-width:1100px;margin:0 auto;padding:28px 24px 60px}
h1{color:#1f5fae;margin:0 0 4px;font-size:28px} h2{color:#1f5fae;font-size:19px;margin:34px 0 10px;border-bottom:2px solid #d0d5dd;padding-bottom:6px}
.meta{color:#6b7280;font-size:13px;margin-bottom:22px}
.kacheln{display:flex;gap:14px;flex-wrap:wrap;margin:18px 0}
.kachel{flex:1;min-width:150px;background:#fff;border-radius:10px;padding:14px 18px;border-left:8px solid #999;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.kachel b{display:block;font-size:30px} .kachel span{font-size:13px;color:#6b7280}
table{border-collapse:collapse;width:100%;background:#fff;font-size:13px;margin:8px 0 4px}
th{background:#1f5fae;color:#fff;text-align:left;padding:7px 9px;font-weight:600} td{padding:6px 9px;border-bottom:1px solid #e5e7eb;vertical-align:top;overflow-wrap:break-word}
tr:nth-child(even) td{background:#f9fafb}
.marke{display:inline-block;min-width:58px;text-align:center;padding:2px 9px;border-radius:11px;color:#fff;font-size:12px;font-weight:600}
.m-Hoch{background:#d92d20}.m-Mittel{background:#f79009}.m-Niedrig{background:#ca8a04}.m-OK{background:#12b76a}.m-Info{background:#1f5fae}
.befund{background:#fff;border-radius:10px;padding:16px 20px;margin:14px 0;border-left:8px solid #999;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.befund h3{margin:0 0 6px;font-size:16px} .befund p{margin:5px 0;font-size:14px;line-height:1.45} .befund .emp{color:#14532d}
.Hoch{border-color:#d92d20}.Mittel{border-color:#f79009}.Niedrig{border-color:#eab308}.OK{border-color:#12b76a}.Info{border-color:#1f5fae}
.klein{color:#6b7280;font-size:12px} a{color:#1f5fae;text-decoration:none}
pre{background:#fff;border:1px solid #d0d5dd;border-radius:8px;padding:12px 14px;font-size:12px;white-space:pre-wrap;word-break:break-word}
.fuss{margin-top:40px;color:#6b7280;font-size:12px;border-top:1px solid #d0d5dd;padding-top:12px}
@media print{body{background:#fff}.befund,.kachel{box-shadow:none;border:1px solid #d0d5dd;border-left-width:8px;break-inside:avoid}}
'@
    $h = New-Object System.Text.StringBuilder
    [void]$h.Append("<!DOCTYPE html><html lang='de'><head><meta charset='utf-8'><title>$(Html $Name) $(Html $Bezug)</title><style>$css</style></head><body><div class='seite'>")
    [void]$h.Append("<h1>$(Html $Name)</h1><div class='meta'>$(Html $BezugName) <b>$(Html $Bezug)</b> · erstellt am $($jetzt.ToString('dd.MM.yyyy HH:mm')) von $(Html "$env:USERDOMAIN\$env:USERNAME") auf $(Html $env:COMPUTERNAME)</div>")
    [void]$h.Append("<div class='kacheln'>")
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') {
        $text = @{ Hoch = 'Befunde mit hoher Priorität'; Mittel = 'mit mittlerer Priorität'; Niedrig = 'mit niedriger Priorität'; Info = 'Hinweise und nicht geprüft'; OK = 'Prüfpunkte ohne Befund' }[$s]
        [void]$h.Append("<div class='kachel $s'><b>$($zahl[$s])</b><span>$text</span></div>")
    }
    [void]$h.Append("</div><h2>Umgebung</h2><table><tr><th>Eigenschaft</th><th>Wert</th></tr>")
    foreach ($k in $Umgebung.Keys) { [void]$h.Append("<tr><td>$(Html $k)</td><td>$(Html $Umgebung[$k])</td></tr>") }
    [void]$h.Append("</table>")

    [void]$h.Append("<h2>Übersicht der Prüfpunkte</h2><table><tr><th>Nr.</th><th>Bereich</th><th>Prüfpunkt</th><th>Treffer</th><th>Bewertung</th></tr>")
    foreach ($x in $Befunde) {
        [void]$h.Append("<tr><td>$($x.Id)</td><td>$(Html $x.Bereich)</td><td><a href='#$($x.Id)'>$(Html $x.Titel)</a></td><td>$($x.Anzahl)</td><td><span class='marke m-$($x.Stufe)'>$($x.Stufe)</span></td></tr>")
    }
    [void]$h.Append("</table>")

    foreach ($k in $Abschnitte.Keys) {
        [void]$h.Append("<h2>$(Html $k)</h2>")
        $inhalt = $Abschnitte[$k]
        if ($inhalt -is [string]) { [void]$h.Append($inhalt) }
        elseif (@($inhalt).Count) { [void]$h.Append(((Einheitlich $inhalt | ConvertTo-Html -Fragment) -join "`n")) }
        else { [void]$h.Append("<p class='klein'>Keine Einträge.</p>") }
    }

    [void]$h.Append("<h2>Befunde</h2>")
    $offen = @($Befunde | Where-Object { $_.Stufe -ne 'OK' } | Sort-Object { $rang[$_.Stufe] }, { IdText $_.Id }, { IdZahl $_.Id })
    foreach ($x in $offen) {
        [void]$h.Append("<div class='befund $($x.Stufe)' id='$($x.Id)'><h3><span class='marke m-$($x.Stufe)'>$($x.Stufe)</span> $($x.Id) · $(Html $x.Titel) <span class='klein'>($($x.Anzahl) Treffer)</span></h3>")
        [void]$h.Append("<p>$(Html $x.Erklaerung)</p><p class='emp'><b>Empfehlung:</b> $(Html $x.Empfehlung)</p>")
        [void]$h.Append(((Einheitlich $x.Objekte | ConvertTo-Html -Fragment) -join "`n"))
        [void]$h.Append("</div>")
    }
    if (-not $offen.Count) { [void]$h.Append("<p>Keine Auffälligkeiten gefunden.</p>") }
    $ok = @($Befunde | Where-Object Stufe -eq 'OK')
    if ($ok.Count) {
        [void]$h.Append("<h2>Ohne Befund</h2><table><tr><th>Nr.</th><th>Prüfpunkt</th></tr>")
        foreach ($x in $ok) { [void]$h.Append("<tr id='$($x.Id)'><td>$($x.Id)</td><td>$(Html $x.Titel)</td></tr>") }
        [void]$h.Append("</table>")
    }
    foreach ($k in $Anhang.Keys) { [void]$h.Append("<h2>$(Html $k)</h2><pre>$(Html $Anhang[$k])</pre>") }
    [void]$h.Append("<div class='fuss'>$(Html $Name) Version $script:Version · IT Admin Tools by MM · Das Skript hat ausschließlich gelesen und nichts verändert. $(Html $Fuss)</div>")
    [void]$h.Append("</div></body></html>")
    [System.IO.File]::WriteAllText($ziel, $h.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    $basis = [System.IO.Path]::ChangeExtension($ziel, $null).TrimEnd('.')
    if ($script:Csv) {
        foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 })) {
            Einheitlich $x.Objekte | Export-Csv -Path "${basis}_$($x.Id).csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8
        }
    }
    Write-Host ''
    Write-Host ("Fertig: {0} hoch, {1} mittel, {2} niedrig, {3} ohne Befund." -f $zahl.Hoch, $zahl.Mittel, $zahl.Niedrig, $zahl.OK) -ForegroundColor Green
    Write-Host "Bericht: $ziel"
    if (-not $script:NichtOeffnen -and [Environment]::UserInteractive) { try { Start-Process $ziel } catch { } }
    $basis
}

function Support($os, $build) {
    # Liefert Anzeigename und Support-Ende zu Betriebssystemname und Build; $null, wenn nicht in der Tabelle
    foreach ($z in $SupportTabelle) {
        if ($os -match $z[0] -and ($null -eq $z[1] -or $z[1] -eq $build)) {
            # "Pro Education" gilt wie Home/Pro (24 Monate), nicht wie Education
            $ende = if ($os -match 'Enterprise|(?<!Pro )Education|Server') { $z[4] } else { $z[3] }
            return [pscustomobject]@{ Name = $z[2]; Ende = [datetime]$ende }
        }
    }
}

# ------------------------------------------------------------------ Daten lesen
Write-Host 'Rechner-Inventar - es wird nur gelesen, nichts verändert.' -ForegroundColor Cyan
$dom = Get-ADDomain @ad
$dn = $dom.DistinguishedName
$jetzt = Get-Date
$grenze = $jetzt.AddDays(-$InaktivTage)

# LAPS: Welche Attribute kennt das Schema? (ms-Mcs-AdmPwdExpirationTime = altes LAPS, msLAPS-PasswordExpirationTime = Windows LAPS)
$schema = (Get-ADRootDSE @ad).schemaNamingContext
$lapsAttr = @(Get-ADObject -SearchBase $schema -LDAPFilter '(|(lDAPDisplayName=ms-Mcs-AdmPwdExpirationTime)(lDAPDisplayName=msLAPS-PasswordExpirationTime))' -Properties lDAPDisplayName @ad | ForEach-Object { $_.lDAPDisplayName })

Write-Host '  Computerkonten ...'
$eig = @('Enabled', 'OperatingSystem', 'OperatingSystemVersion', 'LastLogonDate', 'PasswordLastSet', 'whenCreated', 'Description',
    'DNSHostName', 'primaryGroupID', 'mS-DS-CreatorSID', 'Location') + $lapsAttr
$suche = @{ Filter = '*'; Properties = $eig }
if ($Suchbasis) { $suche.SearchBase = $Suchbasis }
$computer = @(Get-ADComputer @suche @ad | Sort-Object Name)

$ersteller = @{}
function Ersteller($sid) {
    if (-not $sid) { return '' }
    $s = "$sid"
    if (-not $ersteller.ContainsKey($s)) {
        $o = $null
        try { $o = Get-ADObject -LDAPFilter "(objectSid=$s)" -Properties sAMAccountName @ad } catch { }
        $ersteller[$s] = if ($o) { $o.sAMAccountName } else { $s }
    }
    $ersteller[$s]
}

$inventar = foreach ($c in $computer) {
    $build = $null
    if ("$($c.OperatingSystemVersion)" -match '\((\d+)\)') { $build = [int]$Matches[1] }
    $sup = if ($c.OperatingSystem) { Support $c.OperatingSystem $build }
    $typ = if ($c.primaryGroupID -in 516, 521) { 'Domänencontroller' } elseif (-not $c.OperatingSystem) { 'Unbekannt' }
    elseif ($c.OperatingSystem -match 'Server') { 'Server' } elseif ($c.OperatingSystem -match 'Windows') { 'Arbeitsplatz' } else { 'Sonstiges' }
    $lapsAblauf = $null
    foreach ($a in $lapsAttr) {
        $w = $c.$a
        if ($w) { $t = [datetime]::FromFileTimeUtc([int64]$w).ToLocalTime(); if (-not $lapsAblauf -or $t -gt $lapsAblauf) { $lapsAblauf = $t } }
    }
    $inaktiv = ($c.LastLogonDate -and $c.LastLogonDate -lt $grenze) -or (-not $c.LastLogonDate -and $c.whenCreated -lt $grenze)
    [pscustomobject]@{
        Name = $c.Name; Typ = $typ; Aktiv = $c.Enabled; OS = $c.OperatingSystem; Build = $build
        Version = $(if ($sup) { $sup.Name } else { '' }); SupportEnde = $(if ($sup) { $sup.Ende })
        Anmeldung = $c.LastLogonDate; Kennwort = $c.PasswordLastSet; Angelegt = $c.whenCreated; Inaktiv = $inaktiv
        Laps = $lapsAblauf; Ort = Ort $c.DistinguishedName; Standard = ($c.DistinguishedName -like "CN=$($c.Name),CN=Computers,$dn")
        Ersteller = Ersteller $c.'mS-DS-CreatorSID'; Beschreibung = $c.Description; Eval = ($c.OperatingSystem -match 'Evaluation')
    }
}
$inventar = @($inventar)
$aktive = @($inventar | Where-Object Aktiv)

function InvZeile($i, [hashtable]$Zusatz = @{}) {
    $z = [ordered]@{
        Computer = $i.Name; Typ = $i.Typ; Betriebssystem = $i.OS; Version = $(if ($i.Version) { $i.Version } elseif ($i.Build) { "Build $($i.Build)" } else { '' })
        'Letzte Anmeldung' = Datum $i.Anmeldung; Ort = $i.Ort
    }
    foreach ($k in $Zusatz.Keys) { $z[$k] = $Zusatz[$k] }
    [pscustomobject]$z
}

# ------------------------------------------------------------------ Prüfungen
Write-Host '  Auswertung ...'
Befund 'B1' 'Betriebssystem' 'Betriebssystem wird nicht mehr unterstützt' 'Hoch' `
($aktive | Where-Object { $_.SupportEnde -and $_.SupportEnde -lt $jetzt } | Sort-Object SupportEnde, Name | ForEach-Object { InvZeile $_ @{ 'Support-Ende' = Datum $_.SupportEnde } }) `
    'Für diese Systeme liefert Microsoft keine Sicherheitsupdates mehr (ohne kostenpflichtige Verlängerung ESU). Bewertet wird, was der Rechner zuletzt selbst im Verzeichnis eingetragen hat; bei Home/Pro gilt das frühere, bei Enterprise/Education das spätere Datum.' `
    'Systeme aktualisieren, ablösen oder vom Netz trennen. Wenn es das Gerät nicht mehr gibt: Computerkonto löschen. Bei gebuchter ESU-Verlängerung ist der Hinweis gegenstandslos.'

Befund 'K1' 'Computerkonten' "Aktive Computerkonten ohne Anmeldung seit über $InaktivTage Tagen" 'Mittel' `
($aktive | Where-Object Inaktiv | Sort-Object Anmeldung | ForEach-Object { InvZeile $_ @{ Angelegt = Datum $_.Angelegt } }) `
    'Wahrscheinlich ausgemusterte Geräte, deren Konto nie gelöscht wurde. Solche Konten lassen sich missbrauchen und verfälschen jede Auswertung. Das Anmeldedatum wird vom Verzeichnis mit bis zu 14 Tagen Verzögerung abgeglichen.' `
    'Konten deaktivieren und nach einer Wartefrist löschen. Ausnahme: Geräte, die bewusst selten am Netz sind.'

Befund 'K3' 'Computerkonten' 'Computer im Standardcontainer "Computers"' 'Niedrig' `
($aktive | Where-Object { $_.Standard -and $_.Typ -ne 'Domänencontroller' } | ForEach-Object { InvZeile $_ }) `
    'Neu aufgenommene Rechner landen im Container "Computers". Mit diesem Container lassen sich keine Gruppenrichtlinien verknüpfen - die Rechner erhalten nur die Richtlinien der Domäne.' `
    'Rechner in die passende Organisationseinheit verschieben. Mit "redircmp" lässt sich der Standardort für neue Rechner ändern.'

# ------------------------------------------------------------------ Bericht
$zusammen = $inventar | Group-Object { if ($_.Version) { $_.Version } elseif ($_.OS) { $_.OS } else { '(kein Betriebssystem eingetragen)' } } | Sort-Object Count -Descending | ForEach-Object {
    $ende = ($_.Group | Where-Object SupportEnde | Select-Object -First 1).SupportEnde
    [pscustomobject][ordered]@{ Betriebssystem = $_.Name; Anzahl = $_.Count; 'davon aktiv' = @($_.Group | Where-Object Aktiv).Count
        'davon inaktiv' = @($_.Group | Where-Object { $_.Aktiv -and $_.Inaktiv }).Count; 'Support-Ende (spätestes)' = Datum (($_.Group | Where-Object SupportEnde | Sort-Object SupportEnde -Descending | Select-Object -First 1).SupportEnde) 'unbekannt' }
}
$liste = $inventar | ForEach-Object {
    [pscustomobject][ordered]@{
        Computer = $_.Name; Typ = $_.Typ; Aktiv = JaNein $_.Aktiv; Betriebssystem = $_.OS; Version = $_.Version; Build = $_.Build
        'Support-Ende' = Datum $_.SupportEnde ''; 'Letzte Anmeldung' = Datum $_.Anmeldung; Angelegt = Datum $_.Angelegt
        LAPS = $(if ($_.Laps) { 'Ja' } elseif ($_.Typ -eq 'Domänencontroller') { '-' } else { 'Nein' }); Ort = $_.Ort; Beschreibung = $_.Beschreibung
    }
}
$umgebung = [ordered]@{
    'Domäne (NetBIOS)'          = "$($dom.DNSRoot) ($($dom.NetBIOSName))"
    'Ausgewerteter Bereich'     = $(if ($Suchbasis) { $Suchbasis } else { 'gesamte Domäne' })
    'Computerkonten (aktiv / gesamt)' = "$($aktive.Count) / $($inventar.Count)"
    'Arbeitsplätze / Server / Domänencontroller / Sonstige' = "$(@($inventar | Where-Object Typ -eq 'Arbeitsplatz').Count) / $(@($inventar | Where-Object Typ -eq 'Server').Count) / $(@($inventar | Where-Object Typ -eq 'Domänencontroller').Count) / $(@($inventar | Where-Object { $_.Typ -in 'Sonstiges', 'Unbekannt' }).Count)"
    'LAPS im Schema'            = $(if ($lapsAttr.Count) { ($lapsAttr | ForEach-Object { if ($_ -like 'msLAPS*') { 'Windows LAPS' } else { 'altes LAPS' } }) -join ', ' } else { 'nein' })
    'Schwellen dieses Berichts' = "inaktiv ab $InaktivTage Tagen, Vorwarnung Support-Ende $VorwarnTage Tage, Support-Daten Stand 10.2026"
}
$basis = BerichtSchreiben -Name 'Rechner-Inventar Light' -Datei 'Rechner-Inventar' -Bezug $dom.DNSRoot -Umgebung $umgebung `
    -Abschnitte ([ordered]@{ 'Betriebssysteme im Überblick' = $zusammen; 'Inventar aller Computerkonten' = ($liste | Select-Object * -ExcludeProperty Build, Angelegt) }) `
    -Fuss 'Light-Version mit 3 von 12 Prüfpunkten der Vollversion (https://mueller86.gumroad.com/l/zzycdd). Grundlage sind die Angaben der Computerkonten im Active Directory; die Rechner selbst wurden nicht abgefragt. Support-Daten ohne Gewähr (Stand 10.2026). Der Bericht ersetzt keine Sicherheitsprüfung und erhebt keinen Anspruch auf Vollständigkeit. Er enthält Rechnernamen - bitte vertraulich behandeln.'
if ($Csv) { $liste | Export-Csv -Path "${basis}_Inventar.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8 }
if ($PassThru) { $Befunde }
