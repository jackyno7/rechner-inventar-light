<#
.SYNOPSIS
    LIGHT VERSION (3 of 12 checks, free). Computer inventory: lists all computer accounts in an Active Directory with operating system, version, end of support,
    last logon, and LAPS status, and generates an HTML report with anomalies.

    Deutsch:
    LIGHT-VERSION (3 von 12 Prüfpunkten, kostenlos). Rechner-Inventar: listet alle Computerkonten eines Active Directory mit Betriebssystem, Version, Support-Ende,
    letzter Anmeldung und LAPS-Status und erzeugt einen HTML-Bericht mit Auffälligkeiten.

.DESCRIPTION
    The script operates exclusively in READ-ONLY mode. It does not change anything in Active Directory, does not query the computers themselves,
    and sends no data to third parties or the internet. It writes exactly one HTML file (and CSV files if requested)
    to the specified output folder.

    Prerequisites: Windows PowerShell 5.1, the "ActiveDirectory" module, and a domain account with read permissions.
    Administrator rights are not required. The basis is solely the information that each computer itself records in
    the directory (operating system, version, last logon).

    The end-of-support dates are in the $SupportTabelle table further below (as of 10.2026, source: Microsoft).
    New Windows versions can be added there with a single line.

    Deutsch:
    Das Skript arbeitet ausschließlich LESEND. Es ändert nichts im Active Directory, fragt die Rechner selbst
    nicht ab und sendet keine Daten an Dritte oder ins Internet. Es schreibt genau eine HTML-Datei (und auf Wunsch CSV-Dateien)
    in den angegebenen Ausgabeordner.

    Voraussetzungen: Windows PowerShell 5.1, das Modul "ActiveDirectory" und ein Domänenkonto mit Leserechten.
    Administratorrechte sind nicht erforderlich. Grundlage sind allein die Angaben, die jeder Rechner selbst im
    Verzeichnis einträgt (Betriebssystem, Version, letzte Anmeldung).

    Die Support-Enddaten stehen in der Tabelle $SupportTabelle weiter unten (Stand 10.2026, Quelle: Microsoft).
    Neue Windows-Versionen lassen sich dort mit einer Zeile ergänzen.

.PARAMETER Ausgabe
    Path of the HTML file. Default: Rechner-Inventar_<Domäne>_<Datum>.html in the current folder.

    Deutsch:
    Pfad der HTML-Datei. Standard: Rechner-Inventar_<Domäne>_<Datum>.html im aktuellen Ordner.

.PARAMETER InaktivTage
    How many days without logon before a computer account is considered inactive (default 90).

    Deutsch:
    Ab wie vielen Tagen ohne Anmeldung ein Computerkonto als inaktiv gilt (Standard 90).

.PARAMETER VorwarnTage
    Reports operating systems whose support ends within this period (default 180).

    Deutsch:
    Meldet Betriebssysteme, deren Support innerhalb dieser Frist endet (Standard 180).

.PARAMETER Suchbasis
    Only evaluate computers below this organizational unit (DistinguishedName), e.g. "OU=Standort A,DC=firma,DC=local".

    Deutsch:
    Nur Computer unterhalb dieser Organisationseinheit auswerten (DistinguishedName), z. B. "OU=Standort A,DC=firma,DC=local".

.PARAMETER Server
    Specific domain controller or domain (optional).

    Deutsch:
    Bestimmter Domänencontroller oder Domäne (optional).

.PARAMETER Anmeldung
    Credentials for a domain account (Get-Credential), if the script is not running as a domain user.
    Then also specify -Server.

    Deutsch:
    Anmeldedaten eines Domänenkontos (Get-Credential), falls das Skript nicht als Domänenbenutzer läuft.
    Dann zusätzlich -Server angeben.

.PARAMETER Csv
    Writes the complete inventory and, for each checkpoint with hits, a CSV file next to the report.

    Deutsch:
    Schreibt das vollständige Inventar und je Prüfpunkt mit Treffern eine CSV-Datei neben den Bericht.

.PARAMETER NichtOeffnen
    Does not automatically open the completed report in the browser.

    Deutsch:
    Öffnet den fertigen Bericht nicht automatisch im Browser.

.PARAMETER PassThru
    Additionally returns the findings as objects (for custom evaluations).

    Deutsch:
    Gibt die Befunde zusätzlich als Objekte zurück (für eigene Auswertungen).

.PARAMETER Language
    Language of the report and the console output: 'de' (German) or 'en' (English). Default: German if the Windows display language
    is German, otherwise English.

    Deutsch:
    Sprache von Bericht und Konsolenausgabe: 'de' (Deutsch) oder 'en' (Englisch). Standard: Deutsch, wenn die Anzeigesprache von
    Windows Deutsch ist, sonst Englisch.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Rechner-Inventar.ps1

.EXAMPLE
    .\Rechner-Inventar.ps1 -InaktivTage 180 -Csv -Ausgabe C:\Berichte\rechner.html

.NOTES
    IT Admin Tools by MM - Version 1.0 (10.2026)
    Kein Ersatz für eine Inventarisierungssoftware oder Sicherheitsprüfung. Nutzung auf eigene Verantwortung.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('de', 'en')][string]$Language = $(if ((Get-UICulture).TwoLetterISOLanguageName -eq 'de') { 'de' } else { 'en' }),
    [Alias('Output')][string]$Ausgabe,
    [ValidateRange(1, 3650)][int]$InaktivTage = 90,
    [ValidateRange(0, 3650)][int]$VorwarnTage = 180,
    [string]$Suchbasis,
    [string]$Server,
    [Alias('Credential')][System.Management.Automation.PSCredential]$Anmeldung,
    [switch]$Csv,
    [Alias('NoOpen')][switch]$NichtOeffnen,
    [switch]$PassThru
)
# Sprache des Berichts: de oder en (Parameter -Language). T 'deutsch' 'english' liefert den Text der gewählten Sprache.
$En = ($Language -eq 'en')
function T([string]$de, [string]$en) { if ($script:En) { $en } else { $de } }

$ErrorActionPreference = 'Stop'
$Version = '1.1'

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
    Write-Host (T 'Das PowerShell-Modul "ActiveDirectory" fehlt.' 'The PowerShell module "ActiveDirectory" is missing.') -ForegroundColor Red
    Write-Host (T 'Windows 10/11: Einstellungen > Optionale Features > "RSAT: Active Directory Domain Services" hinzufügen.' 'Windows 10/11: Settings > Optional features > add "RSAT: Active Directory Domain Services".')
    Write-Host 'Windows Server:  Install-WindowsFeature RSAT-AD-PowerShell'
    exit 2
}
Import-Module ActiveDirectory

$ad = @{}
if ($Server) { $ad.Server = $Server }
if ($Anmeldung) { $ad.Credential = $Anmeldung }

#@SPRACHE@
# ------------------------------------------------------------------ Hilfsfunktionen
function Datum($d, $leer = $null) { if ($null -eq $leer) { $leer = T 'nie' 'never' }; if ($d) { ([datetime]$d).ToString((T 'dd.MM.yyyy' 'yyyy-MM-dd')) } else { $leer } }
function JaNein($b) { if ($b) { T 'Ja' 'Yes' } else { T 'Nein' 'No' } }
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
        if ($alle.Count) { $sicher | Microsoft.PowerShell.Utility\Export-Csv -Path $Path -NoTypeInformation:$NoTypeInformation -Delimiter $Delimiter -Encoding $Encoding }   # leere Liste: keine Datei (sie hätte keine Kopfzeile)
    }
}

# Tabellen: ConvertTo-Html und Export-Csv nehmen die Spalten des ersten Objekts. Haben die Zeilen einer Liste verschiedene
# Eigenschaften (z. B. Fehlzeile neben normalen Treffern), gingen Spalten verloren: alle Namen der Reihe nach vereinigen.
function Einheitlich($liste) {
    $liste = @($liste | Where-Object { $null -ne $_ })
    if ($liste.Count -ge 2) {
        $namen = New-Object System.Collections.Generic.List[string]
        foreach ($o in $liste) { foreach ($p in $o.PSObject.Properties.Name) { if (-not $namen.Contains($p)) { $namen.Add($p) } } }
        if ($namen.Count -ne @($liste[0].PSObject.Properties.Name).Count) { $liste = @($liste | Select-Object -Property $namen) }
    }
    SpaltenAnzeige $liste
}
# Englische Berichte: Spaltenüberschriften aus der Übersetzungstabelle (intern bleiben die deutschen Namen, der Code liest sie weiter)
$SpaltenEn = @{
    'Abfrage' = 'Query'
    'Administrator' = 'Administrator'
    'Adresse' = 'Address'
    'Aktiv' = 'Active'
    'Aktiv verknüpft mit' = 'Active linked to'
    'Aktive Verknüpfungen' = 'Active links'
    'Alter in Tagen' = 'Age in days'
    'Angelegt' = 'Created'
    'Anmeldung' = 'Last logon'
    'Anmeldungen' = 'Logons'
    'Anteil (%)' = 'Share (%)'
    'Anzahl' = 'Count'
    'Art' = 'Account type'
    'Auf Domänencontroller' = 'On domain controller'
    'Aufräumen' = 'Cleanup'
    'Auslastung' = 'Utilization'
    'Auslösender Rechner' = 'Triggering computer'
    'Auslösender Rechner (Anzahl)' = 'Triggering computers (count)'
    'Aussteller' = 'Issuer'
    'Beispiele' = 'Examples'
    'Belegt' = 'Used'
    'Belegt (%)' = 'Used (%)'
    'Benutzer und Computer (auch verschachtelt)' = 'Users and computers (also nested)'
    'Benutzereinstellungen' = 'User settings'
    'Benutzername' = 'Username'
    'Berechtigter' = 'Principal'
    'Bereich' = 'Scope'
    'Bereiche' = 'Scopes'
    'Beschreibung' = 'Description'
    'Besitzer' = 'Owner'
    'Betriebssystem' = 'OS'
    'Betroffene Konten' = 'Affected accounts'
    'Bewertung' = 'Rating'
    'Bis' = 'To'
    'BitLocker' = 'BitLocker'
    'Computer' = 'Computer'
    'Computereinstellungen' = 'Computer settings'
    'Darf anwenden' = 'Can apply'
    'Darf delegieren an' = 'Can delegate to'
    'Darf sich hier als Benutzer ausgeben' = 'Can log on as user'
    'Datei' = 'File'
    'Dateien' = 'Files'
    'Dateisystem' = 'File system'
    'Dateityp' = 'File type'
    'Datenbank' = 'Database'
    'Detail' = 'Detail'
    'Details' = 'Details'
    'Direkte Mitglieder' = 'Direct members'
    'Domäne' = 'Domain'
    'Domänencontroller' = 'Domain controller'
    'Domänennamenmaster' = 'Domain naming master'
    'Dynamische Updates' = 'Dynamic updates'
    'E-Mail' = 'Email'
    'Eigenschaft' = 'Property'
    'Einstellung' = 'Setting'
    'Empfohlen' = 'Recommended'
    'Enthaltene Gruppen' = 'Included groups'
    'Ereignisse' = 'Events'
    'Erreichbar' = 'Reachable'
    'Erste' = 'First'
    'Erster' = 'First'
    'Erzwungen' = 'Forced'
    'Failover' = 'Failover'
    'Falsches Kennwort' = 'Wrong password'
    'Fehlanmeldungen am DC' = 'Failed logons at DC'
    'Fehlversuche' = 'Failed attempts'
    'Feststellung' = 'Finding'
    'Fingerabdruck' = 'Fingerprint'
    'Firewall' = 'Firewall'
    'Frei' = 'Free'
    'Freigabe' = 'Share'
    'Freigaberecht' = 'Share permission'
    'Funktion' = 'Function'
    'Gelesen von' = 'Read by'
    'Geprüft' = 'Checked'
    'Gesamtgröße' = 'Total size'
    'Geschützte Benutzer' = 'Protected users'
    'Gesperrt seit' = 'Locked since'
    'Geändert' = 'Changed'
    'Gilt' = 'Applies to'
    'Gilt für' = 'Applies to'
    'Globaler Katalog' = 'Global catalog'
    'Grenze' = 'Limit'
    'Groesse' = 'Size'
    'Grund' = 'Reason'
    'Gruppe' = 'Group'
    'Gruppen' = 'Groups'
    'Gruppen (auch verschachtelt)' = 'Groups (including nested)'
    'Gruppenrichtlinie' = 'Group policy'
    'Größe' = 'Size'
    'Größe (GB)' = 'Size (GB)'
    'Gültig ab' = 'Valid from'
    'Gültig bis' = 'Valid to'
    'Herkunft' = 'Origin'
    'Hersteller' = 'Vendor'
    'Hinweis' = 'Note'
    'Häufigster Grund' = 'Most common reason'
    'IP-Adresse' = 'IP address'
    'Im AD gespeichert' = 'Stored in AD'
    'In der Domäne vorhanden' = 'Present in domain'
    'Infrastrukturmaster' = 'Infrastructure master'
    'Inhalt' = 'Content'
    'Ist' = 'Is'
    'Kennung' = 'ID'
    'Kennwort gesetzt' = 'Password set'
    'Kennwort läuft nie ab' = 'Password never expires'
    'Kerberos' = 'Kerberos'
    'Konten' = 'Accounts'
    'Konten (auch verschachtelt)' = 'Accounts (also nested)'
    'Konto' = 'Account'
    'Konto für DNS-Einträge' = 'Account for DNS entries'
    'Kopie bleibt im Postfach' = 'Copy remains in mailbox'
    'LAPS' = 'LAPS'
    'LAPS-Ablauf' = 'LAPS expiry'
    'Laufwerk' = 'Drive'
    'Lease-Dauer' = 'Lease duration'
    'Leitet weiter an' = 'Forwards to'
    'Leseweg' = 'Read path'
    'Letzte Aktualisierung' = 'Last update'
    'Letzte Anmeldung' = 'Last logon'
    'Letzte Vollsicherung' = 'Last full backup'
    'Letzte inkrementelle Sicherung' = 'Last incremental backup'
    'Letzter' = 'Last'
    'Letzter Aufräumlauf' = 'Last cleanup run'
    'Letzter Start' = 'Last start'
    'Lokale Admins' = 'Local admins'
    'Mindestlänge' = 'Minimum length'
    'Mitglied' = 'Member'
    'Mitglieder' = 'Members'
    'NLA' = 'NLA'
    'NTLM' = 'NTLM'
    'Name' = 'Name'
    'Nr.' = 'No.'
    'Objekt' = 'Object'
    'Ordner' = 'Folder'
    'Organisationseinheit' = 'OU'
    'Ort' = 'Location'
    'PDC-Emulator' = 'PDC emulator'
    'Pfad' = 'Path'
    'Postfach' = 'Mailbox'
    'Postfach-Typ' = 'Mailbox type'
    'Postfachtyp' = 'Mailbox type'
    'Postfächer' = 'Mailboxes'
    'Privilegierte Gruppen' = 'Privileged groups'
    'Programm' = 'Program'
    'Prüfpunkt' = 'Checkpoint'
    'Quelle' = 'Source'
    'Quelle (Anzahl)' = 'Sources (count)'
    'RID-Master' = 'RID master'
    'Recht' = 'Permission'
    'Rechte' = 'Permissions'
    'Regel' = 'Rule'
    'Reihenfolge' = 'Order'
    'SPN' = 'SPN'
    'Schemamaster' = 'Schema master'
    'Schlüssel' = 'Key'
    'Schreibgeschützt' = 'Read-only'
    'Sendesperre ab' = 'Send blocked from'
    'Server' = 'Server'
    'Sid' = 'SID'
    'Signatur' = 'Signature'
    'Speicher' = 'Storage'
    'Sperrungen' = 'Lockouts'
    'Stand' = 'Status'
    'Standort' = 'Site'
    'Status' = 'Status'
    'Support-Ende' = 'End of support'
    'Support-Ende (spätestes)' = 'End of support (latest)'
    'Tage bis Ablauf' = 'Days until expiry'
    'Tage her' = 'Days ago'
    'Tage ohne Anmeldung' = 'Days without logon'
    'Tage seit Ablauf' = 'Days since expiry'
    'Treffer' = 'Hits'
    'Typ' = 'Type'
    'Vererbung' = 'Inheritance'
    'Verschiedene Konten' = 'Distinct accounts'
    'Verschlüsselt' = 'Encrypted'
    'Version' = 'Version'
    'Version C/B' = 'Version C/B'
    'Verweigert für' = 'Denied for'
    'Verwendung' = 'Use'
    'Virenschutz' = 'Antivirus'
    'Von' = 'From'
    'WMI-Filter' = 'WMI filter'
    'Weg' = 'Path'
    'Weiterleitungen' = 'Forwarding'
    'Wert' = 'Value'
    'Wirksame Richtlinie' = 'Effective policy'
    'Wirksames Recht' = 'Effective permission'
    'Wirkung' = 'Effect'
    'Zeit' = 'Time'
    'Zeitpunkt' = 'Time'
    'Zeitraum' = 'Period'
    'Zertifikat' = 'Certificate'
    'Zone' = 'Zone'
    'Zonen' = 'Zones'
    'Zonenübertragung' = 'Zone transfer'
    'Zugriff für große Gruppen' = 'Access for large groups'
    'Zugriffsbasierte Aufzählung' = 'Access-based enumeration'
    'Zuletzt' = 'Last'
    'Zustand' = 'State'
    'davon Größe (GB)' = 'Of which size (GB)'
    'davon aktiv' = 'Of which active'
    'davon inaktiv' = 'Of which inactive'
    'davon unbekannte Konten' = 'Of which unknown accounts'
    'ist Mitglied von' = 'Member of'
    'Älter als 5 Jahre' = 'Older than 5 years'
    'Älter als 5 Jahren' = 'Older than 5 years'
    'Ältester Eintrag' = 'Oldest entry'
    'Über' = 'Via'
    'Überwachungsprotokoll' = 'Audit log'
}
# Werte, die der Code intern auf Deutsch vergleicht (z. B. Art = 'Benutzer'), erscheinen im englischen Bericht übersetzt (nur Anzeige und CSV)
$WerteEn = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::Ordinal)   # Groß-/Kleinschreibung zählt (Verwaist / verwaist)
$WerteEn.Add('Alle erweiterten Rechte (u. a. Kennwort zurücksetzen)', 'All extended rights (including password reset)')
$WerteEn.Add('Kein Serverbetriebssystem (Support nicht bewertet)', 'No server operating system (support not assessed)')
$WerteEn.Add('Rechte je Ordner (nur dort, wo sie gesetzt werden)', 'Permissions per folder (only where set)')
$WerteEn.Add('Mehrfach freigegebene Ordner (nur einmal gezählt)', 'Folders shared multiple times (counted only once)')
$WerteEn.Add('Übersprungene Verknüpfungen (Junctions/Symlinks)', 'Skipped links (Junctions/Symlinks)')
$WerteEn.Add('Der Ordner der Freigabe existiert nicht mehr.', 'The share folder no longer exists.')
$WerteEn.Add('DHCP-Überwachungsprotokoll ist ausgeschaltet', 'DHCP audit log is disabled')
$WerteEn.Add('Vollzugriff und "Senden als" im Überblick', 'Full access and "Send as" overview')
$WerteEn.Add('Ordner der Freigabe existiert nicht mehr', 'Share folder no longer exists')
$WerteEn.Add('Enthält außer Sicherheitseinstellungen:', 'Contains except security settings:')
$WerteEn.Add('Nicht geprüft - Postfach nicht gefunden', 'Not checked - mailbox not found')
$WerteEn.Add('WMI-Filter nicht lesbar - nicht geprüft', 'WMI filter not readable - not checked')
$WerteEn.Add('Abfrage fehlgeschlagen - nicht geprüft', 'Query failed - not checked')
$WerteEn.Add('Nicht im Active Directory gespeichert', 'Not stored in Active Directory')
$WerteEn.Add('keine im Active Directory autorisiert', 'None authorized in Active Directory')
$WerteEn.Add('Zugriff auf den Ordner nicht möglich', 'Access to folder not possible')
$WerteEn.Add('Zertifikate (öffentliche Schlüssel)', 'Certificates (public keys)')
$WerteEn.Add('Domänencontroller der Organisation', 'Organization domain controllers')
$WerteEn.Add('Remotedesktop erlaubt Verbindungen', 'Remote Desktop allows connections')
$WerteEn.Add('autorisierte DHCP-Server ermitteln', 'Determine authorized DHCP servers')
$WerteEn.Add('Windows Server 2008 R2 und älter', 'Windows Server 2008 R2 and older')
$WerteEn.Add('Konto deaktiviert oder gesperrt', 'Account disabled or locked')
$WerteEn.Add('Vollzugriff (Zusatz: ObjectType', 'Full access (additional: ObjectType')
$WerteEn.Add('Kennwort oder Konto abgelaufen', 'Password or account expired')
$WerteEn.Add('Lokale Administratorengruppen', 'Local administrator groups')
$WerteEn.Add('Berechtigungen/Besitz ändern', 'Change permissions/ownership')
$WerteEn.Add('Freigabepfade nicht sichtbar', 'Share paths not visible')
$WerteEn.Add('Aktive privilegierte Konten', 'Active privileged accounts')
$WerteEn.Add('Gruppe hat keine Mitglieder', 'Group has no members')
$WerteEn.Add('Keine Verbindung zum Server', 'No connection to server')
$WerteEn.Add('unerlaubte Zeichen im Namen', 'Illegal characters in name')
$WerteEn.Add('Geprüfte Domänencontroller', 'Checked domain controllers')
$WerteEn.Add('Authentifizierte Benutzer', 'Authenticated Users')
$WerteEn.Add('Schwellen dieses Berichts', 'Thresholds of this report')
$WerteEn.Add('Benutzer (inetOrgPerson)', 'User (inetOrgPerson)')
$WerteEn.Add('Datenbestand je Freigabe', 'Data inventory per share')
$WerteEn.Add('Kontosperrungsrichtlinie', 'Account lockout policy')
$WerteEn.Add('Sicherheitseinstellungen', 'Security settings')
$WerteEn.Add('Alle erweiterten Rechte', 'All extended rights')
$WerteEn.Add('Hinweise zur Auswertung', 'Notes for evaluation')
$WerteEn.Add('Nicht abgefragte Server', 'Servers not queried')
$WerteEn.Add('Unterordner und Dateien', 'Subfolders and files')
$WerteEn.Add('Bestätigtes Schreiben:', 'Validated write:')
$WerteEn.Add('Installierte Programme', 'Installed programs')
$WerteEn.Add('Konto für DNS-Einträge', 'Account for DNS records')
$WerteEn.Add('Kontosperrungsschwelle', 'Account lockout threshold')
$WerteEn.Add('Nicht lesbare Bereiche', 'Unreadable areas')
$WerteEn.Add('Berechtigungen ändern', 'Change permissions')
$WerteEn.Add('Hinweise zum Auslesen', 'Notes for reading')
$WerteEn.Add('Windows 8.1 und älter', 'Windows 8.1 and older')
$WerteEn.Add('Überwachungsprotokoll', 'Audit log')
$WerteEn.Add('Nicht berücksichtigt', 'Not considered')
$WerteEn.Add('Anonymous-Anmeldung', 'Anonymous logon')
$WerteEn.Add('Infrastrukturmaster', 'Infrastructure master')
$WerteEn.Add('Secure Boot ist aus', 'Secure Boot is off')
$WerteEn.Add('Domänennamenmaster', 'Domain naming master')
$WerteEn.Add('Kennwortrichtlinie', 'Password policy')
$WerteEn.Add('Komponentenwartung', 'Component servicing')
$WerteEn.Add('alle Eigenschaften', 'all properties')
$WerteEn.Add('Besitz übernehmen', 'Take ownership')
$WerteEn.Add('Domänencontroller', 'Domain Controllers')
$WerteEn.Add('Falsches Kennwort', 'Incorrect password')
$WerteEn.Add('Konten-Operatoren', 'Account Operators')
$WerteEn.Add('Nur dieser Ordner', 'Only this folder')
$WerteEn.Add('Domänen-Benutzer', 'Domain Users')
$WerteEn.Add('Format unbekannt', 'Format unknown')
$WerteEn.Add('Sonstige Dateien', 'Other files')
$WerteEn.Add('Ältester Eintrag', 'Oldest entry')
$WerteEn.Add('Domänencomputer', 'Domain Computers')
$WerteEn.Add('Konto unbekannt', 'Account unknown')
$WerteEn.Add('Standard (alle)', 'Default (all)')
$WerteEn.Add('Computerkonten', 'Computer accounts')
$WerteEn.Add('Konto gesperrt', 'Account locked')
$WerteEn.Add('Nicht erreicht', 'Not reached')
$WerteEn.Add('DHCP-Bereiche', 'DHCP scopes')
$WerteEn.Add('Dieser Ordner', 'This folder')
$WerteEn.Add('Gruppe/System', 'Group/system')
$WerteEn.Add('Remotedesktop', 'Remote Desktop')
$WerteEn.Add('Arbeitsplatz', 'Workstation')
$WerteEn.Add('Schemamaster', 'Schema master')
$WerteEn.Add('DHCP-Server', 'DHCP server')
$WerteEn.Add('Datenbanken', 'Databases')
$WerteEn.Add('Sonderkonto', 'Special account')
$WerteEn.Add('Unterordner', 'Subfolders')
$WerteEn.Add('Vollzugriff', 'Full control')
$WerteEn.Add('Zugriff von', 'Access from')
$WerteEn.Add('DNS-Server', 'DNS server')
$WerteEn.Add('Inaktiv ab', 'Inactive from')
$WerteEn.Add('Schreiben:', 'Write:')
$WerteEn.Add('Verweigern', 'Deny')
$WerteEn.Add('DNS-Zonen', 'DNS zones')
$WerteEn.Add('Schreiben', 'Write')
$WerteEn.Add('Sonstiges', 'Other')
$WerteEn.Add('Standorte', 'Sites')
$WerteEn.Add('Unbekannt', 'Unknown')
$WerteEn.Add('unbekannt', 'unknown')
$WerteEn.Add('(Wurzel)', '(root)')
$WerteEn.Add('Anlegen:', 'Create:')
$WerteEn.Add('Benutzer', 'User')
$WerteEn.Add('Erlauben', 'Allow')
$WerteEn.Add('Netzwerk', 'Network')
$WerteEn.Add('Verwaist', 'Orphaned')
$WerteEn.Add('verwaist', 'orphaned')
$WerteEn.Add('Dateien', 'Files')
$WerteEn.Add('Gruppen', 'Groups')
$WerteEn.Add('Anonym', 'Anonymous')
$WerteEn.Add('Domäne', 'Domain')
$WerteEn.Add('Gruppe', 'Group')
$WerteEn.Add('Umfang', 'Scope')
$WerteEn.Add('direkt', 'direct')
$WerteEn.Add('Ändern', 'Change')
$WerteEn.Add('Lesen', 'Read')
$WerteEn.Add('über', 'via')
$WerteMuster = @()
foreach ($k in ($WerteEn.Keys | Sort-Object Length -Descending)) { $WerteMuster += , @([regex]::new('(?<![\p{L}\d])' + [regex]::Escape($k) + '(?![\p{L}\d])'), $WerteEn[$k]) }
function WertAnzeige($v) {
    if (-not $script:En -or $v -isnot [string] -or -not $v) { return $v }
    if ($v.Contains('\')) { return $v }   # Kontonamen (DOMÄNE\Name) und Pfade bleiben, wie sie sind
    foreach ($m in $WerteMuster) { $v = $m[0].Replace($v, $m[1]) }
    $v
}
function SpaltenAnzeige($liste) {
    if (-not $script:En -or -not @($liste).Count) { return $liste }
    foreach ($o in @($liste)) {
        $n = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) { $k = $p.Name; if ($SpaltenEn.ContainsKey($k)) { $k = $SpaltenEn[$k] }; $i = 2; $k0 = $k; while ($n.Contains($k)) { $k = "$k0 $i"; $i++ }; $n[$k] = WertAnzeige $p.Value }
        [pscustomobject]$n
    }
}
# Sortierschluessel fuer Nummern wie K2, K10: erst Buchstaben, dann die Zahl als Zahl
function IdZahl($id) { $m = [regex]::Match([string]$id, '\d+'); if ($m.Success) { [int]$m.Value } else { 0 } }
function IdText($id) { [regex]::Replace([string]$id, '\d+', '') }

$Befunde = New-Object System.Collections.Generic.List[object]

# Platzhalterzeilen ("nicht geprüft") erkennt das Skript nur an den vollständigen, festen Texten, die es selbst erzeugt - keine Teilstring-Suche
# in Konto-, Objekt- oder Beschreibungstexten, die ein Angreifer beeinflussen könnte
# Deutsch und Englisch: Skripte erzeugen den Text je nach -Language, erkannt werden beide Fassungen
$PlatzhalterTexte = @('Abfrage fehlgeschlagen - nicht geprüft', 'Berechtigungen nicht lesbar - nicht geprüft', 'SYSVOL nicht lesbar - nicht geprüft', 'WMI-Filter nicht lesbar - nicht geprüft', 'Kein Domänencontroller ermittelbar - SYSVOL nicht geprüft', 'Inhalt des Ordners nicht lesbar - Richtlinie nicht geprüft',
    'Query failed - not checked', 'Permissions not readable - not checked', 'SYSVOL not readable - not checked', 'WMI filter not readable - not checked', 'No domain controller found - SYSVOL not checked', 'Folder content not readable - policy not checked')
function IstPlatzhalter($o) {
    # Ein Skript kann $script:PlatzhalterSpalten setzen: dann zählen nur diese Spalten (Namensspalten sind von Benutzern beeinflussbar)
    foreach ($pr in $o.PSObject.Properties) {
        if ($script:PlatzhalterSpalten -and ($pr.Name -notin $script:PlatzhalterSpalten)) { continue }
        if (($pr.Value -is [string]) -and ($PlatzhalterTexte -ccontains $pr.Value)) { return $true }
    }
    ("$($o.Feststellung)" -clike 'Nicht geprüft - *') -or ("$($o.Feststellung)" -clike 'Not checked - *')
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
function NichtGeprueft([string]$IdMuster, [string[]]$Namen, [string]$Verweis, [string]$Spalte = 'Server', [string]$Grund = $(T 'Abfrage nicht möglich' 'query not possible')) {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster })) {
        $war = $x.Anzahl
        $zeilen = @($Namen | ForEach-Object { [pscustomobject][ordered]@{ $Spalte = $_; Feststellung = (T "Nicht geprüft - $Grund (siehe $Verweis)" "Not checked - $Grund (see $Verweis)"); Details = '' } })
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
    param([string]$Name, [string]$Datei, [string]$Bezug, [string]$BezugName = $(T 'Domäne' 'Domain'), [System.Collections.IDictionary]$Umgebung,
        [System.Collections.IDictionary]$Abschnitte = @{}, [System.Collections.IDictionary]$Anhang = @{}, [string]$Fuss)
    Write-Host (T '  Bericht schreiben ...' '  Writing report ...')
    $jetzt = Get-Date
    $ziel = $script:Ausgabe
    if (-not $ziel) { $ziel = Join-Path (Get-Location).Path ("{0}_{1}_{2:yyyy-MM-dd_HHmm}.html" -f $Datei, ($Bezug -replace '[^\w.\-]+', '-'), $jetzt) }
    $ziel = [System.IO.Path]::GetFullPath($ziel)
    $ordner = Split-Path $ziel -Parent
    if (-not (Test-Path $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }

    $stufeText = @{ Hoch = (T 'Hoch' 'High'); Mittel = (T 'Mittel' 'Medium'); Niedrig = (T 'Niedrig' 'Low'); Info = 'Info'; OK = 'OK' }
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
    [void]$h.Append("<!DOCTYPE html><html lang='$Language'><head><meta charset='utf-8'><title>$(Html $Name) $(Html $Bezug)</title><style>$css</style></head><body><div class='seite'>")
    [void]$h.Append("<h1>$(Html $Name)</h1><div class='meta'>$(Html $BezugName) <b>$(Html $Bezug)</b> · $(T 'erstellt am' 'created') $($jetzt.ToString((T 'dd.MM.yyyy HH:mm' 'yyyy-MM-dd HH:mm'))) $(T 'von' 'by') $(Html "$env:USERDOMAIN\$env:USERNAME") $(T 'auf' 'on') $(Html $env:COMPUTERNAME)</div>")
    [void]$h.Append("<div class='kacheln'>")
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') {
        $text = @{ Hoch = (T 'Befunde mit hoher Priorität' 'findings, high priority'); Mittel = (T 'mit mittlerer Priorität' 'medium priority'); Niedrig = (T 'mit niedriger Priorität' 'low priority'); Info = (T 'Hinweise und nicht geprüft' 'notes and not checked'); OK = (T 'Prüfpunkte ohne Befund' 'checks without findings') }[$s]
        [void]$h.Append("<div class='kachel $s'><b>$($zahl[$s])</b><span>$text</span></div>")
    }
    [void]$h.Append("</div><h2>$(T 'Umgebung' 'Environment')</h2><table><tr><th>$(T 'Eigenschaft' 'Property')</th><th>$(T 'Wert' 'Value')</th></tr>")
    foreach ($k in $Umgebung.Keys) { [void]$h.Append("<tr><td>$(Html (WertAnzeige $k))</td><td>$(Html (WertAnzeige $Umgebung[$k]))</td></tr>") }
    [void]$h.Append("</table>")

    [void]$h.Append("<h2>$(T 'Übersicht der Prüfpunkte' 'Overview of checks')</h2><table><tr><th>$(T 'Nr.' 'No.')</th><th>$(T 'Bereich' 'Area')</th><th>$(T 'Prüfpunkt' 'Check')</th><th>$(T 'Treffer' 'Hits')</th><th>$(T 'Bewertung' 'Rating')</th></tr>")
    foreach ($x in $Befunde) {
        [void]$h.Append("<tr><td>$($x.Id)</td><td>$(Html $x.Bereich)</td><td><a href='#$($x.Id)'>$(Html $x.Titel)</a></td><td>$($x.Anzahl)</td><td><span class='marke m-$($x.Stufe)'>$($stufeText[$x.Stufe])</span></td></tr>")
    }
    [void]$h.Append("</table>")

    foreach ($k in $Abschnitte.Keys) {
        [void]$h.Append("<h2>$(Html (WertAnzeige $k))</h2>")
        $inhalt = $Abschnitte[$k]
        if ($inhalt -is [string]) { [void]$h.Append($inhalt) }
        elseif (@($inhalt).Count) { [void]$h.Append(((Einheitlich $inhalt | ConvertTo-Html -Fragment) -join "`n")) }
        else { [void]$h.Append("<p class='klein'>$(T 'Keine Einträge.' 'No entries.')</p>") }
    }

    [void]$h.Append("<h2>$(T 'Befunde' 'Findings')</h2>")
    $offen = @($Befunde | Where-Object { $_.Stufe -ne 'OK' } | Sort-Object { $rang[$_.Stufe] }, { IdText $_.Id }, { IdZahl $_.Id })
    foreach ($x in $offen) {
        [void]$h.Append("<div class='befund $($x.Stufe)' id='$($x.Id)'><h3><span class='marke m-$($x.Stufe)'>$($stufeText[$x.Stufe])</span> $($x.Id) · $(Html $x.Titel) <span class='klein'>($($x.Anzahl) $(T 'Treffer' 'hits'))</span></h3>")
        [void]$h.Append("<p>$(Html $x.Erklaerung)</p><p class='emp'><b>$(T 'Empfehlung' 'Recommendation'):</b> $(Html $x.Empfehlung)</p>")
        [void]$h.Append(((Einheitlich $x.Objekte | ConvertTo-Html -Fragment) -join "`n"))
        [void]$h.Append("</div>")
    }
    if (-not $offen.Count) { [void]$h.Append("<p>$(T 'Keine Auffälligkeiten gefunden.' 'No anomalies found.')</p>") }
    $ok = @($Befunde | Where-Object Stufe -eq 'OK')
    if ($ok.Count) {
        [void]$h.Append("<h2>$(T 'Ohne Befund' 'No findings')</h2><table><tr><th>$(T 'Nr.' 'No.')</th><th>$(T 'Prüfpunkt' 'Check')</th></tr>")
        foreach ($x in $ok) { [void]$h.Append("<tr id='$($x.Id)'><td>$($x.Id)</td><td>$(Html $x.Titel)</td></tr>") }
        [void]$h.Append("</table>")
    }
    foreach ($k in $Anhang.Keys) { [void]$h.Append("<h2>$(Html $k)</h2><pre>$(Html $Anhang[$k])</pre>") }
    [void]$h.Append("<div class='fuss'>$(Html $Name) Version $script:Version · IT Admin Tools by MM · $(T 'Das Skript hat ausschließlich gelesen und nichts verändert.' 'The script only read and changed nothing.') $(Html $Fuss)</div>")
    [void]$h.Append("</div></body></html>")
    [System.IO.File]::WriteAllText($ziel, $h.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    $basis = [System.IO.Path]::ChangeExtension($ziel, $null).TrimEnd('.')
    if ($script:Csv) {
        foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 })) {
            Einheitlich $x.Objekte | Export-Csv -Path "${basis}_$($x.Id).csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8
        }
    }
    Write-Host ''
    Write-Host ((T "Fertig: {0} hoch, {1} mittel, {2} niedrig, {3} ohne Befund." "Done: {0} high, {1} medium, {2} low, {3} without findings.") -f $zahl.Hoch, $zahl.Mittel, $zahl.Niedrig, $zahl.OK) -ForegroundColor Green
    Write-Host ((T 'Bericht' 'Report') + ": $ziel")
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
Write-Host (T 'Rechner-Inventar - es wird nur gelesen, nichts verändert.' 'Computer inventory - read-only, nothing changed.') -ForegroundColor Cyan
$dom = Get-ADDomain @ad
$dn = $dom.DistinguishedName
$jetzt = Get-Date
$grenze = $jetzt.AddDays(-$InaktivTage)

# LAPS: Welche Attribute kennt das Schema? (ms-Mcs-AdmPwdExpirationTime = altes LAPS, msLAPS-PasswordExpirationTime = Windows LAPS)
$schema = (Get-ADRootDSE @ad).schemaNamingContext
$lapsAttr = @(Get-ADObject -SearchBase $schema -LDAPFilter '(|(lDAPDisplayName=ms-Mcs-AdmPwdExpirationTime)(lDAPDisplayName=msLAPS-PasswordExpirationTime))' -Properties lDAPDisplayName @ad | ForEach-Object { $_.lDAPDisplayName })

Write-Host (T '  Computerkonten ...' '  Computer accounts ...')
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
Write-Host (T '  Auswertung ...' '  Evaluation ...')
Befund 'B1' (T 'Betriebssystem' 'Operating system') (T 'Betriebssystem wird nicht mehr unterstützt' 'Operating system is no longer supported') 'Hoch' `
($aktive | Where-Object { $_.SupportEnde -and $_.SupportEnde -lt $jetzt } | Sort-Object SupportEnde, Name | ForEach-Object { InvZeile $_ @{ 'Support-Ende' = Datum $_.SupportEnde } }) `
    (T 'Für diese Systeme liefert Microsoft keine Sicherheitsupdates mehr (ohne kostenpflichtige Verlängerung ESU). Bewertet wird, was der Rechner zuletzt selbst im Verzeichnis eingetragen hat; bei Home/Pro gilt das frühere, bei Enterprise/Education das spätere Datum.' 'For these systems, Microsoft no longer provides security updates (without a paid ESU extension). The evaluation uses the date the computer most recently recorded itself in the directory; for Home/Pro, the earlier date applies; for Enterprise/Education, the later date.') `
    (T 'Systeme aktualisieren, ablösen oder vom Netz trennen. Wenn es das Gerät nicht mehr gibt: Computerkonto löschen. Bei gebuchter ESU-Verlängerung ist der Hinweis gegenstandslos.' 'Update, replace, or disconnect systems. If the device is no longer present: delete the computer account. If an ESU extension has been purchased, the alert is irrelevant.')

Befund 'K1' (T 'Computerkonten' 'Computer accounts') (T "Aktive Computerkonten ohne Anmeldung seit über $InaktivTage Tagen" "Active computer accounts with no logon for more than $InaktivTage days") 'Mittel' `
($aktive | Where-Object Inaktiv | Sort-Object Anmeldung | ForEach-Object { InvZeile $_ @{ Angelegt = Datum $_.Angelegt } }) `
    (T 'Wahrscheinlich ausgemusterte Geräte, deren Konto nie gelöscht wurde. Solche Konten lassen sich missbrauchen und verfälschen jede Auswertung. Das Anmeldedatum wird vom Verzeichnis mit bis zu 14 Tagen Verzögerung abgeglichen.' 'Probably retired devices whose accounts were never deleted. Such accounts can be abused and skew every report. The logon date is reconciled by the directory with a delay of up to 14 days.') `
    (T 'Konten deaktivieren und nach einer Wartefrist löschen. Ausnahme: Geräte, die bewusst selten am Netz sind.' 'Deactivate accounts and delete them after a waiting period. Exception: devices that are intentionally rarely connected to the network.')

Befund 'K3' (T 'Computerkonten' 'Computer accounts') (T 'Computer im Standardcontainer "Computers"' 'Computers in the default container "Computers"') 'Niedrig' `
($aktive | Where-Object { $_.Standard -and $_.Typ -ne 'Domänencontroller' } | ForEach-Object { InvZeile $_ }) `
    (T 'Neu aufgenommene Rechner landen im Container "Computers". Mit diesem Container lassen sich keine Gruppenrichtlinien verknüpfen - die Rechner erhalten nur die Richtlinien der Domäne.' 'Newly added computers end up in the "Computers" container. No Group Policies can be linked to this container - the computers only receive domain policies.') `
    (T 'Rechner in die passende Organisationseinheit verschieben. Mit "redircmp" lässt sich der Standardort für neue Rechner ändern.' 'Move computers to the appropriate organizational unit. With "redircmp", the default location for new computers can be changed.')

# ------------------------------------------------------------------ Bericht
$zusammen = $inventar | Group-Object { if ($_.Version) { $_.Version } elseif ($_.OS) { $_.OS } else { (T '(kein Betriebssystem eingetragen)' '(no operating system registered)') } } | Sort-Object Count -Descending | ForEach-Object {
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
    (T 'Domäne (NetBIOS)' 'Domain (NetBIOS)')          = "$($dom.DNSRoot) ($($dom.NetBIOSName))"
    (T 'Ausgewerteter Bereich' 'Evaluated scope')     = $(if ($Suchbasis) { $Suchbasis } else { (T 'gesamte Domäne' 'entire domain') })
    (T 'Computerkonten (aktiv / gesamt)' 'Computer accounts (active / total)') = "$($aktive.Count) / $($inventar.Count)"
    (T 'Arbeitsplätze / Server / Domänencontroller / Sonstige' 'Workstations / Servers / Domain Controllers / Other') = "$(@($inventar | Where-Object Typ -eq 'Arbeitsplatz').Count) / $(@($inventar | Where-Object Typ -eq 'Server').Count) / $(@($inventar | Where-Object Typ -eq 'Domänencontroller').Count) / $(@($inventar | Where-Object { $_.Typ -in 'Sonstiges', 'Unbekannt' }).Count)"
    (T 'LAPS im Schema' 'LAPS in schema')            = $(if ($lapsAttr.Count) { ($lapsAttr | ForEach-Object { if ($_ -like 'msLAPS*') { 'Windows LAPS' } else { (T 'altes LAPS' 'old LAPS') } }) -join ', ' } else { 'nein' })
    (T 'Schwellen dieses Berichts' 'Thresholds of this report') = (T "inaktiv ab $InaktivTage Tagen, Vorwarnung Support-Ende $VorwarnTage Tage, Support-Daten Stand 10.2026" "Inactive after $InaktivTage days, advance warning for end of support $VorwarnTage days, support data as of 10.2026")
}
$basis = BerichtSchreiben -Name (T 'Rechner-Inventar Light' 'Computer inventory Light') -Datei 'Rechner-Inventar' -Bezug $dom.DNSRoot -Umgebung $umgebung `
    -Abschnitte ([ordered]@{ (T 'Betriebssysteme im Überblick' 'Operating systems overview') = $zusammen; (T 'Inventar aller Computerkonten' 'Inventory of all computer accounts') = ($liste | Select-Object * -ExcludeProperty Build, Angelegt) }) `
    -Fuss (T 'Light-Version mit 3 von 12 Prüfpunkten der Vollversion (https://mueller86.gumroad.com/l/zzycdd). Grundlage sind die Angaben der Computerkonten im Active Directory; die Rechner selbst wurden nicht abgefragt. Support-Daten ohne Gewähr (Stand 10.2026). Der Bericht ersetzt keine Sicherheitsprüfung und erhebt keinen Anspruch auf Vollständigkeit. Er enthält Rechnernamen - bitte vertraulich behandeln.' 'Light version with 3 of 12 checks of the full version (https://mueller86.gumroad.com/l/zzycdd). The basis is the information from computer accounts in Active Directory; the computers themselves were not queried. Support data is provided without guarantee (as of 10.2026). The report does not replace a security check and makes no claim to completeness. It contains computer names - please treat them confidentially.')
if ($Csv) { SpaltenAnzeige $liste | Export-Csv -Path "${basis}_Inventar.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8 }
if ($PassThru) { $Befunde }
