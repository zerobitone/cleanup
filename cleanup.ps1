<#
.SYNOPSIS
    Sicheres, robustes Cleanup-Skript fuer alte Dateien und leere Verzeichnisse.

.DESCRIPTION
    Dieses Skript verarbeitet einen oder mehrere Startpfade rekursiv und loescht:

    1. Dateien, deren LastWriteTime aelter ist als das Datei-Schwellen-Datum.
       Das Datei-Schwellen-Datum kann ueber -FileCutoffDate im Format YYYYMMDDhh:mm
       angegeben oder interaktiv eingegeben werden.
       Wenn keine Eingabe erfolgt, gilt der Standard:
       Heute minus MinimumAgeInYears Jahre.

    2. Leere Verzeichnisse unabhaengig vom Alter.
       Die Verzeichnispruefung erfolgt nach dem Dateiloeschlauf, damit Verzeichnisse,
       die erst durch das Entfernen alter Dateien leer werden, beruecksichtigt werden.

    Sicherheitsmerkmale:
    - Unterstuetzt -WhatIf und -Confirm ueber SupportsShouldProcess.
    - Unterstuetzt DryRun ohne tatsaechliche Loeschung.
    - Schreibt Kandidaten zuerst streaming-basiert in Kandidatendateien.
    - Erstellt Kandidatendateien im gleichen Ordner wie die Logdatei.
    - Loescht danach aus diesen Kandidatendateien, ohne alle Pfade im RAM zu halten.
    - Loescht Dateien zuerst, danach leere Verzeichnisse.
    - Nutzt Remove-Item mit -LiteralPath, um Sonderzeichen, Leerzeichen und Wildcards
      in Pfaden sicher zu behandeln.
    - Ueberspringt symbolische Links / Reparse Points und traversiert sie nicht.
    - Unterstuetzt mehrere Eingabepfade.
    - Unterstuetzt Ausschlussmuster als Wildcard und Regex.
    - Nutzt Retry mit exponentiellem Backoff bei Loeschfehlern.
    - Protokolliert gefundene Kandidaten, Loeschungen, uebersprungene Eintraege und Fehler.
    - Veraendert keine globalen System-Einstellungen dauerhaft.
    - Verlangt vor echten Loeschlaeufen eine explizite Sicherheitsbestaetigung.
    - Verhindert bewusst nicht-interaktive Ausfuehrung, da dieses Skript nie automatisiert
      laufen soll.

.PARAMETER Path
    Ein oder mehrere Startpfade.
    Pflichtparameter.

.PARAMETER MinimumAgeInYears
    Standard-Mindestalter in Jahren, falls kein FileCutoffDate angegeben wird.
    Dateien werden dann nur als Kandidaten behandelt, wenn deren LastWriteTime
    aelter ist als: heutiges Datum minus MinimumAgeInYears Jahre.

    Standard: 5.

.PARAMETER FileCutoffDate
    Optionaler Datei-Stichtag im Format YYYYMMDDhh:mm.

    Dateien werden als Kandidaten behandelt, wenn deren LastWriteTime
    aelter ist als dieser Stichtag.

    Beispiel:
    2021010100:00

    Wenn der Parameter nicht uebergeben wird, fragt das Skript interaktiv danach.
    Wenn auch interaktiv nichts eingegeben wird, gilt der Standard:
    aelter als MinimumAgeInYears Jahre.

.PARAMETER DryRun
    Wenn gesetzt, werden keine Dateien oder Verzeichnisse geloescht.
    Kandidaten werden nur gesucht und protokolliert.

.PARAMETER LogPath
    Pfad zur Logdatei.
    Wenn ein bestehendes Verzeichnis uebergeben wird, erzeugt das Skript darin automatisch
    eine Logdatei mit Zeitstempel.

    Standard:
    ./to_delete_<yyyyMMdd_HHmmss>.log

.PARAMETER ExcludePattern
    Optionale Ausschlussmuster.
    Jedes Muster wird sowohl als Wildcard ueber -like als auch als Regex ueber -match
    gegen vollstaendigen Pfad und Namen geprueft.

.EXAMPLE
    .\cleanup.ps1 -Path "N:\IT-BL" -DryRun

.EXAMPLE
    .\cleanup.ps1 -Path "N:\IT-BL" -FileCutoffDate "2021010100:00" -DryRun

.EXAMPLE
    .\cleanup.ps1 -Path "N:\IT-BL","E:\Backups" -MinimumAgeInYears 7 -DryRun

.EXAMPLE
    .\cleanup.ps1 -Path "N:\IT-BL" -ExcludePattern "*.pst","*.ost","*.msg","*.oft","DoNotDelete","*\DoNotDelete","*\DoNotDelete\*" -DryRun

.EXAMPLE
    .\cleanup.ps1 -Path "N:\IT-BL" -WhatIf

.EXAMPLE
    .\cleanup.ps1 -Path "N:\IT-BL" -Confirm

.EXAMPLE
    .\cleanup.ps1 -Path "\\ibbads\data\IT-BL" -DryRun

.EXAMPLE
    .\cleanup.ps1 -Path "N:\IT-BL" -LogPath "N:\IT-BL\referent\cleanUp\cleanup_logs" -ExcludePattern "*.pst","*.ost","*.msg","*.oft" -DryRun

.NOTES
    Autor:
    Stefan Homberg

    Dateiname:
    cleanup.ps1

    Kompatibilitaet:
    - Windows PowerShell 5.1
    - PowerShell 7+

    Wichtige Betriebsregel:
    Dieses Skript ist bewusst fuer interaktive Ausfuehrung gebaut und wird an autorisierte
    Personen der IT-Abteilungen weitergegeben.
    Es soll nie automatisiert, unbeaufsichtigt, als geplanter Task oder als Hintergrundjob
    produktiv laufen.

    Namenskonvention:
    - Funktionen verwenden PowerShell-typische Verb-Noun-Namen.
    - Verben orientieren sich an approved verbs.
    - Nomen sind moeglichst spezifisch und im Singular.
    - Parameter verwenden PascalCase und klare fachliche Bezeichnungen.
    - Es werden bewusst keine Parameter-Aliasse verwendet.

    Hinweis zu langen Pfaden:
    - Das Skript versucht unter Windows lokale Pfade mit \\?\ und UNC-Pfade mit
      \\?\UNC\ zu normalisieren.
    - Die tatsaechliche Unterstuetzung langer Pfade haengt zusaetzlich von Windows,
      Gruppenrichtlinien und Anwendungskontext ab.

    Hinweis zu Kandidatendateien:
    - Kandidatendateien werden im gleichen Verzeichnis wie die Logdatei erstellt (*.jsonl).
    - Sie dienen der streaming-basierten Verarbeitung grosser Datenbestaende.
    - Am Ende des Skriptlaufs werden diese Kandidatendateien wieder entfernt.
    - Bei einem harten Abbruch koennen sie im Logordner verbleiben und dadurch zusaetzlich
      bei der Fehleranalyse helfen.

    Hinweis zu -WhatIf:
    - -WhatIf verhindert tatsaechliche Loeschungen ueber ShouldProcess.
    - DryRun ist zusaetzlich implementiert und protokolliert Kandidaten ohne Loeschversuch.

    Hinweis zu -Confirm:
    - Das Skript zeigt bei echten Loeschlaeufen immer eine eigene Sicherheitsabfrage.
    - Der PowerShell-Common-Parameter -Confirm kann zusaetzlich verwendet werden,
      um pro Loeschaktion eine PowerShell-Bestaetigung zu erhalten.

    Pester-Testideen:
    - Alte Datei wird im DryRun gefunden, aber nicht geloescht.
    - Alte Datei wird mit -WhatIf nicht geloescht.
    - FileCutoffDate wird korrekt im Format YYYYMMDDhh:mm verarbeitet.
    - Leere Verzeichnisse werden unabhaengig vom Alter gefunden.
    - ExcludePattern verhindert Kandidatenerfassung.
    - Symbolischer Link wird nicht traversiert.
    - Nicht-leeres Verzeichnis bleibt bestehen.
    - Nicht-interaktive Ausfuehrung wird verhindert.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Path,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1000)]
    [int]$MinimumAgeInYears = 5,

    [Parameter(Mandatory = $false)]
    [AllowEmptyString()]
    [string]$FileCutoffDate,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = $(Join-Path -Path (Get-Location) -ChildPath ("to_delete_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))),

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludePattern
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# --------------------------------------------------------------------------------------
# Skriptweite Variablen
# --------------------------------------------------------------------------------------

$script:IsWindowsLike               = $true
$script:CutoffDate                  = $null
$script:ResolvedLogPath             = $null
$script:CandidateDirectoryPath      = $null
$script:CandidateFileList           = @()
$script:FoundFileCount              = 0
$script:FoundDirectoryCount         = 0
$script:RemovedFileCount            = 0
$script:RemovedDirectoryCount       = 0
$script:SkippedItemCount            = 0
$script:ErrorCount                  = 0
$script:ProcessedDirectoryCount     = 0
$script:ProcessedFileCount          = 0
$script:InvalidRegexPattern         = @{}
$script:EffectiveFileCutoffSource   = $null

# --------------------------------------------------------------------------------------
# Konsolen- und Logfunktionen
# --------------------------------------------------------------------------------------

function Write-CleanupInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "[INFO ] $Message" -ForegroundColor Cyan
}

function Write-CleanupSuccess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "[ OK  ] $Message" -ForegroundColor Green
}

function Write-CleanupWarning {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "[WARN ] $Message" -ForegroundColor Yellow
}

function Write-CleanupError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "[FEHLER] $Message" -ForegroundColor Red
}

function Get-CurrentUserName {
    try {
        return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }
    catch {
        if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) {
            return $env:USERNAME
        }

        if (-not [string]::IsNullOrWhiteSpace($env:USER)) {
            return $env:USER
        }

        return "Unbekannt"
    }
}

function Get-SafeComputerName {
    if (-not [string]::IsNullOrWhiteSpace($env:COMPUTERNAME)) {
        return $env:COMPUTERNAME
    }

    if (-not [string]::IsNullOrWhiteSpace($env:HOSTNAME)) {
        return $env:HOSTNAME
    }

    try {
        return [System.Net.Dns]::GetHostName()
    }
    catch {
        return "Unbekannt"
    }
}

function Add-CleanupLog {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($script:ResolvedLogPath)) {
        return
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $line = "[{0}] {1}" -f $timestamp, $Message

    [System.IO.File]::AppendAllText(
        $script:ResolvedLogPath,
        $line + [Environment]::NewLine,
        [System.Text.Encoding]::UTF8
    )
}

function Resolve-CleanupLogPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetLogPath
    )

    # Der Parameter LogPath darf entweder eine konkrete Datei oder ein bestehender Ordner sein.
    # Wenn ein Ordner uebergeben wird, erzeugt das Skript darin automatisch eine Logdatei.
    if (Test-Path -LiteralPath $TargetLogPath -PathType Container) {
        return Join-Path -Path $TargetLogPath -ChildPath ("to_delete_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    }

    $logDirectory = Split-Path -Path $TargetLogPath -Parent

    if ([string]::IsNullOrWhiteSpace($logDirectory)) {
        return Join-Path -Path (Get-Location).Path -ChildPath $TargetLogPath
    }

    return $TargetLogPath
}

function Initialize-CleanupLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetLogPath
    )

    $effectiveLogPath = Resolve-CleanupLogPath -TargetLogPath $TargetLogPath
    $logDirectory = Split-Path -Path $effectiveLogPath -Parent

    if ([string]::IsNullOrWhiteSpace($logDirectory)) {
        $logDirectory = (Get-Location).Path
        $effectiveLogPath = Join-Path -Path $logDirectory -ChildPath $effectiveLogPath
    }

    if (-not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
        throw "Das Log-Verzeichnis existiert nicht: $logDirectory"
    }

    try {
        # Schreibtest vor Beginn der Verarbeitung.
        # Ein Cleanup ohne funktionierendes Log waere fachlich riskant.
        $testFile = Join-Path -Path $logDirectory -ChildPath (".__cleanup_write_test_{0}.tmp" -f ([Guid]::NewGuid().ToString("N")))
        [System.IO.File]::WriteAllText($testFile, "write-test", [System.Text.Encoding]::UTF8)
        Remove-Item -LiteralPath $testFile -Force -ErrorAction Stop
    }
    catch {
        throw "Keine Schreibrechte fuer das Log-Verzeichnis '$logDirectory'. Fehler: $($_.Exception.Message)"
    }

    $resolvedLogPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($effectiveLogPath)
    $script:CandidateDirectoryPath = Split-Path -Path $resolvedLogPath -Parent

    if ([string]::IsNullOrWhiteSpace($script:CandidateDirectoryPath)) {
        $script:CandidateDirectoryPath = (Get-Location).Path
    }

    $header = @(
        "================================================================================"
        "Cleanup-Log"
        "================================================================================"
        "Skript                    : cleanup.ps1"
        "Autor                     : Stefan Homberg"
        "Startzeit                 : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "Benutzer                  : $(Get-CurrentUserName)"
        "Computer                  : $(Get-SafeComputerName)"
        "PowerShell-Version        : $($PSVersionTable.PSVersion)"
        "PowerShell-Edition        : $($PSVersionTable.PSEdition)"
        "Interaktiver Lauf         : Ja"
        "Automatisierung           : Nicht vorgesehen / wird bewusst blockiert"
        "DryRun                    : $DryRun"
        "WhatIfPreference          : $WhatIfPreference"
        "MinimumAgeInYears         : $MinimumAgeInYears"
        "FileCutoffDate Eingabe    : $FileCutoffDate"
        "Datei-Schwellen-Datum     : $script:CutoffDate"
        "Datei-Schwellen-Quelle    : $script:EffectiveFileCutoffSource"
        "Leere Verzeichnisse       : Unabhaengig vom Alter"
        "LogPath                   : $resolvedLogPath"
        "Kandidatendatei-Ordner    : $script:CandidateDirectoryPath"
        "ACHTUNG                   : Bei echtem Lauf werden protokollierte Eintraege endgueltig geloescht."
        "================================================================================"
        ""
    )

    [System.IO.File]::WriteAllLines($resolvedLogPath, $header, [System.Text.Encoding]::UTF8)

    return $resolvedLogPath
}

# --------------------------------------------------------------------------------------
# Plattform-, Interaktivitaets- und Pfadfunktionen
# --------------------------------------------------------------------------------------

function Initialize-CleanupPlatform {
    if ($PSVersionTable.PSEdition -eq 'Core') {
        try {
            $script:IsWindowsLike = [bool]$IsWindows
        }
        catch {
            $script:IsWindowsLike = ([System.Environment]::OSVersion.Platform -eq 'Win32NT')
        }
    }
    else {
        $script:IsWindowsLike = $true
    }
}

function Test-CleanupInteractiveSession {
    # Dieses Skript soll niemals unbeaufsichtigt laufen.
    # Die Pruefung verhindert typische nicht-interaktive Kontexte wie geplante Tasks.
    if (-not [Environment]::UserInteractive) {
        return $false
    }

    try {
        if ($null -eq $Host -or $null -eq $Host.UI -or $null -eq $Host.UI.RawUI) {
            return $false
        }

        return $true
    }
    catch {
        return $false
    }
}

function ConvertTo-CleanupLongPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath
    )

    if (-not $script:IsWindowsLike) {
        return $InputPath
    }

    if ($InputPath.StartsWith("\\?\")) {
        return $InputPath
    }

    if ($InputPath.StartsWith("\\")) {
        return "\\?\UNC\" + $InputPath.Substring(2)
    }

    if ($InputPath -match '^[a-zA-Z]:\\') {
        return "\\?\" + $InputPath
    }

    return $InputPath
}

function ConvertFrom-CleanupLongPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath
    )

    if ($InputPath.StartsWith("\\?\UNC\")) {
        return "\\" + $InputPath.Substring(8)
    }

    if ($InputPath.StartsWith("\\?\")) {
        return $InputPath.Substring(4)
    }

    return $InputPath
}

function Resolve-CleanupPath {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$InputPath
    )

    $resolvedPathList = New-Object System.Collections.Generic.List[string]

    foreach ($currentPath in $InputPath) {
        if ([string]::IsNullOrWhiteSpace($currentPath)) {
            throw "Ein uebergebener Pfad ist leer."
        }

        if (-not (Test-Path -LiteralPath $currentPath)) {
            throw "Pfad existiert nicht: $currentPath"
        }

        try {
            $resolvedItemList = Resolve-Path -LiteralPath $currentPath -ErrorAction Stop

            foreach ($resolvedItem in $resolvedItemList) {
                if ([string]::IsNullOrWhiteSpace($resolvedItem.ProviderPath)) {
                    throw "Pfad konnte nicht eindeutig aufgeloest werden: $currentPath"
                }

                [void]$resolvedPathList.Add($resolvedItem.ProviderPath)
            }
        }
        catch {
            throw "Pfad konnte nicht aufgeloest werden: $currentPath. Fehler: $($_.Exception.Message)"
        }
    }

    return $resolvedPathList.ToArray()
}

function Resolve-CleanupCutoffDate {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$FileCutoffDateInput,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 1000)]
        [int]$DefaultMinimumAgeInYears
    )

    if ([string]::IsNullOrWhiteSpace($FileCutoffDateInput)) {
        $script:EffectiveFileCutoffSource = "Standard: Heute minus $DefaultMinimumAgeInYears Jahre"
        return (Get-Date).Date.AddYears(-1 * $DefaultMinimumAgeInYears)
    }

    $format = "yyyyMMddHH:mm"
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $dateTimeStyle = [System.Globalization.DateTimeStyles]::None
    $parsedCutoffDate = [datetime]::MinValue

    if (
        -not [datetime]::TryParseExact(
            $FileCutoffDateInput,
            $format,
            $culture,
            $dateTimeStyle,
            [ref]$parsedCutoffDate
        )
    ) {
        throw "Ungueltiges Datei-Stichtagsformat. Erwartet wird YYYYMMDDhh:mm, zum Beispiel 2021010100:00 oder 2024050314:30."
    }

    if ($parsedCutoffDate -gt (Get-Date)) {
        throw "Der Datei-Stichtag liegt in der Zukunft: $parsedCutoffDate. Das ist fuer dieses Cleanup-Skript nicht zulaessig."
    }

    $script:EffectiveFileCutoffSource = "Benutzereingabe: $FileCutoffDateInput"
    return $parsedCutoffDate
}

# --------------------------------------------------------------------------------------
# Pruef- und Filterfunktionen
# --------------------------------------------------------------------------------------

function Test-CleanupReparsePoint {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileSystemInfo]$Item
    )

    # Reparse Points sind z. B. symbolische Links, Junctions oder Mount Points.
    # Sie werden bewusst nicht traversiert, damit das Skript nicht unbemerkt ausserhalb
    # des angegebenen Startpfades arbeitet.
    return (($Item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
}

function Test-CleanupExclusion {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileSystemInfo]$Item,

        [Parameter(Mandatory = $false)]
        [string[]]$Pattern
    )

    if (-not $Pattern -or $Pattern.Count -eq 0) {
        return $false
    }

    $fullName = ConvertFrom-CleanupLongPath -InputPath $Item.FullName
    $name = $Item.Name

    foreach ($currentPattern in $Pattern) {
        if ([string]::IsNullOrWhiteSpace($currentPattern)) {
            continue
        }

        # Wildcard-Pruefung, z. B. *.pst oder *\DoNotDelete\*
        if ($fullName -like $currentPattern -or $name -like $currentPattern) {
            return $true
        }

        # Regex-Pruefung fuer fortgeschrittene Muster.
        # Ungueltige Regex-Muster werden nur einmal protokolliert.
        if ($script:InvalidRegexPattern.ContainsKey($currentPattern)) {
            continue
        }

        try {
            if ($fullName -match $currentPattern -or $name -match $currentPattern) {
                return $true
            }
        }
        catch {
            $script:InvalidRegexPattern[$currentPattern] = $true
            Add-CleanupLog "WARNUNG`tUngueltiges Regex-Ausschlussmuster uebersprungen`tPattern=$currentPattern`tFehler=$($_.Exception.Message)"
        }
    }

    return $false
}

function Test-CleanupDirectoryEmpty {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$Directory
    )

    try {
        # Performant: Es wird nicht der gesamte Verzeichnisinhalt geladen.
        # Es reicht zu pruefen, ob mindestens ein Eintrag existiert.
        $enumerator = [System.IO.Directory]::EnumerateFileSystemEntries($Directory.FullName).GetEnumerator()

        try {
            return -not $enumerator.MoveNext()
        }
        finally {
            if ($null -ne $enumerator -and $enumerator -is [System.IDisposable]) {
                $enumerator.Dispose()
            }
        }
    }
    catch {
        $script:ErrorCount++
        Add-CleanupLog "ERROR`tLeerpruefung fehlgeschlagen`tPfad=$($Directory.FullName)`tFehler=$($_.Exception.Message)"
        return $false
    }
}

function Get-CleanupFileSystemEntry {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$Directory
    )

    try {
        return [System.IO.Directory]::EnumerateFileSystemEntries($Directory.FullName)
    }
    catch {
        $script:ErrorCount++
        Add-CleanupLog "ERROR`tVerzeichnis kann nicht gelesen werden`tPfad=$($Directory.FullName)`tFehler=$($_.Exception.Message)"
        return @()
    }
}

# --------------------------------------------------------------------------------------
# Kandidatendateien
# --------------------------------------------------------------------------------------

function New-CleanupCandidateFile {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("file", "directory")]
        [string]$Kind
    )

    if ([string]::IsNullOrWhiteSpace($script:CandidateDirectoryPath)) {
        throw "Der Kandidatendatei-Ordner wurde nicht initialisiert."
    }

    if (-not (Test-Path -LiteralPath $script:CandidateDirectoryPath -PathType Container)) {
        throw "Der Kandidatendatei-Ordner existiert nicht: $script:CandidateDirectoryPath"
    }

    $candidateFileName = "cleanup_candidates_{0}_{1}.jsonl" -f $Kind, ([Guid]::NewGuid().ToString("N"))
    $candidateFilePath = Join-Path -Path $script:CandidateDirectoryPath -ChildPath $candidateFileName

    [System.IO.File]::WriteAllText($candidateFilePath, "", [System.Text.Encoding]::UTF8)
    $script:CandidateFileList += $candidateFilePath

    Add-CleanupLog "INFO`tKandidatendatei erstellt`tTyp=$Kind`tPfad=$candidateFilePath"

    return $candidateFilePath
}

function Add-CleanupCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CandidateFilePath,

        [Parameter(Mandatory = $true)]
        [ValidateSet("File", "Directory")]
        [string]$Type,

        [Parameter(Mandatory = $true)]
        [string]$FullName,

        [Parameter(Mandatory = $true)]
        [datetime]$LastWriteTime
    )

    $displayPath = ConvertFrom-CleanupLongPath -InputPath $FullName

    $candidate = [PSCustomObject]@{
        Type             = $Type
        FullName         = $FullName
        DisplayPath      = $displayPath
        LastWriteTimeUtc = $LastWriteTime.ToUniversalTime().ToString("o")
    }

    $json = $candidate | ConvertTo-Json -Compress
    [System.IO.File]::AppendAllText($CandidateFilePath, $json + [Environment]::NewLine, [System.Text.Encoding]::UTF8)

    if ($Type -eq "File") {
        $script:FoundFileCount++
        Add-CleanupLog "FOUND_FILE`tPfad=$displayPath`tLastWriteTime=$($LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    }
    else {
        $script:FoundDirectoryCount++
        Add-CleanupLog "FOUND_EMPTY_DIR`tPfad=$displayPath`tLastWriteTime=$($LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    }
}

function Get-CleanupCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CandidateFilePath
    )

    if (-not (Test-Path -LiteralPath $CandidateFilePath -PathType Leaf)) {
        return
    }

    foreach ($line in [System.IO.File]::ReadLines($CandidateFilePath)) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $line | ConvertFrom-Json
        }
        catch {
            $script:ErrorCount++
            Add-CleanupLog "ERROR`tKandidat konnte nicht aus JSONL gelesen werden`tDatei=$CandidateFilePath`tFehler=$($_.Exception.Message)"
        }
    }
}

# --------------------------------------------------------------------------------------
# Bestaetigungs- und Retry-Funktionen
# --------------------------------------------------------------------------------------

function Confirm-CleanupDeletion {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$CutoffDate,

        [Parameter(Mandatory = $true)]
        [string]$TargetLogPath,

        [Parameter(Mandatory = $true)]
        [int]$FoundFileCount
    )

    if ($DryRun -or $WhatIfPreference) {
        return
    }

    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host "WARNUNG: ECHTER LOESCHLAUF" -ForegroundColor Yellow
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host "Es werden endgueltige Loeschvorgaenge ausgefuehrt." -ForegroundColor Yellow
    Write-Host "Dieses Skript darf ausschliesslich interaktiv und beaufsichtigt ausgefuehrt werden." -ForegroundColor Yellow
    Write-Host "Datei-Schwellen-Datum: $CutoffDate" -ForegroundColor Yellow
    Write-Host "Gefundene alte Dateien: $FoundFileCount" -ForegroundColor Yellow
    Write-Host "Nach dem Dateiloeschlauf werden zusaetzlich leere Verzeichnisse gesucht." -ForegroundColor Yellow
    Write-Host "Leere Verzeichnisse werden unabhaengig vom Alter geloescht." -ForegroundColor Yellow
    Write-Host "Logdatei: $TargetLogPath" -ForegroundColor Yellow
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host ""

    $answer = Read-Host "Zum Fortfahren bitte exakt JA eingeben"

    if ($answer -cne "JA") {
        Add-CleanupLog "ABBRUCH`tBenutzer hat die Sicherheitsbestaetigung nicht erteilt."
        throw "Abbruch: Benutzerbestaetigung wurde nicht erteilt."
    }

    Add-CleanupLog "CONFIRM`tBenutzer hat echten Loeschlauf bestaetigt."
}

function Invoke-CleanupRetry {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory = $true)]
        [string]$Description,

        [Parameter(Mandatory = $false)]
        [int]$MaximumAttempt = 3
    )

    for ($attempt = 1; $attempt -le $MaximumAttempt; $attempt++) {
        try {
            & $ScriptBlock

            return [PSCustomObject]@{
                Success = $true
                Attempt = $attempt
                Error   = $null
            }
        }
        catch {
            $message = $_.Exception.Message

            if ($attempt -lt $MaximumAttempt) {
                $delaySecond = [int][Math]::Pow(2, ($attempt - 1))
                Add-CleanupLog "RETRY`t$Description`tVersuch=$attempt/$MaximumAttempt`tWarteSekunden=$delaySecond`tFehler=$message"
                Start-Sleep -Seconds $delaySecond
            }
            else {
                return [PSCustomObject]@{
                    Success = $false
                    Attempt = $attempt
                    Error   = $message
                }
            }
        }
    }
}

# --------------------------------------------------------------------------------------
# Dateisuche: Dateien aelter als Datei-Schwellen-Datum
# --------------------------------------------------------------------------------------

function Find-CleanupOldFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StartPath,

        [Parameter(Mandatory = $true)]
        [string]$CandidateFilePath
    )

    $longPath = ConvertTo-CleanupLongPath -InputPath $StartPath
    $displayPath = ConvertFrom-CleanupLongPath -InputPath $longPath

    Add-CleanupLog "FIND_FILE_START`tPfad=$displayPath"
    Write-CleanupInfo "Suche alte Dateien in: $displayPath"

    try {
        $attribute = [System.IO.File]::GetAttributes($longPath)
        $isDirectory = (($attribute -band [System.IO.FileAttributes]::Directory) -ne 0)

        if (-not $isDirectory) {
            $file = New-Object System.IO.FileInfo($longPath)

            if (Test-CleanupExclusion -Item $file -Pattern $ExcludePattern) {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tDatei ausgeschlossen`tPfad=$displayPath"
                return
            }

            if (Test-CleanupReparsePoint -Item $file) {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tReparsePoint/SymbolicLink-Datei uebersprungen`tPfad=$displayPath"
                return
            }

            $script:ProcessedFileCount++

            if ($file.LastWriteTime -lt $script:CutoffDate) {
                Add-CleanupCandidate -CandidateFilePath $CandidateFilePath -Type "File" -FullName $file.FullName -LastWriteTime $file.LastWriteTime
            }

            return
        }

        $rootDirectory = New-Object System.IO.DirectoryInfo($longPath)

        if (Test-CleanupReparsePoint -Item $rootDirectory) {
            $script:SkippedItemCount++
            Add-CleanupLog "SKIP`tStartverzeichnis ist ReparsePoint/SymbolicLink und wird nicht traversiert`tPfad=$displayPath"
            return
        }

        $directoryStack = New-Object System.Collections.Stack
        $directoryStack.Push($rootDirectory)

        while ($directoryStack.Count -gt 0) {
            $directory = [System.IO.DirectoryInfo]$directoryStack.Pop()
            $script:ProcessedDirectoryCount++

            if (($script:ProcessedDirectoryCount % 250) -eq 0) {
                Write-Progress `
                    -Activity "Suche alte Dateien" `
                    -Status "Aktuell: $(ConvertFrom-CleanupLongPath -InputPath $directory.FullName)" `
                    -CurrentOperation "Gefundene Dateien: $script:FoundFileCount | Verzeichnisse verarbeitet: $script:ProcessedDirectoryCount"
            }

            foreach ($entryPath in (Get-CleanupFileSystemEntry -Directory $directory)) {
                try {
                    $entryAttribute = [System.IO.File]::GetAttributes($entryPath)
                    $entryIsDirectory = (($entryAttribute -band [System.IO.FileAttributes]::Directory) -ne 0)

                    if ($entryIsDirectory) {
                        $childDirectory = New-Object System.IO.DirectoryInfo($entryPath)
                        $childDisplayPath = ConvertFrom-CleanupLongPath -InputPath $childDirectory.FullName

                        if (Test-CleanupExclusion -Item $childDirectory -Pattern $ExcludePattern) {
                            $script:SkippedItemCount++
                            Add-CleanupLog "SKIP_SUBTREE`tVerzeichnis ausgeschlossen; Unterstruktur wird nicht traversiert`tPfad=$childDisplayPath"
                            continue
                        }

                        if (Test-CleanupReparsePoint -Item $childDirectory) {
                            $script:SkippedItemCount++
                            Add-CleanupLog "SKIP`tReparsePoint/SymbolicLink-Verzeichnis nicht traversiert`tPfad=$childDisplayPath"
                            continue
                        }

                        $directoryStack.Push($childDirectory)
                    }
                    else {
                        $file = New-Object System.IO.FileInfo($entryPath)
                        $fileDisplayPath = ConvertFrom-CleanupLongPath -InputPath $file.FullName

                        if (Test-CleanupExclusion -Item $file -Pattern $ExcludePattern) {
                            $script:SkippedItemCount++
                            Add-CleanupLog "SKIP`tDatei ausgeschlossen`tPfad=$fileDisplayPath"
                            continue
                        }

                        if (Test-CleanupReparsePoint -Item $file) {
                            $script:SkippedItemCount++
                            Add-CleanupLog "SKIP`tReparsePoint/SymbolicLink-Datei uebersprungen`tPfad=$fileDisplayPath"
                            continue
                        }

                        $script:ProcessedFileCount++

                        if ($file.LastWriteTime -lt $script:CutoffDate) {
                            Add-CleanupCandidate -CandidateFilePath $CandidateFilePath -Type "File" -FullName $file.FullName -LastWriteTime $file.LastWriteTime
                        }
                    }
                }
                catch {
                    $script:ErrorCount++
                    Add-CleanupLog "ERROR`tEintrag konnte bei Dateisuche nicht verarbeitet werden`tPfad=$entryPath`tFehler=$($_.Exception.Message)"
                }
            }
        }
    }
    catch {
        $script:ErrorCount++
        Add-CleanupLog "ERROR`tStartpfad konnte bei Dateisuche nicht verarbeitet werden`tPfad=$displayPath`tFehler=$($_.Exception.Message)"
    }
}

# --------------------------------------------------------------------------------------
# Verzeichnissuche: leere Verzeichnisse unabhaengig vom Alter
# --------------------------------------------------------------------------------------

function Find-CleanupEmptyDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StartPath,

        [Parameter(Mandatory = $true)]
        [string]$CandidateFilePath
    )

    $longPath = ConvertTo-CleanupLongPath -InputPath $StartPath
    $displayPath = ConvertFrom-CleanupLongPath -InputPath $longPath

    Add-CleanupLog "FIND_DIRECTORY_START`tPfad=$displayPath"
    Write-CleanupInfo "Suche leere Verzeichnisse in: $displayPath"

    try {
        $attribute = [System.IO.File]::GetAttributes($longPath)
        $isDirectory = (($attribute -band [System.IO.FileAttributes]::Directory) -ne 0)

        if (-not $isDirectory) {
            return
        }

        $rootDirectory = New-Object System.IO.DirectoryInfo($longPath)

        if (Test-CleanupReparsePoint -Item $rootDirectory) {
            $script:SkippedItemCount++
            Add-CleanupLog "SKIP`tStartverzeichnis ist ReparsePoint/SymbolicLink und wird nicht fuer Verzeichnisse traversiert`tPfad=$displayPath"
            return
        }

        $directoryStack = New-Object System.Collections.Stack
        $directoryStack.Push([PSCustomObject]@{
                Directory = $rootDirectory
                Visited   = $false
                IsRoot    = $true
            })

        while ($directoryStack.Count -gt 0) {
            $frame = $directoryStack.Pop()
            $directory = [System.IO.DirectoryInfo]$frame.Directory
            $visited = [bool]$frame.Visited
            $isRoot = [bool]$frame.IsRoot

            if (-not $visited) {
                $directoryStack.Push([PSCustomObject]@{
                        Directory = $directory
                        Visited   = $true
                        IsRoot    = $isRoot
                    })

                foreach ($entryPath in (Get-CleanupFileSystemEntry -Directory $directory)) {
                    try {
                        $entryAttribute = [System.IO.File]::GetAttributes($entryPath)
                        $entryIsDirectory = (($entryAttribute -band [System.IO.FileAttributes]::Directory) -ne 0)

                        if (-not $entryIsDirectory) {
                            continue
                        }

                        $childDirectory = New-Object System.IO.DirectoryInfo($entryPath)
                        $childDisplayPath = ConvertFrom-CleanupLongPath -InputPath $childDirectory.FullName

                        if (Test-CleanupExclusion -Item $childDirectory -Pattern $ExcludePattern) {
                            $script:SkippedItemCount++
                            Add-CleanupLog "SKIP_SUBTREE`tVerzeichnis ausgeschlossen bei Leerpruefung; Unterstruktur wird nicht traversiert`tPfad=$childDisplayPath"
                            continue
                        }

                        if (Test-CleanupReparsePoint -Item $childDirectory) {
                            $script:SkippedItemCount++
                            Add-CleanupLog "SKIP`tReparsePoint/SymbolicLink-Verzeichnis nicht traversiert bei Leerpruefung`tPfad=$childDisplayPath"
                            continue
                        }

                        $directoryStack.Push([PSCustomObject]@{
                                Directory = $childDirectory
                                Visited   = $false
                                IsRoot    = $false
                            })
                    }
                    catch {
                        $script:ErrorCount++
                        Add-CleanupLog "ERROR`tEintrag konnte bei Verzeichnissuche nicht verarbeitet werden`tPfad=$entryPath`tFehler=$($_.Exception.Message)"
                    }
                }
            }
            else {
                if ($isRoot) {
                    Add-CleanupLog "SKIP`tStartverzeichnis wird nicht geloescht`tPfad=$(ConvertFrom-CleanupLongPath -InputPath $directory.FullName)"
                    continue
                }

                try {
                    $freshDirectory = New-Object System.IO.DirectoryInfo($directory.FullName)

                    if (-not $freshDirectory.Exists) {
                        continue
                    }

                    if (Test-CleanupExclusion -Item $freshDirectory -Pattern $ExcludePattern) {
                        $script:SkippedItemCount++
                        Add-CleanupLog "SKIP_SUBTREE`tVerzeichnis ausgeschlossen vor Kandidatenerfassung; Unterstruktur wird nicht traversiert`tPfad=$(ConvertFrom-CleanupLongPath -InputPath $freshDirectory.FullName)"
                        continue
                    }

                    if (Test-CleanupReparsePoint -Item $freshDirectory) {
                        $script:SkippedItemCount++
                        Add-CleanupLog "SKIP`tReparsePoint/SymbolicLink-Verzeichnis uebersprungen vor Kandidatenerfassung`tPfad=$(ConvertFrom-CleanupLongPath -InputPath $freshDirectory.FullName)"
                        continue
                    }

                    if (-not (Test-CleanupDirectoryEmpty -Directory $freshDirectory)) {
                        continue
                    }

                    Add-CleanupCandidate `
                        -CandidateFilePath $CandidateFilePath `
                        -Type "Directory" `
                        -FullName $freshDirectory.FullName `
                        -LastWriteTime $freshDirectory.LastWriteTime
                }
                catch {
                    $script:ErrorCount++
                    Add-CleanupLog "ERROR`tVerzeichnis konnte bei finaler Leerpruefung nicht verarbeitet werden`tPfad=$($directory.FullName)`tFehler=$($_.Exception.Message)"
                }
            }
        }
    }
    catch {
        $script:ErrorCount++
        Add-CleanupLog "ERROR`tStartpfad konnte bei Verzeichnissuche nicht verarbeitet werden`tPfad=$displayPath`tFehler=$($_.Exception.Message)"
    }
}

# --------------------------------------------------------------------------------------
# Loeschfunktionen
# --------------------------------------------------------------------------------------

function Remove-CleanupFileCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CandidateFilePath
    )

    foreach ($candidate in (Get-CleanupCandidate -CandidateFilePath $CandidateFilePath)) {
        try {
            $fullName = [string]$candidate.FullName
            $displayPath = [string]$candidate.DisplayPath

            if ([string]::IsNullOrWhiteSpace($fullName)) {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tKandidat ohne Pfad uebersprungen`tCandidateFilePath=$CandidateFilePath"
                continue
            }

            if ($DryRun) {
                Add-CleanupLog "DRYRUN`tDatei wuerde geloescht werden`tPfad=$displayPath"
                continue
            }

            if ($PSCmdlet.ShouldProcess($displayPath, "Datei endgueltig loeschen")) {
                $result = Invoke-CleanupRetry -Description "Datei loeschen: $displayPath" -ScriptBlock {
                    Remove-Item -LiteralPath $fullName -Force -ErrorAction Stop
                }

                if ($result.Success) {
                    $script:RemovedFileCount++
                    Add-CleanupLog "REMOVED_FILE`tPfad=$displayPath`tAttempts=$($result.Attempt)"
                }
                else {
                    $script:ErrorCount++
                    Add-CleanupLog "ERROR`tDatei konnte nicht geloescht werden`tPfad=$displayPath`tAttempts=$($result.Attempt)`tExitCode=N/A`tFehler=$($result.Error)"
                }
            }
            else {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tShouldProcess/WhatIf hat Dateiloeschung verhindert`tPfad=$displayPath"
            }
        }
        catch {
            $script:ErrorCount++
            Add-CleanupLog "ERROR`tDateikandidat konnte nicht geloescht/verarbeitet werden`tFehler=$($_.Exception.Message)"
        }
    }
}

function Remove-CleanupDirectoryCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CandidateFilePath
    )

    foreach ($candidate in (Get-CleanupCandidate -CandidateFilePath $CandidateFilePath)) {
        try {
            $fullName = [string]$candidate.FullName
            $displayPath = [string]$candidate.DisplayPath

            if ([string]::IsNullOrWhiteSpace($fullName)) {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tVerzeichniskandidat ohne Pfad uebersprungen`tCandidateFilePath=$CandidateFilePath"
                continue
            }

            if (-not (Test-Path -LiteralPath $fullName -PathType Container)) {
                Add-CleanupLog "SKIP`tVerzeichnis existiert nicht mehr`tPfad=$displayPath"
                continue
            }

            $freshDirectory = New-Object System.IO.DirectoryInfo($fullName)

            if (-not (Test-CleanupDirectoryEmpty -Directory $freshDirectory)) {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tVerzeichnis ist nicht mehr leer`tPfad=$displayPath"
                continue
            }

            if ($DryRun) {
                Add-CleanupLog "DRYRUN`tLeeres Verzeichnis wuerde geloescht werden`tPfad=$displayPath"
                continue
            }

            if ($PSCmdlet.ShouldProcess($displayPath, "Leeres Verzeichnis endgueltig loeschen")) {
                $result = Invoke-CleanupRetry -Description "Leeres Verzeichnis loeschen: $displayPath" -ScriptBlock {
                    Remove-Item -LiteralPath $fullName -Force -ErrorAction Stop
                }

                if ($result.Success) {
                    $script:RemovedDirectoryCount++
                    Add-CleanupLog "REMOVED_DIRECTORY`tPfad=$displayPath`tAttempts=$($result.Attempt)"
                }
                else {
                    $script:ErrorCount++
                    Add-CleanupLog "ERROR`tLeeres Verzeichnis konnte nicht geloescht werden`tPfad=$displayPath`tAttempts=$($result.Attempt)`tExitCode=N/A`tFehler=$($result.Error)"
                }
            }
            else {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tShouldProcess/WhatIf hat Verzeichnisloeschung verhindert`tPfad=$displayPath"
            }
        }
        catch {
            $script:ErrorCount++
            Add-CleanupLog "ERROR`tVerzeichniskandidat konnte nicht geloescht/verarbeitet werden`tFehler=$($_.Exception.Message)"
        }
    }
}

# --------------------------------------------------------------------------------------
# Aufraeumen der Kandidatendateien
# --------------------------------------------------------------------------------------

function Remove-CleanupCandidateFile {
    foreach ($candidateFilePath in $script:CandidateFileList) {
        try {
            if (Test-Path -LiteralPath $candidateFilePath -PathType Leaf) {
                Remove-Item -LiteralPath $candidateFilePath -Force -ErrorAction Stop
                Add-CleanupLog "INFO`tKandidatendatei entfernt`tPfad=$candidateFilePath"
            }
        }
        catch {
            Add-CleanupLog "WARNUNG`tKandidatendatei konnte nicht geloescht werden`tPfad=$candidateFilePath`tFehler=$($_.Exception.Message)"
        }
    }
}

# --------------------------------------------------------------------------------------
# Hauptprogramm
# --------------------------------------------------------------------------------------

try {
    Initialize-CleanupPlatform

    if (-not (Test-CleanupInteractiveSession)) {
        throw "Nicht-interaktive Ausfuehrung erkannt. Dieses Skript darf nicht automatisiert, unbeaufsichtigt, als geplanter Task oder als Hintergrundjob ausgefuehrt werden."
    }

    if (-not $PSBoundParameters.ContainsKey('FileCutoffDate')) {
        $FileCutoffDate = Read-Host "Bitte Datei-Stichtag eingeben im Format YYYYMMDDhh:mm. Keine Eingabe = Standard: aelter als $MinimumAgeInYears Jahre"
    }

    $script:CutoffDate = Resolve-CleanupCutoffDate `
        -FileCutoffDateInput $FileCutoffDate `
        -DefaultMinimumAgeInYears $MinimumAgeInYears

    $resolvedPathList = Resolve-CleanupPath -InputPath $Path
    $script:ResolvedLogPath = Initialize-CleanupLog -TargetLogPath $LogPath

    Add-CleanupLog "INFO`tAufgeloeste Eingabepfade:"
    foreach ($resolvedPath in $resolvedPathList) {
        Add-CleanupLog "INFO`tPATH`t$resolvedPath"
    }

    if ($ExcludePattern -and $ExcludePattern.Count -gt 0) {
        Add-CleanupLog "INFO`tAusschlussmuster:"
        foreach ($currentExcludePattern in $ExcludePattern) {
            Add-CleanupLog "INFO`tEXCLUDE_PATTERN`t$currentExcludePattern"
        }
    }

    Write-Host ""
    Write-CleanupInfo "Datei-Schwellen-Datum: $script:CutoffDate"
    Write-CleanupInfo "Quelle Datei-Schwellen-Datum: $script:EffectiveFileCutoffSource"
    Write-CleanupInfo "Leere Verzeichnisse werden unabhaengig vom Alter beruecksichtigt."
    Write-CleanupInfo "Logdatei: $script:ResolvedLogPath"
    Write-CleanupInfo "Kandidatendateien werden erstellt in: $script:CandidateDirectoryPath"

    if ($DryRun) {
        Write-CleanupWarning "DryRun ist aktiv. Es werden keine Dateien oder Verzeichnisse geloescht."
        Add-CleanupLog "MODE`tDryRun aktiv. Keine Loeschungen."
    }
    elseif ($WhatIfPreference) {
        Write-CleanupWarning "WhatIf ist aktiv. PowerShell verhindert tatsaechliche Loeschvorgaenge."
        Add-CleanupLog "MODE`tWhatIf aktiv. Keine tatsaechlichen Loeschungen."
    }

    # Phase 1: Alte Dateien streaming-basiert entdecken und in Kandidatendatei schreiben.
    $fileCandidatePath = New-CleanupCandidateFile -Kind "file"

    foreach ($resolvedPath in $resolvedPathList) {
        Find-CleanupOldFile -StartPath $resolvedPath -CandidateFilePath $fileCandidatePath
    }

    Write-Progress -Activity "Suche alte Dateien" -Completed

    Add-CleanupLog "PHASE_SUMMARY`tDateisuche abgeschlossen`tFoundFileCount=$script:FoundFileCount`tProcessedFileCount=$script:ProcessedFileCount`tProcessedDirectoryCount=$script:ProcessedDirectoryCount"

    # Sicherheitsabfrage erst nach Kandidatenermittlung, aber vor der ersten echten Loeschung.
    Confirm-CleanupDeletion `
        -CutoffDate $script:CutoffDate `
        -TargetLogPath $script:ResolvedLogPath `
        -FoundFileCount $script:FoundFileCount

    # Phase 2: Dateien loeschen beziehungsweise im DryRun/WhatIf protokollieren.
    Write-CleanupInfo "Verarbeite Dateikandidaten..."
    Remove-CleanupFileCandidate -CandidateFilePath $fileCandidatePath

    # Phase 3: Leere Verzeichnisse nach Dateiloeschung entdecken.
    $directoryCandidatePath = New-CleanupCandidateFile -Kind "directory"

    foreach ($resolvedPath in $resolvedPathList) {
        Find-CleanupEmptyDirectory -StartPath $resolvedPath -CandidateFilePath $directoryCandidatePath
    }

    Add-CleanupLog "PHASE_SUMMARY`tVerzeichnissuche abgeschlossen`tFoundDirectoryCount=$script:FoundDirectoryCount"

    # Phase 4: Leere Verzeichnisse loeschen beziehungsweise im DryRun/WhatIf protokollieren.
    Write-CleanupInfo "Verarbeite Verzeichniskandidaten..."
    Remove-CleanupDirectoryCandidate -CandidateFilePath $directoryCandidatePath

    Write-Progress -Activity "Cleanup" -Completed

    Add-CleanupLog ""
    Add-CleanupLog "================================================================================"
    Add-CleanupLog "ZUSAMMENFASSUNG"
    Add-CleanupLog "================================================================================"
    Add-CleanupLog "Gefundene alte Dateien              : $script:FoundFileCount"
    Add-CleanupLog "Gefundene leere Verzeichnisse       : $script:FoundDirectoryCount"
    Add-CleanupLog "Geloeschte Dateien                  : $script:RemovedFileCount"
    Add-CleanupLog "Geloeschte Verzeichnisse            : $script:RemovedDirectoryCount"
    Add-CleanupLog "Uebersprungene Eintraege            : $script:SkippedItemCount"
    Add-CleanupLog "Fehler                              : $script:ErrorCount"
    Add-CleanupLog "Endzeit                             : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Add-CleanupLog "================================================================================"

    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "ZUSAMMENFASSUNG" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "Gefundene alte Dateien             : $script:FoundFileCount"
    Write-Host "Gefundene leere Verzeichnisse      : $script:FoundDirectoryCount"
    Write-Host "Geloeschte Dateien                 : $script:RemovedFileCount"
    Write-Host "Geloeschte Verzeichnisse           : $script:RemovedDirectoryCount"
    Write-Host "Uebersprungene Eintraege           : $script:SkippedItemCount"
    Write-Host "Fehler                             : $script:ErrorCount"
    Write-Host "Logdatei                           : $script:ResolvedLogPath"
    Write-Host "Kandidatendatei-Ordner             : $script:CandidateDirectoryPath"
    Write-Host "================================================================================" -ForegroundColor Cyan

    Remove-CleanupCandidateFile

    if ($script:ErrorCount -gt 0) {
        Write-CleanupError "Abgeschlossen mit Fehlern. Details stehen in der Logdatei."
        exit 1
    }

    Write-CleanupSuccess "Abgeschlossen ohne protokollierte Fehler."
    exit 0
}
catch {
    try {
        if (-not [string]::IsNullOrWhiteSpace($script:ResolvedLogPath)) {
            Add-CleanupLog "FATAL`t$($_.Exception.Message)"
        }
    }
    catch {
        # Falls selbst das Loggen fehlschlaegt, wird nur die Konsole genutzt.
    }

    try {
        Remove-CleanupCandidateFile
    }
    catch {
        # Keine Stoerungen beim finalen Cleanup der Kandidatendateien.
    }

    Write-CleanupError $_.Exception.Message
    exit 2
}