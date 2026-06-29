<#
.SYNOPSIS
    Sicheres, robustes und speicherschonendes Cleanup-Skript fuer alte Dateien.

.DESCRIPTION
    Findet und loescht ausschliesslich Dateien, deren LastWriteTime aelter ist als
    ein Datei-Schwellen-Datum.

    Unterstuetzt:
    - Path
    - PathListFile
    - ExcludePattern
    - ExcludePatternFile
    - FileCutoffDate
    - MinimumAgeInYears
    - DryRun
    - WhatIf / Confirm
    - LogPath

    Verzeichnisse werden NICHT geloescht.

.PARAMETER Path
    Ein oder mehrere Startpfade. Ein Startpfad darf eine Datei oder ein Verzeichnis sein.

.PARAMETER PathListFile
    Textdatei mit Startpfaden, eine Zeile pro Pfad.
    Leere Zeilen und Kommentarzeilen mit # werden ignoriert.

.PARAMETER ExcludePattern
    Direkte Ausschlussmuster auf der Kommandozeile.

.PARAMETER ExcludePatternFile
    Textdatei mit Startpfaden, eine Zeile pro Pfad.
    Leere Zeilen und Kommentarzeilen mit # werden ignoriert.
    
    # Einzelne Dateiendungen
    *.pst
    *.ost
    *.bak
    *.tmp
    
    # Einzelne Dateien
    Thumbs.db
    desktop.ini
    
    # Vollständige Ordner
    N:\IT-BL\ProjektA
    N:\IT-BL\ProjektB\Archiv
    \\ibbads\data\IT-BL\Test
    
    # Ordner mit Inhalt
    N:\IT-BL\ProjektC\DoNotDelete\*
    \\ibbads\data\IT-BL\Archiv\*
    
    # Ganze Projekte
    N:\IT-BL\ProjektD\*
    N:\IT-BL\ProjektE\*
    
    # Einzelne Datei
    N:\IT-BL\ProjektF\test.xlsx

.PARAMETER MinimumAgeInYears
    Standard-Mindestalter in Jahren, falls kein FileCutoffDate angegeben wird.
    Standard: 5.

.PARAMETER FileCutoffDate
    Optionaler Datei-Stichtag im Format YYYYMMDDhh:mm.
    Beispiel: 2021010100:00

.PARAMETER DryRun
    Wenn gesetzt, werden keine Dateien geloescht.

.PARAMETER LogPath
    Pfad zur Logdatei oder zu einem bestehenden Log-Verzeichnis.

.EXAMPLE
    .\cleanup.ps1 -Path "N:\IT-BL" -DryRun

.EXAMPLE
    .\cleanup.ps1 -Path "N:\IT-BL" -FileCutoffDate "2021010100:00" -DryRun

.EXAMPLE
    .\cleanup.ps1 -PathListFile "C:\Cleanup\paths.txt" -ExcludePatternFile "C:\Cleanup\exclude.txt" -DryRun

.EXAMPLE
    .\cleanup.ps1 -Path "N:\IT-BL" -ExcludePattern "*.pst","*.ost","*\DoNotDelete" -DryRun

.NOTES
    Autor: Stefan Homberg
    Dateiname: cleanup.ps1
    Version: 1.2.0

    Kompatibilitaet:
    - Windows PowerShell 5.1
    - PowerShell 7+

    Betriebsregel:
    Dieses Skript ist bewusst fuer interaktive Ausfuehrung gebaut.
    Es soll nie automatisiert, unbeaufsichtigt, als geplanter Task oder als Hintergrundjob laufen.

    Performance-Strategie:
    - Streaming statt RAM-Listen.
    - .NET EnumerateFileSystemEntries.
    - Gepufferte Writer fuer Log und Kandidatendatei.
    - Keine Regex.
    - Keine In-Memory-Deduplizierung aller Kandidaten.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Path,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$PathListFile,

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludePattern,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ExcludePatternFile,

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
    [string]$LogPath = $(Join-Path -Path (Get-Location) -ChildPath ("to_delete_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")))
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:IsWindowsLike              = $true
$script:CutoffDate                 = $null
$script:ResolvedLogPath            = $null
$script:CandidateDirectoryPath     = $null
$script:CandidateFileList          = @()
$script:LogWriter                  = $null
$script:CandidateWriter            = $null
$script:FoundFileCount             = 0
$script:RemovedFileCount           = 0
$script:SkippedItemCount           = 0
$script:ErrorCount                 = 0
$script:ProcessedDirectoryCount    = 0
$script:ProcessedFileCount         = 0
$script:ProcessedPathListLineCount = 0
$script:LogLineCount               = 0
$script:CandidateLineCount         = 0
$script:EffectiveFileCutoffSource  = $null
$script:ExcludeRuleList            = New-Object System.Collections.Generic.List[object]

function Write-CleanupInfo {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "[INFO ] $Message" -ForegroundColor Cyan
}

function Write-CleanupSuccess {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "[ OK  ] $Message" -ForegroundColor Green
}

function Write-CleanupWarning {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "[WARN ] $Message" -ForegroundColor Yellow
}

function Write-CleanupError {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "[FEHLER] $Message" -ForegroundColor Red
}

function Get-CurrentUserName {
    try {
        return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }
    catch {
        if (-not [string]::IsNullOrWhiteSpace($env:USERNAME)) { return $env:USERNAME }
        if (-not [string]::IsNullOrWhiteSpace($env:USER)) { return $env:USER }
        return "Unbekannt"
    }
}

function Get-SafeComputerName {
    if (-not [string]::IsNullOrWhiteSpace($env:COMPUTERNAME)) { return $env:COMPUTERNAME }
    if (-not [string]::IsNullOrWhiteSpace($env:HOSTNAME)) { return $env:HOSTNAME }

    try {
        return [System.Net.Dns]::GetHostName()
    }
    catch {
        return "Unbekannt"
    }
}

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

function Resolve-CleanupLogPath {
    param([Parameter(Mandatory = $true)][string]$TargetLogPath)

    if (Test-Path -LiteralPath $TargetLogPath -PathType Container) {
        return Join-Path -Path $TargetLogPath -ChildPath ("to_delete_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    }

    $logDirectory = Split-Path -Path $TargetLogPath -Parent

    if ([string]::IsNullOrWhiteSpace($logDirectory)) {
        return Join-Path -Path (Get-Location).Path -ChildPath $TargetLogPath
    }

    return $TargetLogPath
}

function Open-CleanupLog {
    param([Parameter(Mandatory = $true)][string]$TargetLogPath)

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
        $testFile = Join-Path -Path $logDirectory -ChildPath (".__cleanup_write_test_{0}.tmp" -f ([Guid]::NewGuid().ToString("N")))
        [System.IO.File]::WriteAllText($testFile, "write-test", [System.Text.Encoding]::UTF8)
        Remove-Item -LiteralPath $testFile -Force -ErrorAction Stop
    }
    catch {
        throw "Keine Schreibrechte fuer das Log-Verzeichnis '$logDirectory'. Fehler: $($_.Exception.Message)"
    }

    $script:ResolvedLogPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($effectiveLogPath)
    $script:CandidateDirectoryPath = Split-Path -Path $script:ResolvedLogPath -Parent

    if ([string]::IsNullOrWhiteSpace($script:CandidateDirectoryPath)) {
        $script:CandidateDirectoryPath = (Get-Location).Path
    }

    $script:LogWriter = New-Object System.IO.StreamWriter($script:ResolvedLogPath, $false, [System.Text.Encoding]::UTF8, 65536)
    $script:LogWriter.AutoFlush = $false
}

function Close-CleanupLog {
    try {
        if ($null -ne $script:LogWriter) {
            $script:LogWriter.Flush()
            $script:LogWriter.Close()
            $script:LogWriter.Dispose()
            $script:LogWriter = $null
        }
    }
    catch {
        # Fehler beim finalen Schliessen des Log-Writers werden nicht weitergegeben,
        # damit sie keinen urspruenglichen Cleanup-Fehler ueberdecken.
    }
}

function Add-CleanupLog {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message
    )

    if ($null -eq $script:LogWriter) {
        return
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $line = "[{0}] {1}" -f $timestamp, $Message

    $script:LogWriter.WriteLine($line)
    $script:LogLineCount++

    if (($script:LogLineCount % 1000) -eq 0) {
        $script:LogWriter.Flush()
    }
}

function Write-CleanupLogHeader {
    $header = @(
        "================================================================================"
        "Cleanup-Log"
        "================================================================================"
        "Skript                    : cleanup.ps1"
        "Version                   : 1.2.0"
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
        "PathListFile              : $PathListFile"
        "ExcludePatternFile        : $ExcludePatternFile"
        "ExcludeRuleCount          : $($script:ExcludeRuleList.Count)"
        "LogPath                   : $script:ResolvedLogPath"
        "Kandidatendatei-Ordner    : $script:CandidateDirectoryPath"
        "Verzeichnisloeschung      : Nein"
        "Ausschlussmuster          : PowerShell-Wildcards und optionale Praefixe FILE:, DIR:, PATH:"
        "RAM-Strategie             : Streaming, Kandidatendatei, keine In-Memory-Kandidatenliste"
        "ACHTUNG                   : Bei echtem Lauf werden protokollierte Dateien endgueltig geloescht."
        "================================================================================"
        ""
    )

    foreach ($line in $header) {
        Add-CleanupLog $line
    }

    $script:LogWriter.Flush()
}

function ConvertTo-CleanupLongPath {
    param([Parameter(Mandatory = $true)][string]$InputPath)

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
    param([Parameter(Mandatory = $true)][string]$InputPath)

    if ($InputPath.StartsWith("\\?\UNC\")) {
        return "\\" + $InputPath.Substring(8)
    }

    if ($InputPath.StartsWith("\\?\")) {
        return $InputPath.Substring(4)
    }

    return $InputPath
}

function Resolve-CleanupPath {
    param([Parameter(Mandatory = $true)][string]$InputPath)

    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        throw "Ein uebergebener Pfad ist leer."
    }

    if (-not (Test-Path -LiteralPath $InputPath)) {
        throw "Pfad existiert nicht: $InputPath"
    }

    try {
        $resolvedItem = Resolve-Path -LiteralPath $InputPath -ErrorAction Stop | Select-Object -First 1

        if ($null -eq $resolvedItem -or [string]::IsNullOrWhiteSpace($resolvedItem.ProviderPath)) {
            throw "Pfad konnte nicht eindeutig aufgeloest werden: $InputPath"
        }

        return $resolvedItem.ProviderPath
    }
    catch {
        throw "Pfad konnte nicht aufgeloest werden: $InputPath. Fehler: $($_.Exception.Message)"
    }
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

function Test-CleanupInvalidControlCharacter {
    param([Parameter(Mandatory = $true)][string]$Value)

    foreach ($character in $Value.ToCharArray()) {
        if ([int][char]$character -lt 32) {
            return $true
        }
    }

    return $false
}

function Test-CleanupReparsePoint {
    param([Parameter(Mandatory = $true)][System.IO.FileSystemInfo]$Item)

    return (($Item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
}

function New-CleanupExcludeRule {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RawPattern,

        [Parameter(Mandatory = $true)]
        [string]$Source
    )

    $pattern = $RawPattern.Trim()

    if ([string]::IsNullOrWhiteSpace($pattern)) {
        return $null
    }

    if (Test-CleanupInvalidControlCharacter -Value $pattern) {
        throw "Ausschlussmuster enthaelt unzulaessige Steuerzeichen. Quelle=$Source"
    }

    $scope = "Any"
    $value = $pattern

    if ($pattern -match '^(FILE|DIR|PATH):(.*)$') {
        $scope = $matches[1]
        $value = $matches[2].Trim()
    }

    if ([string]::IsNullOrWhiteSpace($value)) {
        return $null
    }

    return [PSCustomObject]@{
        Scope     = $scope
        Pattern   = $value
        HasWildcard = ($value.IndexOfAny([char[]]@('*', '?', '[')) -ge 0)
        Source    = $Source
    }
}

function Add-CleanupExcludeRule {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Pattern,

        [Parameter(Mandatory = $true)]
        [string]$Source
    )

    $rule = New-CleanupExcludeRule -RawPattern $Pattern -Source $Source

    if ($null -ne $rule) {
        [void]$script:ExcludeRuleList.Add($rule)
    }
}

function Initialize-CleanupExcludeRule {
    if ($ExcludePattern -and $ExcludePattern.Count -gt 0) {
        foreach ($currentPattern in $ExcludePattern) {
            Add-CleanupExcludeRule -Pattern $currentPattern -Source "ExcludePattern"
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ExcludePatternFile)) {
        if (-not (Test-Path -LiteralPath $ExcludePatternFile -PathType Leaf)) {
            throw "ExcludePatternFile existiert nicht oder ist keine Datei: $ExcludePatternFile"
        }

        $resolvedExcludePatternFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExcludePatternFile)
        $lineNumber = 0

        foreach ($line in [System.IO.File]::ReadLines($resolvedExcludePatternFile)) {
            $lineNumber++
            $currentLine = $line.Trim()

            if ([string]::IsNullOrWhiteSpace($currentLine)) {
                continue
            }

            if ($currentLine.StartsWith("#")) {
                continue
            }

            if (
                ($currentLine.StartsWith('"') -and $currentLine.EndsWith('"')) -or
                ($currentLine.StartsWith("'") -and $currentLine.EndsWith("'"))
            ) {
                $currentLine = $currentLine.Substring(1, $currentLine.Length - 2)
            }

            Add-CleanupExcludeRule -Pattern $currentLine -Source ("ExcludePatternFile:{0}" -f $lineNumber)
        }
    }
}

function Test-CleanupExclusion {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileSystemInfo]$Item
    )

    if ($script:ExcludeRuleList.Count -eq 0) {
        return $false
    }

    $fullName = ConvertFrom-CleanupLongPath -InputPath $Item.FullName
    $name = $Item.Name
    $isDirectory = (($Item.Attributes -band [System.IO.FileAttributes]::Directory) -ne 0)

    foreach ($rule in $script:ExcludeRuleList) {
        if ($rule.Scope -eq "FILE" -and $isDirectory) {
            continue
        }

        if ($rule.Scope -eq "DIR" -and -not $isDirectory) {
            continue
        }

        $pattern = [string]$rule.Pattern

        switch ($rule.Scope) {
            "FILE" {
                if ($rule.HasWildcard) {
                    if ($name -like $pattern -or $fullName -like $pattern) { return $true }
                }
                else {
                    if ($name -eq $pattern -or $fullName -eq $pattern) { return $true }
                }
            }

            "DIR" {
                if ($rule.HasWildcard) {
                    if ($fullName -like $pattern -or $name -like $pattern) { return $true }
                }
                else {
                    if ($fullName -eq $pattern -or $name -eq $pattern) { return $true }
                }
            }

            "PATH" {
                if ($rule.HasWildcard) {
                    if ($fullName -like $pattern) { return $true }
                }
                else {
                    if ($fullName -eq $pattern) { return $true }
                }
            }

            default {
                if ($rule.HasWildcard) {
                    if ($fullName -like $pattern -or $name -like $pattern) { return $true }
                }
                else {
                    if ($fullName -eq $pattern -or $name -eq $pattern) { return $true }
                }
            }
        }
    }

    return $false
}

function Get-CleanupFileSystemEntry {
    param([Parameter(Mandatory = $true)][string]$DirectoryPath)

    try {
        return [System.IO.Directory]::EnumerateFileSystemEntries($DirectoryPath)
    }
    catch {
        $script:ErrorCount++
        Add-CleanupLog "ERROR`tVerzeichnis kann nicht gelesen werden`tPfad=$DirectoryPath`tFehler=$($_.Exception.Message)"
        return @()
    }
}

function New-CleanupCandidateFile {
    if ([string]::IsNullOrWhiteSpace($script:CandidateDirectoryPath)) {
        throw "Der Kandidatendatei-Ordner wurde nicht initialisiert."
    }

    if (-not (Test-Path -LiteralPath $script:CandidateDirectoryPath -PathType Container)) {
        throw "Der Kandidatendatei-Ordner existiert nicht: $script:CandidateDirectoryPath"
    }

    $candidateFileName = "cleanup_candidates_file_{0}.jsonl" -f ([Guid]::NewGuid().ToString("N"))
    $candidateFilePath = Join-Path -Path $script:CandidateDirectoryPath -ChildPath $candidateFileName

    [System.IO.File]::WriteAllText($candidateFilePath, "", [System.Text.Encoding]::UTF8)
    $script:CandidateFileList += $candidateFilePath

    Add-CleanupLog "INFO`tKandidatendatei erstellt`tPfad=$candidateFilePath"

    return $candidateFilePath
}

function Open-CleanupCandidateWriter {
    param([Parameter(Mandatory = $true)][string]$CandidateFilePath)

    $script:CandidateWriter = New-Object System.IO.StreamWriter($CandidateFilePath, $false, [System.Text.Encoding]::UTF8, 65536)
    $script:CandidateWriter.AutoFlush = $false
}

function Close-CleanupCandidateWriter {
    try {
        if ($null -ne $script:CandidateWriter) {
            $script:CandidateWriter.Flush()
            $script:CandidateWriter.Close()
            $script:CandidateWriter.Dispose()
            $script:CandidateWriter = $null
        }
    }
    catch {
        Add-CleanupLog "WARNUNG`tKandidatendatei-Writer konnte nicht sauber geschlossen werden`tFehler=$($_.Exception.Message)"
    }
}

function Add-CleanupCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FullName,

        [Parameter(Mandatory = $true)]
        [datetime]$LastWriteTime
    )

    if ($null -eq $script:CandidateWriter) {
        throw "Kandidatendatei-Writer ist nicht geoeffnet."
    }

    $displayPath = ConvertFrom-CleanupLongPath -InputPath $FullName

    $candidate = [PSCustomObject]@{
        Type             = "File"
        FullName         = $FullName
        DisplayPath      = $displayPath
        LastWriteTimeUtc = $LastWriteTime.ToUniversalTime().ToString("o")
    }

    $json = $candidate | ConvertTo-Json -Compress
    $script:CandidateWriter.WriteLine($json)
    $script:CandidateLineCount++

    $script:FoundFileCount++
    Add-CleanupLog "FOUND_FILE`tPfad=$displayPath`tLastWriteTime=$($LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"

    if (($script:CandidateLineCount % 10000) -eq 0) {
        $script:CandidateWriter.Flush()
    }
}

function Get-CleanupCandidate {
    param([Parameter(Mandatory = $true)][string]$CandidateFilePath)

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
    Write-Host "Verzeichnisse werden nicht geloescht." -ForegroundColor Yellow
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

function Add-CleanupFileCandidateIfOld {
    param([Parameter(Mandatory = $true)][string]$FilePath)

    try {
        $longPath = ConvertTo-CleanupLongPath -InputPath $FilePath
        $file = New-Object System.IO.FileInfo($longPath)
        $displayPath = ConvertFrom-CleanupLongPath -InputPath $file.FullName

        if (-not $file.Exists) {
            $script:SkippedItemCount++
            Add-CleanupLog "SKIP`tDatei existiert nicht`tPfad=$FilePath"
            return
        }

        if (Test-CleanupExclusion -Item $file) {
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
            Add-CleanupCandidate -FullName $file.FullName -LastWriteTime $file.LastWriteTime
        }
    }
    catch {
        $script:ErrorCount++
        Add-CleanupLog "ERROR`tDatei konnte nicht verarbeitet werden`tPfad=$FilePath`tFehler=$($_.Exception.Message)"
    }
}

function Find-CleanupOldFile {
    param([Parameter(Mandatory = $true)][string]$StartPath)

    $resolvedStartPath = Resolve-CleanupPath -InputPath $StartPath
    $longPath = ConvertTo-CleanupLongPath -InputPath $resolvedStartPath
    $displayPath = ConvertFrom-CleanupLongPath -InputPath $longPath

    Add-CleanupLog "FIND_FILE_START`tPfad=$displayPath"
    Write-CleanupInfo "Suche alte Dateien in: $displayPath"

    try {
        $attribute = [System.IO.File]::GetAttributes($longPath)
        $isDirectory = (($attribute -band [System.IO.FileAttributes]::Directory) -ne 0)

        if (-not $isDirectory) {
            Add-CleanupFileCandidateIfOld -FilePath $resolvedStartPath
            return
        }

        if (($attribute -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            $script:SkippedItemCount++
            Add-CleanupLog "SKIP`tStartverzeichnis ist ReparsePoint/SymbolicLink und wird nicht traversiert`tPfad=$displayPath"
            return
        }

        $directoryStack = New-Object System.Collections.Stack
        $directoryStack.Push($longPath)

        while ($directoryStack.Count -gt 0) {
            $directoryPath = [string]$directoryStack.Pop()
            $script:ProcessedDirectoryCount++

            if (($script:ProcessedDirectoryCount % 250) -eq 0) {
                Write-Progress `
                    -Activity "Suche alte Dateien" `
                    -Status "Aktuell: $(ConvertFrom-CleanupLongPath -InputPath $directoryPath)" `
                    -CurrentOperation "Gefundene Dateien: $script:FoundFileCount | Verzeichnisse verarbeitet: $script:ProcessedDirectoryCount"
            }

            if (($script:ProcessedDirectoryCount % 5000) -eq 0) {
                Add-CleanupLog "PROGRESS`tVerzeichnisse verarbeitet=$script:ProcessedDirectoryCount`tDateien verarbeitet=$script:ProcessedFileCount`tKandidaten=$script:FoundFileCount"
            }

            foreach ($entryPath in (Get-CleanupFileSystemEntry -DirectoryPath $directoryPath)) {
                try {
                    $entryAttribute = [System.IO.File]::GetAttributes($entryPath)
                    $entryIsDirectory = (($entryAttribute -band [System.IO.FileAttributes]::Directory) -ne 0)

                    if ($entryIsDirectory) {
                        $childDirectory = New-Object System.IO.DirectoryInfo($entryPath)
                        $childDisplayPath = ConvertFrom-CleanupLongPath -InputPath $childDirectory.FullName

                        if (Test-CleanupExclusion -Item $childDirectory) {
                            $script:SkippedItemCount++
                            Add-CleanupLog "SKIP_SUBTREE`tVerzeichnis ausgeschlossen; Unterstruktur wird nicht traversiert`tPfad=$childDisplayPath"
                            continue
                        }

                        if (($entryAttribute -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                            $script:SkippedItemCount++
                            Add-CleanupLog "SKIP`tReparsePoint/SymbolicLink-Verzeichnis nicht traversiert`tPfad=$childDisplayPath"
                            continue
                        }

                        $directoryStack.Push($entryPath)
                    }
                    else {
                        Add-CleanupFileCandidateIfOld -FilePath $entryPath
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

function Find-CleanupOldFileFromList {
    param([Parameter(Mandatory = $true)][string]$InputListPath)

    if (-not (Test-Path -LiteralPath $InputListPath -PathType Leaf)) {
        throw "PathListFile existiert nicht oder ist keine Datei: $InputListPath"
    }

    $resolvedListPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($InputListPath)

    Add-CleanupLog "FIND_FILE_LIST_START`tPfad=$resolvedListPath"
    Write-CleanupInfo "Verarbeite Pfadliste: $resolvedListPath"

    $lineNumber = 0

    foreach ($line in [System.IO.File]::ReadLines($resolvedListPath)) {
        $lineNumber++
        $script:ProcessedPathListLineCount++

        $currentPath = $line.Trim()

        if ([string]::IsNullOrWhiteSpace($currentPath)) {
            continue
        }

        if ($currentPath.StartsWith("#")) {
            continue
        }

        if (
            ($currentPath.StartsWith('"') -and $currentPath.EndsWith('"')) -or
            ($currentPath.StartsWith("'") -and $currentPath.EndsWith("'"))
        ) {
            $currentPath = $currentPath.Substring(1, $currentPath.Length - 2)
        }

        if (Test-CleanupInvalidControlCharacter -Value $currentPath) {
            $script:SkippedItemCount++
            Add-CleanupLog "SKIP`tPathListFile-Zeile enthaelt unzulaessige Steuerzeichen`tZeile=$lineNumber"
            continue
        }

        if (($lineNumber % 5000) -eq 0) {
            Write-Progress `
                -Activity "Verarbeite PathListFile" `
                -Status "Zeile $lineNumber" `
                -CurrentOperation "Kandidaten: $script:FoundFileCount"
        }

        try {
            Find-CleanupOldFile -StartPath $currentPath
        }
        catch {
            $script:ErrorCount++
            Add-CleanupLog "ERROR`tPfad aus PathListFile konnte nicht verarbeitet werden`tZeile=$lineNumber`tPfad=$currentPath`tFehler=$($_.Exception.Message)"
        }
    }

    Write-Progress -Activity "Verarbeite PathListFile" -Completed
}

function Remove-CleanupFileCandidate {
    param([Parameter(Mandatory = $true)][string]$CandidateFilePath)

    foreach ($candidate in (Get-CleanupCandidate -CandidateFilePath $CandidateFilePath)) {
        try {
            $fullName = [string]$candidate.FullName
            $displayPath = [string]$candidate.DisplayPath

            if ([string]::IsNullOrWhiteSpace($fullName)) {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tKandidat ohne Pfad uebersprungen`tCandidateFilePath=$CandidateFilePath"
                continue
            }

            if (-not (Test-Path -LiteralPath $fullName -PathType Leaf)) {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tDatei existiert nicht mehr oder ist keine Datei`tPfad=$displayPath"
                continue
            }

            $freshFile = New-Object System.IO.FileInfo($fullName)

            if (Test-CleanupReparsePoint -Item $freshFile) {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tReparsePoint/SymbolicLink-Datei vor Loeschung uebersprungen`tPfad=$displayPath"
                continue
            }

            if (Test-CleanupExclusion -Item $freshFile) {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tDatei vor Loeschung ausgeschlossen`tPfad=$displayPath"
                continue
            }

            if ($freshFile.LastWriteTime -ge $script:CutoffDate) {
                $script:SkippedItemCount++
                Add-CleanupLog "SKIP`tDatei ist nicht mehr aelter als Datei-Schwellen-Datum`tPfad=$displayPath`tLastWriteTime=$($freshFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
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

try {
    Initialize-CleanupPlatform

    if (-not (Test-CleanupInteractiveSession)) {
        throw "Nicht-interaktive Ausfuehrung erkannt. Dieses Skript darf nicht automatisiert, unbeaufsichtigt, als geplanter Task oder als Hintergrundjob ausgefuehrt werden."
    }

    if ((-not $Path -or $Path.Count -eq 0) -and [string]::IsNullOrWhiteSpace($PathListFile)) {
        throw "Es muss mindestens -Path oder -PathListFile angegeben werden."
    }

    Initialize-CleanupExcludeRule

    if (-not $PSBoundParameters.ContainsKey('FileCutoffDate')) {
        $FileCutoffDate = Read-Host "Bitte Datei-Stichtag eingeben im Format YYYYMMDDhh:mm. Keine Eingabe = Standard: aelter als $MinimumAgeInYears Jahre"
    }

    $script:CutoffDate = Resolve-CleanupCutoffDate `
        -FileCutoffDateInput $FileCutoffDate `
        -DefaultMinimumAgeInYears $MinimumAgeInYears

    Open-CleanupLog -TargetLogPath $LogPath
    Write-CleanupLogHeader

    if ($Path -and $Path.Count -gt 0) {
        Add-CleanupLog "INFO`tEingabepfade aus -Path:"
        foreach ($inputPath in $Path) {
            Add-CleanupLog "INFO`tPATH`t$inputPath"
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($PathListFile)) {
        Add-CleanupLog "INFO`tPathListFile`t$PathListFile"
    }

    if (-not [string]::IsNullOrWhiteSpace($ExcludePatternFile)) {
        Add-CleanupLog "INFO`tExcludePatternFile`t$ExcludePatternFile"
    }

    Add-CleanupLog "INFO`tAusschlussregeln geladen`tCount=$($script:ExcludeRuleList.Count)"

    Write-Host ""
    Write-CleanupInfo "Datei-Schwellen-Datum: $script:CutoffDate"
    Write-CleanupInfo "Quelle Datei-Schwellen-Datum: $script:EffectiveFileCutoffSource"
    Write-CleanupInfo "Ausschlussregeln geladen: $($script:ExcludeRuleList.Count)"
    Write-CleanupInfo "Es werden ausschliesslich Dateien geloescht. Verzeichnisse werden nicht geloescht."
    Write-CleanupInfo "Logdatei: $script:ResolvedLogPath"
    Write-CleanupInfo "Kandidatendateien werden erstellt in: $script:CandidateDirectoryPath"

    if ($DryRun) {
        Write-CleanupWarning "DryRun ist aktiv. Es werden keine Dateien geloescht."
        Add-CleanupLog "MODE`tDryRun aktiv. Keine Loeschungen."
    }
    elseif ($WhatIfPreference) {
        Write-CleanupWarning "WhatIf ist aktiv. PowerShell verhindert tatsaechliche Loeschvorgaenge."
        Add-CleanupLog "MODE`tWhatIf aktiv. Keine tatsaechlichen Loeschungen."
    }

    $fileCandidatePath = New-CleanupCandidateFile
    Open-CleanupCandidateWriter -CandidateFilePath $fileCandidatePath

    if ($Path -and $Path.Count -gt 0) {
        foreach ($inputPath in $Path) {
            Find-CleanupOldFile -StartPath $inputPath
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($PathListFile)) {
        Find-CleanupOldFileFromList -InputListPath $PathListFile
    }

    Write-Progress -Activity "Suche alte Dateien" -Completed
    Close-CleanupCandidateWriter

    Add-CleanupLog "PHASE_SUMMARY`tDateisuche abgeschlossen`tFoundFileCount=$script:FoundFileCount`tProcessedFileCount=$script:ProcessedFileCount`tProcessedDirectoryCount=$script:ProcessedDirectoryCount`tPathListLines=$script:ProcessedPathListLineCount"

    Confirm-CleanupDeletion `
        -CutoffDate $script:CutoffDate `
        -TargetLogPath $script:ResolvedLogPath `
        -FoundFileCount $script:FoundFileCount

    Write-CleanupInfo "Verarbeite Dateikandidaten..."
    Remove-CleanupFileCandidate -CandidateFilePath $fileCandidatePath

    Write-Progress -Activity "Cleanup" -Completed

    Add-CleanupLog ""
    Add-CleanupLog "================================================================================"
    Add-CleanupLog "ZUSAMMENFASSUNG"
    Add-CleanupLog "================================================================================"
    Add-CleanupLog "Gefundene alte Dateien              : $script:FoundFileCount"
    Add-CleanupLog "Geloeschte Dateien                  : $script:RemovedFileCount"
    Add-CleanupLog "Uebersprungene Eintraege            : $script:SkippedItemCount"
    Add-CleanupLog "Verarbeitete Dateien                : $script:ProcessedFileCount"
    Add-CleanupLog "Verarbeitete Verzeichnisse          : $script:ProcessedDirectoryCount"
    Add-CleanupLog "Verarbeitete PathListFile-Zeilen    : $script:ProcessedPathListLineCount"
    Add-CleanupLog "Ausschlussregeln                    : $($script:ExcludeRuleList.Count)"
    Add-CleanupLog "Fehler                              : $script:ErrorCount"
    Add-CleanupLog "Endzeit                             : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Add-CleanupLog "================================================================================"

    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "ZUSAMMENFASSUNG" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "Gefundene alte Dateien             : $script:FoundFileCount"
    Write-Host "Geloeschte Dateien                 : $script:RemovedFileCount"
    Write-Host "Uebersprungene Eintraege           : $script:SkippedItemCount"
    Write-Host "Verarbeitete Dateien               : $script:ProcessedFileCount"
    Write-Host "Verarbeitete Verzeichnisse         : $script:ProcessedDirectoryCount"
    Write-Host "Verarbeitete PathListFile-Zeilen   : $script:ProcessedPathListLineCount"
    Write-Host "Ausschlussregeln                   : $($script:ExcludeRuleList.Count)"
    Write-Host "Fehler                             : $script:ErrorCount"
    Write-Host "Logdatei                           : $script:ResolvedLogPath"
    Write-Host "Kandidatendatei-Ordner             : $script:CandidateDirectoryPath"
    Write-Host "================================================================================" -ForegroundColor Cyan

    Remove-CleanupCandidateFile

    if ($script:ErrorCount -gt 0) {
        Write-CleanupError "Abgeschlossen mit Fehlern. Details stehen in der Logdatei."
        Close-CleanupLog
        exit 1
    }

    Write-CleanupSuccess "Abgeschlossen ohne protokollierte Fehler."
    Close-CleanupLog
    exit 0
}
catch {
    try { Close-CleanupCandidateWriter } catch {}

    try {
        if ($null -ne $script:LogWriter) {
            Add-CleanupLog "FATAL`t$($_.Exception.Message)"
        }
    }
    catch {}

    try { Remove-CleanupCandidateFile } catch {}

    Close-CleanupLog

    Write-CleanupError $_.Exception.Message
    exit 2
}