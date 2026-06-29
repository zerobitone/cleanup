<#
.SYNOPSIS
    Sicheres, robustes und speicherschonendes Skript zum Finden und Entfernen leerer Verzeichnisse.

.DESCRIPTION
    Dieses Skript sucht rekursiv ab einem oder mehreren Startverzeichnissen nach leeren
    Verzeichnissen und entfernt diese optional.

    Dateien werden niemals geloescht.

    Die Verarbeitung erfolgt bottom-up:
    Erst werden Unterverzeichnisse verarbeitet, danach das jeweilige Elternverzeichnis.
    Dadurch koennen auch Verzeichnisse geloescht werden, die erst leer werden, nachdem
    darunterliegende leere Verzeichnisse entfernt wurden.

    Sicherheitsmerkmale:
    - Unterstuetzt -DryRun.
    - Unterstuetzt -WhatIf und -Confirm ueber SupportsShouldProcess.
    - Loescht ausschliesslich leere Verzeichnisse.
    - Loescht niemals Dateien.
    - Loescht niemals die angegebenen Startverzeichnisse.
    - Traversiert keine Reparse Points / Symlinks / Junctions.
    - Nutzt Remove-Item -LiteralPath.
    - Unterstuetzt Ausschlussmuster direkt per -ExcludePattern.
    - Unterstuetzt grosse Ausschlusslisten per -ExcludePatternFile.
    - Ausschlussmuster nutzen PowerShell-Wildcards, keine Regex.
    - Optionale Praefixe: DIR:, PATH:
    - Arbeitet iterativ, nicht rekursiv, um StackOverflow zu vermeiden.
    - Nutzt leichte Stack-Frames als Strings statt PSCustomObject.
    - Nutzt gepuffertes Logging.
    - Verhindert bewusst nicht-interaktive Ausfuehrung.

.PARAMETER Path
    Ein oder mehrere Startverzeichnisse.

.PARAMETER DryRun
    Wenn gesetzt, werden keine Verzeichnisse geloescht.
    Leere Verzeichnisse werden nur gefunden und protokolliert.

.PARAMETER LogPath
    Pfad zur Logdatei oder zu einem bestehenden Log-Verzeichnis.

.PARAMETER ExcludePattern
    Optionale Ausschlussmuster als PowerShell-Wildcards.

    Beispiele:
    "*\DoNotDelete"
    "*\Archiv"
    "DIR:N:\IT-BL\ProjektA"
    "DIR:N:\IT-BL\ProjektA\*"
    "PATH:\\ibbads\data\IT-BL\Test"

.PARAMETER ExcludePatternFile
    Textdatei mit Ausschlussmustern, eine Zeile pro Muster.
    Leere Zeilen und Kommentarzeilen mit # werden ignoriert.

    Beispiele:
    DIR:N:\IT-BL\ProjektA
    DIR:N:\IT-BL\ProjektA\*
    PATH:\\ibbads\data\IT-BL\Archiv
    *\DoNotDelete

.EXAMPLE
    .\rm_empty_dirs.ps1 -Path "N:\IT-BL" -DryRun

.EXAMPLE
    .\rm_empty_dirs.ps1 -Path "N:\IT-BL" -ExcludePattern "*\DoNotDelete" -DryRun

.EXAMPLE
    .\rm_empty_dirs.ps1 -Path "N:\IT-BL" -ExcludePatternFile "C:\Cleanup\exclude_dirs.txt" -DryRun

.EXAMPLE
    .\rm_empty_dirs.ps1 -Path "N:\IT-BL" -LogPath "C:\CleanupLogs" -DryRun

.EXAMPLE
    .\rm_empty_dirs.ps1 -Path "\\ibbads\data\IT-BL" -WhatIf

.NOTES
    Autor:
    Stefan Homberg

    Dateiname:
    cleanup_empty_dirs.ps1

    Version:
    1.1.0

    Kompatibilitaet:
    - Windows PowerShell 5.1
    - PowerShell 7+

    Wichtige Betriebsregel:
    Dieses Skript ist bewusst fuer interaktive Ausfuehrung gebaut.
    Es soll nie automatisiert, unbeaufsichtigt, als geplanter Task oder als Hintergrundjob laufen.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Path,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = $(Join-Path -Path (Get-Location) -ChildPath ("empty_dirs_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))),

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludePattern,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ExcludePatternFile
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# --------------------------------------------------------------------------------------
# Skriptweite Variablen
# --------------------------------------------------------------------------------------

$script:IsWindowsLike              = $true
$script:ResolvedLogPath            = $null
$script:LogWriter                  = $null
$script:LogLineCount               = 0
$script:FoundDirectoryCount        = 0
$script:RemovedDirectoryCount      = 0
$script:SkippedDirectoryCount      = 0
$script:ProcessedDirectoryCount    = 0
$script:ErrorCount                 = 0
$script:ExcludeRuleList            = New-Object System.Collections.Generic.List[object]

# --------------------------------------------------------------------------------------
# Konsolenfunktionen
# --------------------------------------------------------------------------------------

function Write-CleanupInfo {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "[INFO ] $Message" -ForegroundColor Cyan
}

function Write-CleanupWarning {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "[WARN ] $Message" -ForegroundColor Yellow
}

function Write-CleanupSuccess {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "[ OK  ] $Message" -ForegroundColor Green
}

function Write-CleanupError {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "[FEHLER] $Message" -ForegroundColor Red
}

# --------------------------------------------------------------------------------------
# Plattform und Laufkontext
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
    # Das Skript ist absichtlich nicht fuer unbeaufsichtigte Automatisierung gedacht.
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

# --------------------------------------------------------------------------------------
# Logging
# --------------------------------------------------------------------------------------

function Resolve-CleanupLogPath {
    param([Parameter(Mandatory = $true)][string]$TargetLogPath)

    if (Test-Path -LiteralPath $TargetLogPath -PathType Container) {
        return Join-Path -Path $TargetLogPath -ChildPath ("empty_dirs_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
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
        # Schreibtest vor dem Start. Ohne Log waere der Lauf nicht sauber nachvollziehbar.
        $testFile = Join-Path -Path $logDirectory -ChildPath (".__empty_dir_cleanup_write_test_{0}.tmp" -f ([Guid]::NewGuid().ToString("N")))
        [System.IO.File]::WriteAllText($testFile, "write-test", [System.Text.Encoding]::UTF8)
        Remove-Item -LiteralPath $testFile -Force -ErrorAction Stop
    }
    catch {
        throw "Keine Schreibrechte fuer das Log-Verzeichnis '$logDirectory'. Fehler: $($_.Exception.Message)"
    }

    $script:ResolvedLogPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($effectiveLogPath)

    # Gepufferter Writer: deutlich schneller als pro Logzeile AppendAllText.
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

    # Wichtig:
    # Formatierung zuerst in eine Variable schreiben.
    # Nicht direkt in WriteLine() mit -f arbeiten, weil PowerShell sonst
    # die Argumente in Methodenaufrufen falsch binden kann.
    $line = "[{0}] {1}" -f $timestamp, $Message

    $script:LogWriter.WriteLine($line)
    $script:LogLineCount++

    if (($script:LogLineCount % 1000) -eq 0) {
        $script:LogWriter.Flush()
    }
}

function Write-CleanupLogHeader {
    Add-CleanupLog "================================================================================"
    Add-CleanupLog "Empty-Directory-Cleanup-Log"
    Add-CleanupLog "================================================================================"
    Add-CleanupLog "Skript                 : cleanup_empty_dirs.ps1"
    Add-CleanupLog "Version                : 1.1.0"
    Add-CleanupLog "Autor                  : Stefan Homberg"
    Add-CleanupLog "Startzeit              : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Add-CleanupLog "Benutzer               : $(Get-CurrentUserName)"
    Add-CleanupLog "Computer               : $(Get-SafeComputerName)"
    Add-CleanupLog "PowerShell-Version     : $($PSVersionTable.PSVersion)"
    Add-CleanupLog "PowerShell-Edition     : $($PSVersionTable.PSEdition)"
    Add-CleanupLog "DryRun                 : $DryRun"
    Add-CleanupLog "WhatIfPreference       : $WhatIfPreference"
    Add-CleanupLog "LogPath                : $script:ResolvedLogPath"
    Add-CleanupLog "Dateien loeschen       : Nein"
    Add-CleanupLog "Verzeichnisse          : Nur leere Verzeichnisse"
    Add-CleanupLog "Startverzeichnis       : Wird nicht geloescht"
    Add-CleanupLog "Reparse Points         : Werden nicht traversiert"
    Add-CleanupLog "ExcludePatternFile     : $ExcludePatternFile"
    Add-CleanupLog "Ausschlussregeln       : $($script:ExcludeRuleList.Count)"
    Add-CleanupLog "RAM-Strategie          : Iterativer Stack mit leichten String-Frames"
    Add-CleanupLog "================================================================================"
    Add-CleanupLog ""
}

# --------------------------------------------------------------------------------------
# Pfadfunktionen
# --------------------------------------------------------------------------------------

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
        throw "Ein uebergebener Startpfad ist leer."
    }

    if (-not (Test-Path -LiteralPath $InputPath -PathType Container)) {
        throw "Startpfad existiert nicht oder ist kein Verzeichnis: $InputPath"
    }

    $resolvedItem = Resolve-Path -LiteralPath $InputPath -ErrorAction Stop | Select-Object -First 1

    if ($null -eq $resolvedItem -or [string]::IsNullOrWhiteSpace($resolvedItem.ProviderPath)) {
        throw "Startpfad konnte nicht eindeutig aufgeloest werden: $InputPath"
    }

    return $resolvedItem.ProviderPath
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

# --------------------------------------------------------------------------------------
# Ausschlussregeln
# --------------------------------------------------------------------------------------

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

    $scope = "DIR"
    $value = $pattern

    if ($pattern -match '^(DIR|PATH):(.*)$') {
        $scope = $matches[1]
        $value = $matches[2].Trim()
    }

    if ([string]::IsNullOrWhiteSpace($value)) {
        return $null
    }

    return [PSCustomObject]@{
        Scope       = $scope
        Pattern     = $value
        HasWildcard = ($value.IndexOfAny([char[]]@('*', '?', '[')) -ge 0)
        Source      = $Source
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

function Test-CleanupExclusionByPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath
    )

    if ($script:ExcludeRuleList.Count -eq 0) {
        return $false
    }

    $displayPath = ConvertFrom-CleanupLongPath -InputPath $DirectoryPath
    $name = [System.IO.Path]::GetFileName($displayPath.TrimEnd('\'))

    foreach ($rule in $script:ExcludeRuleList) {
        $pattern = [string]$rule.Pattern

        switch ($rule.Scope) {
            "PATH" {
                if ($rule.HasWildcard) {
                    if ($displayPath -like $pattern) {
                        return $true
                    }
                }
                else {
                    if ($displayPath -eq $pattern) {
                        return $true
                    }
                }
            }

            default {
                if ($rule.HasWildcard) {
                    if ($displayPath -like $pattern -or $name -like $pattern) {
                        return $true
                    }
                }
                else {
                    if ($displayPath -eq $pattern -or $name -eq $pattern) {
                        return $true
                    }
                }
            }
        }
    }

    return $false
}

# --------------------------------------------------------------------------------------
# Dateisystemoperationen
# --------------------------------------------------------------------------------------

function Test-CleanupReparsePointByAttributes {
    param([Parameter(Mandatory = $true)][System.IO.FileAttributes]$Attributes)

    return (($Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
}

function Test-CleanupDirectoryEmpty {
    param([Parameter(Mandatory = $true)][string]$DirectoryPath)

    try {
        # Streaming-Leerpruefung:
        # Es wird nur getestet, ob mindestens ein Eintrag existiert.
        $enumerator = [System.IO.Directory]::EnumerateFileSystemEntries($DirectoryPath).GetEnumerator()

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
        Add-CleanupLog "ERROR`tLeerpruefung fehlgeschlagen`tPfad=$DirectoryPath`tFehler=$($_.Exception.Message)"
        return $false
    }
}

function Get-CleanupDirectoryEntry {
    param([Parameter(Mandatory = $true)][string]$DirectoryPath)

    try {
        # EnumerateDirectories streamt die Unterordner und baut keine komplette Baumstruktur im RAM auf.
        return [System.IO.Directory]::EnumerateDirectories($DirectoryPath)
    }
    catch {
        $script:ErrorCount++
        Add-CleanupLog "ERROR`tVerzeichnis kann nicht gelesen werden`tPfad=$DirectoryPath`tFehler=$($_.Exception.Message)"
        return @()
    }
}

# --------------------------------------------------------------------------------------
# Bestaetigung und Retry
# --------------------------------------------------------------------------------------

function Confirm-CleanupDeletion {
    param([Parameter(Mandatory = $true)][int]$FoundDirectoryCount)

    if ($DryRun -or $WhatIfPreference) {
        return
    }

    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host "WARNUNG: ECHTER LOESCHLAUF FUER LEERE VERZEICHNISSE" -ForegroundColor Yellow
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host "Es werden ausschliesslich leere Verzeichnisse geloescht." -ForegroundColor Yellow
    Write-Host "Dateien werden nicht geloescht." -ForegroundColor Yellow
    Write-Host "Startverzeichnisse werden nicht geloescht." -ForegroundColor Yellow
    Write-Host "Gefundene leere Verzeichnisse: $FoundDirectoryCount" -ForegroundColor Yellow
    Write-Host "Logdatei: $script:ResolvedLogPath" -ForegroundColor Yellow
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
# Stack-Frames
# --------------------------------------------------------------------------------------

function New-CleanupDirectoryFrame {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,

        [Parameter(Mandatory = $true)]
        [bool]$Visited,

        [Parameter(Mandatory = $true)]
        [bool]$IsRoot
    )

    # Leichter Frame als String:
    # visited|isRoot|path
    # Dadurch entstehen bei sehr vielen Ordnern weniger PowerShell-Objekte.
    $visitedFlag = if ($Visited) { "1" } else { "0" }
    $rootFlag = if ($IsRoot) { "1" } else { "0" }

    return "{0}|{1}|{2}" -f $visitedFlag, $rootFlag, $DirectoryPath
}

function Split-CleanupDirectoryFrame {
    param([Parameter(Mandatory = $true)][string]$Frame)

    $firstSeparator = $Frame.IndexOf('|')
    $secondSeparator = $Frame.IndexOf('|', $firstSeparator + 1)

    if ($firstSeparator -lt 0 -or $secondSeparator -lt 0) {
        throw "Ungueltiger Stack-Frame: $Frame"
    }

    return [PSCustomObject]@{
        Visited = ($Frame.Substring(0, $firstSeparator) -eq "1")
        IsRoot  = ($Frame.Substring($firstSeparator + 1, $secondSeparator - $firstSeparator - 1) -eq "1")
        Path    = $Frame.Substring($secondSeparator + 1)
    }
}

# --------------------------------------------------------------------------------------
# Hauptlogik
# --------------------------------------------------------------------------------------

function Find-CleanupEmptyDirectory {
    param([Parameter(Mandatory = $true)][string]$StartPath)

    $resolvedStartPath = Resolve-CleanupPath -InputPath $StartPath
    $longStartPath = ConvertTo-CleanupLongPath -InputPath $resolvedStartPath
    $displayStartPath = ConvertFrom-CleanupLongPath -InputPath $longStartPath

    Add-CleanupLog "FIND_START`tPfad=$displayStartPath"
    Write-CleanupInfo "Suche leere Verzeichnisse in: $displayStartPath"

    try {
        $startAttributes = [System.IO.File]::GetAttributes($longStartPath)

        if (Test-CleanupReparsePointByAttributes -Attributes $startAttributes) {
            Add-CleanupLog "SKIP`tStartverzeichnis ist ReparsePoint/SymbolicLink`tPfad=$displayStartPath"
            $script:SkippedDirectoryCount++
            return
        }

        if (Test-CleanupExclusionByPath -DirectoryPath $longStartPath) {
            Add-CleanupLog "SKIP`tStartverzeichnis ist ausgeschlossen`tPfad=$displayStartPath"
            $script:SkippedDirectoryCount++
            return
        }

        $stack = New-Object System.Collections.Stack
        $stack.Push((New-CleanupDirectoryFrame -DirectoryPath $longStartPath -Visited $false -IsRoot $true))

        while ($stack.Count -gt 0) {
            $frame = Split-CleanupDirectoryFrame -Frame ([string]$stack.Pop())
            $directoryPath = [string]$frame.Path
            $visited = [bool]$frame.Visited
            $isRoot = [bool]$frame.IsRoot

            if (-not $visited) {
                # Post-Order-Verarbeitung:
                # Erst Kinder, dann Eltern. Nur so werden nachtraeglich leer gewordene Eltern erkannt.
                $stack.Push((New-CleanupDirectoryFrame -DirectoryPath $directoryPath -Visited $true -IsRoot $isRoot))

                foreach ($childPath in (Get-CleanupDirectoryEntry -DirectoryPath $directoryPath)) {
                    try {
                        $childAttributes = [System.IO.File]::GetAttributes($childPath)
                        $childDisplayPath = ConvertFrom-CleanupLongPath -InputPath $childPath

                        if (Test-CleanupReparsePointByAttributes -Attributes $childAttributes) {
                            Add-CleanupLog "SKIP`tReparsePoint/SymbolicLink-Verzeichnis nicht traversiert`tPfad=$childDisplayPath"
                            $script:SkippedDirectoryCount++
                            continue
                        }

                        if (Test-CleanupExclusionByPath -DirectoryPath $childPath) {
                            Add-CleanupLog "SKIP_SUBTREE`tVerzeichnis ausgeschlossen; Unterstruktur wird nicht traversiert`tPfad=$childDisplayPath"
                            $script:SkippedDirectoryCount++
                            continue
                        }

                        $stack.Push((New-CleanupDirectoryFrame -DirectoryPath $childPath -Visited $false -IsRoot $false))
                    }
                    catch {
                        $script:ErrorCount++
                        Add-CleanupLog "ERROR`tUnterverzeichnis konnte nicht verarbeitet werden`tPfad=$childPath`tFehler=$($_.Exception.Message)"
                    }
                }
            }
            else {
                $script:ProcessedDirectoryCount++

                if (($script:ProcessedDirectoryCount % 250) -eq 0) {
                    Write-Progress `
                        -Activity "Suche leere Verzeichnisse" `
                        -Status "Aktuell: $(ConvertFrom-CleanupLongPath -InputPath $directoryPath)" `
                        -CurrentOperation "Gefunden: $script:FoundDirectoryCount | Verarbeitet: $script:ProcessedDirectoryCount"
                }

                if (($script:ProcessedDirectoryCount % 5000) -eq 0) {
                    Add-CleanupLog "PROGRESS`tVerzeichnisse verarbeitet=$script:ProcessedDirectoryCount`tGefunden=$script:FoundDirectoryCount`tGeloescht=$script:RemovedDirectoryCount"
                }

                if ($isRoot) {
                    Add-CleanupLog "SKIP`tStartverzeichnis wird nicht geloescht`tPfad=$(ConvertFrom-CleanupLongPath -InputPath $directoryPath)"
                    continue
                }

                if (-not (Test-CleanupDirectoryEmpty -DirectoryPath $directoryPath)) {
                    continue
                }

                $displayPath = ConvertFrom-CleanupLongPath -InputPath $directoryPath
                $script:FoundDirectoryCount++

                if ($DryRun) {
                    Add-CleanupLog "DRYRUN`tLeeres Verzeichnis wuerde geloescht werden`tPfad=$displayPath"
                    continue
                }

                # Revalidierung direkt vor dem Loeschen:
                # Zwischen Suche und Loeschung kann sich der Ordnerinhalt geaendert haben.
                if (-not (Test-CleanupDirectoryEmpty -DirectoryPath $directoryPath)) {
                    $script:SkippedDirectoryCount++
                    Add-CleanupLog "SKIP`tVerzeichnis ist vor Loeschung nicht mehr leer`tPfad=$displayPath"
                    continue
                }

                if ($PSCmdlet.ShouldProcess($displayPath, "Leeres Verzeichnis endgueltig loeschen")) {
                    $result = Invoke-CleanupRetry -Description "Leeres Verzeichnis loeschen: $displayPath" -ScriptBlock {
                        Remove-Item -LiteralPath $directoryPath -Force -ErrorAction Stop
                    }

                    if ($result.Success) {
                        $script:RemovedDirectoryCount++
                        Add-CleanupLog "REMOVED_DIRECTORY`tPfad=$displayPath`tAttempts=$($result.Attempt)"
                    }
                    else {
                        $script:ErrorCount++
                        Add-CleanupLog "ERROR`tLeeres Verzeichnis konnte nicht geloescht werden`tPfad=$displayPath`tAttempts=$($result.Attempt)`tFehler=$($result.Error)"
                    }
                }
                else {
                    $script:SkippedDirectoryCount++
                    Add-CleanupLog "SKIP`tShouldProcess/WhatIf hat Verzeichnisloeschung verhindert`tPfad=$displayPath"
                }
            }
        }
    }
    catch {
        $script:ErrorCount++
        Add-CleanupLog "ERROR`tStartpfad konnte nicht verarbeitet werden`tPfad=$displayStartPath`tFehler=$($_.Exception.Message)"
    }
}

# --------------------------------------------------------------------------------------
# Hauptprogramm
# --------------------------------------------------------------------------------------

try {
    Initialize-CleanupPlatform

    if (-not (Test-CleanupInteractiveSession)) {
        throw "Nicht-interaktive Ausfuehrung erkannt. Dieses Skript darf nicht automatisiert oder unbeaufsichtigt ausgefuehrt werden."
    }

    Initialize-CleanupExcludeRule

    Open-CleanupLog -TargetLogPath $LogPath
    Write-CleanupLogHeader

    Add-CleanupLog "INFO`tEingabepfade:"
    foreach ($inputPath in $Path) {
        Add-CleanupLog "INFO`tPATH`t$inputPath"
    }

    if ($ExcludePattern -and $ExcludePattern.Count -gt 0) {
        Add-CleanupLog "INFO`tAusschlussmuster aus -ExcludePattern:"
        foreach ($pattern in $ExcludePattern) {
            Add-CleanupLog "INFO`tEXCLUDE_PATTERN`t$pattern"
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ExcludePatternFile)) {
        Add-CleanupLog "INFO`tExcludePatternFile`t$ExcludePatternFile"
    }

    Add-CleanupLog "INFO`tAusschlussregeln geladen`tCount=$($script:ExcludeRuleList.Count)"

    Write-Host ""
    Write-CleanupInfo "Es werden ausschliesslich leere Verzeichnisse verarbeitet."
    Write-CleanupInfo "Dateien werden nicht geloescht."
    Write-CleanupInfo "Startverzeichnisse werden nicht geloescht."
    Write-CleanupInfo "Ausschlussregeln geladen: $($script:ExcludeRuleList.Count)"
    Write-CleanupInfo "Logdatei: $script:ResolvedLogPath"

    if ($DryRun) {
        Write-CleanupWarning "DryRun ist aktiv. Es werden keine Verzeichnisse geloescht."
        Add-CleanupLog "MODE`tDryRun aktiv. Keine Loeschungen."
    }
    elseif ($WhatIfPreference) {
        Write-CleanupWarning "WhatIf ist aktiv. PowerShell verhindert tatsaechliche Loeschvorgaenge."
        Add-CleanupLog "MODE`tWhatIf aktiv. Keine tatsaechlichen Loeschungen."
    }

    foreach ($inputPath in $Path) {
        Find-CleanupEmptyDirectory -StartPath $inputPath
    }

    Write-Progress -Activity "Suche leere Verzeichnisse" -Completed

    Confirm-CleanupDeletion -FoundDirectoryCount $script:FoundDirectoryCount

    Add-CleanupLog ""
    Add-CleanupLog "================================================================================"
    Add-CleanupLog "ZUSAMMENFASSUNG"
    Add-CleanupLog "================================================================================"
    Add-CleanupLog "Gefundene leere Verzeichnisse       : $script:FoundDirectoryCount"
    Add-CleanupLog "Geloeschte leere Verzeichnisse      : $script:RemovedDirectoryCount"
    Add-CleanupLog "Uebersprungene Verzeichnisse        : $script:SkippedDirectoryCount"
    Add-CleanupLog "Verarbeitete Verzeichnisse          : $script:ProcessedDirectoryCount"
    Add-CleanupLog "Ausschlussregeln                    : $($script:ExcludeRuleList.Count)"
    Add-CleanupLog "Fehler                              : $script:ErrorCount"
    Add-CleanupLog "Endzeit                             : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Add-CleanupLog "================================================================================"

    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "ZUSAMMENFASSUNG" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "Gefundene leere Verzeichnisse      : $script:FoundDirectoryCount"
    Write-Host "Geloeschte leere Verzeichnisse     : $script:RemovedDirectoryCount"
    Write-Host "Uebersprungene Verzeichnisse       : $script:SkippedDirectoryCount"
    Write-Host "Verarbeitete Verzeichnisse         : $script:ProcessedDirectoryCount"
    Write-Host "Ausschlussregeln                   : $($script:ExcludeRuleList.Count)"
    Write-Host "Fehler                             : $script:ErrorCount"
    Write-Host "Logdatei                           : $script:ResolvedLogPath"
    Write-Host "================================================================================" -ForegroundColor Cyan

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
    try {
        if ($null -ne $script:LogWriter) {
            Add-CleanupLog "FATAL`t$($_.Exception.Message)"
        }
    }
    catch {
        # Keine Eskalation beim Fehlerlogging im Fatal-Pfad.
    }

    Close-CleanupLog

    Write-CleanupError $_.Exception.Message
    exit 2
}