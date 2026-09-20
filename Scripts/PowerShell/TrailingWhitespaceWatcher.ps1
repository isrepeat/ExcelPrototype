[CmdletBinding()]
param(
    [string]$ConfigurationPath
)

<#
TrailingWhitespaceWatcher.json is discovered by walking from this script's
directory to the drive root. The first file found is used.

The preferred configuration format is:
{
    "rules": [
        {
            "path": "relative-or-absolute-directory",
            "maxDepth": -1,
            "extensions": [".cpp", ".h"],
            "fileNames": ["CMakeLists.txt"]
        }
    ]
}

`path` is resolved relative to the configuration file when it is not absolute.
`maxDepth` is 0 for the directory itself, 1 for direct child directories, and
-1 for every descendant. A rule needs at least one extension or exact file
name. The legacy `roots` array is also accepted and uses the former default
extensions and CMakeLists.txt file name.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Find-ConfigurationPath {
    param([string]$StartDirectory)

    $directory = [System.IO.DirectoryInfo]::new($StartDirectory)
    while ($null -ne $directory) {
        $candidate = Join-Path $directory.FullName 'TrailingWhitespaceWatcher.json'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
        $directory = $directory.Parent
    }
    return $null
}

function New-WatcherRule {
    param(
        [string]$ConfigurationDirectory,
        [object]$ConfiguredRule
    )

    if ($null -eq $ConfiguredRule.path -or [string]::IsNullOrWhiteSpace([string]$ConfiguredRule.path)) {
        throw 'Each watcher rule must contain a non-empty path.'
    }
    $candidate = [string]$ConfiguredRule.path
    if (-not [System.IO.Path]::IsPathRooted($candidate)) {
        $candidate = Join-Path $ConfigurationDirectory $candidate
    }
    if (-not (Test-Path -LiteralPath $candidate -PathType Container)) {
        throw "Configured root does not exist: $candidate"
    }
    $root = (Resolve-Path -LiteralPath $candidate).Path.TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    if ($root -eq [System.IO.Path]::GetPathRoot($root).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)) {
        throw "Configured root must not be a drive root: $root"
    }

    $maxDepth = -1
    if ($null -ne $ConfiguredRule.maxDepth) {
        if ($ConfiguredRule.maxDepth -isnot [int] -and $ConfiguredRule.maxDepth -isnot [long]) {
            throw "Rule '$root' has a non-integer maxDepth."
        }
        $maxDepth = [int]$ConfiguredRule.maxDepth
        if ($maxDepth -lt -1) {
            throw "Rule '$root' has maxDepth below -1."
        }
    }

    $extensions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($extension in $ConfiguredRule.extensions) {
        if ([string]::IsNullOrWhiteSpace([string]$extension) -or -not ([string]$extension).StartsWith('.')) {
            throw "Rule '$root' contains an invalid extension: $extension"
        }
        [void]$extensions.Add([string]$extension)
    }
    $fileNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($fileName in $ConfiguredRule.fileNames) {
        if ([string]::IsNullOrWhiteSpace([string]$fileName) -or [System.IO.Path]::GetFileName([string]$fileName) -ne [string]$fileName) {
            throw "Rule '$root' contains an invalid file name: $fileName"
        }
        [void]$fileNames.Add([string]$fileName)
    }
    if ($extensions.Count -eq 0 -and $fileNames.Count -eq 0) {
        throw "Rule '$root' must contain extensions or fileNames."
    }
    return [pscustomobject]@{
        Root = $root
        MaxDepth = $maxDepth
        Extensions = $extensions
        FileNames = $fileNames
    }
}

function Get-WatcherRules {
    param([string]$Path)

    $configuration = Get-Content -LiteralPath $Path -Raw -Encoding utf8 | ConvertFrom-Json
    $configurationDirectory = [System.IO.Path]::GetDirectoryName($Path)
    $configuredRules = $configuration.rules
    if ($null -eq $configuredRules) {
        $configuredRules = @($configuration.roots | ForEach-Object {
            [pscustomobject]@{
                path = $_
                maxDepth = -1
                extensions = @('.bat', '.cmd', '.cmake', '.cpp', '.cs', '.h', '.hpp', '.json', '.kt', '.md', '.ps1', '.properties', '.sln', '.txt', '.vcxproj', '.xaml', '.xml', '.yml', '.yaml')
                fileNames = @('CMakeLists.txt')
            }
        })
    }
    if ($null -eq $configuredRules -or $configuredRules.Count -eq 0) {
        throw "Configuration '$Path' must contain a non-empty 'rules' array."
    }
    return @($configuredRules | ForEach-Object { New-WatcherRule $configurationDirectory $_ })
}

if ([string]::IsNullOrWhiteSpace($ConfigurationPath)) {
    $ConfigurationPath = Find-ConfigurationPath $PSScriptRoot
} elseif (Test-Path -LiteralPath $ConfigurationPath -PathType Leaf) {
    $ConfigurationPath = (Resolve-Path -LiteralPath $ConfigurationPath).Path
} else {
    throw "Configuration file does not exist: $ConfigurationPath"
}
if ($null -eq $ConfigurationPath) {
    throw "TrailingWhitespaceWatcher.json was not found while searching from '$PSScriptRoot' to the drive root."
}
$watchRules = Get-WatcherRules $ConfigurationPath

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;
using System.Collections.Concurrent;
using System.IO;

public sealed class WatcherEventRecord {
    public string Path;
    public string Kind;
    public DateTime ReceivedAtUtc;
}

public sealed class DirectFileWatcher : IDisposable {
    private readonly ConcurrentQueue<WatcherEventRecord> events = new ConcurrentQueue<WatcherEventRecord>();
    private readonly FileSystemWatcher watcher;

    public DirectFileWatcher(string root) {
        watcher = new FileSystemWatcher(root) {
            IncludeSubdirectories = true,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size,
            InternalBufferSize = 65536,
            EnableRaisingEvents = true
        };
        watcher.Changed += OnChanged;
        watcher.Created += OnChanged;
        watcher.Renamed += OnRenamed;
    }

    public bool TryDequeue(out WatcherEventRecord value) {
        return events.TryDequeue(out value);
    }

    public void Dispose() {
        watcher.Dispose();
    }

    private void OnChanged(object sender, FileSystemEventArgs args) {
        events.Enqueue(new WatcherEventRecord { Path = args.FullPath, Kind = args.ChangeType.ToString(), ReceivedAtUtc = DateTime.UtcNow });
    }

    private void OnRenamed(object sender, RenamedEventArgs args) {
        events.Enqueue(new WatcherEventRecord { Path = args.FullPath, Kind = args.ChangeType.ToString(), ReceivedAtUtc = DateTime.UtcNow });
    }
}
'@

$ignoredDirectories = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
@('.git', 'Build', 'bin', 'Logs', 'obj') | ForEach-Object { [void]$ignoredDirectories.Add($_) }

$logDirectory = 'C:\Logs'
New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
$logPath = Join-Path $logDirectory 'trailing-whitespace-watcher.log'

function Write-WatcherLog {
    param([string]$Level, [string]$Message)

    $timestamp = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss.fff')
    Add-Content -LiteralPath $logPath -Value "[$timestamp] [$Level] $Message" -Encoding utf8
}

function Test-WatchedTextFile {
    param([string]$Path)

    foreach ($segment in $Path.Split([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)) {
        if ($ignoredDirectories.Contains($segment)) {
            return $false
        }
    }
    foreach ($rule in $watchRules) {
        $rootWithSeparator = $rule.Root + [System.IO.Path]::DirectorySeparatorChar
        if (-not $Path.StartsWith($rootWithSeparator, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        $relativePath = $Path.Substring($rootWithSeparator.Length)
        $depth = $relativePath.Split([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar).Count - 1
        if ($rule.MaxDepth -ne -1 -and $depth -gt $rule.MaxDepth) {
            continue
        }
        if ($rule.FileNames.Contains([System.IO.Path]::GetFileName($Path)) -or $rule.Extensions.Contains([System.IO.Path]::GetExtension($Path))) {
            return $true
        }
    }
    return $false
}

function Remove-TerminalWhitespace {
    param([string]$Path, [string]$EventName, [DateTime]$ReceivedAtUtc)

    if (-not (Test-WatchedTextFile $Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        if ($bytes.Length -eq 0) {
            Write-WatcherLog 'INFO' "$EventName ignored empty file: $Path"
            return
        }
        $length = $bytes.Length
        while ($length -gt 0) {
            $value = [int]$bytes[$length - 1]
            if ($value -ne 9 -and $value -ne 10 -and $value -ne 13 -and $value -ne 32) {
                break
            }
            --$length
        }
        if ($length -eq $bytes.Length) {
            return
        }
        if ($length -eq 0) {
            Write-WatcherLog 'WARN' "$EventName did not modify whitespace-only file: $Path"
            return
        }
        $removed = @($bytes[$length..($bytes.Length - 1)] | ForEach-Object { [int]$_ }) -join ', '
        $prefix = [System.Text.Encoding]::UTF8.GetString($bytes, 0, $length)
        $line = 1 + ([regex]::Matches($prefix, "`n")).Count
        [System.IO.File]::WriteAllBytes($Path, $bytes[0..($length - 1)])
        $delay = [Math]::Round(([DateTime]::UtcNow - $ReceivedAtUtc).TotalMilliseconds)
        Write-WatcherLog 'FIXED' "$EventName changed $Path; line $line; removed terminal byte(s): $removed; event-to-fix delay: ${delay}ms"
    } catch {
        Write-WatcherLog 'ERROR' "${EventName} could not process ${Path}: $($_.Exception.Message)"
    }
}

$watchers = @($watchRules | ForEach-Object { [DirectFileWatcher]::new($_.Root) })
$pending = @{}

$notifyIcon = [System.Windows.Forms.NotifyIcon]::new()
$notifyIcon.Icon = [System.Drawing.SystemIcons]::Shield
$notifyIcon.Text = 'Trailing whitespace watcher'
$notifyIcon.Visible = $true
$menu = [System.Windows.Forms.ContextMenuStrip]::new()
$checkItem = $menu.Items.Add('Check all files')
$openLogItem = $menu.Items.Add('Open log')
$exitItem = $menu.Items.Add('Exit')
$notifyIcon.ContextMenuStrip = $menu
$running = $true
$checkItem.add_Click({ $watchRules | ForEach-Object { Get-ChildItem -LiteralPath $_.Root -Recurse -File | ForEach-Object { Remove-TerminalWhitespace $_.FullName 'ManualCheck' [DateTime]::UtcNow } } })
$openLogItem.add_Click({ Start-Process notepad.exe -ArgumentList $logPath })
$exitItem.add_Click({ $script:running = $false })

Write-WatcherLog 'INFO' "Watcher started from configuration $ConfigurationPath; rules: $($watchRules.Root -join '; '); log: $logPath"
try {
    while ($running) {
        foreach ($watcher in $watchers) {
            $event = $null
            while ($watcher.TryDequeue([ref]$event)) {
                if (-not (Test-WatchedTextFile $event.Path)) {
                    continue
                }
                $pending[$event.Path] = $event
                Write-WatcherLog 'EVENT' "$($event.Kind) received: $($event.Path)"
            }
        }
        foreach ($path in @($pending.Keys)) {
            $event = $pending[$path]
            if (([DateTime]::UtcNow - $event.ReceivedAtUtc).TotalMilliseconds -ge 250) {
                Remove-TerminalWhitespace $path $event.Kind $event.ReceivedAtUtc
                $pending.Remove($path)
            }
        }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 25
    }
} finally {
    $watchers | ForEach-Object { $_.Dispose() }
    $notifyIcon.Dispose()
    Write-WatcherLog 'INFO' 'Watcher stopped.'
}