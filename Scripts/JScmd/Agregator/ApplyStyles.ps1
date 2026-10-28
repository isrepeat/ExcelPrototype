param(
    [Parameter(Mandatory = $true)][string]$ConfigPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml.Linq
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -Path (Join-Path $PSScriptRoot 'SnapshotStyles.cs') -ReferencedAssemblies System.Core, System.Xml, System.Drawing, System.Xml.Linq, System.Windows.Forms, System.IO.Compression, System.IO.Compression.FileSystem
[SnapshotStyles]::Apply($ConfigPath, $OutputPath)