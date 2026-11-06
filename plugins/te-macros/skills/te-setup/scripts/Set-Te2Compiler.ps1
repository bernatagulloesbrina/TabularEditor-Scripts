<#
.SYNOPSIS
  Points Tabular Editor 2 at a Roslyn C# compiler, so it compiles macros written in modern C# (string interpolation,
  tuples, ??=, local functions...).

.DESCRIPTION
  Out of the box TE2 compiles scripts and macros with the .NET Framework's built-in compiler, which stops at C# 5.
  TE2 compiles all the macros of MacroActions.json together, so one macro with modern C# makes every macro fail to load.

  Writes ScriptCompilerDirectoryPath and ScriptCompilerOptions in %LOCALAPPDATA%\TabularEditor\Preferences.json (TE2's
  File > Preferences > General > Compiler path / options). The GUI and the TabularEditor.exe command line both read them.
  Compiler used, first found:
    1. Visual Studio / Build Tools: <VS>\MSBuild\Current\Bin\Roslyn (a .NET Framework csc.exe, no extra options)
    2. the latest .NET SDK: <dotnet>\sdk\<version>\Roslyn\bincore (a .NET csc.exe: it needs -lib and the core
       references, or it can't find mscorlib)
  If TE2 is given a folder without a working csc.exe, it silently falls back to the C# 5 compiler, so the .NET SDK path
  breaks when that SDK version is uninstalled (e.g. by an SDK update): run this again then (Test-TeEnvironment.ps1 flags it).

  Tabular Editor 2 must be closed (it writes Preferences.json on exit). Backs up Preferences.json first.

.EXAMPLE
  .\Set-Te2Compiler.ps1 -WhatIf
  .\Set-Te2Compiler.ps1
#>
[CmdletBinding(SupportsShouldProcess)]
param()

$ErrorActionPreference = 'Stop'
$prefs = "$env:LOCALAPPDATA\TabularEditor\Preferences.json"
if (-not (Test-Path $prefs)) { throw "$prefs not found: start Tabular Editor 2 once and close it" }
if (Get-Process -Name 'TabularEditor' -ErrorAction SilentlyContinue) { throw 'Tabular Editor 2 is open: save your work and close it (it rewrites Preferences.json on exit)' }

$frameworkDir = "$env:windir\Microsoft.NET\Framework64\v4.0.30319"
if (-not (Test-Path $frameworkDir)) { $frameworkDir = "$env:windir\Microsoft.NET\Framework\v4.0.30319" }

$dir = $null; $options = $null; $source = $null
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (Test-Path $vswhere) {
    $vs = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath 2>$null | Select-Object -First 1
    if ($vs -and (Test-Path "$vs\MSBuild\Current\Bin\Roslyn\csc.exe")) {
        $dir = "$vs\MSBuild\Current\Bin\Roslyn"; $options = '-langversion:latest'; $source = 'Visual Studio'
    }
}
if (-not $dir) {
    $dotnet = (Get-Command dotnet -ErrorAction SilentlyContinue).Source
    if (-not $dotnet -and (Test-Path "$env:ProgramFiles\dotnet\dotnet.exe")) { $dotnet = "$env:ProgramFiles\dotnet\dotnet.exe" }
    if ($dotnet) {
        # "10.0.402 [C:\Program Files\dotnet\sdk]" -> latest version first
        $sdk = @(& $dotnet --list-sdks 2>$null) | Where-Object { $_ -match '^(\S+) \[(.+)\]$' } |
            ForEach-Object { [pscustomobject]@{ Version = [version](($Matches[1] -split '-')[0]); Path = Join-Path $Matches[2] $Matches[1] } } |
            Sort-Object Version -Descending | Where-Object { Test-Path "$($_.Path)\Roslyn\bincore\csc.exe" } | Select-Object -First 1
        if ($sdk) {
            $dir = "$($sdk.Path)\Roslyn\bincore"; $source = ".NET SDK $($sdk.Version)"
            $options = "-langversion:latest -lib:$frameworkDir -r:mscorlib.dll -r:System.dll -r:System.Core.dll"
        }
    }
}
if (-not $dir) { throw 'No Roslyn compiler found: install the .NET SDK (winget install --id Microsoft.DotNet.SDK.8 -e) and run this again' }

Write-Host "Compiler ($source): $dir"
Write-Host "Options: $options"
if (-not $PSCmdlet.ShouldProcess($prefs, 'set ScriptCompilerDirectoryPath and ScriptCompilerOptions')) { return }

# Edit the two values in place (keeps the rest of the file as TE2 wrote it), JSON-escaping them
$text = [IO.File]::ReadAllText($prefs)
function Set-JsonString([string] $json, [string] $key, [string] $value) {
    $escaped = $value.Replace('\', '\\').Replace('"', '\"')
    $pattern = '("' + $key + '"\s*:\s*)(?:"(?:[^"\\]|\\.)*"|null)'
    if ([regex]::IsMatch($json, $pattern)) {
        return [regex]::Replace($json, $pattern, { param($m) $m.Groups[1].Value + '"' + $escaped + '"' })
    }
    return [regex]::Replace($json, '^\s*\{', { param($m) $m.Value + "`r`n  `"$key`": `"$escaped`"," }, 'None')
}
$text = Set-JsonString $text 'ScriptCompilerDirectoryPath' $dir
$text = Set-JsonString $text 'ScriptCompilerOptions' $options
$null = $text | ConvertFrom-Json   # still valid JSON

$backup = "$prefs.$(Get-Date -Format 'yyyyMMdd-HHmmss').bak"
Copy-Item $prefs $backup
[IO.File]::WriteAllText($prefs, $text, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Updated $prefs (backup: $backup)"
