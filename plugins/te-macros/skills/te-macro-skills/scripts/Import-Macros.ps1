<#
.SYNOPSIS
  Imports skill-ready macros from a GitHub folder (or a local folder) into Tabular Editor 3's MacroActions.json.

.DESCRIPTION
  Reads every *.csx / *.cs file in the source and imports those with a "// MACRO: Name:" line. The whole file
  becomes the macro code; Name, ValidContexts and Tooltip come from the "// MACRO:" lines.

  An existing macro is updated when it carries the same "// SKILL: name:" as the file or, failing that, the same
  macro name. Its name, contexts, tooltip and code are then replaced by the file's. Other macros are not touched.
  The macros file is backed up before it is written.

  Source can be:
    https://github.com/<owner>/<repo>                      (repository root, default branch)
    https://github.com/<owner>/<repo>/tree/<ref>/<folder>
    https://github.com/<owner>/<repo>/blob/<ref>/<file>
    a local folder or file
  It defaults to the "Skill-Ready Macros" folder of github.com/bernatagulloesbrina/TabularEditor-Scripts.
  GitHub requests use $env:GITHUB_TOKEN / $env:GH_TOKEN, or `gh auth token`, when available (needed for private
  repositories and to avoid the 60 requests/hour anonymous limit).

.EXAMPLE
  .\Import-Macros.ps1 -WhatIf
  .\Import-Macros.ps1 https://github.com/owner/repo/tree/main/macros
  .\Import-Macros.ps1 C:\src\macros -Recurse -Only "Sum of Columns"
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0)]
    [string] $Source = 'https://github.com/bernatagulloesbrina/TabularEditor-Scripts/tree/main/Advanced/Skill-Ready%20Macros',
    [string] $MacrosFile = "$env:LOCALAPPDATA\TabularEditor3\MacroActions.json",
    [string[]] $Include = @('*.csx', '*.cs'),
    # Include subfolders
    [switch] $Recurse,
    # Import only these macro or skill names
    [string[]] $Only,
    # Write even though Tabular Editor 3 is running
    [switch] $Force
)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object Text.UTF8Encoding($false)
$utf8Bom = New-Object Text.UTF8Encoding($true)
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

if (-not (Test-Path $MacrosFile)) { throw "Macros file not found: $MacrosFile" }

# Header lines of one kind ("SKILL" / "MACRO"), without the prefix and its first space
function Get-HeaderLines([string] $code, [string] $kind) {
    foreach ($line in $code -split "\r?\n") {
        if ($line -match "^\s*//\s*${kind}:(.*)$") { $Matches[1] -replace '^ ', '' }
    }
}

function Get-HeaderValue([string] $code, [string] $kind, [string] $key) {
    foreach ($l in @(Get-HeaderLines $code $kind)) {
        if ($l -match "^${key}:\s*(.*?)\s*$") { return $Matches[1] }
    }
}

function Get-Contexts([string] $s) { @($s -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ', ' }

# --- Read the source files ----------------------------------------------------------------------------
function Test-Included([string] $name) { @($Include | Where-Object { $name -like $_ }).Count -gt 0 }

$ghHeaders = @{ 'User-Agent' = 'te-macro-skills'; 'Accept' = 'application/vnd.github+json' }
$token = if ($env:GITHUB_TOKEN) { $env:GITHUB_TOKEN } elseif ($env:GH_TOKEN) { $env:GH_TOKEN }
if (-not $token -and (Get-Command gh -ErrorAction SilentlyContinue)) {
    $token = & { $ErrorActionPreference = 'Continue'; gh auth token 2>$null }
    if ($LASTEXITCODE -ne 0) { $token = $null }
}
if ($token) { $ghHeaders['Authorization'] = "Bearer $($token.Trim())" }

function Get-GitHubContent([string] $api, [string] $path, [string] $ref) {
    $uri = "https://api.github.com/repos/$api/contents/$path"
    if ($ref) { $uri += '?ref=' + [uri]::EscapeDataString($ref) }
    Invoke-RestMethod -Uri $uri -Headers $ghHeaders
}

# $listing: a folder's entries or a single file, as returned by the contents API
function Get-GitHubFiles([string] $api, [string] $ref, $listing) {
    foreach ($i in @($listing)) {
        if ($i.type -eq 'dir') {
            if ($Recurse) { Get-GitHubFiles $api $ref (Get-GitHubContent $api $i.path $ref) }
        }
        elseif ($i.type -eq 'file' -and (Test-Included $i.name)) {
            $f = if ($i.content) { $i } else { Get-GitHubContent $api $i.path $ref }
            $text = $utf8.GetString([Convert]::FromBase64String(($f.content -replace '\s', '')))
            [pscustomobject]@{ File = $i.path; Text = $text }
        }
    }
}

if ($Source -match '^https?://github\.com/([^/]+)/([^/?#]+?)(?:\.git)?(?:/(?:tree|blob)/([^?#]+?))?/?(?:[?#].*)?$') {
    $api = "$($Matches[1])/$($Matches[2])"
    $rest = $Matches[3]
    $item = $null; $ref = $null
    if (-not $rest) { $item = Get-GitHubContent $api '' $null }
    else {
        # The ref can contain slashes (feature/x): try the shortest prefix first
        $segments = $rest -split '/'
        for ($n = 1; $n -le $segments.Count -and -not $item; $n++) {
            $ref = [uri]::UnescapeDataString(($segments[0..($n - 1)] -join '/'))
            $path = ($segments | Select-Object -Skip $n) -join '/'
            try { $item = Get-GitHubContent $api $path $ref }
            catch {
                $code = $_.Exception.Response.StatusCode.value__
                if ($code -ne 404) { throw "GitHub request failed ($code): $($_.Exception.Message)" }
            }
        }
        if (-not $item) { throw "Not found on GitHub (or private without a token): $Source" }
    }
    $files = @(Get-GitHubFiles $api $ref $item)
}
elseif (Test-Path -LiteralPath $Source) {
    $root = (Resolve-Path -LiteralPath $Source).Path
    if (Test-Path -LiteralPath $root -PathType Leaf) { $root = Split-Path $root }
    $files = @(Get-ChildItem -LiteralPath $Source -File -Recurse:$Recurse | Where-Object { Test-Included $_.Name } | ForEach-Object {
        [pscustomobject]@{ File = $_.FullName.Substring($root.Length).TrimStart('\'); Text = [IO.File]::ReadAllText($_.FullName) }
    })
}
else { throw "Source must be a github.com URL or an existing local path: $Source" }

if ($files.Count -eq 0) { Write-Host "No $($Include -join ' / ') files in $Source"; exit 0 }

# --- Parse ---------------------------------------------------------------------------------------------
$data = Get-Content $MacrosFile -Raw -Encoding UTF8 | ConvertFrom-Json
$actions = New-Object Collections.Generic.List[object]
foreach ($a in @($data.Actions)) { $actions.Add($a) }

$imports = foreach ($f in $files) {
    $code = ($f.Text.TrimStart([char]0xFEFF) -replace "\r\n", "`n")
    $i = [pscustomobject]@{
        File = $f.File; Code = $code; Errors = @(); Warnings = @()
        Name = Get-HeaderValue $code 'MACRO' 'Name'
        Contexts = Get-Contexts (Get-HeaderValue $code 'MACRO' 'ValidContexts')
        Tooltip = Get-HeaderValue $code 'MACRO' 'Tooltip'
        Skill = Get-HeaderValue $code 'SKILL' 'name'
    }
    if (-not $i.Name) { $i.Warnings += 'no "// MACRO: Name:" line: not a skill-ready macro' }
    if (-not $i.Contexts) { $i.Contexts = 'Model'; $i.Warnings += 'no "// MACRO: ValidContexts:" line, using Model' }
    if ($null -eq $i.Tooltip) { $i.Tooltip = '' }
    if (-not $i.Skill) { $i.Warnings += 'no "// SKILL: name:" line: imported as a macro only, no skill will be generated' }
    $i
}
if ($Only) { $imports = @($imports | Where-Object { $Only -contains $_.Name -or $Only -contains $_.Skill }) }

foreach ($dup in $imports | Where-Object Name | Group-Object Name | Where-Object Count -gt 1) {
    foreach ($i in $dup.Group) { $i.Errors += "macro name '$($dup.Name)' is in several files" }
}
foreach ($dup in $imports | Where-Object Skill | Group-Object Skill | Where-Object Count -gt 1) {
    foreach ($i in $dup.Group) { $i.Errors += "skill name '$($dup.Name)' is in several files" }
}

# --- Merge ---------------------------------------------------------------------------------------------
$tePath = (Resolve-Path $MacrosFile).Path
$teRunning = $tePath -like "$env:LOCALAPPDATA\TabularEditor3\*" -and (Get-Process TabularEditor3 -ErrorAction SilentlyContinue)
$report = New-Object Collections.Generic.List[object]
$dirty = $false

foreach ($i in $imports) {
    $status = 'skipped'; $changes = @()
    if (-not $i.Name) {
        $report.Add([pscustomobject]@{ Macro = $null; Skill = $null; Status = 'ignored'; Changes = ''; File = $i.File; Problems = $i.Warnings[0] })
        continue
    }
    if (-not $i.Errors) {
        $match = @()
        if ($i.Skill) { $match = @($actions | Where-Object { (Get-HeaderValue $_.Execute 'SKILL' 'name') -eq $i.Skill }) }
        if ($match.Count -eq 0) { $match = @($actions | Where-Object { $_.Name -eq $i.Name }) }
        if ($match.Count -gt 1) { $i.Errors += "several macros match: $(($match | ForEach-Object { $_.Name }) -join ', ')" }
    }
    if (-not $i.Errors) {
        $target = $match | Select-Object -First 1
        if ($target) {
            if ($target.Name -ne $i.Name) { $changes += "name (was '$($target.Name)')" }
            if ((Get-Contexts $target.ValidContexts) -ne $i.Contexts) { $changes += "contexts (was '$($target.ValidContexts)')" }
            if ("$($target.Tooltip)" -ne $i.Tooltip) { $changes += 'tooltip' }
            if (($target.Execute -replace "\r\n", "`n") -ne $i.Code) { $changes += 'code' }
            if (-not $changes) { $status = 'unchanged' }
            elseif ($PSCmdlet.ShouldProcess($target.Name, "Update macro ($($changes -join ', '))")) {
                $target.Name = $i.Name; $target.ValidContexts = $i.Contexts; $target.Tooltip = $i.Tooltip; $target.Execute = $i.Code
                $status = 'updated'; $dirty = $true
            }
            else { $status = 'would update' }
        }
        elseif ($PSCmdlet.ShouldProcess($i.Name, 'Create macro')) {
            $id = 1 + ($actions | ForEach-Object { [int]$_.Id } | Measure-Object -Maximum).Maximum
            $actions.Add([pscustomobject][ordered]@{
                Id = [int]$id; Name = $i.Name; Enabled = 'true'; Execute = $i.Code; Tooltip = $i.Tooltip; ValidContexts = $i.Contexts
            })
            $status = 'created'; $dirty = $true
        }
        else { $status = 'would create' }
    }

    $report.Add([pscustomobject]@{
        Macro = $i.Name; Skill = $i.Skill; Status = $status; Changes = $changes -join ', '
        File = $i.File; Problems = (@($i.Errors) + @($i.Warnings)) -join '; '
    })
}

if ($dirty) {
    if ($teRunning -and -not $Force) {
        throw "Tabular Editor 3 is running. Close it and run again (or pass -Force): it only loads macros at startup, and saving a macro there would overwrite the imported ones."
    }
    $backup = "$tePath.$(Get-Date -Format yyyyMMdd-HHmmss).bak"
    Copy-Item $tePath $backup
    $data.Actions = $actions.ToArray()
    [IO.File]::WriteAllText($tePath, ($data | ConvertTo-Json -Depth 5), $utf8Bom)
    Write-Host "Saved $tePath (backup: $backup)"
}
elseif ($teRunning -and @($report | Where-Object Status -like 'would*').Count -gt 0) {
    Write-Warning 'Tabular Editor 3 is running: close it before importing.'
}

$report
exit [int](@($report | Where-Object Status -eq 'skipped').Count -gt 0)
