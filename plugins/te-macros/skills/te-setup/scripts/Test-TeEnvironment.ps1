<#
.SYNOPSIS
  Checks what the Tabular Editor macro environments need on this machine and reports what is missing, with the fix.

.DESCRIPTION
  Read-only: nothing is installed or changed (except the -Build check, which builds the authoring solution).

  Consumption environment (run skill-ready macros as Claude skills, from TabularEditor-Scripts):
    Claude Code CLI, Tabular Editor 3 and/or 2, the te CLI, MacroActions.json, the te-macros and macro-skills plugins.
  Authoring environment (-Authoring; write macros in Visual Studio with TE-Scripting-in-Visual-Studio):
    git, the .NET SDK, Tabular Editor 2 (referenced by the projects), the authoring and publication repo clones,
    TE_PUBLISH_REPO, the CopyMacro macro in Tabular Editor and, with -Build, a build of the solution.

  One row per check: Area, Check, Status (ok / missing / warning / info), Detail, Fix and Who (claude: the agent can do
  it; user: only the user can, e.g. a license, a UAC prompt or a command inside Claude Code).

.EXAMPLE
  .\Test-TeEnvironment.ps1
  .\Test-TeEnvironment.ps1 -Authoring -Build
  .\Test-TeEnvironment.ps1 -Authoring -AuthoringRepo D:\repos\TE-Scripting-in-Visual-Studio
#>
param(
    [switch] $Authoring,
    [string] $AuthoringRepo,
    [string] $PublishRepo,
    # Build the authoring solution ("TE Scripts" project) to prove the toolchain works
    [switch] $Build
)

$ErrorActionPreference = 'Continue'
$results = New-Object System.Collections.Generic.List[object]
function Add-Result($area, $check, $status, $detail, $fix = '', $who = '') {
    $results.Add([pscustomobject]@{ Area = $area; Check = $check; Status = $status; Detail = $detail; Fix = $fix; Who = $who })
}
function Find-Exe([string] $name, [string[]] $paths) {
    $c = Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($c) { return $c.Source }
    $paths | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}
function Get-Version([string] $path) { try { (Get-Item $path).VersionInfo.ProductVersion } catch { '' } }
$hasWinget = [bool](Get-Command winget -ErrorAction SilentlyContinue)
function Get-InstallFix([string] $id, [string] $url) {
    if ($hasWinget) { "winget install --id $id -e" } else { "download and install from $url" }
}

# ---------------------------------------------------------------- Consumption -------------------------------------
$area = 'consumption'
Add-Result $area 'PowerShell' 'ok' ("{0} {1}" -f $PSVersionTable.PSEdition, $PSVersionTable.PSVersion)

$claude = Find-Exe 'claude' @()
if ($claude) { Add-Result $area 'Claude Code CLI' 'ok' $claude }
else { Add-Result $area 'Claude Code CLI' 'missing' 'needed to register the macro-skills plugin' (Get-InstallFix 'Anthropic.ClaudeCode' 'https://claude.com/claude-code') 'claude' }

$te3Exe = Find-Exe 'TabularEditor3.exe' @("$env:ProgramFiles\Tabular Editor 3\TabularEditor3.exe", "${env:ProgramFiles(x86)}\Tabular Editor 3\TabularEditor3.exe", "$env:LOCALAPPDATA\Programs\Tabular Editor 3\TabularEditor3.exe")
$te2Exe = Find-Exe 'TabularEditor.exe' @($(if ($env:TabularEditor2Dir) { [IO.Path]::Combine($env:TabularEditor2Dir, 'TabularEditor.exe') }), "${env:ProgramFiles(x86)}\Tabular Editor\TabularEditor.exe", "$env:ProgramFiles\Tabular Editor\TabularEditor.exe")
$teCli = Find-Exe 'te.exe' @("$env:LOCALAPPDATA\Programs\TabularEditor3\te.exe")

if ($te3Exe) { Add-Result $area 'Tabular Editor 3' 'ok' "$te3Exe ($(Get-Version $te3Exe))" }
# TE3 is optional: used when present, never something to install (TE2 covers everything)
else { Add-Result $area 'Tabular Editor 3' 'info' 'not installed (optional: Tabular Editor 2 is used instead)' }

if ($te2Exe) { Add-Result $area 'Tabular Editor 2' 'ok' "$te2Exe ($(Get-Version $te2Exe))" }
else { Add-Result $area 'Tabular Editor 2' $(if ($Authoring -or -not $te3Exe) { 'missing' } else { 'info' }) $(if ($Authoring -or -not $te3Exe) { 'not installed (free; needed for authoring, and to run the macros when TE3 is not installed)' } else { 'not installed (optional while TE3 is installed; needed for authoring)' })(Get-InstallFix 'TabularEditor.TabularEditor.2' 'https://github.com/TabularEditor/TabularEditor/releases') 'claude' }

if ($teCli) {
    $readme = Join-Path (Split-Path $teCli) 'README.txt'
    $expires = if (Test-Path $readme) { (Select-String -Path $readme -Pattern 'EXPIRES:\s*(\S+)' | Select-Object -First 1).Matches.Groups[1].Value }
    $status = 'ok'; $detail = $teCli
    if ($expires) {
        $detail += " (preview build, expires $expires)"
        try { if ([datetime]::ParseExact($expires, 'yyyy-MM-dd', $null) -lt (Get-Date).AddDays(14)) { $status = 'info'; $detail += ': renew soon, or the skills fall back to TE2' } } catch {}
    }
    Add-Result $area 'te CLI (te.exe)' $status $detail $(if ($status -ne 'ok') { 'optional: download a newer te-win-x64.zip from https://tabulareditor.com/downloads and overwrite te.exe' }) 'user'
}
else {
    # te is optional and licensed separately: mention it, never block on it (the skills run through TabularEditor.exe)
    Add-Result $area 'te CLI (te.exe)' 'info' 'not installed (optional, licensed separately): the macro skills run through Tabular Editor 2, without DAX validation' "optional: if the user has a license, download te-win-x64.zip from https://tabulareditor.com/downloads, extract it to $env:LOCALAPPDATA\Programs\TabularEditor3 and add that folder to the user PATH" 'user'
}

foreach ($e in @(@{ Name = 'TE3'; Folder = "$env:LOCALAPPDATA\TabularEditor3"; Exe = $te3Exe }, @{ Name = 'TE2'; Folder = "$env:LOCALAPPDATA\TabularEditor"; Exe = $te2Exe })) {
    if (-not $e.Exe) { continue }
    $file = Join-Path $e.Folder 'MacroActions.json'
    if (-not (Test-Path $e.Folder)) {
        Add-Result $area "$($e.Name) settings folder" 'missing' "$($e.Folder) does not exist: $($e.Name) has never been started" "start $($e.Name) once (TE3: activate the license) and close it" 'user'
        continue
    }
    if (-not (Test-Path $file)) { Add-Result $area "$($e.Name) macros" 'info' 'no MacroActions.json yet (the importer creates it)'; continue }
    try {
        $actions = @((Get-Content $file -Raw -Encoding UTF8 | ConvertFrom-Json).Actions)
        $skillReady = @($actions | Where-Object { $_.Execute -match '(?m)^\s*//\s*SKILL:\s*name:' }).Count
        Add-Result $area "$($e.Name) macros" 'ok' "$($actions.Count) macros, $skillReady skill-ready ($file)"
    }
    catch { Add-Result $area "$($e.Name) macros" 'warning' "MacroActions.json can't be read: $($_.Exception.Message)" 'restore the latest MacroActions.json.*.bak next to it' 'claude' }
}

$pluginsDir = "$env:USERPROFILE\.claude\plugins"
$installed = @()
if (Test-Path "$pluginsDir\installed_plugins.json") {
    try { $installed = @((Get-Content "$pluginsDir\installed_plugins.json" -Raw | ConvertFrom-Json).plugins.PSObject.Properties.Name) } catch {}
}
if ($installed -match '^te-macros@') { Add-Result $area 'te-macros plugin' 'ok' (($installed -match '^te-macros@') -join ', ') }
else { Add-Result $area 'te-macros plugin' 'missing' 'not installed in Claude Code' 'claude plugin marketplace add bernatagulloesbrina/TabularEditor-Scripts; claude plugin install te-macros@tabular-editor-scripts --scope user' 'claude' }

$macroSkillsDir = "$env:USERPROFILE\.claude\macro-skills\plugins\macro-skills\skills"
$generated = @(Get-ChildItem $macroSkillsDir -Directory -ErrorAction SilentlyContinue | Where-Object { Test-Path (Join-Path $_.FullName 'macro.json') }).Count
if ($installed -contains 'macro-skills@local-macro-skills') { Add-Result $area 'macro-skills plugin' 'ok' "registered, $generated generated skills" }
elseif ($generated) { Add-Result $area 'macro-skills plugin' 'info' "$generated skills generated but the plugin isn't registered yet" 'next step (not setup): New-MacroSkills.ps1 -Register' 'user' }
else { Add-Result $area 'macro-skills plugin' 'info' 'no macro skills generated yet (next step after setup)' 'close Tabular Editor and ask "import the skill-ready macros and generate their skills"' 'user' }

# ready to import and generate: an edition with a settings folder, an engine, and claude for -Register
$ready = (Test-Path "$env:LOCALAPPDATA\TabularEditor3") -or (Test-Path "$env:LOCALAPPDATA\TabularEditor")
$engine = $teCli -or $te2Exe
if ($ready -and $engine -and $claude) { Add-Result $area 'Ready to import macros' 'ok' 'Tabular Editor settings folder, engine (te / TE2) and claude found' }
else {
    $why = @()
    if (-not $ready) { $why += 'start Tabular Editor once' }
    if (-not $engine) { $why += 'install te or Tabular Editor 2' }
    if (-not $claude) { $why += 'install Claude Code CLI' }
    Add-Result $area 'Ready to import macros' 'missing' ($why -join '; ') 'see the rows above' ''
}

$token = $env:GITHUB_TOKEN -or $env:GH_TOKEN
if (-not $token -and (Get-Command gh -ErrorAction SilentlyContinue)) { & gh auth status *> $null; $token = $LASTEXITCODE -eq 0 }
Add-Result $area 'GitHub token' $(if ($token) { 'ok' } else { 'info' }) $(if ($token) { 'available (GITHUB_TOKEN / gh)' } else { 'none: anonymous GitHub limit is 60 requests/hour (about one per macro file)' }) $(if (-not $token) { 'optional: gh auth login, or set a GITHUB_TOKEN user environment variable' }) 'user'

# ---------------------------------------------------------------- Authoring ---------------------------------------
if ($Authoring) {
    $area = 'authoring'

    $git = Find-Exe 'git' @("$env:ProgramFiles\Git\cmd\git.exe")
    if ($git) { Add-Result $area 'git' 'ok' $git } else { Add-Result $area 'git' 'missing' 'needed to clone the repos' (Get-InstallFix 'Git.Git' 'https://git-scm.com/download/win') 'claude' }

    $dotnet = Find-Exe 'dotnet' @("$env:ProgramFiles\dotnet\dotnet.exe")
    $sdks = if ($dotnet) { @(& $dotnet --list-sdks 2>$null) } else { @() }
    $sdkOk = @($sdks | Where-Object { $_ -match '^(\d+)\.' -and [int]$Matches[1] -ge 6 })
    if ($sdkOk) { Add-Result $area '.NET SDK' 'ok' (($sdkOk | ForEach-Object { ($_ -split ' ')[0] }) -join ', ') }
    else { Add-Result $area '.NET SDK' 'missing' $(if ($dotnet) { 'dotnet found but no SDK 6 or later (a runtime alone cannot build)' } else { 'dotnet not found' }) (Get-InstallFix 'Microsoft.DotNet.SDK.8' 'https://dotnet.microsoft.com/download') 'claude' }

    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    # only Visual Studio itself (SSMS is built on the VS shell and shows up with -products *)
    $vs = if (Test-Path $vswhere) { & $vswhere -latest -products Microsoft.VisualStudio.Product.Community Microsoft.VisualStudio.Product.Professional Microsoft.VisualStudio.Product.Enterprise -property displayName 2>$null }
    if ($vs) { $vs = ($vs -replace '[^\x20-\x7E]', ' ') }
    if ($vs) { Add-Result $area 'Visual Studio' 'ok' $vs }
    else { Add-Result $area 'Visual Studio' 'info' 'not found (optional: VS Code with C# Dev Kit or Rider also work, the build only needs the .NET SDK)' (Get-InstallFix 'Microsoft.VisualStudio.2022.Community' 'https://visualstudio.microsoft.com') 'user' }

    # TE2 is referenced by every project (TE.References.props): TabularEditor.exe and TOMWrapper.dll
    $te2Dir = if ($env:TabularEditor2Dir) { $env:TabularEditor2Dir } else { "${env:ProgramFiles(x86)}\Tabular Editor" }
    if (Test-Path ([IO.Path]::Combine($te2Dir, 'TabularEditor.exe'))) {
        $tom = Test-Path ([IO.Path]::Combine($te2Dir, 'TOMWrapper.dll'))
        Add-Result $area 'TE2 build references' $(if ($tom) { 'ok' } else { 'missing' }) "$te2Dir$(if ($env:TabularEditor2Dir) { ' (TabularEditor2Dir)' })$(if (-not $tom) { ': TOMWrapper.dll missing' })" $(if (-not $tom) { 'reinstall Tabular Editor 2' }) 'claude'
    }
    elseif ($te2Exe) {
        Add-Result $area 'TE2 build references' 'missing' "TE2 is in $(Split-Path $te2Exe), not in the default folder the projects look in" "[Environment]::SetEnvironmentVariable('TabularEditor2Dir', '$(Split-Path $te2Exe)', 'User') (then restart Visual Studio / the terminal)" 'claude'
    }
    else { Add-Result $area 'TE2 build references' 'missing' 'Tabular Editor 2 is not installed; the projects reference TabularEditor.exe and TOMWrapper.dll' (Get-InstallFix 'TabularEditor.TabularEditor.2' 'https://github.com/TabularEditor/TabularEditor/releases') 'claude' }

    # repos
    $reposRoot = "$env:USERPROFILE\source\repos"
    if (-not $AuthoringRepo) {
        $AuthoringRepo = @(
            $(if ($PublishRepo) { Join-Path (Split-Path $PublishRepo) 'TE-Scripting-in-Visual-Studio' }),
            (Join-Path (Get-Location) 'TE-Scripting-in-Visual-Studio'),
            (Get-Location).Path,
            (Join-Path (Split-Path (Get-Location)) 'TE-Scripting-in-Visual-Studio'),
            (Join-Path $reposRoot 'TE-Scripting-in-Visual-Studio')
        ) | Where-Object { $_ -and (Test-Path (Join-Path $_ 'TE Scripts\TE Scripts.cs')) } | Select-Object -First 1
    }
    if ($AuthoringRepo -and (Test-Path (Join-Path $AuthoringRepo 'TE Scripts\TE Scripts.cs'))) {
        $AuthoringRepo = (Resolve-Path $AuthoringRepo).Path
        Add-Result $area 'Authoring repo' 'ok' $AuthoringRepo
    }
    else {
        $AuthoringRepo = $null
        Add-Result $area 'Authoring repo' 'missing' 'TE-Scripting-in-Visual-Studio clone not found (pass -AuthoringRepo if it is elsewhere)' "git clone https://github.com/bernatagulloesbrina/TE-Scripting-in-Visual-Studio `"$reposRoot\TE-Scripting-in-Visual-Studio`"" 'claude'
    }

    $publishSource = ''
    if (-not $PublishRepo) {
        if ($env:TE_PUBLISH_REPO) { $PublishRepo = $env:TE_PUBLISH_REPO; $publishSource = 'TE_PUBLISH_REPO' }
        elseif ($AuthoringRepo) {
            $settings = Join-Path $AuthoringRepo '.claude\settings.local.json'
            if (Test-Path $settings) { try { $p = (Get-Content $settings -Raw | ConvertFrom-Json).env.TE_PUBLISH_REPO; if ($p) { $PublishRepo = $p; $publishSource = '.claude\settings.local.json' } } catch {} }
            if (-not $PublishRepo) { $PublishRepo = Join-Path (Split-Path $AuthoringRepo) 'TabularEditor-Scripts'; $publishSource = 'next to the authoring repo' }
        }
        else { $PublishRepo = Join-Path $reposRoot 'TabularEditor-Scripts' }
    }
    if (Test-Path (Join-Path $PublishRepo 'plugins\te-macros')) {
        Add-Result $area 'Publication repo' 'ok' "$PublishRepo ($publishSource)"
        $origin = & git -C $PublishRepo remote get-url origin 2>$null
        if ($origin) { Add-Result $area 'Publication repo remote' 'info' "origin = $origin (publishing needs push rights there, or a fork plus pull requests)" }
    }
    else {
        Add-Result $area 'Publication repo' 'missing' "TabularEditor-Scripts clone not found at $PublishRepo" "git clone https://github.com/bernatagulloesbrina/TabularEditor-Scripts next to the authoring repo (or a fork), or set TE_PUBLISH_REPO in <authoring repo>\.claude\settings.local.json" 'claude'
    }

    # CopyMacro ("Get Macro") in each installed Tabular Editor (TE3 is optional), pointing at this authoring repo
    $install = "& `"$AuthoringRepo\.claude\skills\publish-macro\scripts\Install-CopyMacro.ps1`" (close Tabular Editor first)"
    foreach ($e in @(@{ Name = 'TE3'; Folder = "$env:LOCALAPPDATA\TabularEditor3"; Exe = $te3Exe }, @{ Name = 'TE2'; Folder = "$env:LOCALAPPDATA\TabularEditor"; Exe = $te2Exe })) {
        if (-not $AuthoringRepo -or -not $e.Exe -or -not (Test-Path $e.Folder)) { continue }
        $check = "CopyMacro in $($e.Name)"
        $macrosFile = Join-Path $e.Folder 'MacroActions.json'
        $copyMacro = $null
        if (Test-Path $macrosFile) {
            try { $copyMacro = @((Get-Content $macrosFile -Raw -Encoding UTF8 | ConvertFrom-Json).Actions | Where-Object { $_.Execute -match 'Copy-Macro\.ps1|CopyMacroFromVSFile|macroFilePath' }) | Select-Object -First 1 } catch {}
        }
        if (-not $copyMacro) { Add-Result $area $check 'missing' "no CopyMacro macro in $($e.Name)" $install 'claude'; continue }
        $path = if ($copyMacro.Execute -match '(?:solutionFolder|baseFolderPath)\s*=\s*@"([^"]*)"') { $Matches[1] } else { '' }
        if ($copyMacro.Execute -notmatch 'Copy-Macro\.ps1') { Add-Result $area $check 'warning' "'$($copyMacro.Name)' is the old inline version (folder '$path'): it lacks the TECLI guards skill-ready macros need" $install 'claude' }
        elseif ($path.TrimEnd('\') -ne $AuthoringRepo.TrimEnd('\')) { Add-Result $area $check 'warning' "'$($copyMacro.Name)' points at '$path'" $install 'claude' }
        else { Add-Result $area $check 'ok' "'$($copyMacro.Name)' -> $path" }
    }

    if ($Build) {
        if ($AuthoringRepo -and $sdkOk) {
            $proj = Join-Path $AuthoringRepo 'TE Scripts\TE Scripts.csproj'
            $out = & $dotnet build $proj -nologo -v q 2>&1
            if ($LASTEXITCODE -eq 0) { Add-Result $area 'Build (TE2/TE3)' 'ok' 'TE Scripts.csproj builds' }
            else {
                $err = @($out | Where-Object { $_ -match ': error ' } | ForEach-Object { ($_ -replace '\s*\[[^\]]*\]$', '').Trim() } | Select-Object -Unique -First 3) -join ' | '
                Add-Result $area 'Build (TE2/TE3)' 'missing' $err 'see the error: usually TE2 references or a missing NuGet source' 'claude'
            }
            $out = & $dotnet build $proj -nologo -v q -p:DefineConstants=TECLI "-p:OutDir=$env:TEMP\tecli-build\" 2>&1
            Add-Result $area 'Build (te / TECLI)' $(if ($LASTEXITCODE -eq 0) { 'ok' } else { 'missing' }) $(if ($LASTEXITCODE -eq 0) { 'builds with TECLI defined' } else { (@($out | Where-Object { $_ -match ': error ' } | Select-Object -Unique -First 2) -join ' | ') })
        }
        else { Add-Result $area 'Build' 'info' 'skipped: needs the authoring repo and the .NET SDK' }
    }
}

foreach ($r in $results) { if ($r.Status -eq 'ok') { $r.Fix = ''; $r.Who = '' } }
$results
