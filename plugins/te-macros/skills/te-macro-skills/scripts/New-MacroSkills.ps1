<#
.SYNOPSIS
  Generates one Claude skill per Tabular Editor macro that carries a "// SKILL:" header.

.DESCRIPTION
  Reads MacroActions.json, and for every macro whose code has "// SKILL: name:" and "// SKILL: description:"
  lines, writes <SkillsDir>\<name>\SKILL.md (frontmatter + the header text + a standard "How to run" section),
  macro.json (metadata the runner uses) and scripts\Invoke-MacroSkill.ps1 (the runner).

  The macros come from Tabular Editor 3 (%LOCALAPPDATA%\TabularEditor3\MacroActions.json) or, when TE3 has no
  macros file or -Edition TE2 is passed, from Tabular Editor 2 (%LOCALAPPDATA%\TabularEditor\MacroActions.json).
  The skills run the macros with the Tabular Editor CLI (te) when it is installed, otherwise with the Tabular Editor 2
  command line (TabularEditor.exe); -Engine picks one.

  By default the skills go into a local Claude Code plugin, "macro-skills", in $PluginDir (a local marketplace
  with that one plugin), so they can be enabled or disabled together with /plugin. Claude Code loads it in place:
  after regenerating, /reload-plugins picks up the changes. -Register adds the marketplace and installs the plugin
  (once per machine). -SkillsDir writes plain skills to a skills folder instead (e.g. .\.claude\skills).

  Only folders it generated itself (macro.json with generator = te-macro-skills) are ever updated or deleted.
  Each macro is also compiled with the engine (te: `te script --validate`, TECLI defined; te2: a run on an empty
  model) to report whether it can run there.

.EXAMPLE
  .\New-MacroSkills.ps1 -WhatIf
  .\New-MacroSkills.ps1 -Edition TE2 -Engine te2
  .\New-MacroSkills.ps1 -Register
  .\New-MacroSkills.ps1 -Only sum-of-columns -SkillsDir ".\.claude\skills"
  .\New-MacroSkills.ps1 -Prune
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    # Tabular Editor whose macros become skills. Default: TE3 if it has a macros file, otherwise TE2
    [ValidateSet('TE3', 'TE2')]
    [string] $Edition,
    # Read this macros file instead (its edition comes from its folder)
    [string] $MacrosFile,
    # How the skills run the macros: te (Tabular Editor CLI) or te2 (TabularEditor.exe). Default: te if installed
    [ValidateSet('te', 'te2')]
    [string] $Engine,
    [string] $PluginDir = "$env:USERPROFILE\.claude\macro-skills",
    # Write plain skills to this folder instead of the local plugin
    [string] $SkillsDir,
    # Add the local marketplace and install the plugin in Claude Code (user scope)
    [switch] $Register,
    # Generate only these skill names
    [string[]] $Only,
    # Delete generated skills whose macro no longer exists
    [switch] $Prune,
    # Skip the compile check
    [switch] $NoCompileCheck,
    # te.exe or TabularEditor.exe, for the compile check
    [string] $TabularEditor
)

$ErrorActionPreference = 'Stop'
$generator = 'te-macro-skills'
$runner = Join-Path $PSScriptRoot 'Invoke-MacroSkill.ps1'
$utf8 = New-Object Text.UTF8Encoding($false)

$pluginName = 'macro-skills'
$marketplaceName = 'local-macro-skills'

$editionFiles = [ordered]@{
    TE3 = "$env:LOCALAPPDATA\TabularEditor3\MacroActions.json"
    TE2 = "$env:LOCALAPPDATA\TabularEditor\MacroActions.json"
}
$editionNames = @{ TE3 = 'Tabular Editor 3'; TE2 = 'Tabular Editor 2' }
if ($MacrosFile) {
    if (-not (Test-Path $MacrosFile)) { throw "Macros file not found: $MacrosFile" }
    $MacrosFile = (Resolve-Path $MacrosFile).Path
    if (-not $Edition) { $Edition = if ((Split-Path $MacrosFile) -eq (Split-Path $editionFiles.TE2)) { 'TE2' } else { 'TE3' } }
}
else {
    if (-not $Edition) { $Edition = if (Test-Path $editionFiles.TE3) { 'TE3' } elseif (Test-Path $editionFiles.TE2) { 'TE2' } else { 'TE3' } }
    $MacrosFile = $editionFiles[$Edition]
    if (-not (Test-Path $MacrosFile)) { throw "Macros file not found: $MacrosFile. $(if ($Edition -eq 'TE3') { 'Is Tabular Editor 3 installed? -Edition TE2 reads the Tabular Editor 2 macros.' } else { 'Save a macro in Tabular Editor 2 first, or import some with Import-Macros.ps1.' })" }
}
$editionName = $editionNames[$Edition]

# Engine: te if installed (or asked for), otherwise TabularEditor.exe
function Find-TE { $c = Get-Command te.exe -ErrorAction SilentlyContinue; if ($c) { $c.Source } else { @("$env:LOCALAPPDATA\Programs\TabularEditor3\te.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1 } }
function Find-TE2 {
    $c = Get-Command TabularEditor.exe -ErrorAction SilentlyContinue
    if ($c) { $c.Source } else { @("${env:ProgramFiles(x86)}\Tabular Editor\TabularEditor.exe", "$env:ProgramFiles\Tabular Editor\TabularEditor.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1 }
}
if (-not $Engine) { $Engine = if (Find-TE) { 'te' } elseif (Find-TE2) { 'te2' } else { 'te' } }
if (-not $TabularEditor) { $TabularEditor = if ($Engine -eq 'te') { Find-TE } else { Find-TE2 } }
if (-not $TabularEditor) {
    Write-Warning "$(if ($Engine -eq 'te') { 'Tabular Editor CLI (te.exe)' } else { 'Tabular Editor 2 (TabularEditor.exe)' }) not found: the skills will need it to run."
}

$asPlugin = -not $SkillsDir
if ($asPlugin) {
    New-Item -ItemType Directory -Force $PluginDir -WhatIf:$false | Out-Null
    $PluginDir = (Resolve-Path $PluginDir).Path
    $pluginRoot = Join-Path $PluginDir "plugins\$pluginName"
    $SkillsDir = Join-Path $pluginRoot 'skills'
    $manifests = @{
        (Join-Path $PluginDir '.claude-plugin\marketplace.json') = [ordered]@{
            name = $marketplaceName; owner = [ordered]@{ name = $env:USERNAME }
            metadata = [ordered]@{ description = "Skills generated by $generator from this machine's Tabular Editor macros." }
            plugins = @([ordered]@{ name = $pluginName; source = "./plugins/$pluginName"; description = 'Run Tabular Editor macros from the command line.' })
        }
        (Join-Path $pluginRoot '.claude-plugin\plugin.json') = [ordered]@{
            name = $pluginName; description = 'Run Tabular Editor macros from the command line.'
        }
    }
    foreach ($f in $manifests.Keys) {
        if (-not (Test-Path $f) -and -not $WhatIfPreference) {
            New-Item -ItemType Directory -Force (Split-Path $f) -WhatIf:$false | Out-Null
            [IO.File]::WriteAllText($f, ($manifests[$f] | ConvertTo-Json -Depth 4), $utf8)
        }
    }
}
New-Item -ItemType Directory -Force $SkillsDir -WhatIf:$false | Out-Null
$SkillsDir = (Resolve-Path $SkillsDir).Path

if (-not $NoCompileCheck -and -not ($TabularEditor -and (Test-Path $TabularEditor))) {
    Write-Warning "Skipping the compile check: no $(if ($Engine -eq 'te') { 'te.exe' } else { 'TabularEditor.exe' }) found (pass -TabularEditor)."
    $NoCompileCheck = $true
}

# Header lines of one kind ("SKILL" / "MACRO"), without the prefix and its first space
function Get-HeaderLines([string] $code, [string] $kind) {
    foreach ($line in $code -split "\r?\n") {
        if ($line -match "^\s*//\s*${kind}:(.*)$") { $Matches[1] -replace '^ ', '' }
    }
}

function ConvertTo-YamlString([string] $s) { '"' + ($s -replace '\\', '\\' -replace '"', '\"') + '"' }

# Path as the agent should type it: under the profile, written with $env:USERPROFILE
function Get-DisplayPath([string] $p) {
    if ($p.StartsWith($env:USERPROFILE, [StringComparison]::OrdinalIgnoreCase)) { '$env:USERPROFILE' + $p.Substring($env:USERPROFILE.Length) } else { $p }
}

function Test-MacroCompiles([string] $code) {
    $tmp = Join-Path $env:TEMP ("$generator-" + [guid]::NewGuid() + '.csx')
    [IO.File]::WriteAllText($tmp, $code, $utf8)
    try {
        $ErrorActionPreference = 'Continue'
        if ($Engine -eq 'te') {
            $out = & $TabularEditor script --file $tmp --validate --non-interactive 2>&1 | ForEach-Object { "$_" } |
                Where-Object { $_ -match '\): error' }
            return [pscustomobject]@{ Ok = ($LASTEXITCODE -eq 0); Errors = @($out | ForEach-Object { $_.Trim() }) }
        }
        # TabularEditor.exe has no validate-only mode: the macro runs on an empty model, which is never saved. Its
        # inputs aren't set, so a skill-ready macro stops at the first one with Error(...); only compile errors count
        $bim = "$tmp.bim"
        [IO.File]::WriteAllText($bim, '{ "name": "CompileCheck", "compatibilityLevel": 1702, "model": { "culture": "en-US" } }', $utf8)
        try {
            $psi = New-Object Diagnostics.ProcessStartInfo $TabularEditor, "`"$bim`" -S `"$tmp`""
            $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true
            $p = [Diagnostics.Process]::Start($psi); $out = $p.StandardOutput.ReadToEndAsync()
            if (-not $p.WaitForExit(60000)) {
                $p.Kill()
                return [pscustomobject]@{ Ok = $false; Errors = @('it waits on a dialog in the TE2 command line (no Fx.IsCliMode() branch)') }
            }
            $errors = @($out.Result -split "\r?\n" | Where-Object { $_ -match '^Error CS\d+' } | ForEach-Object { $_.Trim() })
            [pscustomobject]@{ Ok = ($errors.Count -eq 0); Errors = $errors }
        }
        finally { Remove-Item $bim -Force -ErrorAction SilentlyContinue -WhatIf:$false }
    }
    finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue -WhatIf:$false }
}

# --- Parse every macro with a SKILL header ---------------------------------------------------------
$macros = (Get-Content $MacrosFile -Raw -Encoding UTF8 | ConvertFrom-Json).Actions
$skills = foreach ($m in $macros) {
    $lines = @(Get-HeaderLines $m.Execute 'SKILL')
    if ($lines.Count -eq 0) { continue }

    $s = [pscustomobject]@{
        Name = $null; Description = $null; Body = $null; MacroName = $m.Name; MacroId = $m.Id
        Title = $m.Name; ValidContexts = $m.ValidContexts; Inputs = @(); Errors = @(); Warnings = @(); Code = $m.Execute
    }
    $body = New-Object Collections.Generic.List[string]
    foreach ($l in $lines) {
        if (-not $s.Name -and $l -match '^name:\s*(.*?)\s*$') { $s.Name = $Matches[1] }
        elseif (-not $s.Description -and $l -match '^description:\s*(.*?)\s*$') { $s.Description = $Matches[1] }
        else { $body.Add($l.TrimEnd()) }
    }
    $s.Body = ($body -join "`n").Trim("`n")

    foreach ($l in @(Get-HeaderLines $m.Execute 'MACRO')) {
        if ($l -match '^Name:\s*(.+?)\s*$') { $s.Title = $Matches[1] }
    }

    if (-not $s.Name) { $s.Errors += 'no "// SKILL: name:" line' }
    elseif ($s.Name -notmatch '^[a-z0-9]+(-[a-z0-9]+)*$' -or $s.Name.Length -gt 64) { $s.Errors += "name '$($s.Name)' must be lowercase letters, digits and hyphens (max 64)" }
    if (-not $s.Description) { $s.Errors += 'no "// SKILL: description:" line' }
    elseif ($s.Description.Length -gt 1024) { $s.Errors += "description is $($s.Description.Length) characters (max 1024)" }
    if (-not $s.Body) { $s.Errors += 'no skill text after name and description' }

    # Inputs documented under "## Inputs" vs. env variables read in the code
    $inInputs = $false
    foreach ($l in $body) {
        if ($l -match '^##\s') { $inInputs = $l -match '^##\s+Inputs' ; continue }
        if ($inInputs -and $l -match '^-\s+`?([A-Za-z0-9_]+)`?\s*\((required|optional)') {
            $s.Inputs += [pscustomobject]@{ name = $Matches[1]; required = ($Matches[2] -eq 'required') }
        }
    }
    $used = @([regex]::Matches($m.Execute, '(?:FromEnv\(\s*(?:\w+\s*,\s*)?|GetEnvironmentVariable\(\s*)"([A-Za-z0-9_]+)"') |
        ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    $documented = @($s.Inputs | ForEach-Object { $_.name })
    foreach ($v in $used | Where-Object { $documented -notcontains $_ }) { $s.Warnings += "$v is read in the code but not listed under ## Inputs" }
    foreach ($v in $documented | Where-Object { $used -notcontains $_ }) { $s.Warnings += "$v is listed under ## Inputs but not read in the code" }
    if (-not (@($body) -match '^##\s+Inputs')) { $s.Warnings += 'no "## Inputs" section' }
    # Skills run on either engine (default, -Engine or fallback), so check both guards
    if ($m.Execute -notmatch '#if\s+TECLI') { $s.Warnings += 'no #if TECLI block: in te the macro will run its interactive code' }
    if ($m.Execute -notmatch 'IsCliMode\(') { $s.Warnings += 'no Fx.IsCliMode() check: in the TE2 command line the macro will run its interactive code and wait on dialogs' }
    if ($m.Enabled -eq 'false') { $s.Warnings += "macro is disabled in $Edition" }
    $s
}

foreach ($dup in $skills | Where-Object Name | Group-Object Name | Where-Object Count -gt 1) {
    foreach ($s in $dup.Group) { $s.Errors += "skill name '$($dup.Name)' is used by several macros: $(($dup.Group.MacroName) -join ', ')" }
}
if ($Only) { $skills = @($skills | Where-Object { $Only -contains $_.Name }) }

# --- Write ------------------------------------------------------------------------------------------
$report = New-Object Collections.Generic.List[object]
foreach ($s in $skills) {
    $dir = if ($s.Name) { Join-Path $SkillsDir $s.Name }
    $status = $null

    if (-not $s.Errors -and (Test-Path $dir)) {
        $existing = Join-Path $dir 'macro.json'
        $ours = (Test-Path $existing) -and ((Get-Content $existing -Raw | ConvertFrom-Json).generator -eq $generator)
        if (-not $ours) { $s.Errors += "$dir exists and was not generated by $generator; not overwritten" }
    }

    if ($s.Errors) { $status = 'skipped' }
    else {
        if (-not $NoCompileCheck) {
            $c = Test-MacroCompiles $s.Code
            if (-not $c.Ok) {
                $first = if ($c.Errors) { $c.Errors[0] } else { 'see te script --validate' }
                $s.Warnings += "does not compile in $Engine ($($c.Errors.Count) error(s), first: $first)"
            }
        }

        $runnerPath = Get-DisplayPath (Join-Path $dir 'scripts\Invoke-MacroSkill.ps1')
        $envExample = if ($s.Inputs) {
            '@{ ' + (($s.Inputs | Where-Object required | ForEach-Object { "$($_.name) = `"...`"" }) -join '; ') + ' }'
        } else { '@{}' }
        if ($envExample -eq '@{  }') { $envExample = '@{}' }

        $macrosFileText = '%LOCALAPPDATA%\' + (Split-Path (Split-Path $editionFiles[$Edition]) -Leaf) + '\MacroActions.json'
        $defaultName = if ($Engine -eq 'te') { 'the Tabular Editor CLI (`te`)' } else { 'the Tabular Editor 2 command line (`TabularEditor.exe`)' }
        $otherName = if ($Engine -eq 'te') { 'the Tabular Editor 2 command line (`TabularEditor.exe`)' } else { 'the Tabular Editor CLI (`te`)' }
        $otherEngine = if ($Engine -eq 'te') { 'te2' } else { 'te' }
        $askOther = if ($Engine -eq 'te') { 'Tabular Editor 2 or `TabularEditor.exe`' } else { 'the Tabular Editor CLI or `te`' }
        $engineText = "through $defaultName by default. The macro code is read from ``$macrosFileText`` on every run, so edits made in $Edition apply at once. In ``te`` the macro is compiled with ``TECLI`` defined, and in ``TabularEditor.exe`` it detects the command line at runtime; either way it reads the inputs above from environment variables instead of the selection or dialogs.

**Engine.** Add ``-Engine $otherEngine`` to every command below when the user asks to use $askOther; the runner then uses $otherName and doesn't fall back. Without ``-Engine``, if $defaultName isn't installed the runner warns and uses the other one: mention that to the user."
        $inspectText = 'for example `te ls -m <path>`, or by reading the model''s TMDL files or `model.bim`'
        $dryRunText = 'Nothing is saved. With `te` the runner prints te''s diff of every object the macro would create or change; with `TabularEditor.exe` it prints every object created (`+`), changed (`~`) or deleted (`-`):'
        $saveText = 'With `te`, add `-Force` only if the user accepts saving with DAX validation errors. `TabularEditor.exe` doesn''t validate DAX, so after saving with it ask the user to check the new or changed objects in Power BI.'
        $engineGotchas = @(
            '- Saving a PBIP folder can rewrite more lines than the macro changed (with `te`: `compatibilityMode` in `database.tmdl`, blank lines, line endings; with `TabularEditor.exe`: more TMDL lines). Review `git diff` afterwards.'
            '- `TabularEditor.exe` with `-Local`: only one Power BI Desktop may be open; otherwise pass `-Server localhost:<port> -Database <name>`.'
            '- If the runner stops `TabularEditor.exe` after the timeout, the macro is waiting on a dialog: one of its inputs isn''t CLI-ready. Tell the user; don''t retry.'
        )

        $skillMd = @"
---
name: $($s.Name)
description: $(ConvertTo-YamlString $s.Description)
---

# $($s.Title)

$($s.Body)

## How to run

This skill runs the $editionName macro **$($s.MacroName)** $engineText

1. **Pick the model.** Ask the user if the request doesn't make it clear.

   | Parameter | Target |
   |---|---|
   | ``-Path <folder or .bim>`` | ``*.SemanticModel`` / TMDL folder or ``model.bim`` |
   | ``-Local [-Database <name>]`` | Running Power BI Desktop / SSAS instance |
   | ``-Server <workspace or endpoint> -Database <name>`` | Fabric / Power BI workspace or XMLA endpoint |

2. **Work out the inputs** listed above. Inspect the model when needed ($inspectText). Ask the user when a value is unclear instead of guessing.

3. **Dry run.** $dryRunText

   ``````powershell
   & "$runnerPath" -Path "<model>" -MacrosFile "$(Get-DisplayPath ((Resolve-Path $MacrosFile).Path))" -Env $envExample
   ``````

   A non-zero exit code means the engine or the macro reported an error (a missing or unresolved input, for example). Fix the inputs and retry, or tell the user.

4. **Summarize the diff for the user and get confirmation**, then run the same command with ``-Save``. $saveText

## Gotchas

- Close Power BI Desktop before using ``-Path`` on a PBIP it has open, or use ``-Local`` instead.
$($engineGotchas -join "`n")
- If the runner says no macro carries this skill's name, the macro was removed or its ``// SKILL: name:`` changed in $($Edition): regenerate the skills with te-macro-skills.

<!-- Generated by $generator from the "$($s.MacroName)" macro's // SKILL: header. Do not edit; change the macro and regenerate. -->
"@
        $macroJson = [ordered]@{
            generator     = $generator
            skillName     = $s.Name
            macroName     = $s.MacroName
            edition       = $Edition
            engine        = $Engine
            validContexts = $s.ValidContexts
            inputs        = @($s.Inputs)
        } | ConvertTo-Json -Depth 4

        $files = @{
            'SKILL.md'                     = $skillMd -replace "`r`n", "`n"
            'macro.json'                   = $macroJson -replace "`r`n", "`n"
            'scripts\Invoke-MacroSkill.ps1' = [IO.File]::ReadAllText($runner)
        }
        $isNew = -not (Test-Path $dir)
        $changed = $isNew -or @($files.Keys | Where-Object {
            $f = Join-Path $dir $_
            -not (Test-Path $f) -or [IO.File]::ReadAllText($f) -cne $files[$_]
        }).Count -gt 0

        if (-not $changed) { $status = 'unchanged' }
        elseif ($PSCmdlet.ShouldProcess($dir, $(if ($isNew) { 'Create skill' } else { 'Update skill' }))) {
            New-Item -ItemType Directory -Force (Join-Path $dir 'scripts') | Out-Null
            foreach ($k in $files.Keys) { [IO.File]::WriteAllText((Join-Path $dir $k), $files[$k], $utf8) }
            $status = if ($isNew) { 'created' } else { 'updated' }
        }
        else { $status = if ($isNew) { 'would create' } else { 'would update' } }
    }

    $report.Add([pscustomobject]@{
        Skill = $s.Name; Macro = $s.MacroName; Status = $status
        Inputs = (@($s.Inputs | ForEach-Object { $_.name + $(if ($_.required) { '*' }) }) -join ', ')
        Problems = (@($s.Errors) + @($s.Warnings)) -join '; '
    })
}

# --- Orphans: generated skills whose macro is gone ---------------------------------------------------
if (-not $Only) {
    $current = @($skills | ForEach-Object { $_.Name })
    foreach ($mj in Get-ChildItem $SkillsDir -Directory | ForEach-Object { Join-Path $_.FullName 'macro.json' } | Where-Object { Test-Path $_ }) {
        $meta = Get-Content $mj -Raw | ConvertFrom-Json
        if ($meta.generator -ne $generator -or $current -contains $meta.skillName) { continue }
        $dir = Split-Path $mj
        $status = 'orphaned (rerun with -Prune to delete)'
        if ($Prune) {
            $status = if ($PSCmdlet.ShouldProcess($dir, 'Delete orphaned skill')) { Remove-Item $dir -Recurse -Force; 'deleted' } else { 'would delete' }
        }
        $report.Add([pscustomobject]@{ Skill = $meta.skillName; Macro = $meta.macroName; Status = $status; Inputs = ''; Problems = 'no macro carries this skill name any more' })
    }
}

if ($report.Count -eq 0) { Write-Host "No macro in $MacrosFile has a // SKILL: header." }

# --- Local plugin registration -----------------------------------------------------------------------
if ($asPlugin -and -not $WhatIfPreference) {
    $plugins = "$env:USERPROFILE\.claude\plugins"
    $known = (Test-Path "$plugins\known_marketplaces.json") -and
        (Get-Content "$plugins\known_marketplaces.json" -Raw | ConvertFrom-Json).PSObject.Properties.Name -contains $marketplaceName
    $installed = (Test-Path "$plugins\installed_plugins.json") -and
        (Get-Content "$plugins\installed_plugins.json" -Raw | ConvertFrom-Json).plugins.PSObject.Properties.Name -contains "$pluginName@$marketplaceName"
    if ($Register) {
        if (-not $known) {
            & claude plugin marketplace add $PluginDir | Write-Host
            if ($LASTEXITCODE -ne 0) { throw "Failed to add the $marketplaceName marketplace (exit code $LASTEXITCODE)." }
        }
        if (-not $installed) {
            & claude plugin install "$pluginName@$marketplaceName" --scope user | Write-Host
            if ($LASTEXITCODE -ne 0) { throw "Failed to install $pluginName@$marketplaceName (exit code $LASTEXITCODE)." }
        }
    }
    elseif (-not ($known -and $installed)) {
        Write-Warning "The $pluginName plugin is not registered in Claude Code yet. Rerun with -Register, or run: claude plugin marketplace add `"$PluginDir`"; claude plugin install $pluginName@$marketplaceName"
    }
    Write-Host "Skills are in the $pluginName plugin ($pluginRoot). Run /reload-plugins (or start a new session) to load changes; enable or disable them all with /plugin."
}
$report
exit [int](@($report | Where-Object Status -eq 'skipped').Count -gt 0)
