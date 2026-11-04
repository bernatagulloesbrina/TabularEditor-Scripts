<#
.SYNOPSIS
  Runs the Tabular Editor macro behind a generated skill through the Tabular Editor CLI (te) or the Tabular Editor 2
  command line (TabularEditor.exe).

.DESCRIPTION
  Copied by te-macro-skills into every generated skill. It reads macro.json (next to SKILL.md) for the
  skill name, finds the macro whose code carries "// SKILL: name: <skill>" in the macros file (so renaming
  the macro in Tabular Editor doesn't break the skill), sets the input environment variables and runs the macro's
  code.

  Engine te: `te script <source>`. The code is compiled with TECLI defined, so it reads its inputs from those
  variables instead of the selection or dialogs. Without -Save te prints the change diff and saves nothing.

  Engine te2: `TabularEditor.exe <model> -S <script>`. The macro detects the command line at runtime (Fx.IsCliMode)
  and reads the same variables. Without -Save the model is exported before and after the macro and the runner prints
  the objects created, changed or deleted; nothing is saved. With -Save the model is written back: -B for a .bim
  file, -TMDL for a TMDL / PBIP folder, and an incremental SaveChanges for a server or Power BI Desktop. TE2 doesn't
  validate DAX.

  The engine comes from macro.json (chosen when the skill was generated); -Engine overrides it.

.EXAMPLE
  .\Invoke-MacroSkill.ps1 -Path ".\Sales.SemanticModel" -Env @{ SUM_COLUMNS = "'Sales'[Amount];'Sales'[Qty]" }
  .\Invoke-MacroSkill.ps1 -Path ".\Sales.SemanticModel" -Env @{ SUM_COLUMNS = "'Sales'[Amount];'Sales'[Qty]" } -Save
#>
param(
    # Model source: one of -Path, -Local or -Server
    [string] $Path,
    [switch] $Local,
    [string] $Server,
    [string] $Database,
    # Input environment variables for the macro, e.g. @{ SUM_COLUMNS = "'Sales'[Amount]" }
    [hashtable] $Env = @{},
    [switch] $Save,
    # Save even if the macro introduces DAX validation errors (te only)
    [switch] $Force,
    [string] $MacrosFile,
    [ValidateSet('te', 'te2')]
    [string] $Engine,
    # te.exe or TabularEditor.exe
    [string] $TabularEditor,
    # te2: TabularEditor.exe is stopped after this long (a macro waiting on a dialog never ends)
    [int] $TimeoutSeconds = 600
)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object Text.UTF8Encoding($false)

$meta = Get-Content (Join-Path $PSScriptRoot '..\macro.json') -Raw | ConvertFrom-Json
if (-not $Engine) { $Engine = if ($meta.engine) { $meta.engine } else { 'te' } }
if (-not $MacrosFile) {
    $MacrosFile = if ($meta.edition -eq 'TE2') { "$env:LOCALAPPDATA\TabularEditor\MacroActions.json" } else { "$env:LOCALAPPDATA\TabularEditor3\MacroActions.json" }
}

# Executable: parameter, then PATH, then the default install folders
if (-not $TabularEditor) {
    if ($Engine -eq 'te') {
        $onPath = Get-Command te.exe -ErrorAction SilentlyContinue
        $TabularEditor = if ($onPath) { $onPath.Source } else { "$env:LOCALAPPDATA\Programs\TabularEditor3\te.exe" }
    }
    else {
        $onPath = Get-Command TabularEditor.exe -ErrorAction SilentlyContinue
        $TabularEditor = if ($onPath) { $onPath.Source } else {
            @("${env:ProgramFiles(x86)}\Tabular Editor\TabularEditor.exe", "$env:ProgramFiles\Tabular Editor\TabularEditor.exe") |
                Where-Object { Test-Path $_ } | Select-Object -First 1
        }
    }
}
if (-not $TabularEditor -or -not (Test-Path $TabularEditor)) {
    throw "$(if ($Engine -eq 'te') { 'Tabular Editor CLI (te.exe)' } else { 'Tabular Editor 2 (TabularEditor.exe)' }) not found. Pass -TabularEditor, or -Engine $(if ($Engine -eq 'te') { 'te2' } else { 'te' }) to use the other one."
}
if (-not ($Path -or $Local -or $Server)) { throw "Pass one of -Path, -Local or -Server." }
if ($Server -and -not $Database) { throw "-Database is required with -Server." }

# Find the macro by its skill name, not by its name in Tabular Editor
if (-not (Test-Path $MacrosFile)) { throw "Macros file not found: $MacrosFile" }
$pattern = '(?m)^\s*//\s*SKILL:\s*name:\s*' + [regex]::Escape($meta.skillName) + '\s*$'
$macros = @((Get-Content $MacrosFile -Raw -Encoding UTF8 | ConvertFrom-Json).Actions | Where-Object { $_.Execute -match $pattern })
if ($macros.Count -eq 0) { throw "No macro in $MacrosFile has '// SKILL: name: $($meta.skillName)'. Was it removed or renamed in Tabular Editor? Regenerate the skills with te-macro-skills." }
if ($macros.Count -gt 1) { throw "Several macros carry '// SKILL: name: $($meta.skillName)': $(($macros.Name) -join ', '). Keep it on one only." }
$macro = $macros[0]

# Inputs: required ones must be set, unknown ones are reported
$known = @($meta.inputs | ForEach-Object { $_.name })
foreach ($i in $meta.inputs | Where-Object required) {
    if (-not $Env.ContainsKey($i.name) -or [string]::IsNullOrWhiteSpace([string]$Env[$i.name])) { throw "Required input $($i.name) is missing." }
}
foreach ($k in $Env.Keys) {
    if ($known -notcontains $k) { Write-Warning "$k is not an input of $($meta.skillName) (inputs: $($known -join ', '))." }
}

# --- te2 helpers ------------------------------------------------------------------------------------------
# Runs TabularEditor.exe and returns its exit code and output without the banner and progress lines
function Invoke-TE2([string[]] $arguments) {
    $psi = New-Object Diagnostics.ProcessStartInfo $TabularEditor, (($arguments | ForEach-Object { if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\"') + '"' } else { $_ } }) -join ' ')
    $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $p = [Diagnostics.Process]::Start($psi)
    $out = $p.StandardOutput.ReadToEndAsync(); $err = $p.StandardError.ReadToEndAsync()
    if (-not $p.WaitForExit($TimeoutSeconds * 1000)) {
        $p.Kill()
        throw "TabularEditor.exe didn't finish in $TimeoutSeconds s and was stopped. The macro is most likely waiting on a dialog: it has an input that isn't read from an environment variable in the TE2 command line (no Fx.IsCliMode() branch). Make the macro skill-ready again with the current pattern."
    }
    $lines = @(($out.Result + $err.Result) -split "\r?\n" | Where-Object {
        $_ -and $_ -notmatch '^(Tabular Editor \d|-{5,}|Loading macros from|Loading model|Loaded script|Executing script|Building Model\.bim|Saving Model\.bim)'
    })
    [pscustomobject]@{ ExitCode = $p.ExitCode; Lines = $lines }
}

# The model's objects as "<type> '<name>' / ..." -> properties (child collections excluded; annotations kept)
function Get-ModelObjects($node, [string] $prefix, [hashtable] $into) {
    $props = [ordered]@{}
    foreach ($p in $node.PSObject.Properties) {
        $v = $p.Value
        $isChildren = $v -is [array] -and $v.Count -gt 0 -and $v[0] -is [pscustomobject] -and $null -ne $v[0].name -and
            $p.Name -notin @('annotations', 'extendedProperties')
        if ($isChildren) {
            $type = $p.Name -replace 'ies$', 'y' -replace 's$', ''
            foreach ($c in $v) { Get-ModelObjects $c ("$prefix$(if ($prefix) { ' / ' })$type '$($c.name)'") $into }
        }
        elseif ($p.Name -ne 'name') { $props[$p.Name] = ($v | ConvertTo-Json -Depth 20 -Compress) }
    }
    $into[$(if ($prefix) { $prefix } else { 'model' })] = $props
}

function Format-Value([string] $json) {
    $s = try { $v = $json | ConvertFrom-Json; if ($v -is [string]) { $v } else { $json } } catch { $json }
    $s = ($s -replace '\s+', ' ').Trim()
    if ($s.Length -gt 120) { $s.Substring(0, 117) + '...' } else { $s }
}

# Created / deleted / changed objects between two .bim exports
function Write-ModelDiff([string] $beforeBim, [string] $afterBim) {
    $before = @{}; $after = @{}
    Get-ModelObjects (Get-Content $beforeBim -Raw -Encoding UTF8 | ConvertFrom-Json).model '' $before
    Get-ModelObjects (Get-Content $afterBim -Raw -Encoding UTF8 | ConvertFrom-Json).model '' $after
    $n = 0
    foreach ($k in $after.Keys | Sort-Object) {
        if (-not $before.ContainsKey($k)) {
            $n++; Write-Host "+ $k"
            foreach ($p in $after[$k].Keys | Where-Object { $_ -ne 'lineageTag' }) { Write-Host "    $p = $(Format-Value $after[$k][$p])" }
        }
        else {
            $changed = @(@($after[$k].Keys) + @($before[$k].Keys) | Sort-Object -Unique | Where-Object { $after[$k][$_] -cne $before[$k][$_] })
            if ($changed) {
                $n++; Write-Host "~ $k"
                foreach ($p in $changed) {
                    $old = if ($null -ne $before[$k][$p]) { Format-Value $before[$k][$p] } else { '(none)' }
                    $new = if ($null -ne $after[$k][$p]) { Format-Value $after[$k][$p] } else { '(none)' }
                    Write-Host "    ${p}: $old -> $new"
                }
            }
        }
    }
    foreach ($k in $before.Keys | Where-Object { -not $after.ContainsKey($_) } | Sort-Object) { $n++; Write-Host "- $k" }
    if ($n -eq 0) { Write-Host 'No changes.' } else { Write-Host "$n object(s) changed." }
}

# --- Run --------------------------------------------------------------------------------------------------
# Set the inputs only for this call, restoring whatever was there before
$previous = @{}
foreach ($k in $Env.Keys) {
    $previous[$k] = [Environment]::GetEnvironmentVariable($k, 'Process')
    [Environment]::SetEnvironmentVariable($k, [string]$Env[$k], 'Process')
}
$tmpDir = Join-Path $env:TEMP ("$($meta.skillName)-" + [guid]::NewGuid())
New-Item -ItemType Directory $tmpDir | Out-Null
try {
    Write-Host "=== macro '$($macro.Name)' ($($meta.skillName), $Engine)$(if (-not $Save) { ' - dry run' })"
    $ErrorActionPreference = 'Continue'

    if ($Engine -eq 'te') {
        # The macro's code runs through `te script` rather than `te macro run`: same code and TECLI define, but
        # te script prints the change diff, which the dry run needs
        if ($Path)       { $source = @('-m', (Resolve-Path $Path).Path) }
        elseif ($Local)  { $source = @('--local') }
        else             { $source = @('-s', $Server) }
        if ($Database) { $source += @('-d', $Database) }

        $tmp = Join-Path $tmpDir 'macro.csx'
        [IO.File]::WriteAllText($tmp, $macro.Execute, $utf8)
        $teArgs = @('script', '--file', $tmp) + $source + @('--non-interactive')
        if ($Save)  { $teArgs += '--save' }
        if ($Force) { $teArgs += '--force' }

        $out = & $TabularEditor @teArgs 2>&1 | ForEach-Object { "$_" } |
            Where-Object { $_ -notmatch 'early preview|tabulareditor\.com/downloads' }
        $exitCode = $LASTEXITCODE
        $out | Write-Host
        # te exits 0 when the macro itself calls Error(...), e.g. for a missing or wrong input
        if ($exitCode -eq 0 -and ($out -match '^Error\b')) {
            Write-Warning "The macro reported an error (see above)."
            $exitCode = 1
        }
    }
    else {
        # TabularEditor.exe <file or folder> | <server> <database>, and how -Save writes the model back
        $code = $macro.Execute
        if ($Path) {
            $p = (Resolve-Path $Path).Path
            if (Test-Path $p -PathType Container) {
                if (Test-Path (Join-Path $p 'definition')) { $p = Join-Path $p 'definition' }
                elseif (Test-Path (Join-Path $p 'model.bim')) { $p = Join-Path $p 'model.bim' }
            }
            $source = @($p)
            $saveArgs = if (Test-Path $p -PathType Container) { @('-TMDL', $p) } else { @('-B', $p) }
        }
        else {
            if ($Local) {
                # Power BI Desktop: the port of each running instance is in its workspace folder
                $ports = @(Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces",
                    "$env:USERPROFILE\Microsoft\Power BI Desktop Store App\AnalysisServicesWorkspaces" -Recurse -Filter msmdsrv.port.txt -ErrorAction SilentlyContinue |
                    ForEach-Object { (Get-Content $_.FullName -Raw -Encoding Unicode) -replace '\D', '' } | Where-Object { $_ } | Sort-Object -Unique)
                if ($ports.Count -eq 0) { throw "No running Power BI Desktop found." }
                if ($ports.Count -gt 1) { throw "Several Power BI Desktop instances are running (ports $($ports -join ', ')). Close the others, or pass -Server localhost:<port> -Database <name>." }
                $Server = "localhost:$($ports[0])"
            }
            if (-not $Database) {
                # TabularEditor.exe needs the database name; Power BI Desktop has one database, named with a GUID
                Add-Type -Path (Join-Path (Split-Path $TabularEditor) 'Microsoft.AnalysisServices.Tabular.dll')
                $as = New-Object Microsoft.AnalysisServices.Tabular.Server
                $as.Connect("Data Source=$Server")
                try { $Database = @($as.Databases)[0].Name } finally { $as.Disconnect() }
            }
            $source = @($Server, $Database)
            # Incremental save on the open connection, as TE2's own Save does. Appended to the macro (TE2 runs a
            # single -S script); a macro that stops with Error(...) and return never reaches it
            if ($Save) { $code += "`r`nModel.Database.TOMDatabase.Model.SaveChanges();`r`n" }
            $saveArgs = @()
        }

        $tmp = Join-Path $tmpDir 'macro.csx'
        [IO.File]::WriteAllText($tmp, $code, $utf8)
        if ($Save) {
            $r = Invoke-TE2 ($source + @('-S', $tmp) + $saveArgs)
            $r.Lines | Write-Host
            $exitCode = $r.ExitCode
            if ($exitCode -eq 0) { Write-Host "Saved. TE2 doesn't validate DAX: check the new or changed objects in Power BI." }
        }
        else {
            $beforeBim = Join-Path $tmpDir 'before.bim'; $afterBim = Join-Path $tmpDir 'after.bim'
            $r = Invoke-TE2 ($source + @('-S', $tmp, '-B', $afterBim))
            $r.Lines | Write-Host
            $exitCode = $r.ExitCode
            if ($exitCode -eq 0) {
                $b = Invoke-TE2 ($source + @('-B', $beforeBim))
                if ($b.ExitCode -ne 0) { $b.Lines | Write-Host; $exitCode = $b.ExitCode }
                else { Write-ModelDiff $beforeBim $afterBim; Write-Host "Dry run: nothing was saved. TE2 doesn't validate DAX." }
            }
        }
        if ($exitCode -ne 0) { Write-Warning "TabularEditor.exe or the macro reported an error (see above)." }
    }
}
finally {
    Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
    foreach ($k in $previous.Keys) { [Environment]::SetEnvironmentVariable($k, $previous[$k], 'Process') }
}
exit $exitCode
