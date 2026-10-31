<#
.SYNOPSIS
  Runs the Tabular Editor 3 macro behind a generated skill through the Tabular Editor CLI (te).

.DESCRIPTION
  Copied by te-macro-skills into every generated skill. It reads macro.json (next to SKILL.md) for the
  skill name, finds the macro whose code carries "// SKILL: name: <skill>" in the macros file (so renaming
  the macro in TE3 doesn't break the skill), sets the input environment variables and runs the macro's
  code with `te script <source>`. In te the code is compiled with TECLI defined, so it reads its inputs
  from those variables instead of the selection or dialogs.

  Without -Save it is a dry run: te runs the macro and prints the change diff, but saves nothing.

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
    # Save even if the macro introduces DAX validation errors
    [switch] $Force,
    [string] $MacrosFile = "$env:LOCALAPPDATA\TabularEditor3\MacroActions.json",
    [string] $TabularEditor
)

$ErrorActionPreference = 'Stop'

$meta = Get-Content (Join-Path $PSScriptRoot '..\macro.json') -Raw | ConvertFrom-Json

# te.exe: parameter, then PATH, then the default TE3 install folder
if (-not $TabularEditor) {
    $onPath = Get-Command te.exe -ErrorAction SilentlyContinue
    $TabularEditor = if ($onPath) { $onPath.Source } else { "$env:LOCALAPPDATA\Programs\TabularEditor3\te.exe" }
}
if (-not (Test-Path $TabularEditor)) { throw "Tabular Editor CLI not found at '$TabularEditor'. Pass -TabularEditor." }

# te flags selecting the model: -m <path> | --local [-d db] | -s <server> -d <db>
if ($Path)       { $source = @('-m', (Resolve-Path $Path).Path) }
elseif ($Local)  { $source = @('--local') }
elseif ($Server) {
    if (-not $Database) { throw "-Database is required with -Server." }
    $source = @('-s', $Server)
}
else { throw "Pass one of -Path, -Local or -Server." }
if ($Database) { $source += @('-d', $Database) }

# Find the macro by its skill name, not by its TE3 name
if (-not (Test-Path $MacrosFile)) { throw "Macros file not found: $MacrosFile" }
$pattern = '(?m)^\s*//\s*SKILL:\s*name:\s*' + [regex]::Escape($meta.skillName) + '\s*$'
$macros = @((Get-Content $MacrosFile -Raw | ConvertFrom-Json).Actions | Where-Object { $_.Execute -match $pattern })
if ($macros.Count -eq 0) { throw "No macro in $MacrosFile has '// SKILL: name: $($meta.skillName)'. Was it removed or renamed in TE3? Regenerate the skills with te-macro-skills." }
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

# The macro's code runs through `te script` rather than `te macro run`: same code and TECLI define, but
# te script prints the change diff, which the dry run needs
$tmp = Join-Path $env:TEMP ("$($meta.skillName)-" + [guid]::NewGuid() + '.csx')
[IO.File]::WriteAllText($tmp, $macro.Execute, (New-Object Text.UTF8Encoding($false)))
$teArgs = @('script', '--file', $tmp) + $source + @('--non-interactive')
if ($Save)  { $teArgs += '--save' }
if ($Force) { $teArgs += '--force' }

# Set the inputs only for this call, restoring whatever was there before
$previous = @{}
foreach ($k in $Env.Keys) {
    $previous[$k] = [Environment]::GetEnvironmentVariable($k, 'Process')
    [Environment]::SetEnvironmentVariable($k, [string]$Env[$k], 'Process')
}
try {
    Write-Host "=== macro '$($macro.Name)' ($($meta.skillName))$(if (-not $Save) { ' - dry run' })"
    $ErrorActionPreference = 'Continue'
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
finally {
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    foreach ($k in $previous.Keys) { [Environment]::SetEnvironmentVariable($k, $previous[$k], 'Process') }
}
exit $exitCode
