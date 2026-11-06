---
name: te-setup
description: Set up, from scratch or to repair, the Tabular Editor macro environments on a Windows machine, up to the point where the user can import the macros and generate their skills (it doesn't import or generate them) - the consumption environment (what's needed to import skill-ready macros from TabularEditor-Scripts into Tabular Editor and run them as Claude skills through te / TE2) and the authoring environment (TE-Scripting-in-Visual-Studio: write macros in Visual Studio, build with dotnet, copy them with CopyMacro, publish them to TabularEditor-Scripts). Checks prerequisites (Tabular Editor 3/2, te CLI, .NET SDK, git, repos, environment variables, plugins), does every step it can, and tells the user exactly what only they can do. Use when the user asks to "set up / install / configure the environment", "set up TE-Scripting-in-Visual-Studio", "prepare a new machine for the macros", "onboard a colleague", "why doesn't the solution build", "fix the TabularEditor references", or "check my setup".
---

# Set up the Tabular Editor macro environments

Two environments, and the second builds on the first:

| Environment | For | Lives in |
|---|---|---|
| **Consumption** | Anyone who wants to run the skill-ready macros (in Tabular Editor, or by asking Claude) | Tabular Editor 2 (and Tabular Editor 3 if the user already has it), the `te` CLI, the `te-macros` plugin (this one), the generated `macro-skills` plugin |
| **Authoring** | Whoever writes or changes macros | A clone of TE-Scripting-in-Visual-Studio (Visual Studio solution) next to a clone of TabularEditor-Scripts (the publication repo), the .NET SDK, Tabular Editor 2 (the projects compile against it), the CopyMacro macro in Tabular Editor |

The check script is `scripts\Test-TeEnvironment.ps1` next to this file. It changes nothing (except `-Build`, which builds the solution) and returns one row per check: `Area`, `Check`, `Status` (ok / missing / warning / info), `Detail`, `Fix`, `Who` (`claude`: you can do it; `user`: only the user can).

## Step 1 - Scope and check

If the request doesn't make it clear, ask with AskUserQuestion: **run macros only** (consumption) or **also write macros** (consumption + authoring). Then run:

```powershell
& "${CLAUDE_SKILL_DIR}\scripts\Test-TeEnvironment.ps1" [-Authoring] [-Build]
```

Pass `-Authoring -Build` for authoring (`-AuthoringRepo <path>` / `-PublishRepo <path>` if the clones are somewhere unusual). Show the user a compact table of what is not `ok`, split into **I can do this** and **you need to do this**. If everything is ok, say so and stop.

## Step 2 - Do it, in this order

Group the actions that need the user's go-ahead (installs, clones, environment variables, writing to Tabular Editor's macros) into one confirmation, listing the exact commands, instead of asking one by one. Items marked `user` go in a short numbered checklist; wait for the user to say they're done, then re-run the check.

### Software (consumption and authoring)

| Need | Who | How |
|---|---|---|
| Tabular Editor 3 | - | **Optional. Never install it or ask the user to install it.** If it's installed, it's used too (macros are imported into it as well, skills are generated from its macros). If it's installed but `%LOCALAPPDATA%\TabularEditor3` doesn't exist yet, it just hasn't been started; that's fine, TE2 covers everything. |
| Tabular Editor 2 | claude | `winget install --id TabularEditor.TabularEditor.2 -e` (free). Required when TE3 isn't installed, and always for authoring (every project references `TabularEditor.exe` and `TOMWrapper.dll`). Then the user starts it once and closes it, which creates `%LOCALAPPDATA%\TabularEditor`, where macros are imported. |
| te CLI | - | **Optional and licensed separately: mention it once, never wait for it or block on it.** If it's found, the skills run with it (with DAX validation). If not, tell the user it exists (download from https://tabulareditor.com/downloads if they have a license) and carry on: the skills run through TabularEditor.exe (TE2). If the user installs it later: extract `te-win-x64.zip` to `%LOCALAPPDATA%\Programs\TabularEditor3` and add that folder to the **user** PATH (below). Preview builds expire: the check mentions it 14 days ahead. |
| git | claude | `winget install --id Git.Git -e` (authoring) |
| .NET SDK 6 or later | claude | `winget install --id Microsoft.DotNet.SDK.8 -e` (authoring). This is all the build needs: every project is SDK-style and targets net48, so the .NET Framework reference assemblies come from NuGet (no Developer Pack). |
| Visual Studio | user | Optional: `winget install --id Microsoft.VisualStudio.2022.Community -e` with the ".NET desktop development" workload, or VS Code + C# Dev Kit, or Rider. Only needed to edit with IntelliSense; publish-macro builds with `dotnet`. |

winget installs can show a UAC prompt: tell the user to approve it. After an install, the current shell doesn't see the new PATH; refresh it before using the tool:

```powershell
$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
```

### Environment variables

Set user variables with `[Environment]::SetEnvironmentVariable(<name>, <value>, 'User')`, never `setx` (it truncates PATH at 1024 characters). They apply to processes started afterwards: tell the user to restart Visual Studio, terminals and Claude Code.

| Variable | When | Value |
|---|---|---|
| `Path` (user) | only if the user installs the te CLI | append `%LOCALAPPDATA%\Programs\TabularEditor3` if it isn't there (read the current user Path, append, write back) |
| `TabularEditor2Dir` (user) | only if TE2 is **not** in `C:\Program Files (x86)\Tabular Editor` | the folder holding `TabularEditor.exe` (read by `TE.References.props` in the authoring solution) |
| `TE_PUBLISH_REPO` | only if TabularEditor-Scripts is **not** cloned next to TE-Scripting-in-Visual-Studio | the clone's path. Put it in `<authoring repo>\.claude\settings.local.json` as `{ "env": { "TE_PUBLISH_REPO": "D:\\repos\\TabularEditor-Scripts" } }` (merge into the file if it exists; it's git-ignored), or as a user variable to use it everywhere |
| `GITHUB_TOKEN` | optional | Only to import from private repos or to avoid GitHub's 60 requests/hour anonymous limit. Better: the user runs `gh auth login`. Never ask the user to paste a token into the chat. |

### Consumption

Setup ends when the machine is **ready** to import macros and generate skills. Don't import macros or generate skills here: that's the te-macro-skills skill, and the user starts it when they choose.

1. **Plugin**: if `te-macros` is missing (someone running this skill from a copy), `claude plugin marketplace add bernatagulloesbrina/TabularEditor-Scripts` then `claude plugin install te-macros@tabular-editor-scripts --scope user`.
2. **Ready check**: at least one edition whose settings folder exists (`%LOCALAPPDATA%\TabularEditor3` or `%LOCALAPPDATA%\TabularEditor`, created by starting it once), an engine to run the skills (te or TabularEditor.exe), and `claude` on PATH (the generator uses it to register the macro-skills plugin). The `macro-skills plugin` row is `info` until then; that's expected, not a gap to fix.
3. **Next steps for the user**: close Tabular Editor, then ask "import the skill-ready macros and generate their skills" (te-macro-skills does the preview, the import, the skill generation and the first-time `-Register`), then `/reload-plugins`.

### Authoring

1. **Clone both repos side by side** (ask where; default `%USERPROFILE%\source\repos`). Contributors without push rights to bernatagulloesbrina's repos fork TabularEditor-Scripts first and clone the fork (publish-macro writes there; they open pull requests).

   ```powershell
   git clone https://github.com/bernatagulloesbrina/TE-Scripting-in-Visual-Studio "<root>\TE-Scripting-in-Visual-Studio"
   git clone https://github.com/bernatagulloesbrina/TabularEditor-Scripts "<root>\TabularEditor-Scripts"
   ```

   Side by side, publish-macro finds the publication repo with no configuration. Otherwise set `TE_PUBLISH_REPO` (above).
2. **TE2 references**: the projects find TE2 in `C:\Program Files (x86)\Tabular Editor` or in `TabularEditor2Dir`. A build error "Tabular Editor 2 not found in ..." means one of them is wrong. If the clone has old project files with `..\..\..\..\..\..\Program Files (x86)\...` HintPaths (they only work when the repo is exactly six folders below `C:\`), pull the latest version, which uses `TE.References.props`.
3. **Build both ways** (the check's `-Build` does this): `dotnet build "TE Scripts\TE Scripts.csproj"` and the te variant with `-p:DefineConstants=TECLI -p:OutDir=$env:TEMP\tecli-build\`. The first build restores NuGet packages (needs internet).
4. **Install CopyMacro in Tabular Editor** (close it first): the macro lists the methods of `TE Scripts.cs` and puts the chosen one on the clipboard (building with dotnet and running CopyMacro behind the scenes). The repo path is written into the macro, so reinstall it if the clone moves.

   ```powershell
   & "<authoring repo>\.claude\skills\publish-macro\scripts\Install-CopyMacro.ps1" -WhatIf
   & "<authoring repo>\.claude\skills\publish-macro\scripts\Install-CopyMacro.ps1"
   ```

   It replaces a macro with the same name (`-Name`, default "Get Macro"). Point out in the report if it would replace one the user wrote.
5. **Open the solution** (user): `TE Scripting in Visual Studio.sln`. Start Claude Code **inside the authoring repo** for the make-skill-ready and publish-macro skills (they're project skills there).

## Step 3 - Verify and report

Re-run the check (with `-Authoring -Build` when relevant) and show the final table. Then summarise:
- what you installed or changed, and every environment variable you set (name, scope, value)
- what the user still has to do (restart apps for new variables, te CLI renewal date), and the next step: close Tabular Editor and ask to import the macros and generate their skills
- the daily loop: **consume**: ask Claude to run a macro skill on a model, or run the macro in Tabular Editor. **Author**: edit `TE Scripts.cs` in Visual Studio, make it skill-ready (make-skill-ready), "publish <METHOD>" (publish-macro), commit and push TabularEditor-Scripts, then on any machine "update the macros" (te-macro-skills).

Don't commit, push or fork anything unless the user asks.
