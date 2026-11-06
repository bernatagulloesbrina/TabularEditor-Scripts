---
name: te-setup
description: >-
  Set up, from scratch or to repair, the Tabular Editor macro environments on a Windows machine, up to the point where the user can import the macros and generate their skills (it doesn't import or generate them) - the consumption environment (what's needed to import skill-ready macros from TabularEditor-Scripts into Tabular Editor and run them as Claude skills through te / TE2) and the authoring environment (TE-Scripting-in-Visual-Studio: write macros in Visual Studio, build with dotnet, copy them with CopyMacro, publish them to TabularEditor-Scripts). Checks prerequisites (Tabular Editor 3/2, te CLI, .NET SDK, git, repos, environment variables, plugins), does every step it can, and tells the user exactly what only they can do. Use when the user asks to "set up / install / configure the environment", "set up TE-Scripting-in-Visual-Studio", "prepare a new machine for the macros", "onboard a colleague", "why doesn't the solution build", "fix the TabularEditor references", or "check my setup".
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

### Ask what you can't check

Check first and ask afterwards. Only ask what the check can't answer, and only when the answer changes what you do. Put the questions in **one** AskUserQuestion call (up to 4 questions), with the default first and marked "(Recommended)". Never ask about something the check already answered.

| Ask | When | Why |
|---|---|---|
| Where to clone the repos (default `%USERPROFILE%\source\repos`) | authoring, and a repo is `missing` | the default folder may not exist, or the user keeps repos elsewhere |
| Push rights to bernatagulloesbrina/TabularEditor-Scripts, or work from a fork | authoring, and the publication repo is `missing` | without push rights they clone their fork and open pull requests |
| Who installs the software: the user approves the UAC prompts, or IT does it | `Admin rights` is `warning` (standard user), or the machine is managed | a standard user can't approve the UAC prompts. Offer the portable Tabular Editor 2 zip plus `TabularEditor2Dir` |
| Whether the te CLI is licensed | only if the user brings it up | it's optional: mention it once, don't ask by default |
| Whether to replace an existing CopyMacro | Install-CopyMacro `-WhatIf` would replace a macro the user wrote | the macro is overwritten |

### Things Windows blocks by default: always confirm first

Windows blocks or protects each of these by default. Don't change any of them without the user's explicit OK, even when the check marks it `claude`. Explain in one line what it is, and wait for the answer. Ask for all of them in the same confirmation as the Step 2 actions. If a group policy or IT rule blocks something, report it and stop. Never work around it.

| What | Check row | What to say and do |
|---|---|---|
| **Script execution policy** | `Script execution policy` | New Windows machines are `Restricted`: the user can't run `.ps1` files in their own PowerShell. Claude Code's PowerShell bypasses this, so it only matters for SETUP.md's manual steps. Propose `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force` (this user only; local scripts run, downloaded scripts must be signed). If it's set by group policy, only IT can change it. |
| **Installs that need admin (UAC)** | `Admin rights`, every winget row | Before each winget install, tell the user that a UAC prompt will appear and they have to approve it **on their screen**. Until they do, winget waits. A standard user needs IT to approve. |
| **Environment variables and the user PATH** | `TE2 build references`, te CLI | Show the exact name, value and scope before writing it. Read the user PATH, append, write it back. Never write the machine PATH. |
| **Writing Tabular Editor's MacroActions.json or TE2's Preferences.json** | `Tabular Editor open`, CopyMacro rows, `TE2 C# compiler` | Ask the user to save and close Tabular Editor. **Never stop the process yourself**: they would lose unsaved work. If it's open when the file is written, it overwrites the file on exit. |
| **Downloaded files** (te CLI zip, portable TE2 zip) | - | Windows marks downloaded files as coming from the internet, and SmartScreen can block the extracted `.exe`. Only if it's blocked, propose `Unblock-File` on the zip **before** extracting it. Ask first, and only for a file the user downloaded from the official site. |
| **Network and corporate proxies** | `Build`, plugin install | If a clone, a NuGet restore or a winget install fails with a network or certificate error, don't retry in a loop or turn off TLS checks. Ask the user about a proxy or a blocked site. |

## Step 2 - Do it, in this order

Group the actions that need the user's go-ahead (installs, clones, the execution policy, environment variables, writing to Tabular Editor's macros) into one confirmation, listing the exact commands, instead of asking one by one. If the user turns one down, skip it and anything that depends on it, and list it as still to do in the report. Items marked `user` go in a short numbered checklist; wait for the user to say they're done, then re-run the check.

### Software (consumption and authoring)

| Need | Who | How |
|---|---|---|
| Tabular Editor 3 | - | **Optional. Never install it or ask the user to install it.** If it's installed, it's used too (macros are imported into it as well, skills are generated from its macros). If it's installed but `%LOCALAPPDATA%\TabularEditor3` doesn't exist yet, it just hasn't been started; that's fine, TE2 covers everything. |
| Tabular Editor 2 | claude | `winget install --id TabularEditor.TabularEditor.2 -e` (free). Required when TE3 isn't installed, and always for authoring (every project references `TabularEditor.exe` and `TOMWrapper.dll`). Then the user starts it once and closes it, which creates `%LOCALAPPDATA%\TabularEditor`, where macros are imported. |
| te CLI | - | **Optional and licensed separately: mention it once, never wait for it or block on it.** If it's found, the skills run with it (with DAX validation). If not, tell the user it exists (download from https://tabulareditor.com/downloads if they have a license) and carry on: the skills run through TabularEditor.exe (TE2). If the user installs it later: extract `te-win-x64.zip` to `%LOCALAPPDATA%\Programs\TabularEditor3` and add that folder to the **user** PATH (below). Preview builds expire: the check mentions it 14 days ahead. |
| git | claude | `winget install --id Git.Git -e`. Consumption too: Claude Code adds and updates the te-macros plugin marketplace with git |
| .NET SDK 6 or later | claude | `winget install --id Microsoft.DotNet.SDK.8 -e`. **Consumption, when TE2 is installed and Visual Studio isn't**: it provides the Roslyn compiler TE2 needs (below). **Authoring**: it's all the build needs: every project is SDK-style and targets net48, so the .NET Framework reference assemblies come from NuGet (no Developer Pack). |
| TE2 C# compiler | claude | Row `TE2 C# compiler`. TE2 compiles macros with the .NET Framework's C# 5 compiler unless it's given a Roslyn one, and compiles **all** macros together: one macro with modern C# (`$"..."`, tuples, `??=`, e.g. Make Measures From Functions) and **every** macro fails to load ("Unexpected character '$'", "Type expected"), and those macros can't run through TabularEditor.exe either. With the .NET SDK (or Visual Studio) installed and TE2 closed, run `& "${CLAUDE_SKILL_DIR}\scripts\Set-Te2Compiler.ps1" -WhatIf`, then without `-WhatIf`. It writes TE2's Preferences.json (File > Preferences > General > Compiler path / options): ask first, like the other Tabular Editor files. It points at a versioned SDK folder: if that SDK version is uninstalled later, TE2 silently goes back to C# 5, so re-run it when the check flags the row again. |
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
2. **Ready check**: at least one edition whose settings folder exists (`%LOCALAPPDATA%\TabularEditor3` or `%LOCALAPPDATA%\TabularEditor`, created by starting it once), an engine to run the skills (te or TabularEditor.exe), TE2's Roslyn compiler when TE2 is installed, and `claude` on PATH (the generator uses it to register the macro-skills plugin). The `macro-skills plugin` row is `info` until then; that's expected, not a gap to fix.
3. **Next steps for the user**: close Tabular Editor, then ask "import the skill-ready macros and generate their skills" (te-macro-skills does the preview, the import, the skill generation and the first-time `-Register`), then `/reload-plugins`.

### Authoring

1. **Clone both repos side by side** (ask where; default `%USERPROFILE%\source\repos`). Contributors without push rights to bernatagulloesbrina's repos fork TabularEditor-Scripts first and clone the fork (publish-macro writes there; they open pull requests). Create the folder if it doesn't exist. Without Visual Studio, `source\repos` usually doesn't.

   ```powershell
   New-Item -ItemType Directory -Force "<root>" | Out-Null
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
