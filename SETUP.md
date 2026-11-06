# Setting up the macro environments

There are two environments. You need the first to use the macros, and both to write them.

| | Consumption | Authoring |
|---|---|---|
| **For** | Running the skill-ready macros, in Tabular Editor or by asking Claude | Writing and publishing macros |
| **Repos** | none (installs from this repo on GitHub) | [TE-Scripting-in-Visual-Studio](https://github.com/bernatagulloesbrina/TE-Scripting-in-Visual-Studio) + this repo, cloned side by side |
| **Software** | Tabular Editor 2, Claude Code, git (Claude Code uses it to install the plugin). Tabular Editor 3 and the `te` CLI are used too if you have them | everything on the left, plus the .NET SDK and optionally Visual Studio |

## The quick way: let Claude do it

1. Install git (`winget install --id Git.Git -e`) and [Claude Code](https://claude.com/claude-code) (`winget install --id Anthropic.ClaudeCode -e`). Open a new terminal so it picks up the new PATH, then start Claude Code with `claude`.
2. Add the plugin:

   ```
   /plugin marketplace add bernatagulloesbrina/TabularEditor-Scripts
   /plugin install te-macros@tabular-editor-scripts
   ```

3. Say **"set up my Tabular Editor environment"** (or "...including the authoring environment").

The `te-setup` skill checks the machine first. It asks you only what it can't find out itself, such as where to clone the repos or whether you work from a fork. Before anything Windows protects by default, it asks for your OK: installing software (you approve each UAC prompt on screen), changing environment variables or the PowerShell script policy, and writing Tabular Editor's macros. For authoring it also clones and builds the solution and installs CopyMacro. It gives you a checklist for what only you can do, such as starting Tabular Editor once. Ask "check my setup" at any time to re-run the check.

Setup doesn't import any macros. When it's done, close Tabular Editor and ask **"import the skill-ready macros and generate their skills"**, then run `/reload-plugins`.

## The manual way (or with another AI agent)

These steps work as instructions for any coding agent too: give it the link to this page. *Agents: ask the user before you install software, change the execution policy or environment variables, or write to Tabular Editor's macros. Ask the user to close Tabular Editor; never stop it yourself. Never work around a group policy.*

### Before you start

- **Allow PowerShell scripts.** New Windows machines block `.ps1` files ("running scripts is disabled on this system"), and the steps below run some. Allow local scripts for your account only:

  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
  ```

  If it says the setting is overridden by a policy, your IT department controls it. Ask them, or use the quick way: Claude Code isn't affected.
- **Admin rights.** winget installs show a UAC prompt you have to approve. Without admin rights, ask IT to install the software. Alternatively, extract the [portable Tabular Editor 2 zip](https://github.com/TabularEditor/TabularEditor/releases) anywhere and set `TabularEditor2Dir` (see Environment variables).

Optionally, once you have a clone of this repo (authoring step 4), you can see what's already in place and what's missing by running the read-only check in PowerShell from the clone's folder:

```powershell
& ".\plugins\te-macros\skills\te-setup\scripts\Test-TeEnvironment.ps1" -Authoring -Build   # drop -Authoring -Build to check consumption only
```

### Consumption

1. **Tabular Editor 2** (free): `winget install --id TabularEditor.TabularEditor.2 -e`. Start it once and close it.
   Tabular Editor 3 is optional. If you already have it, the macros are imported into it as well. You don't need it for anything here.
2. **te CLI** (optional, licensed separately): without it, the skills run through TabularEditor.exe (TE2) and don't validate DAX. If you have a license, download `te-win-x64.zip` from [tabulareditor.com/downloads](https://tabulareditor.com/downloads), extract it to `%LOCALAPPDATA%\Programs\TabularEditor3`, and add that folder to your **user** PATH (System Properties > Environment Variables). Preview builds expire; download a new one when they do.
3. **Claude Code** and the plugin (steps 1-2 of the quick way).
4. **Close Tabular Editor**, then ask Claude to **"import the skill-ready macros and generate their skills"**. That adds the macros to Tabular Editor and creates one skill per macro in a local `macro-skills` plugin. Run `/reload-plugins`.

### Authoring

1. **git** (`winget install --id Git.Git -e`) and the **.NET SDK** 6 or later (`winget install --id Microsoft.DotNet.SDK.8 -e`). The projects target .NET Framework 4.8, but they're SDK-style, so no Developer Pack is needed.
2. **Tabular Editor 2**, even if you use TE3: every project compiles against `TabularEditor.exe` and `TOMWrapper.dll` from its install folder.
3. **Visual Studio** (optional, for editing): Community edition with the ".NET desktop development" workload, or VS Code with C# Dev Kit, or Rider.
4. **Clone both repos into the same folder.** If you can't push to this repo, fork it and clone your fork:

   ```powershell
   New-Item -ItemType Directory -Force "$env:USERPROFILE\source\repos" | Out-Null
   cd "$env:USERPROFILE\source\repos"
   git clone https://github.com/bernatagulloesbrina/TE-Scripting-in-Visual-Studio
   git clone https://github.com/bernatagulloesbrina/TabularEditor-Scripts
   ```

5. **Build** (from the same folder as step 4): `dotnet build "TE-Scripting-in-Visual-Studio\TE Scripts\TE Scripts.csproj"`.
6. **Install CopyMacro** in Tabular Editor (save your work and close it first). This macro copies any macro from `TE Scripts.cs` to the clipboard, ready to paste and save as a macro. If you already have a macro called "Get Macro", it is replaced. Add `-WhatIf` to see what would change first:

   ```powershell
   & "TE-Scripting-in-Visual-Studio\.claude\skills\publish-macro\scripts\Install-CopyMacro.ps1"
   ```

   The repo path is written into the macro, so run it again if you move the clone.
7. Open `TE Scripting in Visual Studio.sln`. To use the make-skill-ready and publish-macro skills, start Claude Code inside the TE-Scripting-in-Visual-Studio folder.

## Environment variables

You only need these when the defaults don't fit. Set them as **user** variables, then restart Visual Studio, your terminals and Claude Code.

| Variable | Needed when | Value | Used by |
|---|---|---|---|
| `Path` | you installed the optional te CLI | add `%LOCALAPPDATA%\Programs\TabularEditor3` | the macro skills |
| `TabularEditor2Dir` | TE2 isn't in `C:\Program Files (x86)\Tabular Editor` | the folder holding `TabularEditor.exe` | the authoring build (`TE.References.props`) |
| `TE_PUBLISH_REPO` | this repo isn't cloned next to TE-Scripting-in-Visual-Studio | path of your TabularEditor-Scripts clone | publish-macro. It can also go in `TE-Scripting-in-Visual-Studio\.claude\settings.local.json` as `{ "env": { "TE_PUBLISH_REPO": "D:\\repos\\TabularEditor-Scripts" } }` |
| `GITHUB_TOKEN` | importing from a private repo, or many files at once (anonymous GitHub limit: 60 requests/hour) | a GitHub token; `gh auth login` works instead | the macro importer |

## Troubleshooting

- **"running scripts is disabled on this system"**: see [Before you start](#before-you-start).
- **winget waits and nothing happens**: a UAC prompt is waiting for you, possibly behind other windows. Approve it.
- **`/plugin marketplace add` fails**: git isn't installed, or this terminal was opened before git was installed. Install git, open a new terminal and start `claude` again.

- **Build: "Tabular Editor 2 not found in ..."**: install TE2 or set `TabularEditor2Dir`.
- **Build: references to `..\..\..\..\..\..\Program Files (x86)\...`**: your clone is older than `TE.References.props`. Pull the latest version.
- **Imported macros don't show up in Tabular Editor**: it was open during the import. Close it and import again.
- **A macro skill says "does not compile in te"**: the macro was copied with an old CopyMacro. Reinstall CopyMacro (authoring step 6), copy the macro again, save it, and regenerate the skills.
