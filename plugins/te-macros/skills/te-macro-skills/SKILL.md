---
name: te-macro-skills
description: Import skill-ready Tabular Editor macros from a GitHub (or local) folder into every installed Tabular Editor (TE3 and TE2), and generate (or refresh) one Claude skill per macro (from TE3, or TE2 when TE3 isn't installed or on request) that carries a "// SKILL:" header, so an agent can run those macros on a semantic model through the Tabular Editor CLI (te) or the TE2 command line. Use when the user asks to "import macros from <GitHub URL>", "install / update the macros in this repo", "generate skills from my macros", "create skills for my TE3 / TE2 macros", "refresh / sync macro skills", "update the skills after changing a macro", or after saving a new skill-ready macro in Tabular Editor.
---

# Skills from Tabular Editor macros

Each macro in `%LOCALAPPDATA%\TabularEditor3\MacroActions.json` (TE3), or `%LOCALAPPDATA%\TabularEditor\MacroActions.json` (TE2) when TE3 has no macros file or the user asks for TE2, whose code has a `// SKILL:` header becomes a skill. The header is written by the macro's author, for example:

```csharp
// SKILL: name: sum-of-columns
// SKILL: description: Create "Sum of <column>" measures ... Use when the user asks to ...
// SKILL:
// SKILL: <what the macro does>
// SKILL:
// SKILL: ## Inputs (environment variables)
// SKILL: - SUM_COLUMNS (required): 'Table'[Column] refs separated by ';'. ...
// SKILL: - SUM_NAME (optional, default "x"): ...
// SKILL:
// SKILL: ## Notes
// SKILL: ...
#if TECLI
bool cliMode = true;
#else
bool cliMode = Fx.IsCliMode();
#endif
List<Column> columns = null;
if (cliMode) { columns = Fx.GetColumnsFromEnv(Model, "SUM_COLUMNS"); }
#if !TECLI
else { columns = Selected.Columns.ToList(); }
#endif
```

In te the macro is compiled with `TECLI` defined, so it takes its inputs from environment variables. In the TE2 command line (`TabularEditor.exe <model> -S <macro>`) `Fx.IsCliMode()` detects the same at runtime. In the Tabular Editor window it keeps using the selection and dialogs.

The generated skills go into a local Claude Code plugin, **macro-skills**, in `%USERPROFILE%\.claude\macro-skills` (a local marketplace, `local-macro-skills`, holding that one plugin). Claude Code loads it in place, so the user can enable or disable all macro skills at once with `/plugin`, and the skills show up as `macro-skills:<name>`.

For each macro, the generator writes `<skills folder>\<name>\`:

- `SKILL.md`: frontmatter from `name` / `description`, the rest of the header as the body, and a standard "How to run" section (dry run, confirmation, save).
- `macro.json`: the skill name, macro name, edition (TE3 / TE2), engine (te / te2) and inputs. Its presence marks the folder as generated.
- `scripts\Invoke-MacroSkill.ps1`: the runner. It finds the macro by its `// SKILL: name:` (so renaming the macro in Tabular Editor doesn't matter), sets the inputs as environment variables, and runs the macro's live code with the skill's default engine (from `macro.json`). If that engine isn't installed it warns and uses the other one. `-Engine te|te2` forces one, and the generated `SKILL.md` tells the agent to pass it when the user asks for the other engine:
  - **te**: `te script`. The dry run prints te's change diff.
  - **te2**: `TabularEditor.exe <model> -S <macro>`. The dry run exports the model before and after the macro and prints the objects created (`+`), changed (`~`) or deleted (`-`). `-Save` writes the model back: `-B` for a `.bim`, `-TMDL` for a TMDL / PBIP folder, an incremental `SaveChanges` for a server or Power BI Desktop (`-Local` finds its port and database). TE2 doesn't validate DAX, and TabularEditor.exe is stopped after `-TimeoutSeconds` (600) in case a macro waits on a dialog.

## Importing macros from GitHub or a folder

When the user wants macros installed or updated, import them first, then generate the skills (next section). The importer is `scripts\Import-Macros.ps1` next to this file. It imports into the `MacroActions.json` of every installed edition: `%LOCALAPPDATA%\TabularEditor3` (TE3) and `%LOCALAPPDATA%\TabularEditor` (TE2), each one whose folder exists. `%LOCALAPPDATA%` is per user, so this works for any user. A missing `MacroActions.json` is created, and each file keeps its edition's format. `-MacrosFile <path>,...` imports into the given files only. Without a source it imports the published skill-ready macros (the `Advanced/Skill-Ready Macros` folder of github.com/bernatagulloesbrina/TabularEditor-Scripts); when the user gives another URL or folder, pass it as `<source>`.

A macro file is a `.csx` / `.cs` file holding the whole macro code (as copied by CopyMacro) with a `// MACRO:` header:

```csharp
// MACRO: Name: Sum of Columns
// MACRO: ValidContexts: Column
// MACRO: Tooltip: Creates a "Sum of <column>" measure for each selected column and hides the column
// SKILL: name: sum-of-columns
// SKILL: ...
```

1. **Preview**:

   ```powershell
   & "${CLAUDE_SKILL_DIR}\scripts\Import-Macros.ps1" ["<source>"] -WhatIf
   ```

   `<source>` (optional) is `https://github.com/<owner>/<repo>`, `.../tree/<branch>/<folder>`, `.../blob/<branch>/<file>`, or a local folder or file. Add `-Recurse` for subfolders and `-Only <macro or skill name>,...` to import some of them. For private repositories, set `$env:GITHUB_TOKEN` or sign in with `gh auth login`.

2. **Show the user the report** (Target, Macro, Skill, Status, Changes, File, Problems), one row per edition and macro. An existing macro is matched by its `// SKILL: name:`, or by macro name if it has none, and is then replaced: its code, and its name, contexts and tooltip from the `// MACRO:` lines. `Changes` lists what would change. Point out renames (`name (was '...')`) and anything that overwrites a macro the user edited locally. Macro code runs against their models, so for a repository the user doesn't own, say so and offer to show the code first.

3. **Close Tabular Editor** (TE3 and TE2) when the report shows changes for it. It only loads macros at startup, and saving a macro there would overwrite the imported ones. The importer refuses to write anything while an edition it would change is running (`-Force` overrides; only use it if the user insists). If the user wants to keep one open, import into the other only with `-MacrosFile`.

4. **Import**: run the same command without `-WhatIf`. Statuses are `created`, `updated`, `unchanged`, `skipped` (problem in the file; exit code 1) or `ignored` (no `// MACRO: Name:` line). Each macros file is backed up next to itself (`MacroActions.json.<timestamp>.bak`) before it is written.

5. **Generate the skills** for the imported macros that have a skill name: the workflow below, with `-Only <their skill names>`. Files without a `// SKILL:` header become plain macros.

## Generating skills

The generator is `scripts\New-MacroSkills.ps1` next to this file.

1. **Preview** what would be created or updated:

   ```powershell
   & "${CLAUDE_SKILL_DIR}\scripts\New-MacroSkills.ps1" -WhatIf
   ```

   By default it reads the TE3 macros (`%LOCALAPPDATA%\TabularEditor3\MacroActions.json`), or the TE2 ones (`%LOCALAPPDATA%\TabularEditor\MacroActions.json`) when TE3 has no macros file, and writes to the macro-skills plugin. The skills' default engine is `te` when it is installed, otherwise `TabularEditor.exe`; each run can still use the other one (see the runner above), so regenerating isn't needed to switch engines for one request. Pass `-Edition TE2` when the user asks for skills from their TE2 macros, `-Engine te2` to make TabularEditor.exe the default even though te is installed, `-SkillsDir ".\.claude\skills"` to generate plain project-level skills instead (no plugin), `-MacrosFile` for another macros file, and `-Only <name>,<name>` to limit it to some skills. All skills share one plugin folder, so generate them from one edition and engine at a time.

2. **Show the user the report** (one row per macro: Skill, Macro, Status, Inputs with `*` for required, Problems) and explain any problems (see below). Ask before generating if anything is unexpected.

3. **Generate**: run the same command without `-WhatIf`. Statuses are `created`, `updated`, `unchanged` or `skipped`. The exit code is 1 if any macro was skipped.

4. **Orphans**: a generated skill whose macro is gone (or whose `// SKILL: name:` changed) is reported as `orphaned`. Delete orphans with `-Prune` only after the user confirms.

5. **Register the plugin (first time only).** If the output warns that the macro-skills plugin is not registered, ask the user, then run the generator again with `-Register`. It adds the local marketplace and installs the plugin for the user (`claude plugin marketplace add` / `claude plugin install`).

6. Tell the user to run `/reload-plugins` (or start a new session) to load new or changed skills, and that `/plugin` enables or disables them all.

## Problems the report can show

**Skipped** (nothing written):

- `name` missing or not kebab-case (lowercase letters, digits, hyphens, max 64), or `description` missing or over 1024 characters.
- The same skill name on several macros.
- A folder with that name exists but was not generated by te-macro-skills (no `macro.json` from this generator). Hand-written skills are never overwritten; the macro author should pick another name.

**Warnings** (the skill is still written, but tell the user):

- **does not compile in te**: every macro is compiled with `te script --validate` (TECLI defined). The most common cause is helper code that uses WinForms or VisualBasic dialogs, which te doesn't have. The CLI branch must not call them, and helpers that only the interactive branch uses must sit inside `#if !TECLI`. If the macro was built with the TE-Scripting-in-Visual-Studio CopyMacro script, copy it again with a recent version, which adds those guards automatically. Then paste it into Tabular Editor and save the macro. `-NoCompileCheck` skips this check (for example, when the engine isn't installed).
- **does not compile in te2**: with the te2 engine the macro runs once on an empty model (never saved) and TabularEditor.exe's compile errors are reported. "waits on a dialog" means the macro has no `Fx.IsCliMode()` branch: make it skill-ready again with the current pattern.
- An env variable is read in the code but not listed under `## Inputs`, or listed but not read. The agent would then miss or invent an input, so the header should be fixed in the macro.
- No `#if TECLI` block (te) or no `Fx.IsCliMode()` check (te2): the macro would run its interactive code (selection, dialogs) and fail, do nothing, or wait on a dialog.
- The macro is disabled in Tabular Editor.

## Requirements

- Tabular Editor 3 or 2 with the macros saved in `MacroActions.json`.
- To run the skills: the Tabular Editor CLI (`te.exe`) on PATH or in `%LOCALAPPDATA%\Programs\TabularEditor3\`, or Tabular Editor 2 (`TabularEditor.exe`) on PATH or in `Program Files (x86)\Tabular Editor` / `Program Files\Tabular Editor`. The scripts accept `-TabularEditor <path>`.
- Windows PowerShell 5.1 or PowerShell 7.
- The Claude Code CLI (`claude`) on PATH, for `-Register`.
- For importing from GitHub: internet access. Anonymous requests are limited to 60 per hour (about one per file), so set a token for large folders.
