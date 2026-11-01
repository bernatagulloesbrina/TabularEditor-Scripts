#r "System.Drawing"
#r "Microsoft.VisualBasic"
using System.Windows.Forms;
using Microsoft.VisualBasic;

// MACRO: Name: Has Data Measures
// MACRO: ValidContexts: Table
// MACRO: Tooltip: Creates a "<table> has data" measure for each selected table, to show only relevant items in slicers
// SKILL: name: has-data-measures
// SKILL: description: Create "<table> has data" measures (INT(NOT ISEMPTY('Table'))) for tables of a semantic model. Use when the user asks to show only relevant items in slicers, filter slicers by a fact table, hide slicer items without data, or create "has data" measures.
// SKILL:
// SKILL: Creates one measure "<table> has data" per table, in that table, with expression INT(NOT ISEMPTY('Table')) and a description. It returns 1 when the table has visible rows in the filter context and 0 otherwise; used as a visual-level filter on a slicer (is 1) it shows only the items that have data in that table. Nothing else in the model is changed.
// SKILL:
// SKILL: ## Inputs (environment variables)
// SKILL: Set these variables before running the macro. Work out each value from the conversation and the model (inspect the model if needed); if a value is unclear, ask the user instead of guessing.
// SKILL: - HD_TABLES (required): 'Table' refs (or plain table names) separated by ';'. The tables to check for data, typically fact tables.
// SKILL:
// SKILL: ## Notes
// SKILL: The measure is named "<table> has data"; if a measure with that name already exists in the model, check with the user before running, since measure names must be unique. After running, tell the user to add the measure as a visual-level filter (is 1) on the slicers they want to restrict.
// '2026-11-01 / B.Agullo / Agent ready: in CLI mode the tables come from the HD_TABLES environment variable
// '2023-06-08 / B.Agullo / Creates a measure to show only relevant items in slicers for a fact table.
#if TECLI
var tables = Fx.GetTablesFromEnv(Model, "HD_TABLES");
if (tables == null) return;
#else
if(Selected.Tables.Count() == 0)
{
    Error("Select at least one table and try again");
    return;
}
var tables = Selected.Tables.ToList();
#endif
foreach(Table table in tables)
{
    string measureExpression = String.Format(@"INT(NOT ISEMPTY({0}))",table.DaxObjectFullName);
    string measureName = table.Name + " has data";
    string measureDescription = String.Format(@"Returns 1 if {0} has visible rows in filter context, 0 otherwise. Can be used to show only relevant slicer items.", table.DaxObjectFullName);
    Measure measure = table.AddMeasure(measureName, measureExpression);
    measure.Description = measureDescription;
}

public static class Fx
{
    public const string DefaultSentinel = "<<USE_DEFAULT>>";
    // ENV helpers: used in the #if TECLI branch of macros so agents can pass inputs through environment variables.
    // Lists are separated by ';'. All return null (after an Error) when the variable is missing or a value can't be resolved,
    // except when a defaultValue is given, which is used if the variable is missing or empty.
    // Raw string value, e.g. a name for a new measure or table
    public static string GetStringFromEnv(string envVar, string defaultValue = null)
    {
        string value = Environment.GetEnvironmentVariable(envVar);
        if (!String.IsNullOrWhiteSpace(value)) return value;
        if (defaultValue != null) return defaultValue;
        Error("Environment variable " + envVar + " is not set.");
        return null;
    }
    // Several string values separated by ';'
    public static List<string> GetStringsFromEnv(string envVar)
    {
        string value = GetStringFromEnv(envVar);
        if (value == null) return null;
        return value.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
    }
    // Resolves each ';'-separated reference against the candidates using the given name forms
    private static List<T> GetObjectsFromEnv<T>(string envVar, IEnumerable<T> candidates, Func<T, string[]> names, string kind)
    {
        List<string> references = GetStringsFromEnv(envVar);
        if (references == null) return null;
        var result = new List<T>();
        foreach (string reference in references)
        {
            T obj = candidates.FirstOrDefault(c => names(c).Contains(reference));
            if (obj == null) { Error(kind + " " + reference + " (from " + envVar + ") not found."); return null; }
            result.Add(obj);
        }
        return result;
    }
    // 'Sales' or Sales
    public static List<Table> GetTablesFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.Tables, t => new[] { t.DaxObjectFullName, t.Name }, "Table");
    }
}
