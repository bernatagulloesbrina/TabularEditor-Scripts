#r "System.Drawing"
#r "Microsoft.VisualBasic"
using System.Windows.Forms;
using Microsoft.VisualBasic;

// MACRO: Name: Sum of Columns
// MACRO: ValidContexts: Column
// MACRO: Tooltip: Creates a "Sum of <column>" measure for each selected column and hides the column
// SKILL: name: sum-of-columns
// SKILL: description: Create "Sum of <column>" measures for numeric columns of a semantic model. Use when the user asks to create sum or base measures from columns.
// SKILL:
// SKILL: Creates one measure "Sum of <column>" per column, in the column's table and display folder, with Properties annotations.
// SKILL:
// SKILL: ## Inputs (environment variables)
// SKILL: Set these variables before running the macro. Work out each value from the conversation and the model (inspect the model if needed); if a value is unclear, ask the user instead of guessing.
// SKILL: - SUM_COLUMNS (required): 'Table'[Column] refs separated by ';'. The numeric columns to sum.
// SKILL:
// SKILL: ## Notes
// SKILL: Copies FormatString and DisplayFolder from the column and hides the base column.
// '2026-10-28 / B.Agullo / Agent ready: in CLI mode the columns come from the SUM_COLUMNS environment variable
// '2026-02-20 / B.Agullo / Added Properties annotation
// '1990-01-01 / D.Otykier
#if TECLI
var columns = Fx.GetColumnsFromEnv(Model, "SUM_COLUMNS");
if (columns == null) return;
#else
var columns = Selected.Columns.ToList();
#endif
// Loop through all target columns:
foreach(var c in columns)
{
    var newMeasure = c.Table.AddMeasure(
        "Sum of " + c.Name,                    // Name
        "SUM(" + c.DaxObjectFullName + ")",    // DAX expression
        c.DisplayFolder                        // Display Folder
    );
    // Set the format string on the new measure:
    newMeasure.FormatString = c.FormatString;
    // Provide some documentation:
    newMeasure.Description = "This measure is the sum of column " + c.DaxObjectFullName;
    newMeasure.SetAnnotation("Properties","SUM|" + c.Name + "|" + c.Table.Name); 
    newMeasure.SetAnnotation("PropertyNames","Aggregation|baseColumn|baseColumnTable"); 
    // Hide the base column:
    c.IsHidden = true;
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
    // 'Sales'[Amount] or Sales[Amount]
    public static List<Column> GetColumnsFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.AllColumns, c => new[] { c.DaxObjectFullName, c.Table.Name + "[" + c.Name + "]" }, "Column");
    }
}
