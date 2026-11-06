#r "System.Drawing"
#r "Microsoft.VisualBasic"
using System.Windows.Forms;
using Microsoft.VisualBasic;

// MACRO: Name: Create Dynamic Measure
// MACRO: ValidContexts: Measure
// MACRO: Tooltip: Creates a dynamic measure and one disconnected selection table per property, based on the Properties annotations of the selected measures
// SKILL: name: create-dynamic-measure
// SKILL: description: Create a dynamic measure that switches between measures based on slicer selections, from measures carrying Properties annotations (e.g. created by Sum of Columns or Make Measures From Functions). Use when the user asks to create a dynamic measure, a measure switcher driven by slicers, let report users pick the aggregation / base column / time intelligence variant of a measure, or build disconnected selection tables from measure properties.
// SKILL:
// SKILL: Every measure must have a "Properties" annotation with pipe-separated values (e.g. "Sum|Sales Amount"), all with the same number of values and no two measures with the same Properties. For each property position the macro creates a disconnected calculated table "Selection <property name>" with the distinct values (label column, hidden-order "<name> Order" column and "<name> Key" column), plus one partially dynamic measure per value in that table (named after the value, or "<value> (dynamic)" if the name is taken). It then creates the dynamic measure in the table of the first measure: a SWITCH over the selected key of each table that returns the matching measure, with a dynamic format string that copies each measure's format.
// SKILL:
// SKILL: ## Inputs (environment variables)
// SKILL: Set these variables before running the macro. Work out each value from the conversation and the model (inspect the model if needed); if a value is unclear, ask the user instead of guessing.
// SKILL: - DM_MEASURES (required): measure refs ([Measure] or 'Table'[Measure]) separated by ';'. At least two measures with Properties annotations; they become the options of the dynamic measure.
// SKILL: - DM_PROPERTY_NAMES (optional): names of the properties, separated by '|' (e.g. "Aggregation|Base Column"), one per Properties value. They name the selection tables ("Selection <name>") and their columns. By default the PropertyNames annotation of the measures is used (the first one if they differ), or Property1, Property2... when there is none; set it when the measures have no PropertyNames, carry different ones, or the user wants other names.
// SKILL: - DM_NAME (optional, default "Dynamic Measure" followed by the property names separated by spaces, or just "Dynamic Measure" when there are none): name of the dynamic measure. Keep the default unless the user asked for a specific name or a measure with that name already exists.
// SKILL:
// SKILL: ## Notes
// SKILL: Check the Properties annotations of the measures before running (same count, unique combinations); the macro stops with an error otherwise. The "Selection <name>" tables must not already exist, so running it twice over measures with the same property names fails. Tell the user to use the label column of each Selection table in a slicer and the dynamic measure in visuals once the model is saved and refreshed.
// 2026-11-04 / B.Agullo / Skill-ready: in CLI mode (te or TE2 -S) the inputs come from the DM_* environment variables
// 2026-05-30 / B.Agullo / FormatString values with embedded double-quotes are now escaped (doubled) when written into a DAX string literal
// 2026-05-30 / B.Agullo / Fixed missing comma between SWITCH condition and its result value in all generated DAX expressions
// 2026-05-30 / B.Agullo / PropertyNames with spaces are now sanitised (spaces removed) before use as DAX variable names; display names are unchanged
// 2026-04-23 / B.Agullo /  semidynamic measures have also propertyNames annotation
// 2026-04-22 / B.Agullo / improvements in the disconnected tables, with distinct key and display values if needed. 
// 2026-03-29 / B.Agullo / added support for propertyNames annotation
// 2025-01-07 / B.Agullo
// Creates a dynamic measure and disconnected tables based on Properties annotations of selected measures.
// Each property position gets its own disconnected table with distinct values across all measures.
// The dynamic measure returns the appropriate measure based on selected values in each disconnected table.
#if TE3
ScriptHelper.WaitFormVisible = false;
#endif
// Configuration: Prefix for disconnected tables (to group them together in Power BI Desktop)
string tablePrefix = "Selection ";
#if TECLI
bool cliMode = true;
#else
bool cliMode = Fx.IsCliMode();
#endif
List<Measure> sourceMeasures = null;
if (cliMode)
{
    sourceMeasures = Fx.GetMeasuresFromEnv(Model, "DM_MEASURES");
    if (sourceMeasures == null) return;
}
#if !TECLI
else
{
    sourceMeasures = Selected.Measures.ToList();
}
#endif
// Validate selection
if (sourceMeasures.Count < 2)
{
    Error("Select at least two measures with Properties annotations and try again.");
    return;
}
// Step 1: Extract and validate Properties annotations
var measuresWithProperties = new List<(Measure measure, List<string> properties)>();
foreach (var measure in sourceMeasures)
{
    string propertiesAnnotation = measure.GetAnnotation("Properties");
    if (string.IsNullOrEmpty(propertiesAnnotation))
    {
        Error(String.Format("Measure '{0}' does not have a Properties annotation.", measure.Name));
        return;
    }
    var properties = propertiesAnnotation.Split('|').ToList();
    measuresWithProperties.Add((measure, properties));
}
// Step 2: Validate all have the same length
int propertyCount = measuresWithProperties[0].properties.Count;
if (measuresWithProperties.Any(m => m.properties.Count != propertyCount))
{
    Error("All Properties annotations must have the same number of pipe-separated values.");
    return;
}
// Step 3: Validate all Properties annotations are unique
var allPropertiesStrings = measuresWithProperties.Select(m => string.Join("|", m.properties)).ToList();
if (allPropertiesStrings.Count != allPropertiesStrings.Distinct().Count())
{
    Error("All Properties annotations must be unique. Some measures have duplicate Properties.");
    return;
}
// Step 3.5: Retrieve PropertyNames annotations from selected measures
var propertyNamesAnnotations = new List<string>();
foreach (var measure in sourceMeasures)
{
    string propertyNamesAnnotation = measure.GetAnnotation("PropertyNames") ?? "";
    if (!string.IsNullOrEmpty(propertyNamesAnnotation))
    {
        propertyNamesAnnotations.Add(propertyNamesAnnotation);
    }
}
// Determine property names for disconnected tables and variables
List<string> propertyNames = new List<string>();
if (cliMode)
{
    // Explicit names, else the first PropertyNames annotation, else Property1, Property2...
    string cliPropertyNames = Fx.GetStringFromEnv("DM_PROPERTY_NAMES", defaultValue: "");
    if (string.IsNullOrWhiteSpace(cliPropertyNames) && propertyNamesAnnotations.Count > 0)
    {
        cliPropertyNames = propertyNamesAnnotations[0];
    }
    if (string.IsNullOrWhiteSpace(cliPropertyNames))
    {
        for (int i = 0; i < propertyCount; i++) propertyNames.Add(String.Format("Property{0}", i + 1));
    }
    else
    {
        propertyNames = cliPropertyNames.Split('|').Select(p => p.Trim()).ToList();
        if (propertyNames.Count != propertyCount)
        {
            Error(String.Format("Property names '{0}' have {1} values, but the Properties annotations have {2}. Set DM_PROPERTY_NAMES with {2} names separated by '|'.", cliPropertyNames, propertyNames.Count, propertyCount));
            return;
        }
    }
}
#if !TECLI
else if (propertyNamesAnnotations.Count > 0)
{
    // Get distinct PropertyNames values
    var distinctPropertyNames = propertyNamesAnnotations.Distinct().ToList();
    if (distinctPropertyNames.Count == 1)
    {
        // All PropertyNames are the same, use them
        propertyNames = distinctPropertyNames[0].Split('|').ToList();
    }
    else
    {
        // Different PropertyNames found, let user choose
        var optionsWithDefine = distinctPropertyNames.ToList();
        optionsWithDefine.Add("<Define now>");
        string chosenPropertyNames = Fx.ChooseString(
            OptionList: optionsWithDefine,
            label: "Different PropertyNames found. Choose one or define new:"
        );
        if (string.IsNullOrEmpty(chosenPropertyNames)) return;
        if (chosenPropertyNames == "<Define now>")
        {
            // Show input box with first PropertyNames as default
            chosenPropertyNames = Fx.GetNameFromUser(
                Prompt: "Enter property names (pipe-separated):",
                Title: "Property Names",
                DefaultResponse: distinctPropertyNames[0]
            );
            if (string.IsNullOrEmpty(chosenPropertyNames)) return;
        }
        propertyNames = chosenPropertyNames.Split('|').ToList();
    }
}
else
{
    // No PropertyNames found, ask one by one
    for (int i = 0; i < propertyCount; i++)
    {
        // Collect distinct values for this property position (up to 3)
        var sampleValues = measuresWithProperties
            .Select(m => m.properties[i])
            .Distinct()
            .Take(3)
            .ToList();
        string sampleText = sampleValues.Count > 0
            ? String.Format(" (found values: '{0}')", string.Join("', '", sampleValues))
            : "";
        string propertyName = Fx.GetNameFromUser(
            Prompt: String.Format("Enter name for property {0}{1}:", i + 1, sampleText),
            Title: "Property Name",
            DefaultResponse: String.Format("Property{0}", i + 1)
        );
        if (string.IsNullOrEmpty(propertyName)) return;
        propertyNames.Add(propertyName);
    }
}
#endif
// Step 4: Ask user for dynamic measure name with smart default based on PropertyNames
string defaultDynamicMeasureName = propertyNames.Count > 0
    ? "Dynamic Measure " + string.Join(" ", propertyNames)
    : "Dynamic Measure";
string dynamicMeasureName = null;
if (cliMode)
{
    dynamicMeasureName = Fx.GetStringFromEnv("DM_NAME", defaultValue: defaultDynamicMeasureName);
}
#if !TECLI
else
{
    dynamicMeasureName = Fx.GetNameFromUser(
        Prompt: "Enter the name for the dynamic measure:",
        Title: "Dynamic Measure Name",
        DefaultResponse: defaultDynamicMeasureName
    );
}
#endif
if (dynamicMeasureName == null) return;
// DAX variable names cannot contain spaces – strip them from propertyNames
List<string> daxSafePropertyNames = propertyNames.Select(p => p.Replace(" ", "")).ToList();
// Step 5: Get destination table (use first measure's table)
Table destinationTable = measuresWithProperties[0].measure.Table;
// Step 6: Create disconnected tables for each property position
var propertyTables = new List<Table>();
var columnBaseNames = new List<string>();
for (int i = 0; i < propertyCount; i++)
{
    // Collect distinct values for this property position across all measures
    var distinctValues = measuresWithProperties
        .Select(m => m.properties[i])
        .Distinct()
        .OrderBy(v => v)
        .ToList();
    // Create table name with prefix and column base name without prefix
    string columnBaseName = propertyNames.Count > i ? propertyNames[i] : String.Format("Property{0}", i + 1);
    string tableName = tablePrefix + columnBaseName;
    // Build DAX expression for calculated table with 3 columns (Label, Key, Order)
    string tableDax = "{" + Environment.NewLine;
    for (int j = 0; j < distinctValues.Count; j++)
    {
        tableDax += String.Format("    (\"{0}\", \"{1}\", {2})", distinctValues[j], distinctValues[j], j);
        if (j < distinctValues.Count - 1)
            tableDax += "," + Environment.NewLine;
    }
    tableDax += Environment.NewLine + "}";
    // Create the calculated table
    var propTable = Model.AddCalculatedTable(tableName, tableDax);
    // Setup columns (handle both TE2 and TE3) - use columnBaseName without prefix
    var te2 = propTable.Columns.Count == 0;
    var labelColumn = te2
        ? propTable.AddCalculatedTableColumn(columnBaseName, "[Value1]")
        : propTable.Columns["Value1"] as CalculatedTableColumn;
    var keyColumn = te2
        ? propTable.AddCalculatedTableColumn(columnBaseName + " Key", "[Value2]")
        : propTable.Columns["Value2"] as CalculatedTableColumn;
    var orderColumn = te2
        ? propTable.AddCalculatedTableColumn(columnBaseName + " Order", "[Value3]")
        : propTable.Columns["Value3"] as CalculatedTableColumn;
    if (!te2)
    {
        labelColumn.IsNameInferred = false;
        labelColumn.Name = columnBaseName;
        keyColumn.IsNameInferred = false;
        keyColumn.Name = columnBaseName + " Key";
        orderColumn.IsNameInferred = false;
        orderColumn.Name = columnBaseName + " Order";
    }
    labelColumn.SortByColumn = orderColumn;
    orderColumn.IsHidden = true;
    propertyTables.Add(propTable);
    columnBaseNames.Add(columnBaseName);
}
// Step 6.5: Create partially dynamic measures in each property table
int totalPartialMeasures = 0;
for (int tableIndex = 0; tableIndex < propertyTables.Count; tableIndex++)
{
    var currentPropertyTable = propertyTables[tableIndex];
    // Get distinct values for this property table
    var distinctValuesForThisProperty = measuresWithProperties
        .Select(m => m.properties[tableIndex])
        .Distinct()
        .ToList();
    foreach (var propertyValue in distinctValuesForThisProperty)
    {
        // Filter measures that have this property value at this position
        var matchingMeasures = measuresWithProperties
            .Where(m => m.properties[tableIndex] == propertyValue)
            .ToList();
        // Build measure expression
        string partialMeasureExpression = "";
        // Add variable declarations for OTHER properties (not current table)
        for (int i = 0; i < propertyCount; i++)
        {
            if (i == tableIndex) continue;
            string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
            partialMeasureExpression += String.Format(
                "VAR __{0} = SELECTEDVALUE( '{1}'[{2} Key] )" + Environment.NewLine,
                varName,
                propertyTables[i].Name,
                columnBaseNames[i]
            );
        }
        if (propertyCount > 1)
        {
            partialMeasureExpression += "RETURN" + Environment.NewLine;
            partialMeasureExpression += "SWITCH(" + Environment.NewLine + "    TRUE()," + Environment.NewLine;
            // Add cases only for matching measures
            foreach (var item in matchingMeasures)
            {
                string condition = "    ";
                bool firstCondition = true;
                for (int i = 0; i < propertyCount; i++)
                {
                    if (i == tableIndex) continue;
                    if (!firstCondition)
                        condition += Environment.NewLine + "        && ";
                    string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
                    condition += String.Format(
                        "__{0} = \"{1}\"",
                        varName,
                        item.properties[i]
                    );
                    firstCondition = false;
                }
                partialMeasureExpression += condition + "," + Environment.NewLine;
                partialMeasureExpression += String.Format("        {0}," + Environment.NewLine, item.measure.DaxObjectFullName);
            }
            partialMeasureExpression += "    BLANK()" + Environment.NewLine + ")";
        }
        else
        {
            // Only one property, just return the measure directly
            partialMeasureExpression += "RETURN" + Environment.NewLine;
            partialMeasureExpression += matchingMeasures[0].measure.DaxObjectFullName;
        }
        // Check if measure name already exists and add suffix if needed
        string partialMeasureName = propertyValue;
        if (Model.AllMeasures.Any(m => m.Name == propertyValue))
        {
            partialMeasureName = propertyValue + " (dynamic)";
        }
        // Create the measure
        Measure partialMeasure = currentPropertyTable.AddMeasure(partialMeasureName, partialMeasureExpression);
        partialMeasure.FormatDax();
        partialMeasure.SetAnnotation("Properties", propertyValue);
        partialMeasure.SetAnnotation("PropertyNames", propertyNames.Count > tableIndex ? propertyNames[tableIndex] : String.Format("Property{0}", tableIndex + 1));
        // Build format string expression
        string partialFormatExpression = "";
        // Add variable declarations for OTHER properties (not current table)
        for (int i = 0; i < propertyCount; i++)
        {
            if (i == tableIndex) continue;
            string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
            partialFormatExpression += String.Format(
                "VAR __{0} = SELECTEDVALUE( '{1}'[{2} Key] )" + Environment.NewLine,
                varName,
                propertyTables[i].Name,
                columnBaseNames[i]
            );
        }
        if (propertyCount > 1)
        {
            partialFormatExpression += "RETURN" + Environment.NewLine;
            partialFormatExpression += "SWITCH(" + Environment.NewLine + "    TRUE()," + Environment.NewLine;
            // Add cases only for matching measures
            foreach (var item in matchingMeasures)
            {
                string condition = "    ";
                bool firstCondition = true;
                for (int i = 0; i < propertyCount; i++)
                {
                    if (i == tableIndex) continue;
                    if (!firstCondition)
                        condition += Environment.NewLine + "        && ";
                    string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
                    condition += String.Format(
                        "__{0} = \"{1}\"",
                        varName,
                        item.properties[i]
                    );
                    firstCondition = false;
                }
                partialFormatExpression += condition + "," + Environment.NewLine;
                // Use format string or format string expression
                string formatValue = !string.IsNullOrEmpty(item.measure.FormatStringExpression)
                    ? item.measure.FormatStringExpression
                    : String.Format("\"{0}\"", item.measure.FormatString.Replace("\"", "\"\""));
                partialFormatExpression += String.Format("        {0}," + Environment.NewLine, formatValue);
            }
            partialFormatExpression += "    \"\"" + Environment.NewLine + ")";
        }
        else
        {
            // Only one property, use the measure's format directly
            partialFormatExpression += "RETURN" + Environment.NewLine;
            string formatValue = !string.IsNullOrEmpty(matchingMeasures[0].measure.FormatStringExpression)
                ? matchingMeasures[0].measure.FormatStringExpression
                : String.Format("\"{0}\"", matchingMeasures[0].measure.FormatString.Replace("\"", "\"\""));
            partialFormatExpression += formatValue;
        }
        // Set format string expression
        partialMeasure.FormatStringExpression = partialFormatExpression;
        totalPartialMeasures++;
    }
}
// Step 7: Build dynamic measure expression
string measureExpression = "";
// Add variable declarations for each property
for (int i = 0; i < propertyCount; i++)
{
    string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
    measureExpression += String.Format(
        "VAR __{0} = SELECTEDVALUE( '{1}'[{2} Key] )" + Environment.NewLine,
        varName,
        propertyTables[i].Name,
        columnBaseNames[i]
    );
}
measureExpression += "RETURN" + Environment.NewLine;
measureExpression += "SWITCH(" + Environment.NewLine + "    TRUE()," + Environment.NewLine;
foreach (var item in measuresWithProperties)
{
    // Build condition for this measure with variables
    string condition = "    ";
    for (int i = 0; i < propertyCount; i++)
    {
        if (i > 0)
            condition += Environment.NewLine + "        && ";
        string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
        condition += String.Format(
            "__{0} = \"{1}\"",
            varName,
            item.properties[i]
        );
    }
    measureExpression += condition + "," + Environment.NewLine;
    measureExpression += String.Format("        {0}," + Environment.NewLine, item.measure.DaxObjectFullName);
}
measureExpression += "    BLANK()" + Environment.NewLine + ")";
// Step 8: Create the dynamic measure
Measure dynamicMeasure = destinationTable.AddMeasure(dynamicMeasureName, measureExpression);
dynamicMeasure.FormatDax();
// Step 9: Build dynamic format string expression
string formatExpression = "";
// Add variable declarations for each property
for (int i = 0; i < propertyCount; i++)
{
    string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
    formatExpression += String.Format(
        "VAR __{0} = SELECTEDVALUE( '{1}'[{2} Key] )" + Environment.NewLine,
        varName,
        propertyTables[i].Name,
        columnBaseNames[i]
    );
}
formatExpression += "RETURN" + Environment.NewLine;
formatExpression += "SWITCH(" + Environment.NewLine + "    TRUE()," + Environment.NewLine;
foreach (var item in measuresWithProperties)
{
    // Build condition for this measure with variables
    string condition = "    ";
    for (int i = 0; i < propertyCount; i++)
    {
        if (i > 0)
            condition += Environment.NewLine + "        && ";
        string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
        condition += String.Format(
            "__{0} = \"{1}\"",
            varName,
            item.properties[i]
        );
    }
    formatExpression += condition + "," + Environment.NewLine;
    // Use format string or format string expression
    string formatValue = !string.IsNullOrEmpty(item.measure.FormatStringExpression)
        ? item.measure.FormatStringExpression
        : String.Format("\"{0}\"", item.measure.FormatString.Replace("\"", "\"\""));
    formatExpression += String.Format("        {0}," + Environment.NewLine, formatValue);
}
formatExpression += "    \"\"" + Environment.NewLine + ")";
// Set format string expression
dynamicMeasure.FormatStringExpression = formatExpression;
Output(String.Format(
    "Created dynamic measure '{0}' with {1} property tables, {2} partial measures, and {3} source measures.",
    dynamicMeasureName,
    propertyTables.Count,
    totalPartialMeasures,
    measuresWithProperties.Count
));

public static class Fx
{
#if !TECLI
    // In TE2 scripts run inside an "Hourglass" that sets Application.UseWaitCursor = true for the
    // whole application, so dialog boxes show a busy cursor (TE3 solves this with WaitFormVisible).
    // Turn the flag off and nudge Windows (WM_SETCURSOR) to refresh the cursor immediately,
    // otherwise it only updates on the next mouse message. Call this before showing any dialog.
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);
    public static void ResetWaitCursor()
    {
        if (!Application.UseWaitCursor) return;
        Application.UseWaitCursor = false;
        Cursor.Current = Cursors.Default;
        Form f = Form.ActiveForm;
        if (f != null && f.Handle != IntPtr.Zero)
            SendMessage(f.Handle, 0x20, f.Handle, (IntPtr)1); // WM_SETCURSOR
    }
    public static string GetNameFromUser(string Prompt, string Title, string DefaultResponse)
    {
        ResetWaitCursor();
        string response = Interaction.InputBox(Prompt, Title, DefaultResponse, 740, 400);
        if (response == null)
        {
            Error("No input provided.");
            return null;
        };
        return response;
    }
#endif
    public const string DefaultSentinel = "<<USE_DEFAULT>>";
#if !TECLI
    public static string ChooseString(IList<string> OptionList, string label = "Choose item", int customWidth = 400, int customHeight = 500)
    {
        return ChooseStringInternal(OptionList, MultiSelect: false, label: label, customWidth: customWidth, customHeight:customHeight) as string;
    }
    private static object ChooseStringInternal(IList<string> OptionList, bool MultiSelect, string label = "Choose item(s)", int customWidth = 400, int customHeight = 500, string defaultValue = null, string description = null)
    {
        Form form = new Form
        {
            Text =label,
            StartPosition = FormStartPosition.CenterScreen,
            Padding = new Padding(20)
        };
        // Keep track of selected items across filtering
        HashSet<string> selectedItemsSet = new HashSet<string>();
        bool isRestoringSelections = false;
        // Search panel at the top
        Panel searchPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 35,
            Padding = new Padding(0, 0, 0, 5)  // Add 5px bottom padding for spacing
        };
        Label searchLabel = new System.Windows.Forms.Label
        {
            Text = "Search:",
            AutoSize = true,
            Location = new System.Drawing.Point(0, 8),
            Width = 70
        };
        TextBox searchBox = new TextBox
        {
            Location = new System.Drawing.Point(75, 5),
            Width = customWidth - 115,
            Anchor = AnchorStyles.Left | AnchorStyles.Right
        };
        searchPanel.Controls.Add(searchLabel);
        searchPanel.Controls.Add(searchBox);
        // Optional info panel: parameter description and "use default" checkbox
        CheckBox useDefaultCheck = null;
        Panel infoPanel = null;
        int infoPanelHeight = 0;
        if (!string.IsNullOrEmpty(description) || !string.IsNullOrEmpty(defaultValue))
        {
            infoPanel = new Panel { Dock = DockStyle.Top, Padding = new Padding(0, 0, 0, 5) };
            int infoY = 0;
            if (!string.IsNullOrEmpty(description))
            {
                Label descLabel = new System.Windows.Forms.Label
                {
                    Text = description,
                    Location = new System.Drawing.Point(0, infoY),
                    Width = customWidth - 60,
                    Height = 36,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
                };
                infoPanel.Controls.Add(descLabel);
                infoY += 40;
                infoPanelHeight += 40;
            }
            if (!string.IsNullOrEmpty(defaultValue))
            {
                useDefaultCheck = new CheckBox
                {
                    Text = String.Format("Use default ({0})", defaultValue),
                    Location = new System.Drawing.Point(0, infoY),
                    AutoSize = true
                };
                infoPanel.Controls.Add(useDefaultCheck);
                infoPanelHeight += 30;
            }
            infoPanel.Height = infoPanelHeight;
        }
        // ListBox panel in the middle
        Panel listBoxPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 35 + infoPanelHeight, 0, 75)  // Top padding = info + searchPanel height, Bottom = buttonPanel height
        };
        ListBox listbox = new ListBox
        {
            Dock = DockStyle.Fill,
            SelectionMode = MultiSelect ? SelectionMode.MultiExtended : SelectionMode.One
        };
        listBoxPanel.Controls.Add(listbox);
        // Initial population
        listbox.Items.AddRange(OptionList.ToArray());
        if (!MultiSelect && OptionList.Count > 0)
            listbox.SelectedItem = OptionList[0];
        // Track manual selection changes
        listbox.SelectedIndexChanged += delegate
        {
            // Skip if we're programmatically restoring selections
            if (isRestoringSelections) return;
            // Get current items in listbox
            HashSet<string> currentItems = new HashSet<string>();
            foreach (object item in listbox.Items)
            {
                currentItems.Add(item.ToString());
            }
            // Remove only currently visible items that are NOT selected
            foreach (string item in currentItems)
            {
                if (!listbox.SelectedItems.Cast<object>().Any(selected => selected.ToString() == item))
                {
                    selectedItemsSet.Remove(item);
                }
            }
            // Add currently selected items
            foreach (object item in listbox.SelectedItems)
            {
                selectedItemsSet.Add(item.ToString());
            }
        };
        // Search/filter functionality
        searchBox.TextChanged += delegate
        {
            string searchText = searchBox.Text;
            // Filter the list
            var filteredList = OptionList
                .Where(item => item.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();
            // Set flag to prevent SelectedIndexChanged from firing during restoration
            isRestoringSelections = true;
            // Repopulate listbox
            listbox.Items.Clear();
            listbox.Items.AddRange(filteredList.ToArray());
            // Restore previous selections
            for (int i = 0; i < listbox.Items.Count; i++)
            {
                if (selectedItemsSet.Contains(listbox.Items[i].ToString()))
                {
                    listbox.SetSelected(i, true);
                }
            }
            // Re-enable selection tracking
            isRestoringSelections = false;
        };
        FlowLayoutPanel buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 75,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(10, 5, 10, 10)  // Add top padding for spacing from listbox
        };
        Button selectAllButton = new Button { Text = "Select All", Visible = MultiSelect , Height = 50, Width = 150};
        Button selectNoneButton = new Button { Text = "Select None", Visible = MultiSelect, Height = 50, Width = 150 };
        Button okButton = new Button { Text = "OK", DialogResult = DialogResult.OK, Height = 50, Width = 100 };
        Button cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Height = 50, Width = 100 };
        selectAllButton.Click += delegate
        {
            for (int i = 0; i < listbox.Items.Count; i++)
                listbox.SetSelected(i, true);
            // Update tracking set with all currently visible items
            foreach (object item in listbox.Items)
            {
                selectedItemsSet.Add(item.ToString());
            }
        };
        selectNoneButton.Click += delegate
        {
            for (int i = 0; i < listbox.Items.Count; i++)
                listbox.SetSelected(i, false);
            // Remove all currently visible items from tracking set
            foreach (object item in listbox.Items)
            {
                selectedItemsSet.Remove(item.ToString());
            }
        };
        buttonPanel.Controls.Add(selectAllButton);
        buttonPanel.Controls.Add(selectNoneButton);
        buttonPanel.Controls.Add(okButton);
        buttonPanel.Controls.Add(cancelButton);
        // Add controls in proper order for Dock: Bottom first, Top second, Fill last
        form.Controls.Add(buttonPanel);    // Bottom - add first
        form.Controls.Add(searchPanel);    // Top - add second
        if (infoPanel != null) form.Controls.Add(infoPanel);  // Top - sits above search panel
        form.Controls.Add(listBoxPanel);   // Fill - add last
        // When "use default" is checked, disable the list/search to make clear the default will be used
        if (useDefaultCheck != null)
        {
            useDefaultCheck.CheckedChanged += delegate
            {
                bool useDefault = useDefaultCheck.Checked;
                listbox.Enabled = !useDefault;
                searchBox.Enabled = !useDefault;
                selectAllButton.Enabled = !useDefault;
                selectNoneButton.Enabled = !useDefault;
            };
        }
        form.Width = customWidth;
        form.Height = customHeight + infoPanelHeight;
        ResetWaitCursor();
        DialogResult result = form.ShowDialog();
        if (result == DialogResult.Cancel)
        {
            Info("You Cancelled!");
            return null;
        }
        // Default checkbox checked: signal default usage to the caller
        if (useDefaultCheck != null && useDefaultCheck.Checked)
        {
            return MultiSelect ? (object)new List<string>() { DefaultSentinel } : (object)DefaultSentinel;
        }
        if (MultiSelect)
        {
            List<string> selectedItems = new List<string>();
            foreach (object item in listbox.SelectedItems)
                selectedItems.Add(item.ToString());
            return selectedItems;
        }
        else
        {
            return listbox.SelectedItem != null ? listbox.SelectedItem.ToString() : null;
        }
    }
    // True when TE2 runs the script from the command line (-S / -SCRIPT). TE2 doesn't define TECLI and shows
    // Select*/Choose* dialogs modally, which hangs an unattended run. Macros call it in the #else branch of #if TECLI
    // (te is always CLI mode), so this method has no #if of its own.
    public static bool IsCliMode()
    {
        return Environment.GetCommandLineArgs().Any(a =>
            a.Equals("-S", StringComparison.OrdinalIgnoreCase) || a.Equals("-SCRIPT", StringComparison.OrdinalIgnoreCase));
    }
#endif
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
    // [Sales Amount], Sales Amount or 'Sales'[Sales Amount]
    public static List<Measure> GetMeasuresFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.AllMeasures, m => new[] { m.DaxObjectFullName, m.Name, m.Table.DaxObjectFullName + m.DaxObjectFullName }, "Measure");
    }
}
