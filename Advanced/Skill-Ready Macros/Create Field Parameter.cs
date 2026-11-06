#r "Microsoft.VisualBasic"
#r "System.Drawing"
using System.Windows.Forms;
using Microsoft.VisualBasic;

// MACRO: Name: Create Field Parameter
// MACRO: ValidContexts: Measure, Column
// MACRO: Tooltip: Creates a field parameter table with the selected measures or columns, adding a column per Properties annotation
// SKILL: name: create-field-parameter
// SKILL: description: Create a field parameter table from measures or columns of a semantic model. Use when the user asks to create a field parameter, let report users switch the measure or column shown in a visual, build a measure/column selector or a dynamic axis.
// SKILL:
// SKILL: Creates a calculated table that works as a Power BI field parameter, with one row per measure (or column): a name column (with the parameter name), a hidden "<name> Fields" column and a hidden "<name> Order" column. If the objects have a "Properties" annotation (pipe-separated values, e.g. from the Sum of Columns or Make Measures From Functions macros), each value becomes an extra column, named after the "PropertyNames" annotation.
// SKILL:
// SKILL: ## Inputs (environment variables)
// SKILL: Set these variables before running the macro. Work out each value from the conversation and the model (inspect the model if needed); if a value is unclear, ask the user instead of guessing.
// SKILL: Set exactly one of FP_MEASURES (measure field parameter) or FP_COLUMNS (column field parameter); the macro fails if neither is set.
// SKILL: - FP_MEASURES (optional, set it unless FP_COLUMNS is set): measure refs ([Measure] or 'Table'[Measure]) separated by ';'. The measures of the field parameter, in this order.
// SKILL: - FP_COLUMNS (optional, set it instead of FP_MEASURES for a column field parameter): 'Table'[Column] refs separated by ';'. When set, FP_MEASURES is ignored.
// SKILL: - FP_ORDER_BY_NAME (optional, default "false"): true to sort the items by name instead of keeping the given order.
// SKILL: - FP_PROPERTY_NAMES (optional): names of the property columns, separated by '|' (e.g. "Aggregation|Base Column|Table"). Only used when the objects have Properties annotations. By default the PropertyNames annotation of the objects is used (the first one if they differ), or Property1, Property2... when there is none; set it when the user wants other column names or the objects carry different PropertyNames.
// SKILL: - FP_NAME (optional, default "Parameter" followed by the property names separated by spaces, or just "Parameter" when there are no properties): name of the field parameter table and of its main column. Keep the default unless the user asked for a specific name or a table with that name already exists (then pick a descriptive unique one, e.g. "Measure Selector").
// SKILL:
// SKILL: ## Notes
// SKILL: Against Power BI Desktop, unsupported features must be enabled in Tabular Editor. The table must not already exist. Tell the user the field parameter can be used in visuals and slicers once the model is saved and refreshed.
// 2026-11-04 / B.Agullo / Skill-ready: in CLI mode (te or TE2 -S) the inputs come from the FP_* environment variables
// 2026-05-13 / B.Agullo / fixed bug to execute in TE2
//2026-04-23 / B.Agullo / improvements to the default parameter name suggestion logic (based on property names if available) 
//2026-03-29 / B.Agullo / added support for propertyNames annotation
//80% comes from Daniel Otykier --> https://github.com/TabularEditor/TabularEditor3/issues/541#issuecomment-1129228481
//20% B.Agullo --> pop-up to choose parameter name + all the properties annotation support (split by pipe delimiter and added as additional columns in the field parameter table)
// Before running the script, select the measures or columns that you
// would like to use as field parameters (hold down CTRL to select multiple
// objects). Also, you may change the name of the field parameter table
// below. NOTE: If used against Power BI Desktop, you must enable unsupported
// features under File > Preferences (TE2) or Tools > Preferences (TE3).
#if TE3
ScriptHelper.WaitFormVisible = false;
#endif
#if TECLI
bool cliMode = true;
#else
bool cliMode = Fx.IsCliMode();
#endif
// Get objects collection
IEnumerable<ITabularTableObject> objects = null;
bool orderByName = false;
if (cliMode)
{
    if (!String.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("FP_COLUMNS")))
    {
        List<Column> fpColumns = Fx.GetColumnsFromEnv(Model, "FP_COLUMNS");
        if (fpColumns == null) return;
        objects = fpColumns.Cast<ITabularTableObject>().ToList();
    }
    else
    {
        List<Measure> fpMeasures = Fx.GetMeasuresFromEnv(Model, "FP_MEASURES");
        if (fpMeasures == null) return;
        objects = fpMeasures.Cast<ITabularTableObject>().ToList();
    }
    if (!objects.Any()) { Error("FP_MEASURES / FP_COLUMNS contain no objects."); return; }
    bool? orderByNameEnv = Fx.GetBoolFromEnv("FP_ORDER_BY_NAME", defaultValue: false);
    if (orderByNameEnv == null) return;
    orderByName = orderByNameEnv.Value;
}
#if !TECLI
else
{
    if (Selected.Columns.Count == 0 && Selected.Measures.Count == 0) throw new Exception("No columns or measures selected!");
    objects = Selected.Columns.Any() ? Selected.Columns.Cast<ITabularTableObject>() : Selected.Measures;
    orderByName = Fx.IsAnswerYes("Do you want to order measures/columns by name?");
}
#endif
if (orderByName)
{
    objects = objects.OrderBy(o => o.Name).ToList();
}
// In CLI mode the property column names can be given explicitly (pipe-separated)
string cliPropertyNames = null;
if (cliMode)
{
    cliPropertyNames = Fx.GetStringFromEnv("FP_PROPERTY_NAMES", defaultValue: "");
}
// Retrieve Properties annotations and split them by pipe delimiter
var objectPropertiesLists = new List<List<string>>();
foreach (var obj in objects)
{
    string propertiesAnnotation = "";
    if (obj is Measure)
    {
        propertiesAnnotation = ((Measure)obj).GetAnnotation("Properties") ?? "";
    }
    else if (obj is Column)
    {
        propertiesAnnotation = ((Column)obj).GetAnnotation("Properties") ?? "";
    }
    var propertiesList = string.IsNullOrEmpty(propertiesAnnotation)
        ? new List<string>()
        : propertiesAnnotation.Split('|').ToList();
    objectPropertiesLists.Add(propertiesList);
}
// Find maximum length to normalize all lists
int maxPropertiesCount = objectPropertiesLists.Any() ? objectPropertiesLists.Max(pl => pl.Count) : 0;
// Normalize all lists to the same length by adding empty strings
for (int i = 0; i < objectPropertiesLists.Count; i++)
{
    while (objectPropertiesLists[i].Count < maxPropertiesCount)
    {
        objectPropertiesLists[i].Add("");
    }
}
// Retrieve PropertyNames annotations from selected objects
var propertyNamesAnnotations = new List<string>();
foreach (var obj in objects)
{
    string propertyNamesAnnotation = "";
    if (obj is Measure)
    {
        propertyNamesAnnotation = ((Measure)obj).GetAnnotation("PropertyNames") ?? "";
    }
    else if (obj is Column)
    {
        propertyNamesAnnotation = ((Column)obj).GetAnnotation("PropertyNames") ?? "";
    }
    if (!string.IsNullOrEmpty(propertyNamesAnnotation))
    {
        propertyNamesAnnotations.Add(propertyNamesAnnotation);
    }
}
// Determine property names for field parameter columns
List<string> propertyNames = new List<string>();
if (cliMode)
{
    var distinctPropertyNames = propertyNamesAnnotations.Distinct().ToList();
    if (maxPropertiesCount > 0 && !string.IsNullOrWhiteSpace(cliPropertyNames))
    {
        propertyNames = cliPropertyNames.Split('|').Select(p => p.Trim()).ToList();
    }
    else if (distinctPropertyNames.Count > 0)
    {
        // Same default as the interactive dialogs: the (first) PropertyNames annotation
        propertyNames = distinctPropertyNames[0].Split('|').ToList();
    }
    else
    {
        for (int i = 0; i < maxPropertiesCount; i++)
        {
            propertyNames.Add(String.Format("Property{0}", i + 1));
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
    for (int i = 0; i < maxPropertiesCount; i++)
    {
        // Collect distinct values for this property position (up to 3)
        var sampleValues = objectPropertiesLists
            .Where(pl => pl.Count > i && !string.IsNullOrEmpty(pl[i]))
            .Select(pl => pl[i])
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
// Build default parameter name based on PropertyNames
string defaultParameterName = propertyNames.Count > 0
    ? "Parameter " + string.Join(" ", propertyNames)
    : "Parameter";
// Ask user for parameter name with smart default
string name = null;
if (cliMode)
{
    name = Fx.GetStringFromEnv("FP_NAME", defaultValue: defaultParameterName);
}
#if !TECLI
else
{
    name = Fx.GetNameFromUser("Provide the name for the field parameter", "Parameter name", defaultParameterName);
}
#endif
if (string.IsNullOrEmpty(name)) return;
if (cliMode && Model.Tables.Any(t => t.Name == name)) { Error("A table named '" + name + "' already exists. Set FP_NAME to another name."); return; }
// Construct the DAX for the calculated table based on the current selection:
string dax = "";
if (maxPropertiesCount > 0)
{
    // Build DAX with additional property columns
    var daxRows = objects.Select((c, i) =>
    {
        var baseValues = string.Format("\"{0}\", NAMEOF('{1}'[{0}]), {2}", c.Name, c.Table.Name, i);
        var propertyValues = string.Join(", ", objectPropertiesLists[i].Select(p => string.Format("\"{0}\"", p)));
        return string.Format("({0}, {1})", baseValues, propertyValues);
    });
    dax = "{\n    " + string.Join(",\n    ", daxRows) + "\n}";
}
else
{
    // No properties found, use original format
    dax = "{\n    " + string.Join(",\n    ", objects.Select((c, i) => string.Format("(\"{0}\", NAMEOF('{1}'[{0}]), {2})", c.Name, c.Table.Name, i))) + "\n}";
}
// Add the calculated table to the model:
var table = Model.AddCalculatedTable(name, dax);
// In TE2 columns are not created automatically from a DAX expression, so 
// we will have to add them manually:
var te2 = table.Columns.Count == 0;
var nameColumn = te2 ? table.AddCalculatedTableColumn(name, "[Value1]") : table.Columns["Value1"] as CalculatedTableColumn;
var fieldColumn = te2 ? table.AddCalculatedTableColumn(name + " Fields", "[Value2]") : table.Columns["Value2"] as CalculatedTableColumn;
var orderColumn = te2 ? table.AddCalculatedTableColumn(name + " Order", "[Value3]") : table.Columns["Value3"] as CalculatedTableColumn;
// Create additional columns for each property level
var propertyColumns = new List<CalculatedTableColumn>();
for (int i = 0; i < maxPropertiesCount; i++)
{
    int valueIndex = 4 + i; // Value4, Value5, etc.
    string columnName = propertyNames.Count > i ? propertyNames[i] : String.Format("Property{0}", i + 1);
    var propColumn = te2
        ? table.AddCalculatedTableColumn(columnName, "[Value" + valueIndex + "]")
        : table.Columns["Value" + valueIndex] as CalculatedTableColumn;
    if (!te2)
    {
        propColumn.IsNameInferred = false;
        propColumn.Name = columnName;
    }
    propertyColumns.Add(propColumn);
}
if (!te2)
{
    // Rename the columns that were added automatically in TE3:
    nameColumn.IsNameInferred = false;
    nameColumn.Name = name;
    fieldColumn.IsNameInferred = false;
    fieldColumn.Name = name + " Fields";
    orderColumn.IsNameInferred = false;
    orderColumn.Name = name + " Order";
}
// Set remaining properties for field parameters to work
// See: https://twitter.com/markbdi/status/1526558841172893696
nameColumn.SortByColumn = orderColumn;
nameColumn.GroupByColumns.Add(fieldColumn);
fieldColumn.SortByColumn = orderColumn;
fieldColumn.SetExtendedProperty("ParameterMetadata", "{\"version\":3,\"kind\":2}", ExtendedPropertyType.Json);
fieldColumn.IsHidden = true;
orderColumn.IsHidden = true;

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
    public static bool IsAnswerYes(string question, string title = "Please confirm")
    {
        ResetWaitCursor();
        var result = MessageBox.Show(question, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        return result == DialogResult.Yes;
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
    // true/false, yes/no or 1/0 (answer to a Yes/No question)
    public static bool? GetBoolFromEnv(string envVar, bool? defaultValue = null)
    {
        string value = GetStringFromEnv(envVar, defaultValue.HasValue ? defaultValue.Value.ToString() : null);
        if (value == null) return null;
        switch (value.Trim().ToLower())
        {
            case "true": case "yes": case "y": case "1": return true;
            case "false": case "no": case "n": case "0": return false;
        }
        Error("Environment variable " + envVar + " must be true or false, not '" + value + "'.");
        return null;
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
    // [Sales Amount], Sales Amount or 'Sales'[Sales Amount]
    public static List<Measure> GetMeasuresFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.AllMeasures, m => new[] { m.DaxObjectFullName, m.Name, m.Table.DaxObjectFullName + m.DaxObjectFullName }, "Measure");
    }
}
