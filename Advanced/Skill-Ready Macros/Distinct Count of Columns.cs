#r "System.Drawing"
#r "Microsoft.VisualBasic"
using System.Windows.Forms;
using Microsoft.VisualBasic;

// MACRO: Name: Distinct Count of Columns
// MACRO: ValidContexts: Column
// MACRO: Tooltip: Creates a "Distinct Count of <column>" measure for each selected column
// SKILL: name: distinct-count-of-columns
// SKILL: description: Create "Distinct Count of <column>" measures (DISTINCTCOUNT) for columns of a semantic model. Use when the user asks to count distinct values, unique customers/products/orders, or create distinct count measures from columns.
// SKILL:
// SKILL: Creates one measure "Distinct Count of <column>" per column, using DISTINCTCOUNT, in the column's table and display folder, with format string "#,0" and a description. The base columns are left visible.
// SKILL:
// SKILL: ## Inputs (environment variables)
// SKILL: Set these variables before running the macro. Work out each value from the conversation and the model (inspect the model if needed); if a value is unclear, ask the user instead of guessing.
// SKILL: - DC_COLUMNS (required): 'Table'[Column] refs separated by ';'. The columns to count distinct values of (any data type, typically keys or IDs).
// SKILL:
// SKILL: ## Notes
// SKILL: The measure is named "Distinct Count of <column>"; if a measure with that name already exists in the model, check with the user before running, since measure names must be unique.
// '2026-10-31 / B.Agullo / Agent ready: in CLI mode the columns come from the DC_COLUMNS environment variable
#if TECLI
var columns = Fx.GetColumnsFromEnv(Model, "DC_COLUMNS");
if (columns == null) return;
#else
var columns = Selected.Columns.ToList();
#endif
// Loop through all target columns:
foreach(var c in columns)
{
    var newMeasure = c.Table.AddMeasure(
        "Distinct Count of " + c.Name,                    // Name
        "DISTINCTCOUNT(" + c.DaxObjectFullName + ")",    // DAX expression
        c.DisplayFolder                        // Display Folder
    );
    // Set the format string on the new measure:
    newMeasure.FormatString = "#,0";
    // Provide some documentation:
    newMeasure.Description = "This measure is the distinct count of column " + c.DaxObjectFullName;
    // Hide the base column:
    //c.IsHidden = true;
}

public static class Fx
{
#if !TECLI
    public static Measure GetSelectedMeasure(IEnumerable<Measure> measures, string label = "Select Measure")
    {
        Measure selectedMeasure = null;
        if (measures.Count() == 1)
        {
            selectedMeasure = measures.First();
        }
        else
        {
            selectedMeasure = SelectMeasure(label: label);
            if (selectedMeasure == null)
            {
                Info("No measure selected.");
                return null;
            }
        }
        return selectedMeasure;
    }
    public static Table GetSelectedTable(Model model, IEnumerable<Table> tables, string label = "Select Table", bool createMeasureTableIfNoneSelected = false, string createTableName = "ReferentialIntegrity" )
    {
        Table selectedTable = null;
        if (tables.Count() == 1)
        {
            selectedTable = tables.First();
        }
        else if (tables.Count() > 1)
        {
            selectedTable = SelectTable(tables, preselect: tables.First(), label: label);
            if (selectedTable == null)
            {
                Info("No table selected.");
                return null;
            }
        } else             {
            if (createMeasureTableIfNoneSelected)
            {
                selectedTable = model.AddCalculatedTable(createTableName, "FILTER({0},FALSE)");
            } 
            else
            {
                selectedTable = SelectTable(tables: tables, label: label);
                if (selectedTable == null)
                {
                    Info("No table selected.");
                    return null;
                }
            }
        }
        return selectedTable;
    }
    public static Dictionary<string, string> SelectCalculationItems(Model model, string label = "Select calculation items (max 1 per group)")
    {
        if (!model.Tables.OfType<CalculationGroupTable>().Any())
        {
            Info("No calculation groups found in the model.");
            return null;
        }
        // Create a TreeView form
        Form form = new Form
        {
            Text = label,
            StartPosition = FormStartPosition.CenterScreen,
            Width = 600,
            Height = 500,
            Padding = new Padding(10)
        };
        TreeView treeView = new TreeView
        {
            Dock = DockStyle.Fill,
            CheckBoxes = true
        };
        // Track selections per calculation group
        var selectionMap = new Dictionary<string, TreeNode>();
        // Populate tree with calculation groups and items
        foreach (var calcGroup in model.Tables.OfType<CalculationGroupTable>())
        {
            TreeNode groupNode = new TreeNode(calcGroup.Name)
            {
                Tag = calcGroup
            };
            foreach (var calcItem in calcGroup.CalculationItems)
            {
                TreeNode itemNode = new TreeNode(calcItem.Name)
                {
                    Tag = calcItem
                };
                groupNode.Nodes.Add(itemNode);
            }
            treeView.Nodes.Add(groupNode);
            groupNode.Expand();
        }
        // Handle BeforeCheck to prevent checking group nodes
        treeView.BeforeCheck += (sender, e) =>
        {
            // Prevent checking calculation group nodes (only allow calculation items)
            if (e.Node.Tag is CalculationGroupTable)
            {
                e.Cancel = true;
            }
        };
        // Handle node check events to enforce "one per group" rule
        treeView.AfterCheck += (sender, e) =>
        {
            if (e.Node.Tag is CalculationItem)
            {
                var calcItem = (CalculationItem)e.Node.Tag;
                string groupName = calcItem.CalculationGroupTable.Name;
                if (e.Node.Checked)
                {
                    // Uncheck any previously selected item in this group
                    if (selectionMap.ContainsKey(groupName))
                    {
                        selectionMap[groupName].Checked = false;
                    }
                    selectionMap[groupName] = e.Node;
                }
                else
                {
                    // Remove from selection if unchecked
                    if (selectionMap.ContainsKey(groupName) && selectionMap[groupName] == e.Node)
                    {
                        selectionMap.Remove(groupName);
                    }
                }
            }
        };
        // Button panel
        FlowLayoutPanel buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 50,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(5)
        };
        Button okButton = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Width = 80,
            Height = 30
        };
        Button cancelButton = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Width = 80,
            Height = 30
        };
        buttonPanel.Controls.Add(okButton);
        buttonPanel.Controls.Add(cancelButton);
        form.Controls.Add(treeView);
        form.Controls.Add(buttonPanel);
        // Show dialog
        ResetWaitCursor();
        DialogResult result = form.ShowDialog();
        if (result == DialogResult.Cancel)
        {
            Info("Selection cancelled.");
            return null;
        }
        // Build result dictionary: CalculationGroupName -> CalculationItemName
        var selectedItems = new Dictionary<string, string>();
        foreach (var kvp in selectionMap)
        {
            if (kvp.Value.Checked)
            {
                selectedItems[kvp.Key] = kvp.Value.Text;
            }
        }
        return selectedItems;
    }
    public static void CheckCompatibilityVersion(Model model, int requiredVersion, string customMessage = "Compatibility level must be raised to {0} to run this script. Do you want raise the compatibility level?")
    {
        if (model.Database.CompatibilityLevel < requiredVersion)
        {
            if (Fx.IsAnswerYes(String.Format("The model compatibility level is below {0}. " + customMessage, requiredVersion)))
            {
                model.Database.CompatibilityLevel = requiredVersion;
            }
            else
            {
                Info("Operation cancelled.");
                return;
            }
        }
    }
    public static Function CreateFunction(
        Model model,
        string name,
        string expression,
        out bool functionCreated,
        string description = null,
        string annotationLabel = null,
        string annotationValue = null,
        string outputType = null,
        string nameTemplate = null,
        string formatString = null,
        string displayFolder = null,
        string outputDestination = null)
    {
        Function function = null as Function;
        functionCreated = false;
        var matchingFunctions = model.Functions.Where(f => f.GetAnnotation(annotationLabel) == annotationValue);
        if (matchingFunctions.Count() == 1)
        {
            return matchingFunctions.First();
        }
        else if (matchingFunctions.Count() == 0)
        {
            function = model.AddFunction(name);
            function.Expression = expression;
            function.Description = description;
            functionCreated = true;
        }
        else
        {
            Error("More than one function found with annoation " + annotationLabel + " value " + annotationValue);
            return null as Function;
        }
        if (!string.IsNullOrEmpty(annotationLabel) && !string.IsNullOrEmpty(annotationValue))
        {
            function.SetAnnotation(annotationLabel, annotationValue);
        }
        if (!string.IsNullOrEmpty(outputType))
        {
            function.SetAnnotation("outputType", outputType);
        }
        if (!string.IsNullOrEmpty(nameTemplate))
        {
            function.SetAnnotation("nameTemplate", nameTemplate);
        }
        if (!string.IsNullOrEmpty(formatString))
        {
            function.SetAnnotation("formatString", formatString);
        }
        if (!string.IsNullOrEmpty(displayFolder))
        {
            function.SetAnnotation("displayFolder", displayFolder);
        }
        if (!string.IsNullOrEmpty(outputDestination))
        {
            function.SetAnnotation("outputDestination", outputDestination);
        }
        return function;
    }
    public static Table CreateCalcTable(Model model, string tableName, string tableExpression = "FILTER({0},FALSE)")
    {
        return model.Tables.FirstOrDefault(t =>
                            string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase)) //case insensitive search
                            ?? model.AddCalculatedTable(tableName, tableExpression);
    }
    public static Measure CreateMeasure(
        Table table, 
        string measureName, 
        string measureExpression,
        out bool measureCreated,
        string formatString = null,
        string displayFolder = null,
        string description = null,
        string annotationLabel = null, 
        string annotationValue = null,
        bool isHidden = false)
    {
        measureCreated = false;
        IEnumerable<Measure> matchingMeasures = null as IEnumerable<Measure>;
        if (!string.IsNullOrEmpty(annotationLabel) && !string.IsNullOrEmpty(annotationValue))
        {
            matchingMeasures = table.Measures.Where(m => m.GetAnnotation(annotationLabel) == annotationValue);
        }
        else
        {
            matchingMeasures = table.Measures.Where(m => m.Name == measureName);
        }
        if (matchingMeasures.Count() == 1)
        {
            return matchingMeasures.First();
        }
        else if (matchingMeasures.Count() == 0)
        {
            Measure measure = table.AddMeasure(measureName, measureExpression);
            measure.Description = description;
            measure.DisplayFolder = displayFolder;
            measure.FormatString = formatString;
            measureCreated = true;
            if (!string.IsNullOrEmpty(annotationLabel) && !string.IsNullOrEmpty(annotationValue))
            {
                measure.SetAnnotation(annotationLabel, annotationValue);
            }
            measure.IsHidden = isHidden;
            return measure;
        }
        else
        {
            Error("More than one measure found with annoation " + annotationLabel + " value " + annotationValue);
            Output(matchingMeasures);
            return null as Measure;
        }
    }
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
    public static (IList<string> Values, string Type) SelectAnyObjects(Model model, string selectionType = null, string prompt1 = "select item type", string prompt2 = "select item(s)", string placeholderValue = "", string defaultValue = null, string description = null)
    {
        var returnEmpty = (Values: new List<string>(), Type: (string)null);
        if (prompt1.Contains("{0}"))
            prompt1 = string.Format(prompt1, placeholderValue ?? "");
        if(prompt2.Contains("{0}"))
            prompt2 = string.Format(prompt2, placeholderValue ?? "");
        if (selectionType == null)
        {
            IList<string> selectionTypeOptions = new List<string> { "Table", "Column", "Measure", "Calendar", "Scalar", "Text" };
            selectionType = ChooseString(selectionTypeOptions, label: prompt1, customWidth: 600);
        }
        if (selectionType == null) return returnEmpty;
        IList<string> selectedValues = new List<string>();
        switch (selectionType)
        {
            case "Table":
                selectedValues = SelectTableMultiple(model, label: prompt2, defaultValue: defaultValue, description: description);
                break;
            case "Column":
                selectedValues = SelectColumnMultiple(model, label: prompt2, defaultValue: defaultValue, description: description);
                break;
            case "Measure":
                selectedValues = SelectMeasureMultiple(model: model, label: prompt2, defaultValue: defaultValue, description: description);
                break;
            case "Calendar":
                selectedValues = SelectCalendarMultiple(model, label: prompt2, defaultValue: defaultValue, description: description);
                break;
            case "Scalar":
                IList<string> scalarList = new List<string>();
                scalarList.Add(GetNameFromUser(prompt2, "Scalar value", defaultValue ?? "0"));
                selectedValues = scalarList;
                break;
            case "Text":
                IList<string> textList = new List<string>();
                string textValue = GetNameFromUser(prompt2, "Text value", defaultValue ?? "");
                if (textValue == null) return returnEmpty;
                // Wrap in double quotes if not already a quoted DAX string literal
                if (!(textValue.StartsWith("\"") && textValue.EndsWith("\"")))
                    textValue = "\"" + textValue + "\"";
                textList.Add(textValue);
                selectedValues = textList;
                break;
            default:
                Error("Invalid selection type");
                return returnEmpty;
        }
        if (selectedValues == null || selectedValues.Count == 0) return returnEmpty;
        // Default-value option selected: return empty value so the argument is left blank (default used)
        if (selectedValues.Count == 1 && selectedValues[0] == DefaultSentinel)
            return (Values: new List<string>() { "" }, Type: selectionType);
        return (Values:selectedValues, Type:selectionType);
    }
    public static string ChooseString(IList<string> OptionList, string label = "Choose item", int customWidth = 400, int customHeight = 500)
    {
        return ChooseStringInternal(OptionList, MultiSelect: false, label: label, customWidth: customWidth, customHeight:customHeight) as string;
    }
    public static List<string> ChooseStringMultiple(IList<string> OptionList, string label = "Choose item(s)", int customWidth = 650, int customHeight = 550, string defaultValue = null, string description = null)
    {
        return ChooseStringInternal(OptionList, MultiSelect:true, label:label, customWidth: customWidth, customHeight: customHeight, defaultValue: defaultValue, description: description) as List<string>;
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
    public static IEnumerable<Table> GetDateTables(Model model)
    {
        var dateTables = model.Tables
            .Where(t => t.DataCategory == "Time" &&
                   t.Columns.Any(c => c.IsKey && c.DataType == DataType.DateTime))
            .ToList();
        if (!dateTables.Any())
        {
            Error("No date table detected in the model. Please mark your date table(s) as date table");
            return null;
        }
        return dateTables;
    }
    public static Table GetDateTable(Model model, string prompt = "Select Date Table")
    {
        var dateTables = GetDateTables(model);
        if (dateTables == null) {
            Table t = SelectTable(model.Tables, label: prompt);
            if(t == null)
            {
                Error("No table selected");
                return null;
            }
            if (IsAnswerYes(String.Format("Mark {0} as date table?",t.DaxObjectFullName)))
            {
                t.DataCategory = "Time";
                var dateColumns = t.Columns
                    .Where(c => c.DataType == DataType.DateTime)
                    .ToList();
                if(dateColumns.Count == 0)
                {
                    Error(String.Format(@"No date column detected in the table {0}. Please check that the table contains a date column",t.Name));
                    return null;
                }
                var keyColumn = SelectColumn(dateColumns, preselect:dateColumns.First(), label: "Select Date Column to be used as key column");
                if(keyColumn == null)
                {
                    Error("No key column selected");
                    return null;
                }
                keyColumn.IsKey = true;
            }
            return t;
        };
        if (dateTables.Count() == 1)
            return dateTables.First();
        Table dateTable = SelectTable(dateTables, label: prompt);
        if(dateTable == null)
        {
            Error("No table selected");
            return null;
        }
        return dateTable;
    }
    public static Column GetDateColumn(Table dateTable, string prompt = "Select Date Column")
    {
        var dateColumns = dateTable.Columns
            .Where(c => c.DataType == DataType.DateTime)
            .ToList();
        if(dateColumns.Count == 0)
        {
            Error(String.Format(@"No date column detected in the table {0}. Please check that the table contains a date column", dateTable.Name));
            return null;
        }
        if(dateColumns.Any(c => c.IsKey))
        {
            return dateColumns.First(c => c.IsKey);
        }
        Column dateColumn = null;
        if (dateColumns.Count() == 1)
        {
            dateColumn = dateColumns.First();
        }
        else
        {
            dateColumn = SelectColumn(dateColumns, label: prompt);
            if (dateColumn == null)
            {
                Error("No column selected");
                return null;
            }
        }
        return dateColumn;
    }
    public static IEnumerable<Table> GetFactTables(Model model)
    {
        IEnumerable<Table> factTables = model.Tables.Where(
            x => model.Relationships.Where(r => r.ToTable == x)
                    .All(r => r.ToCardinality == RelationshipEndCardinality.Many)
                && model.Relationships.Where(r => r.FromTable == x)
                    .All(r => r.FromCardinality == RelationshipEndCardinality.Many)
                && model.Relationships.Where(r => r.ToTable == x || r.FromTable == x).Any()); // at least one relationship
        if (!factTables.Any())
        {
            Error("No fact table detected in the model. Please check that the model contains relationships");
            return null;
        }
        return factTables;
    }
    public static Table GetFactTable(Model model, string prompt = "Select Fact Table")
    {
        Table factTable = null;
        var factTables = GetFactTables(model);
        if (factTables == null)
        {
           factTable = SelectTable(model.Tables, label: "This does not look like a star schema. Choose your fact table manually");
            if (factTable == null)
            {
                Error("No table selected");
                return null;
            }
            return factTable;
        };
        if (factTables.Count() == 1)
            return factTables.First();
        factTable = SelectTable(factTables, label: prompt);
        if (factTable == null)
        {
            Error("No table selected");
            return null;
        }
        return factTable;
    }
    public static Table GetTablesWithAnnotation(IEnumerable<Table> tables, string annotationLabel, string annotationValue)
    {
        Func<Table, bool> lambda = t => t.GetAnnotation(annotationLabel) == annotationValue;
        IEnumerable<Table> matchTables = GetFilteredTables(tables, lambda);
        return GetFilteredTables(tables, lambda).FirstOrDefault();
    }
    public static IEnumerable<Table> GetFilteredTables(IEnumerable<Table> tables, Func<Table, bool> lambda)
    {
        var filteredTables = tables.Where(t => lambda(t));
        return filteredTables.Any() ? filteredTables : null;
    }
    public static IEnumerable<Column> GetFilteredColumns(IEnumerable<Column> columns, Func<Column, bool> lambda, bool returnAllIfNoneFound = true)
    {
        var filteredColumns = columns.Where(c => lambda(c));
        return filteredColumns.Any() || returnAllIfNoneFound ? filteredColumns : null;
    }
    public static IList<string> SelectMeasureMultiple(Model model, IEnumerable<Measure> measures = null, string label = "Select Measure(s)", string defaultValue = null, string description = null)
    {
        measures ??= model.AllMeasures;
        // Create display strings with format: TableName\DisplayFolder\[MeasureName]
        var measureDisplayMap = new Dictionary<string, string>();
        var displayList = new List<string>();
        foreach (var measure in measures.OrderBy(m => m.Table.Name).ThenBy(m => m.DisplayFolder).ThenBy(m => m.Name))
        {
            string displayString;
            if (string.IsNullOrEmpty(measure.DisplayFolder))
            {
                displayString = String.Format("{0}\\[{1}]", measure.Table.Name, measure.Name);
            }
            else
            {
                displayString = String.Format("{0}\\{1}\\[{2}]", measure.Table.Name, measure.DisplayFolder, measure.Name);
            }
            measureDisplayMap[displayString] = measure.DaxObjectFullName;
            displayList.Add(displayString);
        }
        // Show the display list to user
        var selectedDisplayStrings = ChooseStringMultiple(displayList, label: label, defaultValue: defaultValue, description: description);
        if (selectedDisplayStrings == null || selectedDisplayStrings.Count == 0)
            return new List<string>();
        // Pass through the default sentinel unchanged
        if (selectedDisplayStrings.Count == 1 && selectedDisplayStrings[0] == DefaultSentinel)
            return selectedDisplayStrings;
        // Map back to DaxObjectFullName
        var selectedMeasureNames = selectedDisplayStrings
            .Select(display => measureDisplayMap[display])
            .ToList();
        return selectedMeasureNames;
    }
    public static IList<string> SelectColumnMultiple(Model model, IEnumerable<Column> columns = null, string label = "Select Columns(s)", string defaultValue = null, string description = null)
    {
        columns ??= model.AllColumns;
        IList<string> columnNames = columns.Select(m => m.DaxObjectFullName).ToList();
        IList<string> selectedColumnNames = ChooseStringMultiple(columnNames, label: label, defaultValue: defaultValue, description: description);
        return selectedColumnNames;
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
#if !TECLI
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
#endif
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
#if !TECLI
    // [Sales Amount], Sales Amount or 'Sales'[Sales Amount]
    public static List<Measure> GetMeasuresFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.AllMeasures, m => new[] { m.DaxObjectFullName, m.Name, m.Table.DaxObjectFullName + m.DaxObjectFullName }, "Measure");
    }
    // 'Sales' or Sales
    public static List<Table> GetTablesFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.Tables, t => new[] { t.DaxObjectFullName, t.Name }, "Table");
    }
    // Function name, e.g. Local.TimeIntel.YTD
    public static List<Function> GetFunctionsFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.Functions, f => new[] { f.Name }, "Function");
    }
    // Single-object versions: the variable must hold exactly one reference
    public static Column GetColumnFromEnv(Model model, string envVar) { return Single(GetColumnsFromEnv(model, envVar), envVar); }
    public static Measure GetMeasureFromEnv(Model model, string envVar) { return Single(GetMeasuresFromEnv(model, envVar), envVar); }
    public static Table GetTableFromEnv(Model model, string envVar) { return Single(GetTablesFromEnv(model, envVar), envVar); }
    private static T Single<T>(List<T> list, string envVar) where T : class
    {
        if (list == null) return null;
        if (list.Count != 1) { Error("Environment variable " + envVar + " must contain exactly one reference."); return null; }
        return list[0];
    }
    public static IList<string> SelectTableMultiple(Model model, IEnumerable<Table> Tables = null, string label = "Select Tables(s)", int customWidth = 400, string defaultValue = null, string description = null)
    {
        Tables ??= model.Tables;
        IList<string> TableNames = Tables.Select(m => m.DaxObjectFullName).ToList();
        IList<string> selectedTableNames = ChooseStringMultiple(TableNames, label: label, customWidth: customWidth, defaultValue: defaultValue, description: description);
        return selectedTableNames;
    }
    public static IList<string> SelectCalendarMultiple(Model model, string label = "Select Calendar(s)", string defaultValue = null, string description = null)
    {
        // Build a display->reference map: display is the calendar name, reference is quoted 'Name'
        var calendarMap = new Dictionary<string, string>();
        var displayList = new List<string>();
        foreach (var calendar in model.AllCalendars.OrderBy(c => c.Name))
        {
            string display = calendar.Name;
            calendarMap[display] = "'" + calendar.Name + "'";
            displayList.Add(display);
        }
        var selectedDisplay = ChooseStringMultiple(displayList, label: label, defaultValue: defaultValue, description: description);
        if (selectedDisplay == null || selectedDisplay.Count == 0)
            return new List<string>();
        // Pass through the default sentinel unchanged
        if (selectedDisplay.Count == 1 && selectedDisplay[0] == DefaultSentinel)
            return selectedDisplay;
        return selectedDisplay.Select(d => calendarMap[d]).ToList();
    }
#endif
}
