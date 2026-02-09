#r "Microsoft.VisualBasic"
using System.Windows.Forms;

using Microsoft.VisualBasic;
// 2025-01-07 / B.Agullo
// Creates a dynamic measure and disconnected tables based on Properties annotations of selected measures.
// Each property position gets its own disconnected table with distinct values across all measures.
// The dynamic measure returns the appropriate measure based on selected values in each disconnected table.
#if TE3
ScriptHelper.WaitFormVisible = false;
#endif
// Validate selection
if (Selected.Measures.Count() < 2)
{
    Error("Select at least two measures with Properties annotations and try again.");
    return;
}
// Step 1: Extract and validate Properties annotations
var measuresWithProperties = new List<(Measure measure, List<string> properties)>();
foreach (var measure in Selected.Measures)
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
// Step 4: Ask user for dynamic measure name
string dynamicMeasureName = Fx.GetNameFromUser(
    Prompt: "Enter the name for the dynamic measure:",
    Title: "Dynamic Measure Name",
    DefaultResponse: "Dynamic Measure"
);
if (dynamicMeasureName == null) return;
// Step 5: Get destination table (use first measure's table)
Table destinationTable = measuresWithProperties[0].measure.Table;
// Step 6: Create disconnected tables for each property position
var propertyTables = new List<Table>();
for (int i = 0; i < propertyCount; i++)
{
    // Collect distinct values for this property position across all measures
    var distinctValues = measuresWithProperties
        .Select(m => m.properties[i])
        .Distinct()
        .OrderBy(v => v)
        .ToList();
    // Create table name
    string tableName = String.Format("Property{0}", i + 1);
    // Build DAX expression for calculated table
    string tableDax = "{" + Environment.NewLine;
    for (int j = 0; j < distinctValues.Count; j++)
    {
        tableDax += String.Format("    (\"{0}\", {1})", distinctValues[j], j);
        if (j < distinctValues.Count - 1)
            tableDax += "," + Environment.NewLine;
    }
    tableDax += Environment.NewLine + "}";
    // Create the calculated table
    var propTable = Model.AddCalculatedTable(tableName, tableDax);
    // Setup columns (handle both TE2 and TE3)
    var te2 = propTable.Columns.Count == 0;
    var valueColumn = te2 
        ? propTable.AddCalculatedTableColumn(tableName, "[Value1]") 
        : propTable.Columns["Value1"] as CalculatedTableColumn;
    var orderColumn = te2 
        ? propTable.AddCalculatedTableColumn(tableName + " Order", "[Value2]") 
        : propTable.Columns["Value2"] as CalculatedTableColumn;
    if (!te2)
    {
        valueColumn.IsNameInferred = false;
        valueColumn.Name = tableName;
        orderColumn.IsNameInferred = false;
        orderColumn.Name = tableName + " Order";
    }
    valueColumn.SortByColumn = orderColumn;
    orderColumn.IsHidden = true;
    propertyTables.Add(propTable);
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
            partialMeasureExpression += String.Format(
                "VAR __{0} = SELECTEDVALUE( '{1}'[{1}] )" + Environment.NewLine,
                propertyTables[i].Name,
                propertyTables[i].Name
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
                    condition += String.Format(
                        "__{0} = \"{1}\"",
                        propertyTables[i].Name,
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
        // Build format string expression
        string partialFormatExpression = "";
        // Add variable declarations for OTHER properties (not current table)
        for (int i = 0; i < propertyCount; i++)
        {
            if (i == tableIndex) continue;
            partialFormatExpression += String.Format(
                "VAR __{0} = SELECTEDVALUE( '{1}'[{1}] )" + Environment.NewLine,
                propertyTables[i].Name,
                propertyTables[i].Name
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
                    condition += String.Format(
                        "__{0} = \"{1}\"",
                        propertyTables[i].Name,
                        item.properties[i]
                    );
                    firstCondition = false;
                }
                partialFormatExpression += condition + "," + Environment.NewLine;
                // Use format string or format string expression
                string formatValue = !string.IsNullOrEmpty(item.measure.FormatStringExpression)
                    ? item.measure.FormatStringExpression
                    : String.Format("\"{0}\"", item.measure.FormatString);
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
                : String.Format("\"{0}\"", matchingMeasures[0].measure.FormatString);
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
    measureExpression += String.Format(
        "VAR __{0} = SELECTEDVALUE( '{1}'[{1}] )" + Environment.NewLine,
        propertyTables[i].Name,
        propertyTables[i].Name
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
        condition += String.Format(
            "__{0} = \"{1}\"",
            propertyTables[i].Name,
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
    formatExpression += String.Format(
        "VAR __{0} = SELECTEDVALUE( '{1}'[{1}] )" + Environment.NewLine,
        propertyTables[i].Name,
        propertyTables[i].Name
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
        condition += String.Format(
            "__{0} = \"{1}\"",
            propertyTables[i].Name,
            item.properties[i]
        );
    }
    formatExpression += condition + "," + Environment.NewLine;
    // Use format string or format string expression
    string formatValue = !string.IsNullOrEmpty(item.measure.FormatStringExpression)
        ? item.measure.FormatStringExpression
        : String.Format("\"{0}\"", item.measure.FormatString);
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
    public static Measure GetSelectedMeasure(IEnumerable<Measure> measures, string label = "Select Measure")
    {
        Measure selectedMeasure = null;
        if (measures.Count() == 1)
        {
            selectedMeasure = measures.First();
        }
        else
        {
            selectedMeasure = SelectMeasure(preselect: null, label: label);
            if (selectedMeasure == null)
            {
                Info("No measure selected.");
                return null;
            }
        }
        return selectedMeasure;
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
    public static string GetNameFromUser(string Prompt, string Title, string DefaultResponse)
    {
        string response = Interaction.InputBox(Prompt, Title, DefaultResponse, 740, 400);
        return response;
    }
    public static bool IsAnswerYes(string question, string title = "Please confirm")
    {
        var result = MessageBox.Show(question, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        return result == DialogResult.Yes;
    }
    public static (IList<string> Values, string Type) SelectAnyObjects(Model model, string selectionType = null, string prompt1 = "select item type", string prompt2 = "select item(s)", string placeholderValue = "")
    {
        var returnEmpty = (Values: new List<string>(), Type: (string)null);
        if (prompt1.Contains("{0}"))
            prompt1 = string.Format(prompt1, placeholderValue ?? "");
        if(prompt2.Contains("{0}"))
            prompt2 = string.Format(prompt2, placeholderValue ?? "");
        if (selectionType == null)
        {
            IList<string> selectionTypeOptions = new List<string> { "Table", "Column", "Measure", "Scalar" };
            selectionType = ChooseString(selectionTypeOptions, label: prompt1, customWidth: 600);
        }
        if (selectionType == null) return returnEmpty;
        IList<string> selectedValues = new List<string>();
        switch (selectionType)
        {
            case "Table":
                selectedValues = SelectTableMultiple(model, label: prompt2);
                break;
            case "Column":
                selectedValues = SelectColumnMultiple(model, label: prompt2);
                break;
            case "Measure":
                selectedValues = SelectMeasureMultiple(model: model, label: prompt2);
                break;
            case "Scalar":
                IList<string> scalarList = new List<string>();
                scalarList.Add(GetNameFromUser(prompt2, "Scalar value", "0"));
                selectedValues = scalarList;
                break;
            default:
                Error("Invalid selection type");
                return returnEmpty;
        }
        if (selectedValues == null || selectedValues.Count == 0) return returnEmpty; 
        return (Values:selectedValues, Type:selectionType);
    }
    public static string ChooseString(IList<string> OptionList, string label = "Choose item", int customWidth = 400, int customHeight = 500)
    {
        return ChooseStringInternal(OptionList, MultiSelect: false, label: label, customWidth: customWidth, customHeight:customHeight) as string;
    }
    public static List<string> ChooseStringMultiple(IList<string> OptionList, string label = "Choose item(s)", int customWidth = 650, int customHeight = 550)
    {
        return ChooseStringInternal(OptionList, MultiSelect:true, label:label, customWidth: customWidth, customHeight: customHeight) as List<string>;
    }
    private static object ChooseStringInternal(IList<string> OptionList, bool MultiSelect, string label = "Choose item(s)", int customWidth = 400, int customHeight = 500)
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
            Location = new System.Drawing.Point(0, 6)
        };
        TextBox searchBox = new TextBox
        {
            Location = new System.Drawing.Point(60, 3),
            Width = customWidth - 120,
            Anchor = AnchorStyles.Left | AnchorStyles.Right
        };
        searchPanel.Controls.Add(searchLabel);
        searchPanel.Controls.Add(searchBox);
        // ListBox panel in the middle
        Panel listBoxPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 35, 0, 75)  // Top padding = searchPanel height, Bottom = buttonPanel height
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
        form.Controls.Add(listBoxPanel);   // Fill - add last
        form.Width = customWidth;
        form.Height = customHeight;
        DialogResult result = form.ShowDialog();
        if (result == DialogResult.Cancel)
        {
            Info("You Cancelled!");
            return null;
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
    public static IList<string> SelectMeasureMultiple(Model model, IEnumerable<Measure> measures = null, string label = "Select Measure(s)")
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
        var selectedDisplayStrings = ChooseStringMultiple(displayList, label: label);
        if (selectedDisplayStrings == null || selectedDisplayStrings.Count == 0)
            return new List<string>();
        // Map back to DaxObjectFullName
        var selectedMeasureNames = selectedDisplayStrings
            .Select(display => measureDisplayMap[display])
            .ToList();
        return selectedMeasureNames;
    }
    public static IList<string> SelectColumnMultiple(Model model, IEnumerable<Column> columns = null, string label = "Select Columns(s)")
    {
        columns ??= model.AllColumns;
        IList<string> columnNames = columns.Select(m => m.DaxObjectFullName).ToList();
        IList<string> selectedColumnNames = ChooseStringMultiple(columnNames, label: label);
        return selectedColumnNames;
    }
    public static IList<string> SelectTableMultiple(Model model, IEnumerable<Table> Tables = null, string label = "Select Tables(s)", int customWidth = 400)
    {
        Tables ??= model.Tables;
        IList<string> TableNames = Tables.Select(m => m.DaxObjectFullName).ToList();
        IList<string> selectedTableNames = ChooseStringMultiple(TableNames, label: label, customWidth: customWidth);
        return selectedTableNames;
    }
}
