#r "Microsoft.VisualBasic"
//2025-09-26/B.Agullo/ fixed bug that would not store annotations if initialized during runtime
//2025-09-16/B.Agullo/
//Creates measures based on DAX UDFs 
//Check the blog post for futher information: https://www.esbrina-ba.com/automatically-create-measures-with-dax-user-defined-functions/
using System.Windows.Forms;

using Microsoft.VisualBasic;
using System.Text.RegularExpressions;
#if TE3
ScriptHelper.WaitFormVisible = false;
#endif
// PSEUDOCODE / PLAN:
// 1. Verify that the user has selected one or more functions (Selected.Functions).
// 2. If none selected, show error and abort.
// 3. Create FunctionExtended objects for each selected function and keep them in a list.
// 4. Extract all parameters from the selected functions and build a distinct list by name.
// 5. For each distinct parameter name:
//      - Prompt the user once with Fx.SelectAnyObjects to choose the objects to iterate for that parameter.
//      - If the user cancels or selects nothing, abort the whole operation.
//      - Store the resulting IList<string> in a dictionary keyed by parameter name so it can be retrieved later.
// 6. Example usage: create a sample FunctionExtended (as the original example does),
//    then when iterating the function parameters use the previously built dictionary to obtain the list
//    of objects for each parameter name (do not prompt again).
// 7. Build measure names/expressions by iterating the parameter-object combinations and create measures.
//
// NOTES:
// - All Fx.SelectAnyObjects calls must use parameter names on the call.
// - The dictionary is Dictionary<string, IList<string>> parameterObjectsMap.
// - Abort execution if any required selection is cancelled.
// Validate selection
if (Selected.Functions.Count() == 0)
{
    Error("Select one or more functions and try again.");
    return;
}
// Create FunctionExtended objects for each selected function and store them for later iteration
IList<FunctionExtended> selectedFunctions = new List<FunctionExtended>();
foreach (var f in Selected.Functions)
{
    // Create the FunctionExtended and add to list
    FunctionExtended fe = FunctionExtended.CreateFunctionExtended(f);
    selectedFunctions.Add(fe);
}
// Flatten all parameters from selected functions
var allParametersFlat = selectedFunctions
    .SelectMany(sf => sf.Parameters ?? new List<FunctionParameter>())
    .ToList();
// Build distinct FunctionParameter objects (first occurrence per name)
IList<FunctionParameter> distinctParameters = allParametersFlat
    .GroupBy(p => p.Name)
    .Select(g => g.First())
    .ToList();
// For each distinct parameter, ask the user once which objects should be iterated and store mapping
var parameterObjectsMap = new Dictionary<string, (IList<string> Values, string Type)>();
foreach (var param in distinctParameters)
{
    string selectionType = null;
    if (param.Name.ToUpper().Contains("MEASURE"))
    {
        selectionType = "Measure";
    }
    else if (param.Name.ToUpper().Contains("COLUMN"))
    {
        selectionType = "Column";
    }
    else if (param.Name.ToUpper().Contains("TABLE"))
    {
        selectionType = "Table";
    }
    (IList<string> Values,string Type) selectedObjectsForParam = Fx.SelectAnyObjects(
        Model,
        selectionType: selectionType,
        prompt1: String.Format(@"Select object type for {0} parameter", param.Name),
        prompt2: String.Format(@"Select item for {0} parameter", param.Name),
        placeholderValue: param.Name
    );
    if (selectedObjectsForParam.Type == null || selectedObjectsForParam.Values.Count == 0)
    {
        Info(String.Format("No objects selected for parameter '{0}'. Operation cancelled.", param.Name));
        return;
    }
    parameterObjectsMap[param.Name] = selectedObjectsForParam;
}
foreach (var func in selectedFunctions)
{
    string delimiter = "";
    IList<string> previousList = new List<string>() { func.Name + "(" };
    IList<string> currentList = new List<string>();
    IList<string> previousListNames = new List<string>() { func.OutputNameTemplate };
    IList<string> currentListNames = new List<string>();
    IList<string> previousDestinations = new List<string>() { func.OutputDestination };
    IList<string> currentDestinations = new List<string>();
    IList<string> previousDisplayFolders = new List<string>() { func.OutputDisplayFolder };
    IList<string> currentDisplayFolders = new List<string>();
    IList<string> previousFormatStrings = new List<string>() { func.OutputFormatString };
    IList<string> currentFormatStrings = new List<string>();
    // When iterating the parameters of this specific function, use the mapping created for distinct parameters.
    foreach (var param in func.Parameters)
    {
        currentList = new List<string>(); //reset current list
        currentListNames = new List<string>();
        currentFormatStrings = new List<string>();
        currentDestinations = new List<string>();
        currentDisplayFolders = new List<string>();
        // Retrieve the objects list for this parameter name from the map (prompting was done earlier)
        (IList<string> Values, string Type) paramObject;
        if (!parameterObjectsMap.TryGetValue(param.Name, out paramObject) || paramObject.Type == null || paramObject.Values.Count == 0)
        {
            Error(String.Format("No objects were selected earlier for parameter '{0}'.", param.Name));
            return;
        }
        for (int i = 0; i < previousList.Count; i++)
        {
            string s = previousList[i];
            string sName = previousListNames[i];
            string sFormatString = previousFormatStrings[i];
            string sDisplayFolder = previousDisplayFolders[i];
            string sDestination = previousDestinations[i];
            foreach (var o in paramObject.Values)
            {
                //extract original name and format string if the parameter is a measure
                string paramName = o;
                string paramFormatStringFull = "";
                string paramFormatStringRoot = "";
                string paramDisplayFolder = "";
                string paramTable = "";
                //prepare placeholder
                string paramNamePlaceholder = param.Name + "Name";
                string paramFormatStringRootPlaceholder = param.Name + "FormatStringRoot";
                string paramFormatStringFullPlaceholder = param.Name + "FormatStringFull";
                string paramDisplayFolderPlaceholder = param.Name + "DisplayFolder";
                string paramTablePlaceholder = "";
                if (paramObject.Type == "Measure")
                {
                    Measure m = Model.AllMeasures.FirstOrDefault(m => m.DaxObjectFullName == o);
                    paramName = m.Name;
                    paramFormatStringFull = m.FormatString;
                    paramDisplayFolder = m.DisplayFolder;
                    paramTable = m.Table.DaxObjectFullName;
                    paramTablePlaceholder = param.Name + "Table";
                }
                else if (paramObject.Type == "Column")
                {
                    Column c = Model.AllColumns.FirstOrDefault(c => c.DaxObjectFullName == o);
                    paramName = c.Name;
                    paramFormatStringFull = c.FormatString;
                    paramDisplayFolder = c.DisplayFolder;
                    paramTable = c.Table.DaxObjectFullName;
                    paramTablePlaceholder = param.Name + "Table";
                }
                else if (paramObject.Type == "Table")
                {
                    Table t = Model.Tables.FirstOrDefault(t => t.DaxObjectFullName == o);
                    paramName = t.Name;
                    paramFormatStringFull = "";
                    paramDisplayFolder = "";
                    paramTable = t.DaxObjectFullName;
                    paramTablePlaceholder = param.Name;
                }
                if (paramFormatStringFull.Contains(";"))
                {
                    //keep the first part of the format string, strip it of any + sign
                    paramFormatStringRoot = paramFormatStringFull.Split(';')[0].Replace("+","");
                }
                else
                {
                    paramFormatStringRoot = paramFormatStringFull;
                }
                currentList.Add(s + delimiter + o);
                currentListNames.Add(sName.Replace(paramNamePlaceholder, paramName));
                currentFormatStrings.Add(
                    sFormatString
                        .Replace(paramFormatStringFullPlaceholder, paramFormatStringFull)
                        .Replace(paramFormatStringRootPlaceholder, paramFormatStringRoot));
                currentDisplayFolders.Add(
                    sDisplayFolder
                        .Replace(paramNamePlaceholder, paramName)
                        .Replace(paramDisplayFolderPlaceholder, paramDisplayFolder));
                currentDestinations.Add(
                    sDestination.Replace(paramTablePlaceholder, paramTable));
            }
        }
        delimiter = ", ";
        previousList = currentList;
        previousListNames = currentListNames;
        previousDestinations = currentDestinations;
        previousDisplayFolders = currentDisplayFolders;
        previousFormatStrings = currentFormatStrings;
    }
    IList<Table> currentDestinationTables = new List<Table>();
    if(func.OutputType == "Measure" || func.OutputType == "Column")
    {
        for (int i = 0; i < currentDestinations.Count; i++)
        {
            //transform to actual tables, initialize if necessary
            Table destinationTable = Model.Tables.Where(
                t => t.DaxObjectFullName == currentDestinations[i])
                .FirstOrDefault();
            if (destinationTable == null)
            {
                destinationTable = SelectTable(preselect: null, label: $"Select destinatoin table for {func.OutputType} {currentListNames[i]}");
                if (destinationTable == null) return;
            }
            currentDestinationTables.Add(destinationTable);
        }
    }
    if (func.OutputType == "Measure")
    {
        for (int i = 0; i < currentList.Count; i++)
        {
            //It normalizes a folder/display-folder string by collapsing repeated slashes, removing leading/trailing backslashes and trimming whitespace.
            string cleanCurrentDisplayFolder = Regex.Replace(currentDisplayFolders[i], @"[/]+", @"").Trim('\\').Trim();
            Measure measure = currentDestinationTables[i].AddMeasure(currentListNames[i], currentList[i] + ")");
            measure.FormatDax();
            measure.Description = String.Format("Measure created with {0} function. Check function for details.", func.Name);
            measure.DisplayFolder = cleanCurrentDisplayFolder;
            measure.FormatString = currentFormatStrings[i];
        }
    }
    else if (func.OutputType == "Column") 
    {
        for (int i = 0; i < currentList.Count; i++)
        {
            //It normalizes a folder/display-folder string by collapsing repeated slashes, removing leading/trailing backslashes and trimming whitespace.
            string cleanCurrentDisplayFolder = Regex.Replace(currentDisplayFolders[i], @"[/]+", @"").Trim('\\').Trim();
            Column column = currentDestinationTables[i].AddCalculatedColumn(currentListNames[i], currentList[i] + ")");
            //column.FormatDax();
            column.Description = String.Format("Column created with {0} function. Check function for details.", func.Name);
            column.DisplayFolder = cleanCurrentDisplayFolder;
            column.FormatString = currentFormatStrings[i];
        }
    }
    else
    {
        Info("Not implemented yet for output types other than Measure.");
    }
}

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
        if (selectedValues.Count == 0) return returnEmpty; 
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
        ListBox listbox = new ListBox
        {
            Dock = DockStyle.Fill,
            SelectionMode = MultiSelect ? SelectionMode.MultiExtended : SelectionMode.One
        };
        listbox.Items.AddRange(OptionList.ToArray());
        if (!MultiSelect && OptionList.Count > 0)
            listbox.SelectedItem = OptionList[0];
        FlowLayoutPanel buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 70,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(10)
        };
        Button selectAllButton = new Button { Text = "Select All", Visible = MultiSelect , Height = 50, Width = 150};
        Button selectNoneButton = new Button { Text = "Select None", Visible = MultiSelect, Height = 50, Width = 150 };
        Button okButton = new Button { Text = "OK", DialogResult = DialogResult.OK, Height = 50, Width = 100 };
        Button cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Height = 50, Width = 100 };
        selectAllButton.Click += delegate
        {
            for (int i = 0; i < listbox.Items.Count; i++)
                listbox.SetSelected(i, true);
        };
        selectNoneButton.Click += delegate
        {
            for (int i = 0; i < listbox.Items.Count; i++)
                listbox.SetSelected(i, false);
        };
        buttonPanel.Controls.Add(selectAllButton);
        buttonPanel.Controls.Add(selectNoneButton);
        buttonPanel.Controls.Add(okButton);
        buttonPanel.Controls.Add(cancelButton);
        form.Controls.Add(listbox);
        form.Controls.Add(buttonPanel);
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
        IList<string> measureNames = measures.Select(m => m.DaxObjectFullName).OrderBy(t=>t).ToList();
        IList<string> selectedMeasureNames = ChooseStringMultiple(measureNames, label: label);
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

    

    public class FunctionExtended
    {
        public string Name { get; set; }
        public string Expression { get; set; }
        public string Description { get; set; }
        public string OutputFormatString { get; set; }
        public string OutputNameTemplate { get; set; }
        public string OutputType { get; set; }
        public string OutputDisplayFolder { get; set; }

        public string OutputDestination { get; set; } 
        public Function OriginalFunction { get; set; }
        public List<FunctionParameter> Parameters { get; set; }
        private static List<FunctionParameter> ExtractParametersFromExpression(string expression)
        {
            // Find the first set of parentheses before the "=>"
            int arrowIndex = expression.IndexOf("=>");
            if (arrowIndex == -1)
                return new List<FunctionParameter>();

            int openParenIndex = expression.LastIndexOf('(', arrowIndex);
            int closeParenIndex = expression.IndexOf(')', openParenIndex);
            if (openParenIndex == -1 || closeParenIndex == -1 || closeParenIndex > arrowIndex)
                return new List<FunctionParameter>();

            string paramSection = expression.Substring(openParenIndex + 1, closeParenIndex - openParenIndex - 1);
            var paramList = new List<FunctionParameter>();
            var paramStrings = paramSection.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var param in paramStrings)
            {
                var trimmed = param.Trim();
                var nameParams = trimmed.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                //var parts = trimmed.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                var fp = new FunctionParameter();
                fp.Name = nameParams.Length > 0 ? nameParams[0].Trim() : param.Trim();
                fp.Type =
                    (nameParams.Length == 1) ? "ANYVAL" :
                    (nameParams.Length > 1) ?
                        (nameParams[1].IndexOf("anyVal", StringComparison.OrdinalIgnoreCase) >= 0 ? "ANYVAL" :
                         nameParams[1].IndexOf("Scalar", StringComparison.OrdinalIgnoreCase) >= 0 ? "SCALAR" :
                         nameParams[1].IndexOf("Table", StringComparison.OrdinalIgnoreCase) >= 0 ? "TABLE" :
                         nameParams[1].IndexOf("AnyRef", StringComparison.OrdinalIgnoreCase) >= 0 ? "ANYREF" :
                         "ANYVAL")
                    : "ANYVAL";

                fp.Subtype =
                    (fp.Type == "SCALAR" && nameParams.Length > 1) ?
                        (
                            nameParams[1].IndexOf("variant", StringComparison.OrdinalIgnoreCase) >= 0 ? "VARIANT" :
                            nameParams[1].IndexOf("int64", StringComparison.OrdinalIgnoreCase) >= 0 ? "INT64" :
                            nameParams[1].IndexOf("decimal", StringComparison.OrdinalIgnoreCase) >= 0 ? "DECIMAL" :
                            nameParams[1].IndexOf("double", StringComparison.OrdinalIgnoreCase) >= 0 ? "DOUBLE" :
                            nameParams[1].IndexOf("string", StringComparison.OrdinalIgnoreCase) >= 0 ? "STRING" :
                            nameParams[1].IndexOf("datetime", StringComparison.OrdinalIgnoreCase) >= 0 ? "DATETIME" :
                            nameParams[1].IndexOf("boolean", StringComparison.OrdinalIgnoreCase) >= 0 ? "BOOLEAN" :
                            nameParams[1].IndexOf("numeric", StringComparison.OrdinalIgnoreCase) >= 0 ? "NUMERIC" :
                            null
                        )
                    : null;

                // ParameterMode: check for VAL or EXPR (any casing) in the parameter string
                string paramMode = null;
                if (trimmed.IndexOf("VAL", StringComparison.OrdinalIgnoreCase) >= 0)
                    paramMode = "VAL";
                else if (trimmed.IndexOf("EXPR", StringComparison.OrdinalIgnoreCase) >= 0)
                    paramMode = "EXPR";
                else if (fp.Type == "ANYREF")
                {
                    paramMode = "EXPR";
                }else
                    paramMode = "VAL";
                fp.ParameterMode = paramMode;

                paramList.Add(fp);
            }

            return paramList;
        }

        

        public static FunctionExtended CreateFunctionExtended(Function function, bool completeMetadata = true)
        {

            FunctionExtended emptyFunction = null as FunctionExtended;
            List<FunctionParameter> Parameters =  ExtractParametersFromExpression (function.Expression);

            string nameTemplateDefault = "";
            string formatStringDefault = "";
            string displayFolderDefault = "";
            string functionNameShort = function.Name;
            string destinationDefault = ""; 

            if(function.Name.IndexOf(".") > 0)
            {
                functionNameShort = function.Name.Substring(function.Name.LastIndexOf(".") + 1);
            }

            if (Parameters.Count == 0) {
                nameTemplateDefault = function.Name;
                formatStringDefault = "";
                displayFolderDefault = "";
                destinationDefault = "";
            }
            else
            {
                nameTemplateDefault = string.Join(" ", Parameters.Select(p => p.Name + "Name"));
                if(function.Name.Contains("Pct"))
                {
                    formatStringDefault = "+0.0%;-0.0%;-";
                }
                else
                {
                    formatStringDefault = Parameters[0].Name + "FormatStringRoot";
                }


                    
                displayFolderDefault = 
                    String.Format(
                        @"{0}DisplayFolder/{1}Name {2}", 
                        Parameters[0].Name, 
                        Parameters[0].Name,
                        functionNameShort);
                
                if (Parameters[0].Name.ToUpper().Contains("TABLE"))
                {
                    destinationDefault = Parameters[0].Name;
                }
                else if (Parameters[0].Name.ToUpper().Contains("MEASURE") || Parameters[0].Name.ToUpper().Contains("COLUMN"))
                {
                    destinationDefault = Parameters[0].Name + "Table";
                }
                else
                {
                    destinationDefault = "Custom";
                }
                

            };
            
            


            string myOutputType = function.GetAnnotation("outputType"); 
            string myNameTemplate = function.GetAnnotation("nameTemplate");
            string myFormatString = function.GetAnnotation("formatString");
            string myDisplayFolder = function.GetAnnotation("displayFolder");
            string myOutputDestination = function.GetAnnotation("outputDestination");

            if (completeMetadata)
            {
                if (string.IsNullOrEmpty(myOutputType))
                {
                    IList<string> selectionTypeOptions = new List<string> { "Table", "Column", "Measure", "None" };
                    myOutputType =
                        Fx.ChooseString(
                            OptionList: selectionTypeOptions,
                            label: "Choose output type for function" + function.Name,
                            customWidth: 600);
                    if (string.IsNullOrEmpty(myOutputType)) return emptyFunction;
                    function.SetAnnotation("outputType", myOutputType);
                }

                if (string.IsNullOrEmpty(myNameTemplate))
                {
                    myNameTemplate = Fx.GetNameFromUser(Prompt: "Enter output name template for function " + function.Name, "Name Template", nameTemplateDefault);
                    if (string.IsNullOrEmpty(myNameTemplate)) return emptyFunction;
                    function.SetAnnotation("nameTemplate", myNameTemplate);
                }
                if (string.IsNullOrEmpty(myFormatString))
                {
                    myFormatString = Fx.GetNameFromUser(Prompt: "Enter output format string for function " + function.Name, "Format String", formatStringDefault);
                    if (string.IsNullOrEmpty(myFormatString)) return emptyFunction;
                    function.SetAnnotation("formatString", myFormatString);

                }
                if (string.IsNullOrEmpty(myDisplayFolder))
                {
                    myDisplayFolder =
                        Fx.GetNameFromUser(
                            Prompt: "Enter output display folder for function " + function.Name,
                            Title: "Display Folder",
                            DefaultResponse: displayFolderDefault);

                    if (string.IsNullOrEmpty(myDisplayFolder)) return emptyFunction;
                    function.SetAnnotation("displayFolder", myDisplayFolder);
                }

                if (string.IsNullOrEmpty(myOutputDestination))
                {
                    if (myOutputType == "Table")
                    {
                        myOutputDestination = "Model";
                    }
                    else if (myOutputType == "Column" || myOutputType == "Measure")
                    {
                        myOutputDestination =
                            Fx.GetNameFromUser(
                                Prompt: "Enter Destination template for " + function.Name,
                                Title: "Destination",
                                DefaultResponse: destinationDefault);

                        if (string.IsNullOrEmpty(myOutputDestination)) return emptyFunction;
                        function.SetAnnotation("outputDestination", destinationDefault);
                    }
                }
            }

            var functionExtended = new FunctionExtended
            {
                Name = function.Name,
                Expression = function.Expression,
                Description = function.Description,
                Parameters = Parameters,
                OutputFormatString = myFormatString,
                OutputNameTemplate = myNameTemplate,
                OutputType = myOutputType,
                OutputDisplayFolder = myDisplayFolder,
                OutputDestination = myOutputDestination,
                OriginalFunction = function

            };

            return functionExtended;
        }

        
    }


    public class FunctionParameter
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Subtype { get; set; }
        public string ParameterMode { get; set; }
    }
