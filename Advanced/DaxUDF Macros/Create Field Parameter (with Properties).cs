#r "Microsoft.VisualBasic"
using System.Windows.Forms;

using Microsoft.VisualBasic;
using Microsoft.VisualBasic;
//80% comes from Daniel Otykier --> https://github.com/TabularEditor/TabularEditor3/issues/541#issuecomment-1129228481
//20% B.Agullo --> pop-up to choose parameter name + all the properties annotation support (split by pipe delimiter and added as additional columns in the field parameter table)
// Before running the script, select the measures or columns that you
// would like to use as field parameters (hold down CTRL to select multiple
// objects). Also, you may change the name of the field parameter table
// below. NOTE: If used against Power BI Desktop, you must enable unsupported
// features under File > Preferences (TE2) or Tools > Preferences (TE3).
var name = Interaction.InputBox("Provide the name for the field parameter", "Parameter name", "Parameter");
if (Selected.Columns.Count == 0 && Selected.Measures.Count == 0) throw new Exception("No columns or measures selected!");
// Get objects collection
var objects = Selected.Columns.Any() ? Selected.Columns.Cast<ITabularTableObject>() : Selected.Measures;
if (Fx.IsAnswerYes("Do you want to order measures/columns by name?"))
{
    objects = objects.OrderBy(o => o.Name).ToList();
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
    string columnName = name + " Property" + (i + 1);
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
    public static Measure GetSelectedMeasure(IEnumerable<Measure> measures, string label = "Select Measure")
    {
        Measure selectedMeasure = null;
        if (measures.Count() == 1)
        {
            selectedMeasure = measures.First();
        }
        else
        {
            selectedMeasure = SelectMeasure(preselect: measures.First(), label: label);
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
