#r "Microsoft.VisualBasic"
//this script only executes in TE3
using System.Windows.Forms;

using Microsoft.VisualBasic;
using Dax.Analyzer;
using Dax.Formatter; 
using System.IO;
#if TE3
    ScriptHelper.WaitFormVisible = false;
#endif
Measure targetMeasure = Fx.GetSelectedMeasure(Selected.Measures);
if (targetMeasure == null) return;
var selectedItems = Fx.SelectCalculationItems(Model, "Select one calculation item per group");
if (selectedItems == null || selectedItems.Count == 0) return;
// Get only calculation groups that have selected items, ordered by precedence (descending)
var calcGroupsByPrecedence = Model.Tables
    .OfType<CalculationGroupTable>()
    .Where(cg => selectedItems.ContainsKey(cg.Name))
    .OrderByDescending(cg => cg.CalculationGroup.Precedence)
    .ToList();
string previousFormatExpression = @"SELECTEDMEASUREFORMATSTRING()";
string currentFormatExpression = @"";
foreach (var calcGroup in calcGroupsByPrecedence)
{
    string selectedItemName = selectedItems[calcGroup.Name];
    //Output($"Calc Group: {calcGroup.Name}, Precedence: {calcGroup.CalculationGroup.Precedence}, Selected Item: {selectedItemName}");
    // Get the actual calculation item object
    var calcItem = calcGroup.CalculationItems.FirstOrDefault(ci => ci.Name == selectedItemName);
    if (calcItem == null) { Error(String.Format("Calc Item {0} not found", selectedItemName)); return; }
    var calcItemFormatTokens = DaxTokenizer.Tokenize(calcItem.FormatStringExpression);
    string calcItemFormatReplaceExpression = String.Format(@"/*{0} - {1}*/" + Environment.NewLine,calcGroup.Name, calcItem.Name); 
    bool closeParenthesisFound = true;
    bool isSelectedMeasureArgs = false;
    //{0} SELECTEDMEASURE
    //{1} SELECTEDMEASURENAME
    //{2} SELECTEDMEASUREFORMATSTRING
    foreach (var calcItemFormatToken in calcItemFormatTokens)
    {
        if (!closeParenthesisFound)
        {
            if (calcItemFormatToken.Type == DaxToken.CLOSE_PARENS)
            {
                closeParenthesisFound = true;
            }
            if (isSelectedMeasureArgs)
            {
                if (calcItemFormatToken.Type == DaxToken.CLOSE_PARENS)
                {
                    calcItemFormatReplaceExpression += "}}";
                    isSelectedMeasureArgs = false;
                }
                else if (calcItemFormatToken.Type == DaxToken.COLUMN_OR_MEASURE)
                {
                    calcItemFormatReplaceExpression += @"""" + calcItemFormatToken.Text + @"""";
                }
                else if (calcItemFormatToken.Type == DaxToken.COMMA)
                {
                    calcItemFormatReplaceExpression += calcItemFormatToken.Text;
                }
            }
        }
        else if (calcItemFormatToken.Type == DaxToken.TABLE)
        {
            calcItemFormatReplaceExpression += "'" + calcItemFormatToken.Text + "'";
        }
        else if (calcItemFormatToken.Type == DaxToken.COLUMN_OR_MEASURE)
        {
            calcItemFormatReplaceExpression += "[" + calcItemFormatToken.Text + "]";
        }
        else if (calcItemFormatToken.Type == DaxToken.SELECTEDMEASURE)
        {
            calcItemFormatReplaceExpression += "{0}";
            closeParenthesisFound = false;
        }
        else if (calcItemFormatToken.Type == DaxToken.ISSELECTEDMEASURE)
        {
            calcItemFormatReplaceExpression += @"""{1}"" IN {{";
            isSelectedMeasureArgs = true;
            closeParenthesisFound = false;
        }
        else if (calcItemFormatToken.Type == DaxToken.SELECTEDMEASURENAME)
        {
            calcItemFormatReplaceExpression += @"""{1}""";
            closeParenthesisFound = false;
        }
        //selectedmeasureformatstring should stay as is in format expression
        //else if (calcItemFormatToken.Type == DaxToken.SELECTEDMEASUREFORMATSTRING)
        //{
        //    calcItemFormatReplaceExpression += @"{2}";
        //    closeParenthesisFound = false;
        //}
        else if (calcItemFormatToken.Type == DaxToken.STRING_LITERAL)
        {
            calcItemFormatReplaceExpression += @"""" + calcItemFormatToken.Text + @"""";
        }
        else
        {
            calcItemFormatReplaceExpression += calcItemFormatToken.Text;
        }
    }
    var expressionFormatTokens = DaxTokenizer.Tokenize(previousFormatExpression);
    currentFormatExpression = "";
    bool closeParenthesisFound2 = true;
    bool isColumn = false;
    foreach (var token in expressionFormatTokens)
    {
        // Check if previous token indicated this is a column
        if (!closeParenthesisFound2)
        {
            //'do nothing'
            if (token.Type == DaxToken.CLOSE_PARENS)
            {
                closeParenthesisFound2 = true;
            }
        } else if (isColumn)
        {
            currentFormatExpression += "[" + token.Text + "]";
            isColumn = false;
        }
        // Check for Table token followed by Column/Measure token
        else if (token.Type == DaxToken.TABLE)
        {
            Table table = Model.Tables.FirstOrDefault(t => t.Name == token.Text);
            if (table == null)
            {
                Error(token.Text + " table not found in the model.");
                return;
            }
            if (token.Next.Type == DaxToken.COLUMN_OR_MEASURE)
            {
                // Set flag to indicate next token is a column
                isColumn = true;
            }
            currentFormatExpression += table.DaxObjectFullName;
        }
        else if (token.Type == DaxToken.COLUMN_OR_MEASURE)
        {
            //currentFormatExpression += "[" + token.Text + "]";
            currentFormatExpression += String.Format(
                calcItemFormatReplaceExpression,
                targetMeasure.DaxObjectFullName,
                targetMeasure.Name);
            closeParenthesisFound2 = false;
        }
        else if(token.Type == DaxToken.SELECTEDMEASURE)
        {
            currentFormatExpression += "{0}";
            closeParenthesisFound2 = false;
        }
        else if (token.Type == DaxToken.SELECTEDMEASURENAME)
        {
            currentFormatExpression += @"""{1}""";
            closeParenthesisFound2 = false;
        }
        else if (token.Type == DaxToken.SELECTEDMEASUREFORMATSTRING)
        {
            currentFormatExpression += String.Format(
                calcItemFormatReplaceExpression,
                targetMeasure.DaxObjectFullName,
                targetMeasure.Name);
            closeParenthesisFound2 = false;
        }
        else if (token.Type == DaxToken.STRING_LITERAL)
        {
            currentFormatExpression += @"""" + token.Text + @"""";
        }
        else
        {
            currentFormatExpression += token.Text;
        }
    }
    previousFormatExpression = currentFormatExpression;
}
//replace selectedmeasureformatstring by the format string of the target measure now 
previousFormatExpression = currentFormatExpression;
var finalTokens = DaxTokenizer.Tokenize(previousFormatExpression);
currentFormatExpression = "";
bool closeParenthesisFound3 = true;
foreach (var token in finalTokens)
{
    if (!closeParenthesisFound3)
    {
        if (token.Type == DaxToken.CLOSE_PARENS)
        {
            closeParenthesisFound3 = true;
        }
    } else if (token.Type == DaxToken.SELECTEDMEASUREFORMATSTRING)
    {
        if (targetMeasure.FormatStringExpression != null)
        {
            currentFormatExpression += targetMeasure.FormatStringExpression;
        }
        else
        {
            currentFormatExpression += @"""" + targetMeasure.FormatString + @"""";
        }
        closeParenthesisFound3 = false;
    }
    else if (token.Type == DaxToken.STRING_LITERAL)
    {
        currentFormatExpression += @"""" + token.Text + @"""";
    }
    else if (token.Type == DaxToken.TABLE)
    {
        currentFormatExpression += "'" + token.Text + "'";
    }
    else if (token.Type == DaxToken.COLUMN_OR_MEASURE)
    {
        currentFormatExpression += "[" + token.Text + "]";
    }
    else
    {
        currentFormatExpression += token.Text;
    }
}
currentFormatExpression = FormatDax(currentFormatExpression);
string fileName = targetMeasure.Name + "_" + calcGroupsByPrecedence.Select(cg=> cg.Name + " " + selectedItems[cg.Name]).Aggregate((a,b) => a + "_" + b) + ".dax";
string tempFile = Path.Combine(Path.GetTempPath(), "AppliedCalcGroupFormatExpression_" + fileName + ".dax");
File.WriteAllText(tempFile, currentFormatExpression);
System.Diagnostics.Process.Start("notepad.exe", tempFile);
Info("Expansion complete. Result saved to: " + tempFile);

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
        IList<string> measureNames = measures.Select(m => m.DaxObjectFullName).ToList();
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
