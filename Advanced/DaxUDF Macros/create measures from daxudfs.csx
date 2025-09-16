#r "Microsoft.VisualBasic"
using System.Windows.Forms;

using Microsoft.VisualBasic;

//2025-09-16/B.Agullo/
//Creates measures based on DAX UDFs 
//Check the blog post for futher information: https://www.esbrina-ba.com/automatically-create-measures-with-dax-user-defined-functions/

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
    (IList<string> Values,string Type) selectedObjectsForParam = Fx.SelectAnyObjects(
        Model,
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
    IList<string> previousList = new List<string>() { func.Name + "(" };
    IList<string> previousListNames = new List<string>() { func.OutputNameTemplate };
    IList<string> currentList = new List<string>();
    IList<string> currentListNames = new List<string>();
    string delimiter = "";
    IList<Table> previousDestinationTables = new List<Table>();
    IList<Table> currentDestinationTables = new List<Table>();
    IList<string> previousDisplayFolders = new List<string>();
    IList<string> currentDisplayFolders = new List<string>();
    IList<string> currentFormatStrings = new List<string>();
    IList<string> previousFormatStrings = new List<string>();
    // When iterating the parameters of this specific function, use the mapping created for distinct parameters.
    foreach (var param in func.Parameters)
    {
        currentList = new List<string>(); //reset current list
        currentListNames = new List<string>();
        currentDestinationTables = new List<Table>();
        currentFormatStrings = new List<string>();
        // Retrieve the objects list for this parameter name from the map (prompting was done earlier)
        (IList<string> Values, string Type) paramObject;
        if (!parameterObjectsMap.TryGetValue(param.Name, out paramObject) || paramObject.Type == null || paramObject.Values.Count == 0)
        {
            Error(String.Format("No objects were selected earlier for parameter '{0}'.", param.Name));
            return;
        }
        bool destinationSet = false;
        for (int i = 0; i < previousList.Count; i++)
        {
            string s = previousList[i];
            string sName = previousListNames[i];
            foreach (var o in paramObject.Values)
            {
                //extract original name and format string if the parameter is a measure
                string paramRawName = o;
                string paramFormatString = "";
                string paramDisplayFolder = "";
                if (paramObject.Type == "Measure")
                {
                    Measure m = Model.AllMeasures.FirstOrDefault(m => m.DaxObjectFullName == o);
                    paramRawName = m.Name;
                    paramFormatString = m.FormatString;
                    paramDisplayFolder = m.DisplayFolder;
                }else if (paramObject.Type == "Column")
                {
                    Column c = Model.AllColumns.FirstOrDefault(c => c.DaxObjectFullName == o);
                    paramRawName = c.Name;
                    paramFormatString = c.FormatString;
                    paramDisplayFolder = c.DisplayFolder;
                }
                else if (paramObject.Type == "Table")
                {
                    Table t = Model.Tables.FirstOrDefault(t => t.Name == o);
                    paramRawName = t.Name;
                    paramFormatString = "";
                    paramDisplayFolder = "";
                }
                if (destinationSet == false)
                {
                    Table destinationTable = null;
                    if (param.Name.ToUpper().Contains("MEASURE"))
                    {
                        var m = Model.AllMeasures.FirstOrDefault(me => me.DaxObjectFullName == o);
                        if (m != null) destinationTable = m.Table;
                    }
                    else
                    {
                        destinationTable = SelectTable(label: "Select destination table for " + func.OutputType + "(s) created for " + o);
                    }
                    currentDestinationTables.Add(destinationTable);
                    string displayFolder = "";
                    if (func.OutputDisplayFolder != null)
                    {
                        displayFolder = 
                            func.OutputDisplayFolder
                                .Replace(param.Name + "Name", paramRawName)
                                .Replace(param.Name + "DisplayFolder", paramDisplayFolder);
                    }
                    ;
                    currentDisplayFolders.Add(displayFolder);
                }
                else
                {
                    currentDestinationTables.Add(previousDestinationTables[i]);
                    currentDisplayFolders.Add(previousDisplayFolders[i]);
                }
                currentList.Add(s + delimiter + o);
                currentListNames.Add(sName.Replace(param.Name + "Name", paramRawName));
                currentFormatStrings.Add(
                    func.OutputFormatString.Replace(
                        param.Name + "FormatString",paramFormatString));
            }
            destinationSet = true;
        }
        delimiter = ", ";
        previousList = currentList;
        previousListNames = currentListNames;
        previousDestinationTables = currentDestinationTables;
        previousDisplayFolders = currentDisplayFolders;
        previousFormatStrings = currentFormatStrings;
    }
    if (func.OutputType == "Measure")
    {
        for (int i = 0; i < currentList.Count; i++)
        {
            Measure measure = currentDestinationTables[i].AddMeasure(currentListNames[i], currentList[i] + ")");
            measure.FormatDax();
            measure.Description = String.Format("Measure created with {0} function. Check function for details.", func.Name);
            measure.DisplayFolder = currentDisplayFolders[i];
            measure.FormatString = currentFormatStrings[i];
        }
    }else
    {
        Info("Not implemented yet for output types other than Measure.");
    }
}

public static class Fx
{
    public static Table CreateCalcTable(Model model, string tableName, string tableExpression)
    {
        return model.Tables.FirstOrDefault(t =>
                            string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase)) //case insensitive search
                            ?? model.AddCalculatedTable(tableName, tableExpression);
    }
    public static string GetNameFromUser(string Prompt, string Title, string DefaultResponse)
    {
        string response = Interaction.InputBox(Prompt, Title, DefaultResponse, 740, 400);
        return response;
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
        switch (selectionType)
        {
            case "Table":
                return (Values:SelectTableMultiple(model, label: prompt2), Type:selectionType);
            case "Column":
                return (Values:SelectColumnMultiple(model, label: prompt2), Type: selectionType);
            case "Measure":
                return (Values: SelectMeasureMultiple(model: model, label: prompt2), Type: selectionType);
            //case "Scalar":
            //    selectedType = "Scalar";
            //    return GetNameFromUser("Enter scalar value", "Scalar value", "0");
            default:
                Error("Invalid selection type");
                return returnEmpty;
        }
    }
    public static string ChooseString(IList<string> OptionList, string label = "Choose item", int customWidth = 400, int customHeight = 500)
    {
        return ChooseStringInternal(OptionList, MultiSelect: false, label: label, customWidth: customWidth, customHeight:customHeight) as string;
    }
    public static List<string> ChooseStringMultiple(IList<string> OptionList, string label = "Choose item(s)", int customWidth = 400, int customHeight = 500)
    {
        return ChooseStringInternal(OptionList, MultiSelect:true, label:label, customWidth: customWidth, customHeight: customHeight) as List<string>;
    }
    private static object ChooseStringInternal(IList<string> OptionList, bool MultiSelect, string label = "Choose item(s)", int customWidth = 400, int customHeight = 500)
    {
        Form form = new Form
        {
            Text =label,
            Width = customWidth,
            Height = customHeight,
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
            Height = 40,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(10)
        };
        Button selectAllButton = new Button { Text = "Select All", Visible = MultiSelect };
        Button selectNoneButton = new Button { Text = "Select None", Visible = MultiSelect };
        Button okButton = new Button { Text = "OK", DialogResult = DialogResult.OK };
        Button cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };
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
    public static IList<string> SelectMeasureMultiple(Model model, IEnumerable<Measure> measures = null, string label = "Select Measure(s)", int customWidth = 400)
    {
        measures ??= model.AllMeasures;
        IList<string> measureNames = measures.Select(m => m.DaxObjectFullName).ToList();
        IList<string> selectedMeasureNames = ChooseStringMultiple(measureNames, label: label, customWidth:customWidth);
        return selectedMeasureNames; 
    }
    public static IList<string> SelectColumnMultiple(Model model, IEnumerable<Column> columns = null, string label = "Select Columns(s)", int customWidth = 400)
    {
        columns ??= model.AllColumns;
        IList<string> columnNames = columns.Select(m => m.DaxObjectFullName).ToList();
        IList<string> selectedColumnNames = ChooseStringMultiple(columnNames, label: label, customWidth: customWidth);
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
                fp.Name = nameParams.Length > 0 ? nameParams[0] : param;
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
                else
                    paramMode = "VAL";
                fp.ParameterMode = paramMode;

                paramList.Add(fp);
            }

            return paramList;
        }
        public static FunctionExtended CreateFunctionExtended(Function function)
        {

            string myOutputType = function.GetAnnotation("outputType"); 
            string myNameTemplate = function.GetAnnotation("nameTemplate");
            string myFormatString = function.GetAnnotation("formatString");
            string myDisplayFolder = function.GetAnnotation("displayFolder");

            if(string.IsNullOrEmpty(myOutputType))
            {
                IList<string> selectionTypeOptions = new List<string> { "Table", "Column", "Measure", "None" };
                myOutputType = Fx.ChooseString(selectionTypeOptions, label: "Choose output type for function" + function.Name, customWidth: 600);
            }

            if (string.IsNullOrEmpty(myNameTemplate))
            {
                myNameTemplate = Fx.GetNameFromUser(Prompt:"Enter output name template for function " + function.Name,  "Name Template", "(use parameternameName as placeholder)");
            }
            if(string.IsNullOrEmpty(myFormatString))
            {
                myFormatString = Fx.GetNameFromUser(Prompt: "Enter output format string for function " + function.Name, "Format String", "(use parameternameFormatString as placeholder)");
            }
            if(string.IsNullOrEmpty(myDisplayFolder))
            {
                myDisplayFolder = Fx.GetNameFromUser(Prompt: "Enter output display folder for function " + function.Name, "Display Folder", "(use parameternameName and parameternameDisplayFolder as placeholder)");
            }

            var functionExtended = new FunctionExtended
            {
                Name = function.Name,
                Expression = function.Expression,
                Description = function.Description,
                Parameters = ExtractParametersFromExpression(function.Expression),
                OutputFormatString = function.GetAnnotation("formatString"),
                OutputNameTemplate = function.GetAnnotation("nameTemplate"),
                OutputType = function.GetAnnotation("outputType"),
                OutputDisplayFolder = function.GetAnnotation("displayFolder"),
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
