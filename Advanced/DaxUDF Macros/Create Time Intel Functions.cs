#r "Microsoft.VisualBasic
using System.Windows.Forms;

using Microsoft.VisualBasic;
// 2025-09-29/B.Agullo
// Creates Time Intelligence functions (CY, PY, YOY, YOYPCT) in the model if they do not exist.
// It also creates a hidden calculated column and measure in the date table to handle cases where the fact table has no data for some dates.
// The script assumes there is a date table and a fact table in the model.
// The script will prompt the user to select the main date column in the fact table if there are multiple date columns.
// More details at: https://www.esbrina-ba.com/industrializing-model-dependent-dax-udfs-with-tabular-editor-c-scripting-time-intel-reloaded/
if(Model.Database.CompatibilityLevel < 1702)
{
    if(Fx.IsAnswerYes("The model compatibility level is below 1702. Time Intelligence functions are only supported in 1702 or higher. Do to change the compatibility level to 1702?"))
    {
        Model.Database.CompatibilityLevel = 1702;
    }
    else
    {
        Info("Operation cancelled.");
        return;
    }
}
Table dateTable = Fx.GetDateTable(model: Model);
if (dateTable == null) return;
Column dateColumn = Fx.GetDateColumn(dateTable);
if (dateColumn == null) return;
Table factTable = Fx.GetFactTable(model: Model);
if (factTable == null) return;
Column factTableDateColumn = null; 
IEnumerable<Column> factTableDateColumns =
    factTable.Columns.Where(
        c => c.DataType == DataType.DateTime); 
if(factTableDateColumns.Count() == 0)
{
    Error("No Date columns found in fact table " + factTable.Name);
    return;
}
if(factTableDateColumns.Count() == 1)
{
    factTableDateColumn = factTableDateColumns.First();
}
else
{
    factTableDateColumn = factTableDateColumns.First(
        c=> Model.Relationships.Any(
            r => ((r.FromColumn == dateColumn && r.ToColumn == c)
                || (r.ToColumn == dateColumn && r.FromColumn == c)
                    && r.IsActive)));
    factTableDateColumn = SelectColumn(factTableDateColumns, factTableDateColumn, "Select main date column from the fact table"); 
}
if(factTableDateColumn == null) return;
string dateTableAuxColumnName = "DateWith" + factTable.Name.Replace(" ", "");
string dateTableAuxColumnExpression = String.Format(@"{0} <= MAX({1})", dateColumn.DaxObjectFullName, factTableDateColumn.DaxObjectFullName);
CalculatedColumn dateTableAuxColumn = dateTable.AddCalculatedColumn(dateTableAuxColumnName, dateTableAuxColumnExpression);
dateTableAuxColumn.FormatDax(); 
dateTableAuxColumn.IsHidden = true;
string dateTableAuxMeasureName = "ShowValueForDates";
string dateTableAuxMeasureExpression =
    String.Format(
        @"VAR LastDateWithData =
            CALCULATE (
                MAX ( {0} ),
                REMOVEFILTERS ()
            )
        VAR FirstDateVisible =
            MIN ( {1} )
        VAR Result =
            FirstDateVisible <= LastDateWithData
        RETURN
            Result",
        factTableDateColumn.DaxObjectFullName,
        dateColumn.DaxObjectFullName);
Measure dateTableAuxMeasure = dateTable.AddMeasure(dateTableAuxMeasureName, dateTableAuxMeasureExpression);
dateTableAuxMeasure.IsHidden = true;
dateTableAuxMeasure.FormatDax();
//CY --just for the sake of completion 
string CYfunctionName = "Local.TimeIntel.CY";
string CYfunctionExpression = "(baseMeasure) => baseMeasure";
Function CYfunction = Model.AddFunction(CYfunctionName);
CYfunction.Expression = CYfunctionExpression;
CYfunction.FormatDax();
CYfunction.SetAnnotation("displayFolder", @"baseMeasureDisplayFolder\baseMeasureName TimeIntel");
CYfunction.SetAnnotation("formatString", "baseMeasureFormatStringFull");
CYfunction.SetAnnotation("outputType", "Measure");
CYfunction.SetAnnotation("nameTemplate", "baseMeasureName CY");
CYfunction.SetAnnotation("outputDestination", "baseMeasureTable");
//PY
string PYfunctionName = "Local.TimeIntel.PY";
string PYfunctionExpression = 
    String.Format(
        @"(baseMeasure: ANYREF) =>
        IF(
            {0},
            CALCULATE(         
                baseMeasure,
                CALCULATETABLE(
                    DATEADD(
                        {1},
                        -1,
                        YEAR
                    ),
                    {2} = TRUE
                )
            )
        )",
        dateTableAuxMeasure.DaxObjectFullName,
        dateColumn.DaxObjectFullName,
        dateTableAuxColumn.DaxObjectFullName);
Function PYfunction = Model.AddFunction(PYfunctionName);
PYfunction.Expression = PYfunctionExpression;
PYfunction.FormatDax();
PYfunction.SetAnnotation("displayFolder", @"baseMeasureDisplayFolder\baseMeasureName TimeIntel");
PYfunction.SetAnnotation("formatString", "baseMeasureFormatStringFull");
PYfunction.SetAnnotation("outputType", "Measure");
PYfunction.SetAnnotation("nameTemplate", "baseMeasureName PY");
PYfunction.SetAnnotation("outputDestination", "baseMeasureTable");
//YOY
string YOYfunctionName = "Local.TimeIntel.YOY";
string YOYfunctionExpression =
    @"(baseMeasure: ANYREF) =>
    VAR ValueCurrentPeriod = Local.TimeIntel.CY(baseMeasure)
    VAR ValuePreviousPeriod = Local.TimeIntel.PY(baseMeasure)
    VAR Result =
	                IF(
		                NOT ISBLANK( ValueCurrentPeriod )
			                && NOT ISBLANK( ValuePreviousPeriod ),
		                ValueCurrentPeriod
			                - ValuePreviousPeriod
	                )
    RETURN
	                Result";
Function YOYfunction = Model.AddFunction(YOYfunctionName);
YOYfunction.Expression = YOYfunctionExpression;
YOYfunction.FormatDax();
YOYfunction.SetAnnotation("displayFolder", @"baseMeasureDisplayFolder\baseMeasureName TimeIntel");
YOYfunction.SetAnnotation("formatString", "+baseMeasureFormatStringRoot;-baseMeasureFormatStringRoot;-");
YOYfunction.SetAnnotation("outputType", "Measure");
YOYfunction.SetAnnotation("nameTemplate", "baseMeasureName YOY");
YOYfunction.SetAnnotation("outputDestination", "baseMeasureTable");
//YOY%
string YOYPfunctionName = "Local.TimeIntel.YOYPCT";
string YOYPfunctionExpression =
    @"(baseMeasure: ANYREF) =>
    VAR ValueCurrentPeriod = Local.TimeIntel.CY(baseMeasure)
    VAR ValuePreviousPeriod = Local.TimeIntel.PY(baseMeasure)
    VAR CurrentMinusPreviousPeriod =
	                IF(
		                NOT ISBLANK( ValueCurrentPeriod )
			                && NOT ISBLANK( ValuePreviousPeriod ),
		                ValueCurrentPeriod
			                - ValuePreviousPeriod
	                )
    VAR Result =
	                DIVIDE(
		                CurrentMinusPreviousPeriod,
		                ValuePreviousPeriod
	                )
    RETURN
	                Result";
Function YOYPfunction = Model.AddFunction(YOYPfunctionName);
YOYPfunction.Expression = YOYPfunctionExpression;
YOYPfunction.FormatDax();
YOYPfunction.SetAnnotation("displayFolder", @"baseMeasureDisplayFolder\baseMeasureName TimeIntel");
YOYPfunction.SetAnnotation("formatString", "+0.0%;-0.0%;-");
YOYPfunction.SetAnnotation("outputType", "Measure");
YOYPfunction.SetAnnotation("nameTemplate", "baseMeasureName YOY%");
YOYPfunction.SetAnnotation("outputDestination", "baseMeasureTable");

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
