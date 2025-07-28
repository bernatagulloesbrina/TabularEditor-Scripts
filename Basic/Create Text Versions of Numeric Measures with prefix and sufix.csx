#r "Microsoft.VisualBasic"
using System.Windows.Forms;
using Microsoft.VisualBasic;

//2025-07-28/B.Agullo
//This script creates text measures based on the selected measures in the model.
//It prompts the user for a prefix and suffix to be added to the text measures.
//It also allows the user to specify a suffix for the names of the new text measures.
if (Selected.Measures.Count() == 0)
{
    Error("No measures selected. Please select at least one measure.");
    return;
}
// Ask user for prefix
string prefix = Fx.GetNameFromUser(
    Prompt: "Enter a prefix for the new text measures:",
    Title: "Text Measure Prefix",
    DefaultResponse: ""
);
if (prefix == null) return;
// Ask user for suffix
string suffix = Fx.GetNameFromUser(
    Prompt: "Enter a suffix for the new text measures:",
    Title: "Text Measure Suffix",
    DefaultResponse: ""
);
if (suffix == null) return;
// Ask user for measure name suffix
string measureNameSuffix = Fx.GetNameFromUser(
    Prompt: "Enter a suffix for the Name of the new text measures:",
    Title: "Suffix for names!",
    DefaultResponse: " Text"
);
if (measureNameSuffix == null) return;
foreach (Measure m in Selected.Measures)
{
    string newMeasureName = m.Name + measureNameSuffix;
    string newMeasureDisplayFolder = (m.DisplayFolder + measureNameSuffix).Trim();
    string newMeasureExpression = String.Format(@"""{2}"" & FORMAT([{0}], ""{1}"") & ""{3}""", m.Name, m.FormatString, prefix, suffix);
    Measure newMeasure = m.Table.AddMeasure(newMeasureName, newMeasureExpression,newMeasureDisplayFolder);
    newMeasure.FormatDax();
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
}
