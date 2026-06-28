// 2026-06-28 / B.Agullo / Grid editor to manage Properties / PropertyNames annotations across selected measures.
//                         Rows = measures, columns = property names. Supports column rename (with merge on
//                         duplicate name), adding columns, and copy/paste. On save every measure stores the
//                         full column list in PropertyNames (empty segments where it has no value); completely
//                         empty columns are dropped and '|' characters are sanitised out of names/values.
using System.Windows.Forms;
using System.Drawing;
#if TE3
ScriptHelper.WaitFormVisible = false;
#endif
const string propertiesAnnotationLabel = "Properties";
const string propertyNamesAnnotationLabel = "PropertyNames";
// Rows of the grid: only the selected measures
List<Measure> orderedMeasures = Selected.Measures.ToList();
if (orderedMeasures.Count == 0)
{
    Error("Select one or more measures and try again.");
    return;
}
// Backing model: ordered column names + per-measure (columnName -> value) dictionary
List<string> columnNames = new List<string>();
Dictionary<Measure, Dictionary<string, string>> values = new Dictionary<Measure, Dictionary<string, string>>();
bool sanitizedAny = false;
string Sanitize(string s)
{
    if (s == null) return "";
    if (s.IndexOf('|') >= 0)
    {
        sanitizedAny = true;
        s = s.Replace("|", "/");
    }
    return s;
}
// Initialise backing model from existing annotations
foreach (Measure measure in orderedMeasures)
{
    string propertiesAnnotation = measure.GetAnnotation(propertiesAnnotationLabel) ?? "";
    string propertyNamesAnnotation = measure.GetAnnotation(propertyNamesAnnotationLabel) ?? "";
    string[] props = propertiesAnnotation.Length == 0 ? new string[0] : propertiesAnnotation.Split('|');
    string[] names = propertyNamesAnnotation.Length == 0 ? new string[0] : propertyNamesAnnotation.Split('|');
    int count = Math.Max(props.Length, names.Length);
    Dictionary<string, string> measureValues = new Dictionary<string, string>();
    for (int i = 0; i < count; i++)
    {
        string name = (i < names.Length && !string.IsNullOrWhiteSpace(names[i]))
            ? names[i]
            : String.Format("Property{0}", i + 1);
        string value = i < props.Length ? props[i] : "";
        if (!columnNames.Contains(name)) columnNames.Add(name);
        measureValues[name] = value; // duplicate name within a measure: last value wins
    }
    values[measure] = measureValues;
}
// Build the dialog
Form form = new Form();
form.Text = "Manage Property Annotations";
form.Width = 900;
form.Height = 520;
form.StartPosition = FormStartPosition.CenterScreen;
System.Windows.Forms.Label info = new System.Windows.Forms.Label();
info.Dock = DockStyle.Top;
info.Height = 40;
info.Padding = new System.Windows.Forms.Padding(8, 6, 8, 0);
info.Text = "Each row is a measure, each column a property name. Double-click a column header (or use 'Rename Column') to rename it; "
    + "renaming to an existing column merges them. Use 'Add Column' to add a property. Copy/paste with Ctrl+C / Ctrl+V.";
DataGridView grid = new DataGridView();
grid.Dock = DockStyle.Fill;
grid.AllowUserToAddRows = false;
grid.AllowUserToDeleteRows = false;
grid.RowHeadersVisible = false;
grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
grid.MultiSelect = true;
grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
// (Re)build the grid from the backing model
void RebuildGrid()
{
    grid.Columns.Clear();
    grid.Rows.Clear();
    DataGridViewTextBoxColumn measureColumn = new DataGridViewTextBoxColumn();
    measureColumn.HeaderText = "Measure";
    measureColumn.ReadOnly = true;
    measureColumn.Frozen = true;
    measureColumn.Width = 220;
    measureColumn.DefaultCellStyle.BackColor = System.Drawing.SystemColors.Control;
    grid.Columns.Add(measureColumn);
    foreach (string columnName in columnNames)
    {
        DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
        col.HeaderText = columnName;
        col.Width = 140;
        grid.Columns.Add(col);
    }
    foreach (Measure measure in orderedMeasures)
    {
        int rowIndex = grid.Rows.Add();
        DataGridViewRow row = grid.Rows[rowIndex];
        row.Cells[0].Value = measure.Name;
        Dictionary<string, string> measureValues = values[measure];
        for (int c = 0; c < columnNames.Count; c++)
        {
            string columnName = columnNames[c];
            string value;
            row.Cells[c + 1].Value = measureValues.TryGetValue(columnName, out value) ? value : "";
        }
    }
}
// Read the grid back into the backing model (grid columns currently match columnNames 1:1)
void SyncGridToModel()
{
    for (int r = 0; r < grid.Rows.Count; r++)
    {
        Measure measure = orderedMeasures[r];
        Dictionary<string, string> measureValues = new Dictionary<string, string>();
        for (int c = 0; c < columnNames.Count; c++)
        {
            string columnName = columnNames[c];
            object cellValue = grid.Rows[r].Cells[c + 1].Value;
            string value = Sanitize(cellValue == null ? "" : cellValue.ToString());
            if (measureValues.ContainsKey(columnName))
            {
                // merged (duplicate header) column: keep the first non-empty value
                if (string.IsNullOrEmpty(measureValues[columnName]) && !string.IsNullOrEmpty(value))
                    measureValues[columnName] = value;
            }
            else
            {
                measureValues[columnName] = value;
            }
        }
        values[measure] = measureValues;
    }
}
// Small input prompt (self-contained, no Error popup on cancel)
string PromptText(string title, string prompt, string defaultValue)
{
    using (Form inputForm = new Form())
    {
        inputForm.Text = title;
        inputForm.Width = 420;
        inputForm.Height = 170;
        inputForm.StartPosition = FormStartPosition.CenterParent;
        inputForm.FormBorderStyle = FormBorderStyle.FixedDialog;
        inputForm.MinimizeBox = false;
        inputForm.MaximizeBox = false;
        System.Windows.Forms.Label label = new System.Windows.Forms.Label();
        label.Text = prompt;
        label.Left = 10;
        label.Top = 10;
        label.Width = 390;
        TextBox textBox = new TextBox();
        textBox.Left = 10;
        textBox.Top = 35;
        textBox.Width = 390;
        textBox.Text = defaultValue;
        Button okButton = new Button();
        okButton.Text = "OK";
        okButton.DialogResult = DialogResult.OK;
        okButton.Left = 230;
        okButton.Top = 75;
        okButton.Width = 80;
        okButton.Height = 30;
        Button cancelButton = new Button();
        cancelButton.Text = "Cancel";
        cancelButton.DialogResult = DialogResult.Cancel;
        cancelButton.Left = 320;
        cancelButton.Top = 75;
        cancelButton.Width = 80;
        cancelButton.Height = 30;
        inputForm.Controls.Add(label);
        inputForm.Controls.Add(textBox);
        inputForm.Controls.Add(okButton);
        inputForm.Controls.Add(cancelButton);
        inputForm.AcceptButton = okButton;
        inputForm.CancelButton = cancelButton;
        textBox.SelectAll();
        return inputForm.ShowDialog() == DialogResult.OK ? textBox.Text : null;
    }
}
// Rename the column at gridColumnIndex (1-based against the data columns, i.e. excludes Measure column)
void RenameColumn(int gridColumnIndex)
{
    if (gridColumnIndex <= 0 || gridColumnIndex > columnNames.Count) return;
    SyncGridToModel();
    int modelIndex = gridColumnIndex - 1;
    string oldName = columnNames[modelIndex];
    string newNameRaw = PromptText("Rename Column", "Enter the new property name:", oldName);
    if (newNameRaw == null) return;
    string newName = Sanitize(newNameRaw).Trim();
    if (string.IsNullOrEmpty(newName) || newName == oldName) return;
    // Update the per-measure values: rename oldName -> newName, merging when needed
    foreach (Measure measure in orderedMeasures)
    {
        Dictionary<string, string> measureValues = values[measure];
        string oldValue;
        bool hadOld = measureValues.TryGetValue(oldName, out oldValue);
        measureValues.Remove(oldName);
        if (measureValues.ContainsKey(newName))
        {
            if (string.IsNullOrEmpty(measureValues[newName]) && hadOld && !string.IsNullOrEmpty(oldValue))
                measureValues[newName] = oldValue;
        }
        else
        {
            measureValues[newName] = hadOld ? oldValue : "";
        }
    }
    // Update the column list: merge if the new name already exists elsewhere
    bool existsElsewhere = false;
    for (int i = 0; i < columnNames.Count; i++)
    {
        if (i != modelIndex && columnNames[i] == newName) { existsElsewhere = true; break; }
    }
    if (existsElsewhere)
        columnNames.RemoveAt(modelIndex);
    else
        columnNames[modelIndex] = newName;
    RebuildGrid();
}
// Paste clipboard (tab/newline delimited) starting at the current cell
void PasteClipboard()
{
    if (!Clipboard.ContainsText()) return;
    string text = Clipboard.GetText();
    if (string.IsNullOrEmpty(text)) return;
    string[] lines = text.Replace("\r\n", "\n").Replace("\r", "\n").TrimEnd('\n').Split('\n');
    // Single copied value pasted across a multi-cell selection: fill every selected (editable) cell
    bool singleValue = lines.Length == 1 && lines[0].IndexOf('\t') < 0;
    if (singleValue && grid.SelectedCells.Count > 1)
    {
        string single = Sanitize(lines[0]);
        foreach (DataGridViewCell selectedCell in grid.SelectedCells)
        {
            if (selectedCell.ColumnIndex == 0) continue; // skip the read-only Measure column
            selectedCell.Value = single;
        }
        return;
    }
    int startRow = grid.CurrentCell != null ? grid.CurrentCell.RowIndex : 0;
    int startCol = grid.CurrentCell != null ? grid.CurrentCell.ColumnIndex : 1;
    if (startCol < 1) startCol = 1; // never write into the read-only Measure column
    for (int r = 0; r < lines.Length; r++)
    {
        int rowIdx = startRow + r;
        if (rowIdx >= grid.Rows.Count) break;
        string[] cells = lines[r].Split('\t');
        for (int c = 0; c < cells.Length; c++)
        {
            int colIdx = startCol + c;
            if (colIdx >= grid.Columns.Count) break;
            if (colIdx == 0) continue;
            grid.Rows[rowIdx].Cells[colIdx].Value = Sanitize(cells[c]);
        }
    }
}
grid.ColumnHeaderMouseDoubleClick += (sender, e) =>
{
    if (e.ColumnIndex >= 1) RenameColumn(e.ColumnIndex);
};
grid.KeyDown += (sender, e) =>
{
    if (e.Control && e.KeyCode == Keys.V)
    {
        PasteClipboard();
        e.Handled = true;
    }
};
// Button bar
Panel buttonPanel = new Panel();
buttonPanel.Dock = DockStyle.Bottom;
buttonPanel.Height = 52;
Button addColumnButton = new Button();
addColumnButton.Text = "Add Column";
addColumnButton.Left = 8;
addColumnButton.Top = 10;
addColumnButton.Width = 110;
addColumnButton.Height = 32;
addColumnButton.Click += (sender, e) =>
{
    SyncGridToModel();
    string defaultName = String.Format("Property{0}", columnNames.Count + 1);
    string nameRaw = PromptText("Add Column", "Enter the new property name:", defaultName);
    if (nameRaw == null) return;
    string name = Sanitize(nameRaw).Trim();
    if (string.IsNullOrEmpty(name)) return;
    if (columnNames.Contains(name))
    {
        MessageBox.Show("A column with that name already exists.", "Add Column", MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
    }
    columnNames.Add(name);
    foreach (Measure measure in orderedMeasures) values[measure][name] = "";
    RebuildGrid();
};
Button renameColumnButton = new Button();
renameColumnButton.Text = "Rename Column";
renameColumnButton.Left = 126;
renameColumnButton.Top = 10;
renameColumnButton.Width = 110;
renameColumnButton.Height = 32;
renameColumnButton.Click += (sender, e) =>
{
    int colIndex = grid.CurrentCell != null ? grid.CurrentCell.ColumnIndex : -1;
    if (colIndex >= 1) RenameColumn(colIndex);
    else MessageBox.Show("Select a cell in the column you want to rename.", "Rename Column", MessageBoxButtons.OK, MessageBoxIcon.Information);
};
Button saveButton = new Button();
saveButton.Text = "Save";
saveButton.DialogResult = DialogResult.OK;
saveButton.Width = 90;
saveButton.Top = 10;
saveButton.Height = 32;
Button cancelButton2 = new Button();
cancelButton2.Text = "Cancel";
cancelButton2.DialogResult = DialogResult.Cancel;
cancelButton2.Width = 90;
cancelButton2.Top = 10;
cancelButton2.Height = 32;
buttonPanel.Resize += (sender, e) =>
{
    cancelButton2.Left = buttonPanel.Width - cancelButton2.Width - 8;
    saveButton.Left = cancelButton2.Left - saveButton.Width - 8;
};
buttonPanel.Controls.Add(addColumnButton);
buttonPanel.Controls.Add(renameColumnButton);
buttonPanel.Controls.Add(saveButton);
buttonPanel.Controls.Add(cancelButton2);
form.Controls.Add(grid);
form.Controls.Add(info);
form.Controls.Add(buttonPanel);
form.AcceptButton = saveButton;
form.CancelButton = cancelButton2;
RebuildGrid();
cancelButton2.Left = buttonPanel.Width - cancelButton2.Width - 8;
saveButton.Left = cancelButton2.Left - saveButton.Width - 8;
if (form.ShowDialog() != DialogResult.OK) return;
// Persist: read final state, drop completely-empty columns, write full column list to every measure
SyncGridToModel();
List<string> keptColumns = new List<string>();
foreach (string columnName in columnNames)
{
    bool anyValue = orderedMeasures.Any(m =>
    {
        string v;
        return values[m].TryGetValue(columnName, out v) && !string.IsNullOrEmpty(v);
    });
    if (anyValue) keptColumns.Add(columnName);
}
int updatedCount = 0;
foreach (Measure measure in orderedMeasures)
{
    if (keptColumns.Count == 0)
    {
        measure.RemoveAnnotation(propertiesAnnotationLabel);
        measure.RemoveAnnotation(propertyNamesAnnotationLabel);
        continue;
    }
    Dictionary<string, string> measureValues = values[measure];
    List<string> valuesList = new List<string>();
    foreach (string columnName in keptColumns)
    {
        string v;
        valuesList.Add(Sanitize(measureValues.TryGetValue(columnName, out v) ? v : ""));
    }
    string propertyNamesAnnotation = string.Join("|", keptColumns.Select(c => Sanitize(c)));
    string propertiesAnnotation = string.Join("|", valuesList);
    measure.SetAnnotation(propertyNamesAnnotationLabel, propertyNamesAnnotation);
    measure.SetAnnotation(propertiesAnnotationLabel, propertiesAnnotation);
    updatedCount++;
}
string summary = keptColumns.Count == 0
    ? String.Format("Cleared Property annotations on {0} measure(s).", orderedMeasures.Count)
    : String.Format("Updated {0} measure(s) with {1} column(s): {2}", updatedCount, keptColumns.Count, string.Join(", ", keptColumns));
if (sanitizedAny)
    summary += Environment.NewLine + Environment.NewLine + "Note: '|' characters are reserved as separators and were replaced with '/'.";
Info(summary);
