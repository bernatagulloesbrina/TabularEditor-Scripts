#r "Microsoft.VisualBasic"
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.IO;
using Newtonsoft.Json.Linq;

// 2025-10-30 / B.Agullo
// Initializes a report, lets user select visuals, iterates through projections to extract display names,
// and stores them as "DisplayName" annotations in the model. Handles multiple display names by prompting user.
#if TE3
ScriptHelper.WaitFormVisible = false;
#endif
// Step 1: Initialize report
ReportExtended report = Rx.InitReport();
if (report == null) return;
// Step 2: Let user select visuals
IList<VisualExtended> selectedVisuals = Rx.SelectVisuals(report);
if (selectedVisuals == null || selectedVisuals.Count == 0)
{
    Info("No visuals selected.");
    return;
}
// Step 3: Collect display names from selected visuals
var measureDisplayNamesDict = new Dictionary<string, HashSet<string>>();
var columnDisplayNamesDict = new Dictionary<string, HashSet<string>>();
foreach (var visual in selectedVisuals)
{
    var queryState = visual.Content?.Visual?.Query?.QueryState;
    if (queryState == null) continue;
    // Create list of all projection sets to iterate
    var projectionSets = new List<VisualDto.ProjectionsSet>
    {
        queryState.Values,
        queryState.Y,
        queryState.Y2,
        queryState.Category,
        queryState.Series,
        queryState.Data,
        queryState.Rows
    };
    // Iterate through each projection set
    foreach (var projectionSet in projectionSets)
    {
        if (projectionSet?.Projections == null) continue;
        foreach (var projection in projectionSet.Projections)
        {
            if (projection?.Field == null) continue;
            string displayName = projection.DisplayName;
            if (string.IsNullOrEmpty(displayName)) continue;
            // Check if it's a measure
            if (projection.Field.Measure != null)
            {
                var measureExpr = projection.Field.Measure;
                if (measureExpr.Expression?.SourceRef?.Entity != null && measureExpr.Property != null)
                {
                    string fullName = String.Format("'{0}'[{1}]", measureExpr.Expression.SourceRef.Entity, measureExpr.Property);
                    if (!measureDisplayNamesDict.ContainsKey(fullName))
                    {
                        measureDisplayNamesDict[fullName] = new HashSet<string>();
                    }
                    // Only add if it's different from the default field name
                    if (displayName != measureExpr.Property)
                    {
                        measureDisplayNamesDict[fullName].Add(displayName);
                    }
                    else if (measureDisplayNamesDict[fullName].Count == 0)
                    {
                        // If no custom display name yet, use the property name
                        measureDisplayNamesDict[fullName].Add(measureExpr.Property);
                    }
                }
            }
            // Check if it's a column
            else if (projection.Field.Column != null)
            {
                var columnExpr = projection.Field.Column;
                if (columnExpr.Expression?.SourceRef?.Entity != null && columnExpr.Property != null)
                {
                    string fullName = String.Format("'{0}'[{1}]", columnExpr.Expression.SourceRef.Entity, columnExpr.Property);
                    if (!columnDisplayNamesDict.ContainsKey(fullName))
                    {
                        columnDisplayNamesDict[fullName] = new HashSet<string>();
                    }
                    // Only add if it's different from the default field name
                    if (displayName != columnExpr.Property && displayName != fullName)
                    {
                        columnDisplayNamesDict[fullName].Add(displayName);
                    }
                }
            }
        }
    }
}
// Step 4: Resolve conflicts (multiple display names for same field)
var measureDisplayNames = new Dictionary<string, string>();
var columnDisplayNames = new Dictionary<string, string>();
foreach (var kvp in measureDisplayNamesDict)
{
    string fieldKey = kvp.Key;
    var displayNames = kvp.Value.ToList();
    if (displayNames.Count == 0) continue;
    if (displayNames.Count == 1)
    {
        measureDisplayNames[fieldKey] = displayNames[0];
    }
    else
    {
        // Multiple display names found - ask user
        string chosen = Fx.ChooseString(
            OptionList: displayNames,
            label: String.Format("Multiple display names found for measure '{0}'. Choose one:", fieldKey)
        );
        if (string.IsNullOrEmpty(chosen))
        {
            Info("Operation cancelled.");
            return;
        }
        measureDisplayNames[fieldKey] = chosen;
    }
}
foreach (var kvp in columnDisplayNamesDict)
{
    string fieldKey = kvp.Key;
    var displayNames = kvp.Value.ToList();
    if (displayNames.Count == 0) continue;
    if (displayNames.Count == 1)
    {
        columnDisplayNames[fieldKey] = displayNames[0];
    }
    else
    {
        // Multiple display names found - ask user
        string chosen = Fx.ChooseString(
            OptionList: displayNames,
            label: String.Format("Multiple display names found for column '{0}'. Choose one:", fieldKey)
        );
        if (string.IsNullOrEmpty(chosen))
        {
            Info("Operation cancelled.");
            return;
        }
        columnDisplayNames[fieldKey] = chosen;
    }
}
// Step 5: Apply annotations to model
int measuresUpdated = 0;
int columnsUpdated = 0;
foreach (var kvp in measureDisplayNames)
{
    var measure = Model.AllMeasures.FirstOrDefault(
        m => m.Table.DaxObjectFullName + m.DaxObjectFullName == kvp.Key);
    if (measure != null)
    {
        if (measure.GetAnnotation("DisplayName") == kvp.Value) continue;
        measure.SetAnnotation("DisplayName", kvp.Value);
        measuresUpdated++;
    }
}
foreach (var kvp in columnDisplayNames)
{
    var column = Model.AllColumns.FirstOrDefault(c => c.DaxObjectFullName == kvp.Key);
    if (column != null)
    {
        if(column.GetAnnotation("DisplayName") == kvp.Value) continue;
        column.SetAnnotation("DisplayName", kvp.Value);
        columnsUpdated++;
    }
}
Output(String.Format("Updated {0} measures and {1} columns with DisplayName annotations.", measuresUpdated, columnsUpdated));

public static class Fx
{
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

public static class Rx

{





    

    

    public static VisualExtended DuplicateVisual(VisualExtended visualExtended)

    {

        // Generate a clean 16-character name from a GUID (no dashes or slashes)

        string newVisualName = Guid.NewGuid().ToString("N").Substring(0, 16);

        string sourceFolder = Path.GetDirectoryName(visualExtended.VisualFilePath);

        string targetFolder = Path.Combine(Path.GetDirectoryName(sourceFolder), newVisualName);

        if (Directory.Exists(targetFolder))

        {

            Error(string.Format("Folder already exists: {0}", targetFolder));

            return null;

        }

        Directory.CreateDirectory(targetFolder);



        // Deep clone the VisualDto.Root object

        string originalJson = JsonConvert.SerializeObject(visualExtended.Content, Newtonsoft.Json.Formatting.Indented);

        VisualDto.Root clonedContent = 

            JsonConvert.DeserializeObject<VisualDto.Root>(

                originalJson, 

                new JsonSerializerSettings {

                    DefaultValueHandling = DefaultValueHandling.Ignore,

                    NullValueHandling = NullValueHandling.Ignore



                });



        // Update the name property if it exists

        if (clonedContent != null && clonedContent.Name != null)

        {

            clonedContent.Name = newVisualName;

        }



        // Set the new file path

        string newVisualFilePath = Path.Combine(targetFolder, "visual.json");



        // Create the new VisualExtended object

        VisualExtended newVisual = new VisualExtended

        {

            Content = clonedContent,

            VisualFilePath = newVisualFilePath

        };



        return newVisual;

    }



    public static VisualExtended GroupVisuals(List<VisualExtended> visualsToGroup, string groupName = null, string groupDisplayName = null)

    {

        if (visualsToGroup == null || visualsToGroup.Count == 0)

        {

            Error("No visuals to group.");

            return null;

        }

        // Generate a clean 16-character name from a GUID (no dashes or slashes) if no group name is provided

        if (string.IsNullOrEmpty(groupName))

        {

            groupName = Guid.NewGuid().ToString("N").Substring(0, 16);

        }

        if (string.IsNullOrEmpty(groupDisplayName))

        {

            groupDisplayName = groupName;

        }



        // Find minimum X and Y

        double minX = visualsToGroup.Min(v => v.Content.Position != null ? (double)v.Content.Position.X : 0);

        double minY = visualsToGroup.Min(v => v.Content.Position != null ? (double)v.Content.Position.Y : 0);



       //Info("minX:" + minX.ToString() + ", minY: " + minY.ToString());



        // Calculate width and height

        double groupWidth = 0;

        double groupHeight = 0;

        foreach (var v in visualsToGroup)

        {

            if (v.Content != null && v.Content.Position != null)

            {

                double visualWidth = v.Content.Position != null ? (double)v.Content.Position.Width : 0;

                double visualHeight = v.Content.Position != null ? (double)v.Content.Position.Height : 0;

                double xOffset = (double)v.Content.Position.X - (double)minX;

                double yOffset = (double)v.Content.Position.Y - (double)minY;

                double totalWidth = xOffset + visualWidth;

                double totalHeight = yOffset + visualHeight;

                if (totalWidth > groupWidth) groupWidth = totalWidth;

                if (totalHeight > groupHeight) groupHeight = totalHeight;

            }

        }



        // Create the group visual content

        var groupContent = new VisualDto.Root

        {

            Schema = visualsToGroup.FirstOrDefault().Content.Schema,

            Name = groupName,

            Position = new VisualDto.Position

            {

                X = minX,

                Y = minY,

                Width = groupWidth,

                Height = groupHeight

            },

            VisualGroup = new VisualDto.VisualGroup

            {

                DisplayName = groupDisplayName,

                GroupMode = "ScaleMode"

            }

        };



        // Set VisualFilePath for the group visual

        // Use the VisualFilePath of the first visual as a template

        string groupVisualFilePath = null;

        var firstVisual = visualsToGroup.FirstOrDefault(v => !string.IsNullOrEmpty(v.VisualFilePath));

        if (firstVisual != null && !string.IsNullOrEmpty(firstVisual.VisualFilePath))

        {

            string originalPath = firstVisual.VisualFilePath;

            string parentDir = Path.GetDirectoryName(Path.GetDirectoryName(originalPath)); // up to 'visuals'

            if (!string.IsNullOrEmpty(parentDir))

            {

                string groupFolder = Path.Combine(parentDir, groupName);

                groupVisualFilePath = Path.Combine(groupFolder, "visual.json");

            }

        }



        // Create the new VisualExtended for the group

        var groupVisual = new VisualExtended

        {

            Content = groupContent,

            VisualFilePath = groupVisualFilePath // Set as described

        };



        // Update grouped visuals: set parentGroupName and adjust X/Y

        foreach (var v in visualsToGroup)

        {

            

            if (v.Content == null) continue;

            v.Content.ParentGroupName = groupName;



            if (v.Content.Position != null)

            {

                v.Content.Position.X = v.Content.Position.X - minX + 0;

                v.Content.Position.Y = v.Content.Position.Y - minY + 0;

            }

        }



        return groupVisual;

    }



    



    private static readonly string RecentPathsFile = Path.Combine(

    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),

    "Tabular Editor Macro Settings", "recentPbirPaths.json");



    public static string GetPbirFilePathWithHistory(string label = "Select definition.pbir file")

    {

        // Load recent paths

        List<string> recentPaths = LoadRecentPbirPaths();



        // Filter out non-existing files

        recentPaths = recentPaths.Where(File.Exists).ToList();



        // Present options to the user

        var options = new List<string>(recentPaths);

        options.Add("Browse for new file...");



        string selected = Fx.ChooseString(options,label:label, customWidth:600, customHeight:300);



        if (selected == null) return null;



        string chosenPath = null;

        if (selected == "Browse for new file..." )

        {

            chosenPath = GetPbirFilePath(label);

        }

        else

        {

            chosenPath = selected;

        }



        if (!string.IsNullOrEmpty(chosenPath))

        {

            // Update recent paths

            UpdateRecentPbirPaths(chosenPath, recentPaths);

        }



        return chosenPath;

    }



    private static List<string> LoadRecentPbirPaths()

    {

        try

        {

            if (File.Exists(RecentPathsFile))

            {

                string json = File.ReadAllText(RecentPathsFile);

                return JsonConvert.DeserializeObject<List<string>>(json) ?? new List<string>();

            }

        }

        catch { }

        return new List<string>();

    }



    private static void UpdateRecentPbirPaths(string newPath, List<string> recentPaths)

    {

        // Remove if already exists, insert at top

        recentPaths.RemoveAll(p => string.Equals(p, newPath, StringComparison.OrdinalIgnoreCase));

        recentPaths.Insert(0, newPath);



        // Keep only the latest 10

        while (recentPaths.Count > 10)

            recentPaths.RemoveAt(recentPaths.Count - 1);



        // Ensure directory exists

        Directory.CreateDirectory(Path.GetDirectoryName(RecentPathsFile));

        File.WriteAllText(RecentPathsFile, JsonConvert.SerializeObject(recentPaths, Newtonsoft.Json.Formatting.Indented));

    }





    public static ReportExtended InitReport(string label = "Please select definition.pbir file of the target report")

    {

        // Get the base path from the user  

        string basePath = Rx.GetPbirFilePathWithHistory(label:label);

        if (basePath == null) return null; 

        

        // Define the target path  

        string baseDirectory = Path.GetDirectoryName(basePath);

        string targetPath = Path.Combine(baseDirectory, "definition", "pages");



        // Check if the target path exists  

        if (!Directory.Exists(targetPath))

        {

            Error(String.Format("The path '{0}' does not exist.", targetPath));

            return null;

        }



        // Get all subfolders in the target path  

        List<string> subfolders = Directory.GetDirectories(targetPath).ToList();



        string pagesFilePath = Path.Combine(targetPath, "pages.json");

        string pagesJsonContent = File.ReadAllText(pagesFilePath);

        

        if (string.IsNullOrEmpty(pagesJsonContent))

        {

            Error(String.Format("The file '{0}' is empty or does not exist.", pagesFilePath));

            return null;

        }



        PagesDto pagesDto = JsonConvert.DeserializeObject<PagesDto>(pagesJsonContent);



        ReportExtended report = new ReportExtended();

        report.PagesFilePath = pagesFilePath;

        report.PagesConfig = pagesDto;



        // Process each folder  

        foreach (string folder in subfolders)

        {

            string pageJsonPath = Path.Combine(folder, "page.json");

            if (File.Exists(pageJsonPath))

            {

                try

                {

                    string jsonContent = File.ReadAllText(pageJsonPath);

                    PageDto page = JsonConvert.DeserializeObject<PageDto>(jsonContent);



                    PageExtended pageExtended = new PageExtended();

                    pageExtended.Page = page;

                    pageExtended.PageFilePath = pageJsonPath;



                    pageExtended.ParentReport = report;



                    string visualsPath = Path.Combine(folder, "visuals");



                    if (!Directory.Exists(visualsPath))

                    {

                        report.Pages.Add(pageExtended); // still add the page

                        continue; // skip visual loading

                    }



                    List<string> visualSubfolders = Directory.GetDirectories(visualsPath).ToList();



                    foreach (string visualFolder in visualSubfolders)

                    {

                        string visualJsonPath = Path.Combine(visualFolder, "visual.json");

                        if (File.Exists(visualJsonPath))

                        {

                            try

                            {

                                string visualJsonContent = File.ReadAllText(visualJsonPath);

                                VisualDto.Root visual = JsonConvert.DeserializeObject<VisualDto.Root>(visualJsonContent);



                                VisualExtended visualExtended = new VisualExtended();

                                visualExtended.Content = visual;

                                visualExtended.VisualFilePath = visualJsonPath;

                                visualExtended.ParentPage = pageExtended; // Set parent page reference

                                pageExtended.Visuals.Add(visualExtended);

                            }

                            catch (Exception ex2)

                            {

                                Output(String.Format("Error reading or deserializing '{0}': {1}", visualJsonPath, ex2.Message));

                                return null;

                            }



                        }

                    }



                    report.Pages.Add(pageExtended);



                }

                catch (Exception ex)

                {

                    Output(String.Format("Error reading or deserializing '{0}': {1}", pageJsonPath, ex.Message));

                }

            }



        }

        return report;

    }





    public static VisualExtended SelectTableVisual(ReportExtended report)

    {

        List<string> visualTypes = new List<string>

        {

            "tableEx","pivotTable"

        };

        return SelectVisual(report: report, visualTypes);

    }







    public static VisualExtended SelectVisual(ReportExtended report, List<string> visualTypeList = null)

    {

        return SelectVisualInternal(report, Multiselect: false, visualTypeList:visualTypeList) as VisualExtended;

    }



    public static List<VisualExtended> SelectVisuals(ReportExtended report, List<string> visualTypeList = null)

    {

        return SelectVisualInternal(report, Multiselect: true, visualTypeList:visualTypeList) as List<VisualExtended>;

    }



    private static object SelectVisualInternal(ReportExtended report, bool Multiselect, List<string> visualTypeList = null)

    {

        // Step 1: Build selection list

        var visualSelectionList = 

            report.Pages

            .SelectMany(p => p.Visuals

                .Where(v =>

                    v?.Content != null &&

                    (

                        // If visualTypeList is null, do not filter at all

                        (visualTypeList == null) ||

                        // If visualTypeList is provided and not empty, filter by it

                        (visualTypeList.Count > 0 && v.Content.Visual != null && visualTypeList.Contains(v.Content?.Visual?.VisualType))

                        // Otherwise, include all visuals and visual groups

                        || (visualTypeList.Count == 0)

                    )

                )

                .Select(v => new

                    {

                        // Use visual type for regular visuals, displayname for groups

                        Display = string.Format(

                            "{0} - {1} ({2}, {3})",

                            p.Page.DisplayName,

                            v?.Content?.Visual?.VisualType

                                ?? v?.Content?.VisualGroup?.DisplayName,

                            (int)(v.Content.Position?.X ?? 0),

                            (int)(v.Content.Position?.Y ?? 0)

                        ),

                        Page = p,

                        Visual = v

                    }

                )

            )

            .ToList();



        if (visualSelectionList.Count == 0)

        {

            if (visualTypeList != null)

            {

                string types = string.Join(", ", visualTypeList);

                Error(string.Format("No visual of type {0} were found", types));



            }else

            {

                Error("No visuals found in the report.");

            }





            return null;

        }



        // Step 2: Let user choose a visual

        var options = visualSelectionList.Select(v => v.Display).ToList();



        if (Multiselect)

        {

            // For multiselect, use ChooseStringMultiple

            var multiSelelected = Fx.ChooseStringMultiple(options);

            if (multiSelelected == null || multiSelelected.Count == 0)

            {

                Info("You cancelled.");

                return null;

            }

            // Find all selected visuals

            var selectedVisuals = visualSelectionList.Where(v => multiSelelected.Contains(v.Display)).Select(v => v.Visual).ToList();



            return selectedVisuals;

        }

        else

        {

            string selected = Fx.ChooseString(options);



            if (string.IsNullOrEmpty(selected))

            {

                Info("You cancelled.");

                return null;

            }



            // Step 3: Find the selected visual

            var selectedVisual = visualSelectionList.FirstOrDefault(v => v.Display == selected);



            if (selectedVisual == null)

            {

                Error("Selected visual not found.");

                return null;

            }



            return selectedVisual.Visual;

        }

    }



    public static PageExtended ReplicateFirstPageAsBlank(ReportExtended report, bool showMessages = false)

    {

        if (report.Pages == null || !report.Pages.Any())

        {

            Error("No pages found in the report.");

            return null;

        }



        PageExtended firstPage = report.Pages[0];



        // Generate a clean 16-character name from a GUID (no dashes or slashes)

        string newPageName = Guid.NewGuid().ToString("N").Substring(0, 16);

        string newPageDisplayName = firstPage.Page.DisplayName + " - Copy";



        string sourceFolder = Path.GetDirectoryName(firstPage.PageFilePath);

        string targetFolder = Path.Combine(Path.GetDirectoryName(sourceFolder), newPageName);

        string visualsFolder = Path.Combine(targetFolder, "visuals");



        if (Directory.Exists(targetFolder))

        {

            Error($"Folder already exists: {targetFolder}");

            return null;

        }



        Directory.CreateDirectory(targetFolder);

        Directory.CreateDirectory(visualsFolder);



        var newPageDto = new PageDto

        {

            Name = newPageName,

            DisplayName = newPageDisplayName,

            DisplayOption = firstPage.Page.DisplayOption,

            Height = firstPage.Page.Height,

            Width = firstPage.Page.Width,

            Schema = firstPage.Page.Schema

        };



        var newPage = new PageExtended

        {

            Page = newPageDto,

            PageFilePath = Path.Combine(targetFolder, "page.json"),

            Visuals = new List<VisualExtended>() // empty visuals

        };



        File.WriteAllText(newPage.PageFilePath, JsonConvert.SerializeObject(newPageDto, Newtonsoft.Json.Formatting.Indented));



        report.Pages.Add(newPage);



        if(showMessages) Info($"Created new blank page: {newPageName}");



        return newPage; 

    }





    public static void SaveVisual(VisualExtended visual)

    {



        // Save new JSON, ignoring nulls

        string newJson = JsonConvert.SerializeObject(

            visual.Content,

            Newtonsoft.Json.Formatting.Indented,

            new JsonSerializerSettings

            {

                //DefaultValueHandling = DefaultValueHandling.Ignore,

                NullValueHandling = NullValueHandling.Ignore



            }

        );

        // Ensure the directory exists before saving

        string visualFolder = Path.GetDirectoryName(visual.VisualFilePath);

        if (!Directory.Exists(visualFolder))

        {

            Directory.CreateDirectory(visualFolder);

        }

        File.WriteAllText(visual.VisualFilePath, newJson);

    }





    public static string ReplacePlaceholders(string pageContents, Dictionary<string, string> placeholders)

    {

        if (placeholders != null)

        {

            foreach (string placeholder in placeholders.Keys)

            {

                string valueToReplace = placeholders[placeholder];



                pageContents = pageContents.Replace(placeholder, valueToReplace);



            }

        }





        return pageContents;

    }





    public static String GetPbirFilePath(string label = "Please select definition.pbir file of the target report")

    {



        // Create an instance of the OpenFileDialog

        OpenFileDialog openFileDialog = new OpenFileDialog

        {

            Title = label,

            // Set filter options and filter index.

            Filter = "PBIR Files (*.pbir)|*.pbir",

            FilterIndex = 1

        };

        // Call the ShowDialog method to show the dialog box.

        DialogResult result = openFileDialog.ShowDialog();

        // Process input if the user clicked OK.

        if (result != DialogResult.OK)

        {

            Error("You cancelled");

            return null;

        }

        return openFileDialog.FileName;



    }





}



   

    

    public class PagesDto
    {
        [Newtonsoft.Json.JsonProperty("$schema")]
        public string Schema { get; set; }

        [Newtonsoft.Json.JsonProperty("pageOrder")]
        public List<string> PageOrder { get; set; }

        [Newtonsoft.Json.JsonProperty("activePageName")]
        public string ActivePageName { get; set; }
        
    }


    public class PageDto
    {
        [Newtonsoft.Json.JsonProperty("$schema")]
        public string Schema { get; set; }

        [Newtonsoft.Json.JsonProperty("name")]
        public string Name { get; set; }

        [Newtonsoft.Json.JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [Newtonsoft.Json.JsonProperty("displayOption")]
        public string DisplayOption { get; set; } // Could create enum if you want stricter typing

        [Newtonsoft.Json.JsonProperty("height")]
        public double? Height { get; set; }

        [Newtonsoft.Json.JsonProperty("width")]
        public double? Width { get; set; }
    }



    public partial class VisualDto
    {
        public class Root
        {
            [JsonProperty("$schema")] public string Schema { get; set; }
            [JsonProperty("name")] public string Name { get; set; }
            [JsonProperty("position")] public Position Position { get; set; }
            [JsonProperty("visual")] public Visual Visual { get; set; }
            

            [JsonProperty("visualGroup")] public VisualGroup VisualGroup { get; set; }
            [JsonProperty("parentGroupName")] public string ParentGroupName { get; set; }
            [JsonProperty("filterConfig")] public FilterConfig FilterConfig { get; set; }
            [JsonProperty("isHidden")] public bool IsHidden { get; set; }

            [JsonExtensionData]
            
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }


        public class VisualContainerObjects
        {
            [JsonProperty("general")]
            public List<VisualContainerObject> General { get; set; }

            // Add other known properties as needed, e.g.:
            [JsonProperty("title")]
            public List<VisualContainerObject> Title { get; set; }

            [JsonProperty("subTitle")]
            public List<VisualContainerObject> SubTitle { get; set; }

            // This will capture any additional properties not explicitly defined above
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualContainerObject
        {
            [JsonProperty("properties")]
            public Dictionary<string, VisualContainerProperty> Properties { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualContainerProperty
        {
            [JsonProperty("expr")]
            public VisualExpr Expr { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualExpr
        {
            [JsonProperty("Literal")]
            public VisualLiteral Literal { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualLiteral
        {
            [JsonProperty("Value")]
            public string Value { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualGroup
        {
            [JsonProperty("displayName")] public string DisplayName { get; set; }
            [JsonProperty("groupMode")] public string GroupMode { get; set; }
        }

        public class Position
        {
            [JsonProperty("x")] public double X { get; set; }
            [JsonProperty("y")] public double Y { get; set; }
            [JsonProperty("z")] public int Z { get; set; }
            [JsonProperty("height")] public double Height { get; set; }
            [JsonProperty("width")] public double Width { get; set; }

            [JsonProperty("tabOrder", NullValueHandling = NullValueHandling.Ignore)]
            public int? TabOrder { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Visual
        {
            [JsonProperty("visualType")] public string VisualType { get; set; }
            [JsonProperty("query")] public Query Query { get; set; }
            [JsonProperty("objects")] public Objects Objects { get; set; }
            [JsonProperty("visualContainerObjects")]
            public VisualContainerObjects VisualContainerObjects { get; set; }
            [JsonProperty("drillFilterOtherVisuals")] public bool DrillFilterOtherVisuals { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Query
        {
            [JsonProperty("queryState")] public QueryState QueryState { get; set; }
            [JsonProperty("sortDefinition")] public SortDefinition SortDefinition { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class  QueryState
        {
            [JsonProperty("Rows", Order = 1)] public VisualDto.ProjectionsSet Rows { get; set; }
            [JsonProperty("Category", Order = 2)] public VisualDto.ProjectionsSet Category { get; set; }
            [JsonProperty("Y", Order = 3)] public VisualDto.ProjectionsSet Y { get; set; }
            [JsonProperty("Y2", Order = 4)] public VisualDto.ProjectionsSet Y2 { get; set; }
            [JsonProperty("Values", Order = 5)] public VisualDto.ProjectionsSet Values { get; set; }
            
            [JsonProperty("Series", Order = 6)] public VisualDto.ProjectionsSet Series { get; set; }
            [JsonProperty("Data", Order = 7)] public VisualDto.ProjectionsSet Data { get; set; }

            
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ProjectionsSet
        {
            [JsonProperty("projections")] public List<VisualDto.Projection> Projections { get; set; }
            [JsonProperty("fieldParameters")] public List<VisualDto.FieldParameter> FieldParameters { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FieldParameter
        {
            [JsonProperty("parameterExpr")]
            public Field ParameterExpr { get; set; }

            [JsonProperty("index")]
            public int Index { get; set; }

            [JsonProperty("length")]
            public int Length { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Projection
        {
            [JsonProperty("field")] public VisualDto.Field Field { get; set; }
            [JsonProperty("queryRef")] public string QueryRef { get; set; }
            [JsonProperty("nativeQueryRef")] public string NativeQueryRef { get; set; }

            [JsonProperty("displayName")] public string DisplayName { get; set; }
            [JsonProperty("active")] public bool? Active { get; set; }
            [JsonProperty("hidden")] public bool? Hidden { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Field
        {
            [JsonProperty("Aggregation")] public VisualDto.Aggregation Aggregation { get; set; }
            [JsonProperty("NativeVisualCalculation")] public NativeVisualCalculation NativeVisualCalculation { get; set; }
            [JsonProperty("Measure")] public VisualDto.MeasureObject Measure { get; set; }
            [JsonProperty("Column")] public VisualDto.ColumnField Column { get; set; }

            [JsonExtensionData] public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Aggregation
        {
            [JsonProperty("Expression")] public VisualDto.Expression Expression { get; set; }
            [JsonProperty("Function")] public int Function { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class NativeVisualCalculation
        {
            [JsonProperty("Language")] public string Language { get; set; }
            [JsonProperty("Expression")] public string Expression { get; set; }
            [JsonProperty("Name")] public string Name { get; set; }

            [JsonProperty("DataType")] public string DataType { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class MeasureObject
        {
            [JsonProperty("Expression")] public VisualDto.Expression Expression { get; set; }
            [JsonProperty("Property")] public string Property { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ColumnField
        {
            [JsonProperty("Expression")] public VisualDto.Expression Expression { get; set; }
            [JsonProperty("Property")] public string Property { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Expression
        {
            [JsonProperty("Column")] public ColumnExpression Column { get; set; }
            [JsonProperty("SourceRef")] public VisualDto.SourceRef SourceRef { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ColumnExpression
        {
            [JsonProperty("Expression")] public VisualDto.SourceRef Expression { get; set; }
            [JsonProperty("Property")] public string Property { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class SourceRef
        {
            [JsonProperty("Schema")] public string Schema { get; set; }
            [JsonProperty("Entity")] public string Entity { get; set; }
            [JsonProperty("Source")] public string Source { get; set; }

            
        }

        public class SortDefinition
        {
            [JsonProperty("sort")] public List<VisualDto.Sort> Sort { get; set; }
            [JsonProperty("isDefaultSort")] public bool IsDefaultSort { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Sort
        {
            [JsonProperty("field")] public VisualDto.Field Field { get; set; }
            [JsonProperty("direction")] public string Direction { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Objects
        {
            [JsonProperty("valueAxis")] public List<VisualDto.ObjectProperties> ValueAxis { get; set; }
            [JsonProperty("general")] public List<VisualDto.ObjectProperties> General { get; set; }
            [JsonProperty("data")] public List<VisualDto.ObjectProperties> Data { get; set; }
            [JsonProperty("title")] public List<VisualDto.ObjectProperties> Title { get; set; }
            [JsonProperty("legend")] public List<VisualDto.ObjectProperties> Legend { get; set; }
            [JsonProperty("labels")] public List<VisualDto.ObjectProperties> Labels { get; set; }
            [JsonProperty("dataPoint")] public List<VisualDto.ObjectProperties> DataPoint { get; set; }
            [JsonProperty("columnFormatting")] public List<VisualDto.ObjectProperties> ColumnFormatting { get; set; }
            [JsonProperty("referenceLabel")] public List<VisualDto.ObjectProperties> ReferenceLabel { get; set; }
            [JsonProperty("referenceLabelDetail")] public List<VisualDto.ObjectProperties> ReferenceLabelDetail { get; set; }
            [JsonProperty("referenceLabelValue")] public List<VisualDto.ObjectProperties> ReferenceLabelValue { get; set; }

            [JsonProperty("values")] public List<VisualDto.ObjectProperties> Values { get; set; }

            [JsonProperty("y1AxisReferenceLine")] public List<VisualDto.ObjectProperties> Y1AxisReferenceLine { get; set; }

            [JsonExtensionData] public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ObjectProperties
        {
            [JsonProperty("properties")]
            [JsonConverter(typeof(PropertiesConverter))]
            public Dictionary<string, object> Properties { get; set; }

            [JsonProperty("selector")]
            public Selector Selector { get; set; }


            [JsonExtensionData] public IDictionary<string, JToken> ExtensionData { get; set; }
        }




        public class VisualObjectProperty
        {
            [JsonProperty("expr")] public VisualPropertyExpr Expr { get; set; }
            [JsonProperty("solid")] public SolidColor Solid { get; set; }
            [JsonProperty("color")] public ColorExpression Color { get; set; }

            [JsonProperty("paragraphs")]
            public List<Paragraph> Paragraphs { get; set; }

            [JsonExtensionData] public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualPropertyExpr
        {
            // Existing Field properties
            [JsonProperty("Measure")] public MeasureObject Measure { get; set; }
            [JsonProperty("Column")] public ColumnField Column { get; set; }
            [JsonProperty("Aggregation")] public Aggregation Aggregation { get; set; }
            [JsonProperty("NativeVisualCalculation")] public NativeVisualCalculation NativeVisualCalculation { get; set; }

            // New properties from JSON
            [JsonProperty("SelectRef")] public SelectRefExpression SelectRef { get; set; }
            [JsonProperty("Literal")] public VisualLiteral Literal { get; set; }

            [JsonExtensionData] public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class SelectRefExpression
        {
            [JsonProperty("ExpressionName")]
            public string ExpressionName { get; set; }
        }

        public class Paragraph
        {
            [JsonProperty("textRuns")]
            public List<TextRun> TextRuns { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class TextRun
        {
            [JsonProperty("value")]
            public string Value { get; set; }

            [JsonProperty("url")]
            public string Url { get; set; }

            [JsonProperty("textStyle")]
            public Dictionary<string, object> TextStyle { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class SolidColor
        {
            [JsonProperty("color")] public ColorExpression Color { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ColorExpression
        {
            [JsonProperty("expr")]
            public VisualColorExprWrapper Expr { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FillRuleExprWrapper
        {
            [JsonProperty("FillRule")] public FillRuleExpression FillRule { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FillRuleExpression
        {
            [JsonProperty("Input")] public VisualDto.Field Input { get; set; }
            [JsonProperty("FillRule")] public Dictionary<string, object> FillRule { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ThemeDataColor
        {
            [JsonProperty("ColorId")] public int ColorId { get; set; }
            [JsonProperty("Percent")] public double Percent { get; set; }
            [JsonExtensionData] public Dictionary<string, JToken> ExtensionData { get; set; }
        }
        public class VisualColorExprWrapper
        {
            [JsonProperty("Measure")]
            public VisualDto.MeasureObject Measure { get; set; }

            [JsonProperty("Column")]
            public VisualDto.ColumnField Column { get; set; }

            [JsonProperty("Aggregation")]
            public VisualDto.Aggregation Aggregation { get; set; }

            [JsonProperty("NativeVisualCalculation")]
            public NativeVisualCalculation NativeVisualCalculation { get; set; }

            [JsonProperty("FillRule")]
            public FillRuleExpression FillRule { get; set; }

            public VisualLiteral Literal { get; set; }

            [JsonProperty("ThemeDataColor")] 
            public ThemeDataColor ThemeDataColor { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }


        

        public class Selector
        {
            

            [JsonProperty("id")]
            public string Id { get; set; }

            [JsonProperty("order")]
            public int? Order { get; set; }

            [JsonProperty("data")]
            public List<DataObject> Data { get; set; }

            [JsonProperty("metadata")]
            public string Metadata { get; set; }

            [JsonProperty("scopeId")]
            public string ScopeId { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class DataObject
        {
            [JsonProperty("dataViewWildcard")]
            public DataViewWildcard DataViewWildcard { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class DataViewWildcard
        {
            [JsonProperty("matchingOption")]
            public int MatchingOption { get; set; }
            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FilterConfig
        {
            [JsonProperty("filters")]
            public List<VisualFilter> Filters { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class VisualFilter
        {
            [JsonProperty("name")] public string Name { get; set; }
            [JsonProperty("field")] public VisualDto.Field Field { get; set; }
            [JsonProperty("type")] public string Type { get; set; }
            [JsonProperty("filter")] public FilterDefinition Filter { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FilterDefinition
        {
            [JsonProperty("Version")] public int Version { get; set; }
            [JsonProperty("From")] public List<FilterFrom> From { get; set; }
            [JsonProperty("Where")] public List<FilterWhere> Where { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FilterFrom
        {
            [JsonProperty("Name")] public string Name { get; set; }
            [JsonProperty("Entity")] public string Entity { get; set; }
            [JsonProperty("Type")] public int Type { get; set; }
            [JsonProperty("Expression")] public FilterExpression Expression { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FilterExpression
        {
            [JsonProperty("Subquery")] public SubqueryExpression Subquery { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class SubqueryExpression
        {
            [JsonProperty("Query")] public SubqueryQuery Query { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class SubqueryQuery
        {
            [JsonProperty("Version")] public int Version { get; set; }
            [JsonProperty("From")] public List<FilterFrom> From { get; set; }
            [JsonProperty("Select")] public List<SelectExpression> Select { get; set; }
            [JsonProperty("OrderBy")] public List<OrderByExpression> OrderBy { get; set; }
            [JsonProperty("Top")] public int? Top { get; set; }

            [JsonProperty("Where")] public List<FilterWhere> Where { get; set; } // 🔹 Added

            [JsonExtensionData] public Dictionary<string, JToken> ExtensionData { get; set; }
        }


        public class SelectExpression
        {
            [JsonProperty("Column")] public ColumnSelect Column { get; set; }
            [JsonProperty("Name")] public string Name { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ColumnSelect
        {
            [JsonProperty("Expression")]
            public VisualDto.Expression Expression { get; set; }  // NOTE: wrapper that contains "SourceRef"

            [JsonProperty("Property")]
            public string Property { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class OrderByExpression
        {
            [JsonProperty("Direction")] public int Direction { get; set; }
            [JsonProperty("Expression")] public OrderByInnerExpression Expression { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class OrderByInnerExpression
        {
            [JsonProperty("Measure")] public VisualDto.MeasureObject Measure { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FilterWhere
        {
            [JsonProperty("Condition")] public Condition Condition { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class Condition
        {
            [JsonProperty("In")] public InExpression In { get; set; }
            [JsonProperty("Not")] public NotExpression Not { get; set; }
            [JsonProperty("Comparison")] public ComparisonExpression Comparison { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class InExpression
        {
            [JsonProperty("Expressions")] public List<ColumnSelect> Expressions { get; set; }
            [JsonProperty("Table")] public InTable Table { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class InTable
        {
            [JsonProperty("SourceRef")] public VisualDto.SourceRef SourceRef { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class NotExpression
        {
            [JsonProperty("Expression")] public Condition Expression { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class ComparisonExpression
        {
            [JsonProperty("ComparisonKind")] public int ComparisonKind { get; set; }
            [JsonProperty("Left")] public FilterOperand Left { get; set; }
            [JsonProperty("Right")] public FilterOperand Right { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class FilterOperand
        {
            [JsonProperty("Measure")] public VisualDto.MeasureObject Measure { get; set; }
            [JsonProperty("Column")] public VisualDto.ColumnField Column { get; set; }
            [JsonProperty("Literal")] public LiteralOperand Literal { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }

        public class LiteralOperand
        {
            [JsonProperty("Value")] public string Value { get; set; }

            [JsonExtensionData]
            public Dictionary<string, JToken> ExtensionData { get; set; }
        }


        public class PropertiesConverter : JsonConverter
        {
            public override bool CanConvert(Type objectType) => objectType == typeof(Dictionary<string, object>);

            public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
            {
                var result = new Dictionary<string, object>();
                var jObj = JObject.Load(reader);

                foreach (var prop in jObj.Properties())
                {
                    if (prop.Name == "paragraphs")
                    {
                        var paragraphs = prop.Value.ToObject<List<Paragraph>>(serializer);
                        result[prop.Name] = paragraphs;
                    }
                    else
                    {
                        var visualProp = prop.Value.ToObject<VisualObjectProperty>(serializer);
                        result[prop.Name] = visualProp;
                    }
                }

                return result;
            }

            public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
            {
                var dict = (Dictionary<string, object>)value;
                writer.WriteStartObject();

                foreach (var kvp in dict)
                {
                    writer.WritePropertyName(kvp.Key);

                    if (kvp.Value is VisualObjectProperty vo)
                        serializer.Serialize(writer, vo);
                    else if (kvp.Value is List<Paragraph> ps)
                        serializer.Serialize(writer, ps);
                    else
                        serializer.Serialize(writer, kvp.Value);
                }

                writer.WriteEndObject();
            }
        }
    }


    public class VisualExtended
    {
        public VisualDto.Root Content { get; set; }

        public string VisualFilePath { get; set; }

        public bool isVisualGroup => Content?.VisualGroup != null;
        public bool isGroupedVisual => Content?.ParentGroupName != null;

        public bool IsBilingualVisualGroup()
        {
            if (!isVisualGroup || string.IsNullOrEmpty(Content.VisualGroup.DisplayName))
                return false;
            return System.Text.RegularExpressions.Regex.IsMatch(Content.VisualGroup.DisplayName, @"^P\d{2}-\d{3}$");
        }

        public PageExtended ParentPage { get; set; }

        public bool IsInBilingualVisualGroup()
        {
            if (ParentPage == null || ParentPage.Visuals == null || Content.ParentGroupName == null)
                return false;
            return ParentPage.Visuals.Any(v => v.IsBilingualVisualGroup() && v.Content.Name == Content.ParentGroupName);
        }

        [JsonIgnore]
        public string AltText
        {
            get
            {
                var general = Content?.Visual?.VisualContainerObjects?.General;
                if (general == null || general.Count == 0)
                    return null;
                if (!general[0].Properties.ContainsKey("altText"))
                    return null;
                return general[0].Properties["altText"]?.Expr?.Literal?.Value?.Trim('\'');
            }
            set
            {
                if (Content?.Visual == null)
                    Content.Visual = new VisualDto.Visual();

                if (Content?.Visual?.VisualContainerObjects == null)
                    Content.Visual.VisualContainerObjects = new VisualDto.VisualContainerObjects();

                if (Content.Visual?.VisualContainerObjects.General == null || Content.Visual?.VisualContainerObjects.General.Count == 0)
                    Content.Visual.VisualContainerObjects.General =
                        new List<VisualDto.VisualContainerObject> {
                        new VisualDto.VisualContainerObject {
                            Properties = new Dictionary<string, VisualDto.VisualContainerProperty>()
                        }
                        };

                var general = Content.Visual.VisualContainerObjects.General[0];

                if (general.Properties == null)
                    general.Properties = new Dictionary<string, VisualDto.VisualContainerProperty>();

                general.Properties["altText"] = new VisualDto.VisualContainerProperty
                {
                    Expr = new VisualDto.VisualExpr
                    {
                        Literal = new VisualDto.VisualLiteral
                        {
                            Value = value == null ? null : "'" + value.Replace("'", "\\'") + "'"
                        }
                    }
                };
            }
        }

        private IEnumerable<VisualDto.Field> GetAllFields()
        {
            var fields = new List<VisualDto.Field>();
            var queryState = Content?.Visual?.Query?.QueryState;

            if (queryState != null)
            {
                fields.AddRange(GetFieldsFromProjections(queryState.Values));
                fields.AddRange(GetFieldsFromProjections(queryState.Y));
                fields.AddRange(GetFieldsFromProjections(queryState.Y2));
                fields.AddRange(GetFieldsFromProjections(queryState.Category));
                fields.AddRange(GetFieldsFromProjections(queryState.Series));
                fields.AddRange(GetFieldsFromProjections(queryState.Data));
                fields.AddRange(GetFieldsFromProjections(queryState.Rows));
            }

            var sortList = Content?.Visual?.Query?.SortDefinition?.Sort;
            if (sortList != null)
                fields.AddRange(sortList.Select(s => s.Field));

            var objects = Content?.Visual?.Objects;
            if (objects != null)
            {
                fields.AddRange(GetFieldsFromObjectList(objects.DataPoint));
                fields.AddRange(GetFieldsFromObjectList(objects.Data));
                fields.AddRange(GetFieldsFromObjectList(objects.Labels));
                fields.AddRange(GetFieldsFromObjectList(objects.Title));
                fields.AddRange(GetFieldsFromObjectList(objects.Legend));
                fields.AddRange(GetFieldsFromObjectList(objects.General));
                fields.AddRange(GetFieldsFromObjectList(objects.ValueAxis));
                fields.AddRange(GetFieldsFromObjectList(objects.Y1AxisReferenceLine));
                fields.AddRange(GetFieldsFromObjectList(objects.ReferenceLabel));
                fields.AddRange(GetFieldsFromObjectList(objects.ReferenceLabelDetail));
                fields.AddRange(GetFieldsFromObjectList(objects.ReferenceLabelValue));
            }

            fields.AddRange(GetFieldsFromFilterConfig(Content?.FilterConfig as VisualDto.FilterConfig));

            return fields.Where(f => f != null);
        }

        public IEnumerable<VisualDto.Field> GetFieldsFromProjections(VisualDto.ProjectionsSet set)
        {
            return set?.Projections?.Select(p => p.Field) ?? Enumerable.Empty<VisualDto.Field>();
        }

        

        private IEnumerable<VisualDto.Field> GetFieldsFromObjectList(List<VisualDto.ObjectProperties> objectList)
        {
            if (objectList == null) yield break;

            foreach (var obj in objectList)
            {
                if (obj.Properties == null) continue;

                foreach (var val in obj.Properties.Values)
                {
                    var prop = val as VisualDto.VisualObjectProperty;
                    if (prop == null) continue;

                    if (prop.Expr != null)
                    {
                        if (prop.Expr.Measure != null)
                            yield return new VisualDto.Field { Measure = prop.Expr.Measure };

                        if (prop.Expr.Column != null)
                            yield return new VisualDto.Field { Column = prop.Expr.Column };
                    }

                    if (prop.Color?.Expr?.FillRule?.Input != null)
                        yield return prop.Color.Expr.FillRule.Input;

                    if (prop.Solid?.Color?.Expr?.FillRule?.Input != null)
                        yield return prop.Solid.Color.Expr.FillRule.Input;

                    var solidExpr = prop.Solid?.Color?.Expr;
                    if (solidExpr?.Measure != null)
                        yield return new VisualDto.Field { Measure = solidExpr.Measure };
                    if (solidExpr?.Column != null)
                        yield return new VisualDto.Field { Column = solidExpr.Column };
                }
            }
        }

        private IEnumerable<VisualDto.Field> GetFieldsFromFilterConfig(VisualDto.FilterConfig filterConfig)
        {
            var fields = new List<VisualDto.Field>();

            if (filterConfig?.Filters == null)
                return fields;

            foreach (var filter in filterConfig.Filters ?? Enumerable.Empty<VisualDto.VisualFilter>())
            {
                if (filter.Field != null)
                    fields.Add(filter.Field);

                if (filter.Filter != null)
                {
                    var aliasMap = BuildAliasMap(filter.Filter.From);

                    foreach (var from in filter.Filter.From ?? Enumerable.Empty<VisualDto.FilterFrom>())
                    {
                        if (from.Expression?.Subquery?.Query != null)
                            ExtractFieldsFromSubquery(from.Expression.Subquery.Query, fields);
                    }

                    foreach (var where in filter.Filter.Where ?? Enumerable.Empty<VisualDto.FilterWhere>())
                        ExtractFieldsFromCondition(where.Condition, fields, aliasMap);
                }
            }

            return fields;
        }

        private void ExtractFieldsFromSubquery(VisualDto.SubqueryQuery query, List<VisualDto.Field> fields)
        {
            var aliasMap = BuildAliasMap(query.From);

            // SELECT columns
            foreach (var sel in query.Select ?? Enumerable.Empty<VisualDto.SelectExpression>())
            {
                var srcRef = sel.Column?.Expression?.SourceRef ?? new VisualDto.SourceRef();
                srcRef.Source = ResolveSource(srcRef.Source, aliasMap);

                var columnExpr = sel.Column ?? new VisualDto.ColumnSelect();
                columnExpr.Expression ??= new VisualDto.Expression();
                columnExpr.Expression.SourceRef ??= new VisualDto.SourceRef();
                columnExpr.Expression.SourceRef.Source = srcRef.Source;

                fields.Add(new VisualDto.Field
                {
                    Column = new VisualDto.ColumnField
                    {
                        Property = sel.Column.Property,
                        Expression = new VisualDto.Expression
                        {
                            SourceRef = columnExpr.Expression.SourceRef
                        }
                    }
                });
            }

            // ORDER BY measures
            foreach (var ob in query.OrderBy ?? Enumerable.Empty<VisualDto.OrderByExpression>())
            {
                var measureExpr = ob.Expression?.Measure?.Expression ?? new VisualDto.Expression();
                measureExpr.SourceRef ??= new VisualDto.SourceRef();
                measureExpr.SourceRef.Source = ResolveSource(measureExpr.SourceRef.Source, aliasMap);

                fields.Add(new VisualDto.Field
                {
                    Measure = new VisualDto.MeasureObject
                    {
                        Property = ob.Expression.Measure.Property,
                        Expression = measureExpr
                    }
                });
            }

            // Nested subqueries
            foreach (var from in query.From ?? Enumerable.Empty<VisualDto.FilterFrom>())
                if (from.Expression?.Subquery?.Query != null)
                    ExtractFieldsFromSubquery(from.Expression.Subquery.Query, fields);

            // WHERE conditions
            foreach (var where in query.Where ?? Enumerable.Empty<VisualDto.FilterWhere>())
                ExtractFieldsFromCondition(where.Condition, fields, aliasMap);
        }
        private Dictionary<string, string> BuildAliasMap(List<VisualDto.FilterFrom> fromList)
        {
            var map = new Dictionary<string, string>();
            foreach (var from in fromList ?? Enumerable.Empty<VisualDto.FilterFrom>())
            {
                if (!string.IsNullOrEmpty(from.Name) && !string.IsNullOrEmpty(from.Entity))
                    map[from.Name] = from.Entity;
            }
            return map;
        }

        private string ResolveSource(string source, Dictionary<string, string> aliasMap)
        {
            if (string.IsNullOrEmpty(source))
                return source;
            return aliasMap.TryGetValue(source, out var entity) ? entity : source;
        }

        private void ExtractFieldsFromCondition(VisualDto.Condition condition, List<VisualDto.Field> fields, Dictionary<string, string> aliasMap)
        {
            if (condition == null) return;

            // IN Expression
            if (condition.In != null)
            {
                foreach (var expr in condition.In.Expressions ?? Enumerable.Empty<VisualDto.ColumnSelect>())
                {
                    var srcRef = expr.Expression?.SourceRef ?? new VisualDto.SourceRef();
                    srcRef.Source = ResolveSource(srcRef.Source, aliasMap);

                    fields.Add(new VisualDto.Field
                    {
                        Column = new VisualDto.ColumnField
                        {
                            Property = expr.Property,
                            Expression = new VisualDto.Expression
                            {
                                SourceRef = srcRef
                            }
                        }
                    });
                }
            }

            // NOT Expression
            if (condition.Not != null)
                ExtractFieldsFromCondition(condition.Not.Expression, fields, aliasMap);

            // COMPARISON Expression
            if (condition.Comparison != null)
            {
                AddOperandField(condition.Comparison.Left, fields, aliasMap);
                AddOperandField(condition.Comparison.Right, fields, aliasMap);
            }
        }
        private void AddOperandField(VisualDto.FilterOperand operand, List<VisualDto.Field> fields, Dictionary<string, string> aliasMap)
        {
            if (operand == null) return;

            // MEASURE
            if (operand.Measure != null)
            {
                var srcRef = operand.Measure.Expression?.SourceRef ?? new VisualDto.SourceRef();
                srcRef.Source = ResolveSource(srcRef.Source, aliasMap);

                fields.Add(new VisualDto.Field
                {
                    Measure = new VisualDto.MeasureObject
                    {
                        Property = operand.Measure.Property,
                        Expression = new VisualDto.Expression
                        {
                            SourceRef = srcRef
                        }
                    }
                });
            }

            // COLUMN
            if (operand.Column != null)
            {
                var srcRef = operand.Column.Expression?.SourceRef ?? new VisualDto.SourceRef();
                srcRef.Source = ResolveSource(srcRef.Source, aliasMap);

                fields.Add(new VisualDto.Field
                {
                    Column = new VisualDto.ColumnField
                    {
                        Property = operand.Column.Property,
                        Expression = new VisualDto.Expression
                        {
                            SourceRef = srcRef
                        }
                    }
                });
            }
        }
        public IEnumerable<string> GetAllReferencedMeasures()
        {
            return GetAllFields()
                .Select(f => f.Measure)
                .Where(m => m?.Expression?.SourceRef?.Entity != null && m.Property != null)
                .Select(m => $"'{m.Expression.SourceRef.Entity}'[{m.Property}]")
                .Distinct();
        }

        public IEnumerable<string> GetAllReferencedColumns()
        {
            return GetAllFields()
                .Select(f => f.Column)
                .Where(c => c?.Expression?.SourceRef?.Entity != null && c.Property != null)
                .Select(c => $"'{c.Expression.SourceRef.Entity}'[{c.Property}]")
                .Distinct();
        }

        public void ReplaceMeasure(string oldFieldKey, Measure newMeasure, HashSet<VisualExtended> modifiedSet = null)
        {
            var newField = new VisualDto.Field
            {
                Measure = new VisualDto.MeasureObject
                {
                    Property = newMeasure.Name,
                    Expression = new VisualDto.Expression
                    {
                        SourceRef = new VisualDto.SourceRef { Entity = newMeasure.Table.Name }
                    }
                }
            };
            ReplaceField(oldFieldKey, newField, isMeasure: true, modifiedSet);
        }

        public void ReplaceColumn(string oldFieldKey, Column newColumn, HashSet<VisualExtended> modifiedSet = null)
        {
            var newField = new VisualDto.Field
            {
                Column = new VisualDto.ColumnField
                {
                    Property = newColumn.Name,
                    Expression = new VisualDto.Expression
                    {
                        SourceRef = new VisualDto.SourceRef { Entity = newColumn.Table.Name }
                    }
                }
            };
            ReplaceField(oldFieldKey, newField, isMeasure: false, modifiedSet);
        }

        private string ToFieldKey(VisualDto.Field f)
        {
            if (f?.Measure?.Expression?.SourceRef?.Entity is string mEntity && f.Measure.Property is string mProp)
                return $"'{mEntity}'[{mProp}]";

            if (f?.Column?.Expression?.SourceRef?.Entity is string cEntity && f.Column.Property is string cProp)
                return $"'{cEntity}'[{cProp}]";

            return null;
        }

        private void ReplaceField(string oldFieldKey, VisualDto.Field newField, bool isMeasure, HashSet<VisualExtended> modifiedSet = null)
        {
            var query = Content?.Visual?.Query;
            var objects = Content?.Visual?.Objects;
            bool wasModified = false;

            void Replace(VisualDto.Field f)
            {
                if (f == null) return;

                if (isMeasure && newField.Measure != null)
                {
                    // Preserve Expression with SourceRef
                    f.Measure ??= new VisualDto.MeasureObject();
                    f.Measure.Property = newField.Measure.Property;
                    f.Measure.Expression ??= new VisualDto.Expression();
                    f.Measure.Expression.SourceRef = newField.Measure.Expression?.SourceRef != null
                        ? new VisualDto.SourceRef
                        {
                            Entity = newField.Measure.Expression.SourceRef.Entity,
                            Source = newField.Measure.Expression.SourceRef.Source
                        }
                        : f.Measure.Expression.SourceRef;
                    f.Column = null;
                    wasModified = true;
                }
                else if (!isMeasure && newField.Column != null)
                {
                    // Preserve Expression with SourceRef
                    f.Column ??= new VisualDto.ColumnField();
                    f.Column.Property = newField.Column.Property;
                    f.Column.Expression ??= new VisualDto.Expression();
                    f.Column.Expression.SourceRef = newField.Column.Expression?.SourceRef != null
                        ? new VisualDto.SourceRef
                        {
                            Entity = newField.Column.Expression.SourceRef.Entity,
                            Source = newField.Column.Expression.SourceRef.Source
                        }
                        : f.Column.Expression.SourceRef;
                    f.Measure = null;
                    wasModified = true;
                }
            }

            void UpdateProjection(VisualDto.Projection proj)
            {
                if (proj == null) return;

                if (ToFieldKey(proj.Field) == oldFieldKey)
                {
                    Replace(proj.Field);

                    string entity = isMeasure
                        ? proj.Field.Measure.Expression?.SourceRef?.Entity
                        : proj.Field.Column.Expression?.SourceRef?.Entity;

                    string prop = isMeasure
                        ? proj.Field.Measure.Property
                        : proj.Field.Column.Property;

                    if (!string.IsNullOrEmpty(entity) && !string.IsNullOrEmpty(prop))
                    {
                        proj.QueryRef = $"{entity}.{prop}";
                    }

                    wasModified = true;
                }
            }

            foreach (var proj in query?.QueryState?.Values?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Y?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Y2?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Category?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Series?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Data?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var proj in query?.QueryState?.Rows?.Projections ?? Enumerable.Empty<VisualDto.Projection>())
                UpdateProjection(proj);

            foreach (var sort in query?.SortDefinition?.Sort ?? Enumerable.Empty<VisualDto.Sort>())
                if (ToFieldKey(sort.Field) == oldFieldKey) Replace(sort.Field);

            string oldMetadata = oldFieldKey.Replace("'", "").Replace("[", ".").Replace("]", "");
            string newMetadata = isMeasure
                ? $"{newField.Measure.Expression.SourceRef.Entity}.{newField.Measure.Property}"
                : $"{newField.Column.Expression.SourceRef.Entity}.{newField.Column.Property}";

            IEnumerable<VisualDto.ObjectProperties> AllObjectProperties() =>
                (objects?.DataPoint ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Data ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Labels ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Title ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Legend ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.General ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.ValueAxis ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.ReferenceLabel ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.ReferenceLabelDetail ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.ReferenceLabelValue ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Values ?? Enumerable.Empty<VisualDto.ObjectProperties>())
                .Concat(objects?.Y1AxisReferenceLine ?? Enumerable.Empty<VisualDto.ObjectProperties>());

            foreach (var obj in AllObjectProperties())
            {
                foreach (var prop in obj.Properties.Values.OfType<VisualDto.VisualObjectProperty>())
                {
                    var field = isMeasure ? new VisualDto.Field { Measure = prop.Expr?.Measure } : new VisualDto.Field { Column = prop.Expr?.Column };
                    if (ToFieldKey(field) == oldFieldKey)
                    {
                        if (prop.Expr != null)
                        {
                            if (isMeasure)
                            {
                                prop.Expr.Measure ??= new VisualDto.MeasureObject();
                                prop.Expr.Measure.Property = newField.Measure.Property;
                                prop.Expr.Measure.Expression ??= new VisualDto.Expression();
                                prop.Expr.Measure.Expression.SourceRef = newField.Measure.Expression?.SourceRef;
                                prop.Expr.Column = null;
                                wasModified = true;
                            }
                            else
                            {
                                prop.Expr.Column ??= new VisualDto.ColumnField();
                                prop.Expr.Column.Property = newField.Column.Property;
                                prop.Expr.Column.Expression ??= new VisualDto.Expression();
                                prop.Expr.Column.Expression.SourceRef = newField.Column.Expression?.SourceRef;
                                prop.Expr.Measure = null;
                                wasModified = true;
                            }
                        }
                    }

                    var fillInput = prop.Color?.Expr?.FillRule?.Input;
                    if (ToFieldKey(fillInput) == oldFieldKey)
                    {
                        if (isMeasure)
                        {
                            fillInput.Measure ??= new VisualDto.MeasureObject();
                            fillInput.Measure.Property = newField.Measure.Property;
                            fillInput.Measure.Expression ??= new VisualDto.Expression();
                            fillInput.Measure.Expression.SourceRef = newField.Measure.Expression?.SourceRef;
                            fillInput.Column = null;
                            wasModified = true;
                        }
                        else
                        {
                            fillInput.Column ??= new VisualDto.ColumnField();
                            fillInput.Column.Property = newField.Column.Property;
                            fillInput.Column.Expression ??= new VisualDto.Expression();
                            fillInput.Column.Expression.SourceRef = newField.Column.Expression?.SourceRef;
                            fillInput.Measure = null;
                            wasModified = true;
                        }
                    }

                    var solidInput = prop.Solid?.Color?.Expr?.FillRule?.Input;
                    if (ToFieldKey(solidInput) == oldFieldKey)
                    {
                        if (isMeasure)
                        {
                            solidInput.Measure ??= new VisualDto.MeasureObject();
                            solidInput.Measure.Property = newField.Measure.Property;
                            solidInput.Measure.Expression ??= new VisualDto.Expression();
                            solidInput.Measure.Expression.SourceRef = newField.Measure.Expression?.SourceRef;
                            solidInput.Column = null;
                            wasModified = true;
                        }
                        else
                        {
                            solidInput.Column ??= new VisualDto.ColumnField();
                            solidInput.Column.Property = newField.Column.Property;
                            solidInput.Column.Expression ??= new VisualDto.Expression();
                            solidInput.Column.Expression.SourceRef = newField.Column.Expression?.SourceRef;
                            solidInput.Measure = null;
                            wasModified = true;
                        }
                    }

                    var solidExpr = prop.Solid?.Color?.Expr;
                    if (solidExpr != null)
                    {
                        var solidField = isMeasure
                            ? new VisualDto.Field { Measure = solidExpr.Measure }
                            : new VisualDto.Field { Column = solidExpr.Column };

                        if (ToFieldKey(solidField) == oldFieldKey)
                        {
                            if (isMeasure)
                            {
                                solidExpr.Measure ??= new VisualDto.MeasureObject();
                                solidExpr.Measure.Property = newField.Measure.Property;
                                solidExpr.Measure.Expression ??= new VisualDto.Expression();
                                solidExpr.Measure.Expression.SourceRef = newField.Measure.Expression?.SourceRef;
                                solidExpr.Column = null;
                                wasModified = true;
                            }
                            else
                            {
                                solidExpr.Column ??= new VisualDto.ColumnField();
                                solidExpr.Column.Property = newField.Column.Property;
                                solidExpr.Column.Expression ??= new VisualDto.Expression();
                                solidExpr.Column.Expression.SourceRef = newField.Column.Expression?.SourceRef;
                                solidExpr.Measure = null;
                                wasModified = true;
                            }
                        }
                    }
                }

                if (obj.Selector?.Metadata == oldMetadata)
                {
                    obj.Selector.Metadata = newMetadata;
                    wasModified = true;
                }
            }

            if (wasModified && modifiedSet != null)
                modifiedSet.Add(this);
        }

    }


    public class PageExtended
    {
        public PageDto Page { get; set; }

        public ReportExtended ParentReport { get; set; }

        public int PageIndex
        {
            get
            {
                if (ParentReport == null || ParentReport.PagesConfig == null || ParentReport.PagesConfig.PageOrder == null)
                    return -1;
                return ParentReport.PagesConfig.PageOrder.IndexOf(Page.Name);
            }
        }


        public IList<VisualExtended> Visuals { get; set; } = new List<VisualExtended>();
        public string PageFilePath { get; set; }
    }


    public class ReportExtended
    {
        public IList<PageExtended> Pages { get; set; } = new List<PageExtended>();
        public string PagesFilePath { get; set; }
        public PagesDto PagesConfig { get; set; }
    }
