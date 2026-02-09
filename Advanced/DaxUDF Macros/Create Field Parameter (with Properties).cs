#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
//80% comes from Daniel Otykier --> https://github.com/TabularEditor/TabularEditor3/issues/541#issuecomment-1129228481
//20% B.Agullo --> pop-up to choose parameter name & properties expansion. 
// Before running the script, select the measures or columns that you
// would like to use as field parameters (hold down CTRL to select multiple
// objects). 
// this script relies on the properties annotaiton generated when creating measures 
// based on DAXUDFs and the follwoing script: https://github.com/bernatagulloesbrina/TabularEditor-Scripts/blob/main/Advanced/DaxUDF%20Macros/create%20measures%20from%20daxudfs.csx


var name = Interaction.InputBox("Provide the name for the field parameter", "Parameter name", "Parameter");
if (Selected.Columns.Count == 0 && Selected.Measures.Count == 0) throw new Exception("No columns or measures selected!");
// Get objects collection
var objects = Selected.Columns.Any() ? Selected.Columns.Cast<ITabularTableObject>() : Selected.Measures;
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
