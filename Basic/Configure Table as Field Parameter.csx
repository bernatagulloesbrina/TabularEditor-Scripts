//'2024-08-09 / B.Agullo /
// Configure a field parameter table
if(Selected.Tables.Count() != 1)
{
    Error("Select a single table and try again");
    return;
}
Table fieldParameterTable = Selected.Table;
if (fieldParameterTable.Columns.Count() < 3)
{
    Error("This script expects at least 3 columns in the table");
    return;
}
Column displayNameColumn = SelectColumn(fieldParameterTable, fieldParameterTable.Columns[0], "Select display name column");
Column fieldColumn = SelectColumn(fieldParameterTable, fieldParameterTable.Columns[1], "Select field table");
Column orderColumn = SelectColumn(fieldParameterTable, fieldParameterTable.Columns[2], "Select order column");
fieldColumn.SetExtendedProperty(name: "ParameterMetadata", value: @"{""version"":3,""kind"":2}", type: ExtendedPropertyType.Json);
displayNameColumn.GroupByColumns.Add(fieldColumn);
displayNameColumn.SortByColumn = orderColumn;
fieldColumn.SortByColumn = orderColumn;
fieldColumn.IsHidden = true;
orderColumn.IsHidden = true; 
