// '2026-02-20 / B.Agullo / Added Properties annotation
// '1990-01-01 / D.Otykier
 
// Loop through all currently selected columns:
foreach(var c in Selected.Columns)
{
    var newMeasure = c.Table.AddMeasure(
        "Sum of " + c.Name,                    // Name
        "SUM(" + c.DaxObjectFullName + ")",    // DAX expression
        c.DisplayFolder                        // Display Folder
    );
    
    // Set the format string on the new measure:
    newMeasure.FormatString = c.FormatString;

    // Provide some documentation:
    newMeasure.Description = "This measure is the sum of column " + c.DaxObjectFullName;
    newMeasure.SetAnnotation("Properties","SUM|" + c.Name + "|" + c.Table.Name); 
    // Hide the base column:
    c.IsHidden = true;
}
