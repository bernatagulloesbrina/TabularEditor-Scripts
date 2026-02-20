// '2026-02-20 / B.Agullo / Added Properties annotation
// '1990-01-01 / D.Otykier
 
 
// Loop through all currently selected columns:
foreach(var c in Selected.Columns)
{
    var newMeasure = c.Table.AddMeasure(
        "Average of " + c.Name,                    // Name
        "AVERAGE(" + c.DaxObjectFullName + ")",    // DAX expression
        c.DisplayFolder                        // Display Folder
    );
    
    // Set the format string on the new measure:
    newMeasure.FormatString = c.FormatString;

    // Provide some documentation:
    newMeasure.Description = "This measure is the average of column " + c.DaxObjectFullName;
    newMeasure.SetAnnotation("Properties","AVERAGE|" + c.Name + "|" + c.Table.Name); 


}
