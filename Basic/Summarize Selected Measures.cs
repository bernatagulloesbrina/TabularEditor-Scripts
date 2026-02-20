#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;

//2026-02-20 / B.Agullo / 

//Select 1 or more measures and will generate a new measure showing the values of each of them concatenating the result and name for each of them



if(Selected.Measures.Count <= 1) {
    Error("Select two or more measures"); 
    return; 
} 

string newMeasureName = Interaction.InputBox("New Measure name", "Name", "Summary of " + Selected.Measures.Count + " measures", 740, 400);

string newMeasureExpression = ""; 
string measureTable = ""; 

foreach(var iMeasure in Selected.Measures) { 
    if(measureTable == "") measureTable = iMeasure.Table.Name; 

    if(newMeasureExpression == "") {
        newMeasureExpression = "IF(ISBLANK([" + iMeasure.Name +"]),BLANK(),FORMAT([" + iMeasure.Name + "],\"#,0\") & \" " + iMeasure.Name + "\")";
    } else {
        newMeasureExpression += " & UNICHAR(10) & IF(ISBLANK([" + iMeasure.Name +"]),BLANK(),FORMAT([" + iMeasure.Name + "],\"#,0\") & \" " + iMeasure.Name + "\")";
    }; 
};
var newMeasure = Model.Tables[measureTable].AddMeasure(newMeasureName,newMeasureExpression); 
