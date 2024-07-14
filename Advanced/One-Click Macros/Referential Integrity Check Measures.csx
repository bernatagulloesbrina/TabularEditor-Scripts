// 2023-02-15 / B.Agulló / Added useRelationship in the expression to check also inactive relationships
// 2023-12-22 / B.Agulló / Added suggestions by Ed Hansberry 
// 2024-07-13 / B.Agulló / Add annotations for the report-layer script, possible to execute without selected table, subtitles measures
//Select the desired table  to store all data quality measures
// or execute just on the model and a new table will be created for you
//See https://www.esbrina-ba.com/easy-management-of-referential-integrity/
//change the resulting variable names if you want
string overallCounterName = "Total Unmapped Items";
string overallDetailName = "Data Problems";
string targetTableNameIfCreated = "Referential Integrity";
//do not modify the script below
string overallCounterExpression = "";
string overallDetailExpression = "\"\"";
string annLabel = "ReferencialIntegrityMeasures";
string annValueTotal = "TotalUnmappedItems";
string annValueDetail = "DataProblems";
string annValueDataQualityMeasures = "DataQualityMeasure";
string annValueDataQualityTitles = "DataQualityTitle";
string annValueDataQualitySubtitles = "DataQualitySubitle";
string annValueFactColumn = "FactColumn";
Table tableToStoreMeasures = null as Table;
if (Selected.Tables.Count() == 0)
{
    tableToStoreMeasures = Model.AddCalculatedTable(targetTableNameIfCreated, "{0}");
}
else
{
    tableToStoreMeasures = Selected.Tables.First();
}
int measureIndex = 0;
foreach (var r in Model.Relationships)
{
    bool isOneToMany =
        r.FromCardinality == RelationshipEndCardinality.One
        && r.ToCardinality == RelationshipEndCardinality.Many;
    bool isManyToOne =
        r.FromCardinality == RelationshipEndCardinality.Many
        && r.ToCardinality == RelationshipEndCardinality.One;
    Column manyColumn = null as Column;
    Column oneColumn = null as Column;
    bool isOneToManyOrManyToOne = true;
    if (isOneToMany)
    {
        manyColumn = r.ToColumn;
        oneColumn = r.FromColumn;
    }
    else if (isManyToOne)
    {
        manyColumn = r.FromColumn;
        oneColumn = r.ToColumn;
    }
    else
    {
        isOneToManyOrManyToOne = false;
    }
    if (isOneToManyOrManyToOne)
    {
        measureIndex++; //increment index
        //add measure counting how many different items in the fact table are not present in the dimension
        string orphanCountExpression =
            "CALCULATE("
                + "SUMX(VALUES(" + manyColumn.DaxObjectFullName + "),1),"
                + "ISBLANK(" + oneColumn.DaxObjectFullName + "),"
                + "USERELATIONSHIP(" + manyColumn.DaxObjectFullName + "," + oneColumn.DaxObjectFullName + "),"
                + "ALLEXCEPT(" + manyColumn.Table.DaxObjectFullName + "," + manyColumn.DaxObjectFullName + ")"
            + ")";
        string orphanMeasureName =
            manyColumn.Name + " not mapped in " + manyColumn.Table.Name;
        Measure newCounter = tableToStoreMeasures.AddMeasure(name: orphanMeasureName, expression: orphanCountExpression, displayFolder: "_Data quality Measures");
        newCounter.FormatString = "#,##0";
        newCounter.FormatDax();
        newCounter.SetAnnotation(annLabel, annValueDataQualityMeasures + "_" + measureIndex.ToString());
        //add annotation to trace back the fact table when building the report
        manyColumn.SetAnnotation(annLabel, annValueFactColumn + "_" + measureIndex.ToString());
        //add measure saying how many are not mapped in the fact table
        string orphanTableTitleMeasureExpression = "FORMAT(" + newCounter.DaxObjectFullName +"+0,\"" + newCounter.FormatString + "\") & \" " + newCounter.Name + "\"";
        string orphanTableTitleMeasureName = newCounter.Name + " Title";
        Measure newTitle = tableToStoreMeasures.AddMeasure(name: orphanTableTitleMeasureName, expression: orphanTableTitleMeasureExpression, displayFolder: "_Data quality Titles");
        newTitle.FormatDax();
        newTitle.SetAnnotation(annLabel, annValueDataQualityTitles + "_" + measureIndex.ToString());
        //add measure for subtitle saying how many need to be added to the dimension table
        string orphanTableSubtitleMeasureExpression =
            String.Format(
                @"FORMAT({0}+0,""{1}"") & "" values missing in "" & ""{2}""", 
                newCounter.DaxObjectFullName, 
                newCounter.FormatString, 
                oneColumn.Table.Name);
        string orphanTableSubitleMeasureName = newCounter.Name + " Subtitle";
        Measure newSubtitle = tableToStoreMeasures.AddMeasure(name: orphanTableSubitleMeasureName, expression: orphanTableSubtitleMeasureExpression, displayFolder: "_Data quality Subtitles");
        newSubtitle.FormatDax();
        newSubtitle.SetAnnotation(annLabel, annValueDataQualitySubtitles + "_" + measureIndex.ToString());
        overallCounterExpression = overallCounterExpression + "+" + newCounter.DaxObjectFullName;
        overallDetailExpression = overallDetailExpression
                + " & IF(" + newCounter.DaxObjectFullName + "> 0,"
                            + newTitle.DaxObjectFullName + " & UNICHAR(10))";
    };
};
Measure counter = tableToStoreMeasures.AddMeasure(name: overallCounterName, expression: overallCounterExpression);
counter.FormatString = "#,##0";
counter.FormatDax();
counter.SetAnnotation(annLabel, annValueTotal);
Measure descr = tableToStoreMeasures.AddMeasure(name: overallDetailName, expression: overallDetailExpression);
descr.FormatDax();
descr.SetAnnotation(annLabel, annValueDetail);
