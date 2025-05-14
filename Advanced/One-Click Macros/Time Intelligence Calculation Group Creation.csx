#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;

//
// CHANGELOG:
// '2021-05-01 / B.Agullo / 
// '2021-05-17 / B.Agullo / added affected measure table
// '2021-06-19 / B.Agullo / data label measures
// '2021-07-10 / B.Agullo / added flag expression to avoid breaking already special format strings
// '2021-09-23 / B.Agullo / added code to prompt for parameters (code credit to Daniel Otykier) 
// '2021-09-27 / B.Agullo / added code for general name 
// '2022-10-11 / B.Agullo / added MMT and MWT calc item groups
// '2023-01-24 / B.Agullo / added Date Range Measure and completed dynamic label for existing items
// '2025-05-14 / B.Agullo / some refactoring + month based standard calculations
//
// by Bernat Agulló
// twitter: @AgulloBernat
// www.esbrina-ba.com/blog
//
// REFERENCE: 
// Check out https://www.esbrina-ba.com/time-intelligence-the-smart-way/ where this script is introduced
// 
// FEATURED: 
// this script featured in GuyInACube https://youtu.be/_j0iTUo2HT0
//
// THANKS:
// shout out to Johnny Winter for the base script and SQLBI for daxpatterns.com
//select the measures that you want to be affected by the calculation group
//before running the script. 
//measure names can also be included in the following array (no need to select them) 
string[] preSelectedMeasures = { }; //include measure names in double quotes, like: {"Profit","Total Cost"};
//AT LEAST ONE MEASURE HAS TO BE AFFECTED!, 
//either by selecting it or typing its name in the preSelectedMeasures Variable
//
// ----- do not modify script below this line -----
//
string affectedMeasures = "{";
int i = 0;
for (i = 0; i < preSelectedMeasures.GetLength(0); i++)
{
    if (affectedMeasures == "{")
    {
        affectedMeasures = affectedMeasures + "\"" + preSelectedMeasures[i] + "\"";
    }
    else
    {
        affectedMeasures = affectedMeasures + ",\"" + preSelectedMeasures[i] + "\"";
    };
};
if (Selected.Measures.Count != 0)
{
    foreach (var m in Selected.Measures)
    {
        if (affectedMeasures == "{")
        {
            affectedMeasures = affectedMeasures + "\"" + m.Name + "\"";
        }
        else
        {
            affectedMeasures = affectedMeasures + ",\"" + m.Name + "\"";
        };
    };
};
//check that by either method at least one measure is affected
if (affectedMeasures == "{")
{
    Error("No measures affected by calc group");
    return;
};
string calcGroupName = String.Empty;
string columnName = String.Empty;
if (Model.CalculationGroups.Any(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group"))
{
    calcGroupName = Model.CalculationGroups.Where(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group").First().Name;
}
else
{
    calcGroupName = Interaction.InputBox("Provide a name for your Calc Group", "Calc Group Name", "Time Intelligence", 740, 400);
};
if (calcGroupName == String.Empty) return;
if (Model.CalculationGroups.Any(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group"))
{
    columnName = Model.Tables.Where(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group").First().Columns.First().Name;
}
else
{
    columnName = Interaction.InputBox("Provide a name for your Calc Group Column", "Calc Group Column Name", calcGroupName, 740, 400);
};
if (columnName == String.Empty) return;
string affectedMeasuresTableName = String.Empty;
if (Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table"))
{
    affectedMeasuresTableName = Model.Tables.Where(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table").First().Name;
}
else
{
    affectedMeasuresTableName = Interaction.InputBox("Provide a name for affected measures table", "Affected Measures Table Name", calcGroupName + " Affected Measures", 740, 400);
};
if (affectedMeasuresTableName ==String.Empty) return;
string affectedMeasuresColumnName = String.Empty;
if (Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table"))
{
    affectedMeasuresColumnName = Model.Tables.Where(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table").First().Columns.First().Name;
}
else
{
    affectedMeasuresColumnName = Interaction.InputBox("Provide a name for affected measures column", "Affected Measures Table Column Name", "Measure", 740, 400);
};
if (affectedMeasuresColumnName == String.Empty) return;
//string affectedMeasuresColumnName = "Measure"; 
string labelAsValueMeasureName = "Label as Value Measure";
string labelAsFormatStringMeasureName = "Label as format string";
// '2021-09-24 / B.Agullo / model object selection prompts! 
var factTable = SelectTable(label: "Select your fact table");
if (factTable == null) return;
var factTableDateColumn = SelectColumn(factTable.Columns, label: "Select the main date column");
if (factTableDateColumn == null) return;
Table dateTableCandidate = null;
if (Model.Tables.Any
    (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table"
        || x.Name == "Date"
        || x.Name == "Calendar"))
{
    dateTableCandidate = Model.Tables.Where
        (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table"
            || x.Name == "Date"
            || x.Name == "Calendar").First();
};
var dateTable =
    SelectTable(
        label: "Select your date table",
        preselect: dateTableCandidate);
if (dateTable == null)
{
    Error("You just aborted the script");
    return;
}
else
{
    dateTable.SetAnnotation("@AgulloBernat", "Time Intel Date Table");
};
Column dateTableDateColumnCandidate = null;
if (dateTable.Columns.Any
            (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Date Column" || x.Name == "Date"))
{
    dateTableDateColumnCandidate = dateTable.Columns.Where
        (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Date Column" || x.Name == "Date").First();
};
var dateTableDateColumn =
    SelectColumn(
        dateTable.Columns,
        label: "Select the date column",
        preselect: dateTableDateColumnCandidate);
if (dateTableDateColumn == null)
{
    Error("You just aborted the script");
    return;
}
else
{
    dateTableDateColumn.SetAnnotation("@AgulloBernat", "Time Intel Date Table Date Column");
};
Column dateTableYearColumnCandidate = null;
if (dateTable.Columns.Any(x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Year Column" || x.Name == "Year"))
{
    dateTable.Columns.Where
        (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Year Column" || x.Name == "Year").First();
};
var dateTableYearColumn =
    SelectColumn(
        dateTable.Columns,
        label: "Select the year column",
        preselect: dateTableYearColumnCandidate);
if (dateTableYearColumn == null)
{
    Error("You just abourted the script");
    return;
}
else
{
    dateTableYearColumn.SetAnnotation("@AgulloBernat", "Time Intel Date Table Year Column");
};
//these names are for internal use only, so no need to be super-fancy, better stick to datpatterns.com model
string ShowValueForDatesMeasureName = "ShowValueForDates";
string dateWithSalesColumnName = "DateWith" + factTable.Name;
// '2021-09-24 / B.Agullo / I put the names back to variables so I don't have to tough the script
string factTableName = factTable.Name;
string factTableDateColumnName = factTableDateColumn.Name;
string dateTableName = dateTable.Name;
string dateTableDateColumnName = dateTableDateColumn.Name;
string dateTableYearColumnName = dateTableYearColumn.Name;
// '2021-09-24 / B.Agullo / this is for internal use only so better leave it as is 
string flagExpression = "UNICHAR( 8204 )";
string calcItemProtection = "<CODE>"; //default value if user has selected no measures
string calcItemFormatProtection = "<CODE>"; //default value if user has selected no measures
// check if there's already an affected measure table
if (Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table"))
{
    //modifying an existing calculated table is not risk-free
    Info("Make sure to include measure names to the table " + affectedMeasuresTableName);
}
else
{
    // create calculated table containing all names of affected measures
    // this is why you need to enable 
    if (affectedMeasures != "{")
    {
        affectedMeasures = affectedMeasures + "}";
        string affectedMeasureTableExpression =
            "SELECTCOLUMNS(" + affectedMeasures + ",\"" + affectedMeasuresColumnName + "\",[Value])";
        var affectedMeasureTable =
            Model.AddCalculatedTable(affectedMeasuresTableName, affectedMeasureTableExpression);
        affectedMeasureTable.FormatDax();
        affectedMeasureTable.Description =
            "Measures affected by " + calcGroupName + " calculation group.";
        affectedMeasureTable.SetAnnotation("@AgulloBernat", "Time Intel Affected Measures Table");
        // this causes error
        // affectedMeasureTable.Columns[affectedMeasuresColumnName].SetAnnotation("@AgulloBernat","Time Intel Affected Measures Table Column");
        affectedMeasureTable.IsHidden = true;
    };
};
//if there where selected or preselected measures, prepare protection code for expresion and formatstring
string affectedMeasuresValues = "VALUES('" + affectedMeasuresTableName + "'[" + affectedMeasuresColumnName + "])";
calcItemProtection =
    "SWITCH(" +
    "   TRUE()," +
    "   SELECTEDMEASURENAME() IN " + affectedMeasuresValues + "," +
    "   <CODE> ," +
    "   ISSELECTEDMEASURE([" + labelAsValueMeasureName + "])," +
    "   <LABELCODE> ," +
    "   SELECTEDMEASURE() " +
    ")";
calcItemFormatProtection =
    "SWITCH(" +
    "   TRUE() ," +
    "   SELECTEDMEASURENAME() IN " + affectedMeasuresValues + "," +
    "   <CODE> ," +
    "   ISSELECTEDMEASURE([" + labelAsFormatStringMeasureName + "])," +
    "   <LABELCODEFORMATSTRING> ," +
    "   SELECTEDMEASUREFORMATSTRING() " +
    ")";
string dateColumnWithTable = "'" + dateTableName + "'[" + dateTableDateColumnName + "]";
string yearColumnWithTable = "'" + dateTableName + "'[" + dateTableYearColumnName + "]";
string factDateColumnWithTable = "'" + factTableName + "'[" + factTableDateColumnName + "]";
string dateWithSalesWithTable = "'" + dateTableName + "'[" + dateWithSalesColumnName + "]";
string calcGroupColumnWithTable = "'" + calcGroupName + "'[" + columnName + "]";
//check to see if a table with this name already exists
//if it doesnt exist, create a calculation group with this name
if (!Model.Tables.Contains(calcGroupName))
{
    var cg = Model.AddCalculationGroup(calcGroupName);
    cg.Description = "Calculation group for time intelligence. Availability of data is taken from " + factTableName + ".";
    cg.SetAnnotation("@AgulloBernat", "Time Intel Calc Group");
};
//set variable for the calc group
Table calcGroup = Model.Tables[calcGroupName];
//if table already exists, make sure it is a Calculation Group type
if (calcGroup.SourceType.ToString() != "CalculationGroup")
{
    Error("Table exists in Model but is not a Calculation Group. Rename the existing table or choose an alternative name for your Calculation Group.");
    return;
};
//adds the two measures that will be used for label as value, label as format string 
var labelAsValueMeasure = calcGroup.AddMeasure(labelAsValueMeasureName, "");
labelAsValueMeasure.Description = "Use this measure to show the year evaluated in tables";
var labelAsFormatStringMeasure = calcGroup.AddMeasure(labelAsFormatStringMeasureName, "0");
labelAsFormatStringMeasure.Description = "Use this measure to show the year evaluated in charts";
//by default the calc group has a column called Name. If this column is still called Name change this in line with specfied variable
if (calcGroup.Columns.Contains("Name"))
{
    calcGroup.Columns["Name"].Name = columnName;
};
calcGroup.Columns[columnName].Description = "Select value(s) from this column to apply time intelligence calculations.";
calcGroup.Columns[columnName].SetAnnotation("@AgulloBernat", "Time Intel Calc Group Column");
//Only create them if not in place yet (reruns)
if (!Model.Tables[dateTableName].Columns.Any(C => C.GetAnnotation("@AgulloBernat") == "Date with Data Column"))
{
    string DateWithSalesCalculatedColumnExpression =
        dateColumnWithTable + " <= MAX ( " + factDateColumnWithTable + ")";
    Column dateWithDataColumn = dateTable.AddCalculatedColumn(dateWithSalesColumnName, DateWithSalesCalculatedColumnExpression);
    dateWithDataColumn.SetAnnotation("@AgulloBernat", "Date with Data Column");
};
if (!Model.Tables[dateTableName].Measures.Any(M => M.Name == ShowValueForDatesMeasureName))
{
    string ShowValueForDatesMeasureExpression = String.Format(
        @"VAR LastDateWithData = 
            CALCULATE(
                MAX({0}),
                REMOVEFILTERS()
            )
        VAR FirstDateVisible = 
            MIN({1})
        VAR Result = 
            FirstDateVisible <= LastDateWithData
        RETURN 
            Result",
        factDateColumnWithTable,
        dateColumnWithTable
    );
    var ShowValueForDatesMeasure = dateTable.AddMeasure(ShowValueForDatesMeasureName, ShowValueForDatesMeasureExpression);
    ShowValueForDatesMeasure.FormatDax();
};
// Defining expressions and format strings for each calc item
string CY = @"/*CY*/ SELECTEDMEASURE()";
string CYlabel = String.Format(@"SELECTEDVALUE({0})", yearColumnWithTable);
string PY = String.Format(
    @"/*PY*/ 
    IF (
        [{0}], 
        CALCULATE ( 
            {1}, 
            CALCULATETABLE ( 
                DATEADD ( {2}, -1, YEAR ), 
                {3} = TRUE 
            ) 
        ) 
    )", 
    ShowValueForDatesMeasureName, CY, dateColumnWithTable, dateWithSalesWithTable);
string PYlabel = String.Format(
    @"/*PY*/ 
    IF (
        [{0}], 
        CALCULATE ( 
            {1}, 
            CALCULATETABLE ( 
                DATEADD ( {2}, -1, YEAR ), 
                {3} = TRUE 
            ) 
        ) 
    )", 
    ShowValueForDatesMeasureName, CYlabel, dateColumnWithTable, dateWithSalesWithTable);
string YOY = String.Format(
    @"/*YOY*/ 
    VAR ValueCurrentPeriod = {0} 
    VAR ValuePreviousPeriod = {1} 
    VAR Result = 
    IF ( 
        NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
        ValueCurrentPeriod - ValuePreviousPeriod
    ) 
    RETURN 
       Result", 
    CY, PY);
string YOYlabel = String.Format(
    @"/*YOY*/ 
    VAR ValueCurrentPeriod = {0} 
    VAR ValuePreviousPeriod = {1} 
    VAR Result = 
    IF ( 
        NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
        ValueCurrentPeriod & "" vs "" & ValuePreviousPeriod
    ) 
    RETURN 
       Result", 
    CYlabel, PYlabel);
string YOYpct = String.Format(
    @"/*YOY%*/ 
    VAR ValueCurrentPeriod = {0} 
    VAR ValuePreviousPeriod = {1} 
    VAR CurrentMinusPreviousPeriod = 
    IF ( 
        NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
        ValueCurrentPeriod - ValuePreviousPeriod
    ) 
    VAR Result = 
    DIVIDE ( 
        CurrentMinusPreviousPeriod,
        ValuePreviousPeriod
    ) 
    RETURN 
      Result", 
    CY, PY);
string YOYpctLabel = String.Format(
    @"/*YOY%*/ 
    VAR ValueCurrentPeriod = {0} 
    VAR ValuePreviousPeriod = {1} 
    VAR Result = 
    IF ( 
        NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
        ValueCurrentPeriod & "" vs "" & ValuePreviousPeriod & "" (%)""
    ) 
    RETURN 
      Result", 
    CYlabel, PYlabel);
string YTD = String.Format(
    @"/*YTD*/
    IF (
        [{0}],
        CALCULATE (
            {1},
            DATESYTD ({2})
       )
    )", 
    ShowValueForDatesMeasureName, CY, dateColumnWithTable);
string YTDlabel = String.Format(@"{0} & "" YTD""", CYlabel);
string PYTD = String.Format(
    @"/*PYTD*/
    IF ( 
        [{0}], 
       CALCULATE ( 
           {1},
        CALCULATETABLE ( 
            DATEADD ( {2}, -1, YEAR ), 
           {3} = TRUE 
           )
       )
    )", 
    ShowValueForDatesMeasureName, YTD, dateColumnWithTable, dateWithSalesWithTable);
string PYTDlabel = String.Format(@"{0} & "" YTD""", PYlabel);
string YOYTD = String.Format(
    @"/*YOYTD*/
    VAR ValueCurrentPeriod = {0} 
    VAR ValuePreviousPeriod = {1} 
    VAR Result = 
    IF ( 
        NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
        ValueCurrentPeriod - ValuePreviousPeriod
    ) 
    RETURN 
       Result", 
    YTD, PYTD);
string YOYTDlabel = String.Format(
    @"/*YOYTD*/
    VAR ValueCurrentPeriod = {0} 
    VAR ValuePreviousPeriod = {1} 
    VAR Result = 
    IF ( 
        NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
        ValueCurrentPeriod & "" vs "" & ValuePreviousPeriod
    ) 
    RETURN 
       Result", 
    YTDlabel, PYTDlabel);
string YOYTDpct = String.Format(
    @"/*YOYTD%*/
    VAR ValueCurrentPeriod = {0} 
    VAR ValuePreviousPeriod = {1} 
    VAR CurrentMinusPreviousPeriod = 
    IF ( 
        NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
        ValueCurrentPeriod - ValuePreviousPeriod
    ) 
    VAR Result = 
    DIVIDE ( 
        CurrentMinusPreviousPeriod,
        ValuePreviousPeriod
    ) 
    RETURN 
      Result", 
    YTD, PYTD);
string YOYTDpctLabel = String.Format(
    @"/*YOY%*/ 
    VAR ValueCurrentPeriod = {0} 
    VAR ValuePreviousPeriod = {1} 
    VAR Result = 
    IF ( 
        NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
        ValueCurrentPeriod & "" vs "" & ValuePreviousPeriod & "" (%)""
    ) 
    RETURN 
      Result", 
    YTDlabel, PYTDlabel);
    string CM = @"/*CM*/ SELECTEDMEASURE()";
    string CMlabel = String.Format(@"SELECTEDVALUE({0}, ""Current Month"")", dateColumnWithTable);
    string PM = String.Format(
        @"/*PM*/ 
        IF (
            [{0}], 
            CALCULATE ( 
                {1}, 
                CALCULATETABLE ( 
                    DATEADD ( {2}, -1, MONTH ), 
                    {3} = TRUE 
                ) 
            ) 
        )", 
        ShowValueForDatesMeasureName, CM, dateColumnWithTable, dateWithSalesWithTable);
    string PMlabel = String.Format(
        @"/*PM*/ 
        IF (
            [{0}], 
            CALCULATE ( 
                {1}, 
                CALCULATETABLE ( 
                    DATEADD ( {2}, -1, MONTH ), 
                    {3} = TRUE 
                ) 
            ) 
        )", 
        ShowValueForDatesMeasureName, CMlabel, dateColumnWithTable, dateWithSalesWithTable);
    string MOM = String.Format(
        @"/*MOM*/ 
        VAR ValueCurrentPeriod = {0} 
        VAR ValuePreviousPeriod = {1} 
        VAR Result = 
        IF ( 
            NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
            ValueCurrentPeriod - ValuePreviousPeriod
        ) 
        RETURN 
           Result", 
        CM, PM);
    string MOMlabel = String.Format(
        @"/*MOM*/ 
        VAR ValueCurrentPeriod = {0} 
        VAR ValuePreviousPeriod = {1} 
        VAR Result = 
        IF ( 
            NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
            ValueCurrentPeriod & "" vs "" & ValuePreviousPeriod
        ) 
        RETURN 
           Result", 
        CMlabel, PMlabel);
    string MOMpct = String.Format(
        @"/*MOM%*/ 
        VAR ValueCurrentPeriod = {0} 
        VAR ValuePreviousPeriod = {1} 
        VAR CurrentMinusPreviousPeriod = 
        IF ( 
            NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
            ValueCurrentPeriod - ValuePreviousPeriod
        ) 
        VAR Result = 
        DIVIDE ( 
            CurrentMinusPreviousPeriod,
            ValuePreviousPeriod
        ) 
        RETURN 
          Result", 
        CM, PM);
    string MOMpctLabel = String.Format(
        @"/*MOM%*/ 
        VAR ValueCurrentPeriod = {0} 
        VAR ValuePreviousPeriod = {1} 
        VAR Result = 
        IF ( 
            NOT ISBLANK ( ValueCurrentPeriod ) && NOT ISBLANK ( ValuePreviousPeriod ), 
            ValueCurrentPeriod & "" vs "" & ValuePreviousPeriod & "" (%)""
        ) 
        RETURN 
          Result", 
        CMlabel, PMlabel);
        string MTD = String.Format(
            @"/*MTD*/
            IF (
                [{0}],
                CALCULATE (
                    SELECTEDMEASURE(),
                    DATESMTD({1})
                )
            )",
            ShowValueForDatesMeasureName, dateColumnWithTable);
        string MTDlabel = String.Format(
            @"/*MTD*/
            IF (
                [{0}],
                CALCULATE (
                    ""Month to date"",
                    DATESMTD({1})
                )
            )",
            ShowValueForDatesMeasureName, dateColumnWithTable);
        string PMTD = String.Format(
            @"/*PMTD*/
            IF (
                [{0}],
                CALCULATE (
                    SELECTEDMEASURE(),
                    DATESMTD(DATEADD({1}, -1, MONTH))
                )
            )",
            ShowValueForDatesMeasureName, dateColumnWithTable);
        string PMTDlabel = String.Format(
            @"/*PMTD*/
            IF (
                [{0}],
                CALCULATE (
                    ""Previous month to date"",
                    DATESMTD(DATEADD({1}, -1, MONTH))
                )
            )",
            ShowValueForDatesMeasureName, dateColumnWithTable);
        string MOMTD = String.Format(
            @"/*MOMTD*/
            VAR ValueCurrentPeriod = {0}
            VAR ValuePreviousPeriod = {1}
            VAR Result =
                IF (
                    NOT ISBLANK(ValueCurrentPeriod) && NOT ISBLANK(ValuePreviousPeriod),
                    ValueCurrentPeriod - ValuePreviousPeriod
                )
            RETURN
                Result",
            MTD, PMTD);
        string MOMTDlabel = String.Format(
            @"/*MOMTD*/
            VAR ValueCurrentPeriod = {0}
            VAR ValuePreviousPeriod = {1}
            VAR Result =
                IF (
                    NOT ISBLANK(ValueCurrentPeriod) && NOT ISBLANK(ValuePreviousPeriod),
                    ValueCurrentPeriod & "" vs "" & ValuePreviousPeriod
                )
            RETURN
                Result",
            MTDlabel, PMTDlabel);
        string MOMTDpct = String.Format(
            @"/*MOMTD%*/
            VAR ValueCurrentPeriod = {0}
            VAR ValuePreviousPeriod = {1}
            VAR CurrentMinusPreviousPeriod =
                IF (
                    NOT ISBLANK(ValueCurrentPeriod) && NOT ISBLANK(ValuePreviousPeriod),
                    ValueCurrentPeriod - ValuePreviousPeriod
                )
            VAR Result =
                DIVIDE(
                    CurrentMinusPreviousPeriod,
                    ValuePreviousPeriod
                )
            RETURN
                Result",
            MTD, PMTD);
        string MOMTDpctlabel = String.Format(
            @"/*MOMTD%*/
            VAR ValueCurrentPeriod = {0}
            VAR ValuePreviousPeriod = {1}
            VAR Result =
                IF (
                    NOT ISBLANK(ValueCurrentPeriod) && NOT ISBLANK(ValuePreviousPeriod),
                    ValueCurrentPeriod & "" vs "" & ValuePreviousPeriod & "" (%)""
                )
            RETURN
                Result",
            MTDlabel, PMTDlabel);
string MAT = String.Format(@"
/*TAM*/
IF (
    [{0}],
    CALCULATE (
        SELECTEDMEASURE(),
        DATESINPERIOD (
{1},
MAX({1}),
-1,
YEAR
        )
    )
)", ShowValueForDatesMeasureName, dateColumnWithTable);
string MATlabel = String.Format(@"
/*TAM*/
IF (
    [{0}],
    CALCULATE (
        ""Year ending "" & FORMAT(MAX('Date'[Date]), ""d-MMM-yyyy"", ""en-US""),
        DATESINPERIOD (
{1},
MAX({1}),
-1,
YEAR
        )
    )
)", ShowValueForDatesMeasureName, dateColumnWithTable);
string MATminus1 = String.Format(@"
/*TAM*/
IF (
    [{0}],
    CALCULATE (
        SELECTEDMEASURE(),
        DATESINPERIOD (
{1},
LASTDATE(DATEADD({1}, -1, YEAR)),
-1,
YEAR
        )
    )
)", ShowValueForDatesMeasureName, dateColumnWithTable);
string MATminus1label = String.Format(@"
/*MAT-1*/
IF (
    [{0}],
    CALCULATE (
        ""Year ending "" & FORMAT(MAX('Date'[Date]), ""d-MMM-yyyy"", ""en-US""),
        DATESINPERIOD (
{1},
LASTDATE(DATEADD({1}, -1, YEAR)),
-1,
YEAR
        )
    )
)", ShowValueForDatesMeasureName, dateColumnWithTable);
string MATvsMATminus1 = String.Format(@"
/*MAT vs MAT-1*/
VAR MAT = {0}
VAR MAT_1 = {1}
RETURN 
    IF(ISBLANK(MAT) || ISBLANK(MAT_1), BLANK(), MAT - MAT_1)
", MAT, MATminus1);
string MATvsMATminus1label = String.Format(@"
/*MAT vs MAT-1*/
VAR MAT = {0}
VAR MAT_1 = {1}
RETURN 
    IF(ISBLANK(MAT) || ISBLANK(MAT_1), BLANK(), MAT & "" vs "" & MAT_1)
", MATlabel, MATminus1label);
string MATvsMATminus1pct = String.Format(@"
/*MAT vs MAT-1(%)*/
VAR MAT = {0}
VAR MAT_1 = {1}
RETURN
    IF(
        ISBLANK(MAT) || ISBLANK(MAT_1),
        BLANK(),
        DIVIDE(MAT - MAT_1, MAT_1)
    )
", MAT, MATminus1);
string MATvsMATminus1pctlabel = String.Format(@"
/*MAT vs MAT-1 (%)*/
VAR MAT = {0}
VAR MAT_1 = {1}
RETURN 
    IF(ISBLANK(MAT) || ISBLANK(MAT_1), BLANK(), MAT & "" vs "" & MAT_1 & "" (%)"")
", MATlabel, MATminus1label);
string MMT = String.Format(
    @"/*MMT*/
        IF(
[{0}],
CALCULATE( SELECTEDMEASURE( ), DATESINPERIOD( {1}, MAX( {1} ), -1, MONTH ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable);
string MMTlabel = String.Format(
    @"/*MMT*/
        IF(
[{0}],
CALCULATE( {2}, DATESINPERIOD( {1}, MAX( {1} ), -1, MONTH ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable, "\"Month ending \" & FORMAT(MAX( 'Date'[Date] ),\"d-MMM-yyyy\",\"en-US\")");
string MMTminus1 = String.Format(
    @"/*MMT*/
        IF(
[{0}],
CALCULATE( SELECTEDMEASURE( ), DATESINPERIOD( {1}, LASTDATE( DATEADD( {1}, -1, MONTH ) ), -1, MONTH ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable);
string MMTminus1label = String.Format(
    @"/*MMT-1*/
        IF(
[{0}],
CALCULATE( {2}, DATESINPERIOD( {1}, LASTDATE( DATEADD( {1}, -1, MONTH ) ), -1, MONTH ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable, "\"Month ending \" & FORMAT(MAX( 'Date'[Date] ),\"d-MMM-yyyy\",\"en-US\")");
string MMTvsMMTminus1 = String.Format(
    @"/*MMT vs MMT-1*/
    VAR MMT = {0}
    VAR MMT_1 = {1}
    RETURN 
        IF( ISBLANK( MMT ) || ISBLANK( MMT_1 ), BLANK(), MMT - MMT_1 )",
    MMT, MMTminus1);
string MMTvsMMTminus1label = String.Format(
    @"/*MMT vs MMT-1*/
    VAR MMT = {0}
    VAR MMT_1 = {1}
    RETURN 
        IF( ISBLANK( MMT ) || ISBLANK( MMT_1 ), BLANK(), MMT & "" vs "" & MMT_1 )",
    MMTlabel, MMTminus1label);
string MMTvsMMTminus1pct = String.Format(
    @"/*MMT vs MMT-1(%)*/
    VAR MMT = {0}
    VAR MMT_1 = {1}
    RETURN 
        IF(
            ISBLANK( MMT ) || ISBLANK( MMT_1 ),
            BLANK(),
            DIVIDE( MMT - MMT_1, MMT_1 )
        )",
    MMT, MMTminus1);
string MMTvsMMTminus1pctlabel = String.Format(
    @"/*MMT vs MMT-1(%)*/
    VAR MMT = {0}
    VAR MMT_1 = {1}
    RETURN 
        IF( ISBLANK( MMT ) || ISBLANK( MMT_1 ), BLANK(), MMT & "" vs "" & MMT_1 & "" (%)"")",
    MMTlabel, MMTminus1label);
string MWT = String.Format(
        @"/*MWT*/
        IF(
[{0}],
CALCULATE( SELECTEDMEASURE( ), DATESINPERIOD( {1}, MAX( {1} ), -7, DAY ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable);
string MWTlabel = "/*MWT*/" +
    String.Format(
        @"/*MWT*/
        IF(
[{0}],
CALCULATE( {2}, DATESINPERIOD( {1}, MAX( {1} ), -7, DAY ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable, "\"Week ending \" & FORMAT(MAX( 'Date'[Date] ),\"d-MMM-yyyy\",\"en-US\")"); ;
string MWTminus1 = String.Format(
        @"/*MWT*/
        IF(
[{0}],
CALCULATE( SELECTEDMEASURE( ), DATESINPERIOD( {1}, LASTDATE( DATEADD( {1}, -7, DAY ) ), -7, DAY ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable);
string MWTminus1label = "/*MWT-1*/" +
    String.Format(
        @"/*MWT*/
        IF(
[{0}],
CALCULATE( {2}, DATESINPERIOD( {1}, LASTDATE( DATEADD( {1}, -7, DAY ) ), -7, DAY ) )
        )", ShowValueForDatesMeasureName, dateColumnWithTable, "\"Week ending \" & FORMAT(MAX( 'Date'[Date] ),\"d-MMM-yyyy\",\"en-US\")");
string MWTvsMWTminus1 = String.Format(
    @"/*MWT vs MWT-1*/
    VAR MWT = {0}
    VAR MWT_1 = {1}
    RETURN 
        IF( ISBLANK( MWT ) || ISBLANK( MWT_1 ), BLANK(), MWT - MWT_1 )",
    MWT, MWTminus1);
string MWTvsMWTminus1label = String.Format(
    @"/*MWT vs MWT-1*/
    VAR MWT = {0}
    VAR MWT_1 = {1}
    RETURN 
        IF( ISBLANK( MWT ) || ISBLANK( MWT_1 ), BLANK(), MWT & "" vs "" & MWT_1 )",
    MWTlabel, MWTminus1label);
string MWTvsMWTminus1pct = String.Format(
    @"/*MWT vs MWT-1(%)*/
    VAR MWT = {0}
    VAR MWT_1 = {1}
    RETURN
        IF(
            ISBLANK( MWT ) || ISBLANK( MWT_1 ),
            BLANK(),
            DIVIDE( MWT - MWT_1, MWT_1 )
        )",
    MWT, MWTminus1);
string MWTvsMWTminus1pctlabel = String.Format(
    @"/*MWT vs MWT-1 (%)*/
    VAR MWT = {0}
    VAR MWT_1 = {1}
    RETURN 
        IF( ISBLANK( MWT ) || ISBLANK( MWT_1 ), BLANK(), MWT & "" vs "" & MWT_1 & "" (%)"")",
    MWTlabel, MWTminus1label);
string defFormatString = "SELECTEDMEASUREFORMATSTRING()";
//if the flag expression is already present in the format string, do not change it, otherwise apply % format. 
string pctFormatString = String.Format(
    @"IF(
        FIND( {0}, SELECTEDMEASUREFORMATSTRING(), 1, -1 ) <> -1,
        SELECTEDMEASUREFORMATSTRING(),
        ""#,##0.# %""
    )",
    flagExpression);
//the order in the array also determines the ordinal position of the item    
string[,] calcItems =
    {
        {"CY",      CY,         defFormatString,    "Current year",             CYlabel},
        {"PY",      PY,         defFormatString,    "Previous year",            PYlabel},
        {"YOY",     YOY,        defFormatString,    "Year-over-year",           YOYlabel},
        {"YOY%",    YOYpct,     pctFormatString,    "Year-over-year%",          YOYpctLabel},
        {"YTD",     YTD,        defFormatString,    "Year-to-date",             YTDlabel},
        {"PYTD",    PYTD,       defFormatString,    "Previous year-to-date",    PYTDlabel},
        {"YOYTD",   YOYTD,      defFormatString,    "Year-over-year-to-date",   YOYTDlabel},
        {"YOYTD%",  YOYTDpct,   pctFormatString,    "Year-over-year-to-date%",  YOYTDpctLabel},
        {"CM",      CM,         defFormatString,    "Current month",             CMlabel},
        {"PM",      PM,         defFormatString,    "Previous month",            PMlabel},
        {"MOM",     MOM,        defFormatString,    "Month-over-month",          MOMlabel},
        {"MOM%",    MOMpct,     pctFormatString,    "Month-over-month%",         MOMpctLabel},
        {"MTD",     MTD,        defFormatString,    "Month-to-date",             MTDlabel},
        {"PMTD",    PMTD,       defFormatString,    "Previous month-to-date",    PMTDlabel},
        {"MOMTD",   MOMTD,      defFormatString,    "Month-over-month-to-date",  MOMTDlabel},
        {"MOMTD%",  MOMTDpct,   pctFormatString,    "Month-over-month-to-date%", MOMTDpctlabel},
        {"MAT",     MAT,        defFormatString,    "Moving Anual Total",       MATlabel},
        {"MAT-1",   MATminus1,  defFormatString,    "Moving Anual Total -1 year", MATminus1label},
        {"MAT vs MAT-1", MATvsMATminus1, defFormatString, "Moving Anual Total vs Moving Anual Total -1 year", MATvsMATminus1label},
        {"MAT vs MAT-1(%)", MATvsMATminus1pct, pctFormatString, "Moving Anual Total vs Moving Anual Total -1 year (%)", MATvsMATminus1pctlabel},
        {"MMT",     MMT,        defFormatString,    "Moving Monthly Total",       MMTlabel},
        {"MMT-1",   MMTminus1,  defFormatString,    "Moving Monthly Total -1 month", MMTminus1label},
        {"MMT vs MMT-1", MMTvsMMTminus1, defFormatString, "Moving Monthly Total vs Moving Monthly Total -1 month", MMTvsMMTminus1label},
        {"MMT vs MMT-1(%)", MMTvsMMTminus1pct, pctFormatString, "Moving Monthly Total vs Moving Monthly Total -1 month (%)", MMTvsMMTminus1pctlabel},
        {"MWT",     MWT,        defFormatString,    "Moving Weekly Total",       MWTlabel},
        {"MWT-1",   MWTminus1,  defFormatString,    "Moving Weekly Total -1 week", MWTminus1label},
        {"MWT vs MWT-1", MWTvsMWTminus1, defFormatString, "Moving Weekly Total vs Moving Weekly Total -1 month", MWTvsMWTminus1label},
        {"MWT vs MWT-1(%)", MWTvsMWTminus1pct, pctFormatString, "Moving Weekly Total vs Moving Weekly Total -1 week (%)", MWTvsMWTminus1pctlabel}
    };
int j = 0;
//create calculation items for each calculation with formatstring and description
foreach (var cg in Model.CalculationGroups)
{
    if (cg.Name == calcGroupName)
    {
        for (j = 0; j < calcItems.GetLength(0); j++)
        {
            string itemName = calcItems[j, 0];
            string itemExpression = calcItemProtection.Replace("<CODE>", calcItems[j, 1]);
            itemExpression = itemExpression.Replace("<LABELCODE>", calcItems[j, 4]);
            string itemFormatExpression = calcItemFormatProtection.Replace("<CODE>", calcItems[j, 2]);
            itemFormatExpression = itemFormatExpression.Replace("<LABELCODEFORMATSTRING>", "\"\"\"\" & " + calcItems[j, 4] + " & \"\"\"\"");
            //if(calcItems[j,2] != defFormatString) {
            //    itemFormatExpression = calcItemFormatProtection.Replace("<CODE>",calcItems[j,2]);
            //};
            string itemDescription = calcItems[j, 3];
            if (!cg.CalculationItems.Contains(itemName))
            {
                var nCalcItem = cg.AddCalculationItem(itemName, itemExpression);
                nCalcItem.FormatStringExpression = itemFormatExpression;
                nCalcItem.FormatDax();
                nCalcItem.Ordinal = j;
                nCalcItem.Description = itemDescription;
            };
        };
    };
};
