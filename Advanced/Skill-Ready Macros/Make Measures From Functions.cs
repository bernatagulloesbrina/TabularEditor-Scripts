#r "System.Drawing"
#r "Microsoft.VisualBasic"
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;

// MACRO: Name: Make Measures From Functions
// MACRO: ValidContexts: Function
// MACRO: Tooltip: Creates measures (or calculated columns) by calling the selected DAX user-defined functions with every combination of the chosen objects
// SKILL: name: make-measures-from-functions
// SKILL: description: Create measures (or calculated columns) from DAX user-defined functions (UDFs), calling each function with every combination of the given measures, columns, tables, calendars or values. Use when the user asks to create measures from functions or UDFs, apply a UDF to several measures, generate time intelligence / variance measures with a function, or bulk-create measures with a function.
// SKILL:
// SKILL: For each function, builds one call per combination of the values given for its parameters (e.g. 3 measures x 2 calendars = 6 measures) and creates a measure, or a calculated column, with expression Function(arg1, arg2, ...). Name, format string, display folder, destination table, Properties / PropertyNames annotations and data category come from templates stored as annotations on the function (nameTemplate, formatString, displayFolder, outputDestination, outputProperties, outputPropertyNames, dataCategory, outputType). Placeholders such as <param>Name, <param>FormatStringFull, <param>FormatStringRoot, <param>DisplayFolder, <param>Table, <param>Properties, <param>PropertyNames and <param>DataCategory are replaced with the properties of each object passed. Measures get their DAX formatted and a description pointing to the function.
// SKILL:
// SKILL: ## Inputs (environment variables)
// SKILL: Set these variables before running the macro. Work out each value from the conversation and the model (inspect the model if needed); if a value is unclear, ask the user instead of guessing.
// SKILL: - MF_FUNCTIONS (required): function names separated by ';', e.g. Local.TimeIntel.YTD. The functions to call.
// SKILL: - MF_PARAM_<PARAM> (required, one per distinct parameter name across the functions; <PARAM> is the parameter name in upper case, e.g. MF_PARAM_BASEMEASURE for baseMeasure): the values to iterate for that parameter, separated by ';'. Format by kind: measures as [Measure] or 'Table'[Measure]; columns as 'Table'[Column]; tables as 'Table'; calendars as the calendar name; Scalar as literal DAX values (e.g. 0.5); Text as plain text (it is quoted automatically). Leave it unset or empty to omit the argument and use the parameter's default value, which is only allowed when the parameter has one (param: TYPE = default).
// SKILL: - MF_PARAM_<PARAM>_TYPE (required only when the kind can't be inferred): Table, Column, Measure, Calendar, Scalar or Text. The kind is inferred from the parameter type (measureRef, columnRef, tableRef, calendarRef) or from its name (contains MEASURE, COLUMN, TABLE, TEXT or STRING, or ends with CALENDAR); set this variable only for parameters where neither applies (e.g. a numeric anyVal parameter named threshold -> Scalar).
// SKILL: - MF_DESTINATION_TABLE (optional, needed only in this case): 'Table' or table name. The table that receives the measures or columns whose destination template doesn't resolve to a model table, e.g. when the destination is "Custom". Leave unset otherwise.
// SKILL: - Output settings, one set per function (optional): <FN> is the function name in upper case with every character that is not a letter or digit replaced by '_', e.g. LOCAL_TIMEINTEL_YTD for Local.TimeIntel.YTD. When a variable is unset, the function's annotation is used, or the default below if the function has no such annotation. The values used are saved as annotations on the function, so the next run reuses them.
// SKILL:   - MF_OUTPUT_TYPE_<FN> (annotation outputType, default "Measure"): Measure, Column, Table or None. Only Measure and Column are implemented.
// SKILL:   - MF_NAME_TEMPLATE_<FN> (annotation nameTemplate): name of each object. Default "<p1>Name <p2>Name ... <short>", or the function name if it has no parameters.
// SKILL:   - MF_FORMAT_STRING_<FN> (annotation formatString): format string. Default "+0.0%;-0.0%;-" if the function name contains Pct, otherwise "<p1>FormatStringFull" (format of the first argument); empty if no parameters.
// SKILL:   - MF_DISPLAY_FOLDER_<FN> (annotation displayFolder): display folder. Default "<p1>DisplayFolder/<p1>Name <group>"; empty if no parameters.
// SKILL:   - MF_DESTINATION_<FN> (annotation outputDestination, Measure and Column only): destination table. Default "<p1>" if the first parameter name contains TABLE, "<p1>Table" (table of the first argument) if it contains MEASURE or COLUMN, otherwise "Custom" (then MF_DESTINATION_TABLE is needed). A literal 'Table' reference also works.
// SKILL:   - MF_PROPERTIES_<FN> (annotation outputProperties): value of the Properties annotation of each object. Default "<p1>Name|<p2>Name|...|<short>", or "<short>" if no parameters.
// SKILL:   - MF_PROPERTY_NAMES_<FN> (annotation outputPropertyNames): value of the PropertyNames annotation. Default "<p1>|<p2>|...|<group>", or "<short>" if no parameters.
// SKILL:   - MF_DATA_CATEGORY_<FN> (annotation dataCategory, Measure only, default "Uncategorized"): Uncategorized, Address, Place, City, County, StateOrProvince, PostalCode, Country, Continent, Latitude, Longitude, WebUrl, ImageUrl, Barcode, or <p>DataCategory for a measure or column parameter (copies the category of the argument).
// SKILL:   In the defaults, <p1>, <p2>... are the parameter names in order, <short> is the part of the function name after the last '.', and <group> is the second dot-separated part of the function name (the whole name if it has no '.'). Templates can use these placeholders for any parameter <p>: <p>Name, <p>FormatStringFull, <p>FormatStringRoot (first section of the format, without '+'), <p>DisplayFolder, <p>Table, <p>Properties, <p>PropertyNames, <p>DataCategory.
// SKILL:
// SKILL: ## Before running: confirm the output settings with the user
// SKILL: Read each function's expression and its annotations. For every function, show the user the output settings that will be used (from the annotation if it exists, otherwise the default above), together with an example of the resulting measure name, format and folder for one combination of arguments. Ask the user to confirm them or say what to change, and set the matching MF_..._<FN> variables only for what they change. Do not run the macro until the user has confirmed.
// SKILL:
// SKILL: ## Notes
// SKILL: The MF_PARAM_... and MF_..._<FN> variable names depend on the functions, so the runner warns that they are "not an input" of this skill; that warning is expected and the variables are still passed to the macro. Check that the generated names don't clash with existing measures: measure names must be unique in the model. The number of objects created is the product of the value counts of all parameters, so tell the user the total before running. Values are split on ';', so a single text value can't contain ';'.
// '2026-11-03 / B.Agullo / CLI mode is also detected at runtime, so the macro runs unattended in the TE2 command line (-S)
// '2026-11-01 / B.Agullo / Agent ready: in CLI mode functions, parameter values, destination table and output settings come from environment variables
// '2026-11-01 / B.Agullo / Measure and Uncategorized are preselected when asking for the output type and data category
//2026-10-15 / B.Agullo / added support for dataCategory annotation (measures only)
//2026-03-29 / B.Agullo / added support for propertyNames annotation
//2026-02-15 / B.Agullo / improved default function annotation values and added support for output properties annotation.
//2025-09-26/B.Agullo/ fixed bug that would not store annotations if initialized during runtime
//2025-09-16/B.Agullo/
//Creates measures based on DAX UDFs 
//Check the blog post for futher information: https://www.esbrina-ba.com/automatically-create-measures-with-dax-user-defined-functions/
#if TE3
ScriptHelper.WaitFormVisible = false;
#endif
// PSEUDOCODE / PLAN:
// 1. Verify that the user has selected one or more functions (Selected.Functions).
// 2. If none selected, show error and abort.
// 3. Create FunctionExtended objects for each selected function and keep them in a list.
// 4. Extract all parameters from the selected functions and build a distinct list by name.
// 5. For each distinct parameter name:
//      - Prompt the user once with Fx.SelectAnyObjects to choose the objects to iterate for that parameter.
//      - If the user cancels or selects nothing, abort the whole operation.
//      - Store the resulting IList<string> in a dictionary keyed by parameter name so it can be retrieved later.
// 6. Example usage: create a sample FunctionExtended (as the original example does),
//    then when iterating the function parameters use the previously built dictionary to obtain the list
//    of objects for each parameter name (do not prompt again).
// 7. Build measure names/expressions by iterating the parameter-object combinations and create measures.
//
// NOTES:
// - All Fx.SelectAnyObjects calls must use parameter names on the call.
// - The dictionary is Dictionary<string, IList<string>> parameterObjectsMap.
// - Abort execution if any required selection is cancelled.
#if TECLI
bool cliMode = true;
#else
bool cliMode = Fx.IsCliMode();
#endif
List<Function> functions = null;
if (cliMode)
{
    functions = Fx.GetFunctionsFromEnv(Model, "MF_FUNCTIONS");
    if (functions == null) return;
}
#if !TECLI
else
{
    // Validate selection
    if (Selected.Functions.Count() == 0)
    {
        Error("Select one or more functions and try again.");
        return;
    }
    functions = Selected.Functions.ToList();
}
#endif
// Create FunctionExtended objects for each selected function and store them for later iteration
IList<FunctionExtended> selectedFunctions = new List<FunctionExtended>();
foreach (var f in functions)
{
    // Create the FunctionExtended and add to list
    FunctionExtended fe = null;
    if (cliMode)
    {
        // output settings come from MF_<KEY>_<FUNCTION> variables
        string outputSettingsPrefix = "MF";
        fe = FunctionExtended.CreateFunctionExtendedFromEnv(f, outputSettingsPrefix);
    }
#if !TECLI
    else
    {
        fe = FunctionExtended.CreateFunctionExtended(f);
    }
#endif
    if (fe == null) return;
    selectedFunctions.Add(fe);
}
// Flatten all parameters from selected functions
var allParametersFlat = selectedFunctions
    .SelectMany(sf => sf.Parameters ?? new List<FunctionParameter>())
    .ToList();
// Build distinct FunctionParameter objects (first occurrence per name)
IList<FunctionParameter> distinctParameters = allParametersFlat
    .GroupBy(p => p.Name)
    .Select(g => g.First())
    .ToList();
// For each distinct parameter, ask the user once which objects should be iterated and store mapping
var parameterObjectsMap = new Dictionary<string, (IList<string> Values, string Type)>();
foreach (var param in distinctParameters)
{
    string selectionType = null;
    if (param.Type == "CALENDARREF" || param.Name.ToUpper().EndsWith("CALENDAR"))
    {
        selectionType = "Calendar";
    }
    else if (param.Type == "MEASUREREF" || param.Name.ToUpper().Contains("MEASURE"))
    {
        selectionType = "Measure";
    }
    else if (param.Type == "COLUMNREF" || param.Name.ToUpper().Contains("COLUMN"))
    {
        selectionType = "Column";
    }
    else if (param.Type == "TABLEREF" || param.Name.ToUpper().Contains("TABLE"))
    {
        selectionType = "Table";
    }
    else if (param.Name.ToUpper().Contains("TEXT") || param.Name.ToUpper().Contains("STRING"))
    {
        selectionType = "Text";
    }
    (IList<string> Values, string Type) selectedObjectsForParam = (null, null);
    if (cliMode)
    {
        string paramEnvVar = "MF_PARAM_" + param.Name.ToUpper();
        if (selectionType == null)
        {
            IList<string> selectionTypeOptions = new List<string> { "Table", "Column", "Measure", "Calendar", "Scalar", "Text" };
            selectionType = Fx.GetStringFromEnv(paramEnvVar + "_TYPE");
            if (selectionType == null) return;
            if (!selectionTypeOptions.Contains(selectionType))
            {
                Error(String.Format("Environment variable {0}_TYPE must be one of {1}, not '{2}'.", paramEnvVar, String.Join(", ", selectionTypeOptions), selectionType));
                return;
            }
        }
        // Unset or empty with a parameter default: the argument is left blank so the default is used
        string paramEnvValue = Fx.GetStringFromEnv(paramEnvVar, defaultValue: param.DefaultValue != null ? "" : null);
        if (paramEnvValue == null) return;
        IList<string> paramValues = null;
        if (paramEnvValue == "")
        {
            paramValues = new List<string>() { "" };
        }
        else if (selectionType == "Measure")
        {
            var measures = Fx.GetMeasuresFromEnv(Model, paramEnvVar);
            if (measures == null) return;
            paramValues = measures.Select(m => m.DaxObjectFullName).ToList();
        }
        else if (selectionType == "Column")
        {
            var columns = Fx.GetColumnsFromEnv(Model, paramEnvVar);
            if (columns == null) return;
            paramValues = columns.Select(c => c.DaxObjectFullName).ToList();
        }
        else if (selectionType == "Table")
        {
            var tables = Fx.GetTablesFromEnv(Model, paramEnvVar);
            if (tables == null) return;
            paramValues = tables.Select(t => t.DaxObjectFullName).ToList();
        }
        else if (selectionType == "Calendar")
        {
            var calendars = Fx.GetCalendarsFromEnv(Model, paramEnvVar);
            if (calendars == null) return;
            paramValues = calendars.Select(c => "'" + c.Name + "'").ToList();
        }
        else if (selectionType == "Text")
        {
            // Wrap in double quotes if not already a quoted DAX string literal
            paramValues = Fx.GetStringsFromEnv(paramEnvVar)
                .Select(t => t.StartsWith("\"") && t.EndsWith("\"") ? t : "\"" + t + "\"")
                .ToList();
        }
        else
        {
            paramValues = Fx.GetStringsFromEnv(paramEnvVar);
        }
        selectedObjectsForParam = (paramValues, selectionType);
    }
#if !TECLI
    else
    {
        selectedObjectsForParam = Fx.SelectAnyObjects(
            Model,
            selectionType: selectionType,
            prompt1: String.Format(@"Select object type for {0} parameter", param.Name),
            prompt2: String.Format(@"Select item for {0} parameter", param.Name),
            placeholderValue: param.Name,
            defaultValue: param.DefaultValue,
            description: param.Description
        );
    }
#endif
    if (selectedObjectsForParam.Values == null || selectedObjectsForParam.Type == null) return; //user cancelled
    parameterObjectsMap[param.Name] = selectedObjectsForParam;
}
int totalMeasuresCreated = 0;
int totalColumnsCreated = 0;
foreach (var func in selectedFunctions)
{
    string delimiter = "";
    IList<string> previousList = new List<string>() { func.Name + "(" };
    IList<string> currentList = new List<string>();
    IList<string> previousListNames = new List<string>() { func.OutputNameTemplate };
    IList<string> currentListNames = new List<string>();
    IList<string> previousDestinations = new List<string>() { func.OutputDestination };
    IList<string> currentDestinations = new List<string>();
    IList<string> previousDisplayFolders = new List<string>() { func.OutputDisplayFolder };
    IList<string> currentDisplayFolders = new List<string>();
    IList<string> previousFormatStrings = new List<string>() { func.OutputFormatString };
    IList<string> currentFormatStrings = new List<string>();
    IList<string> previousOutputProperties = new List<string>() { func.OutputProperties };
    IList<string> currentOutputProperties = new List<string>();
    IList<string> previousOutputPropertyNames = new List<string>() { func.OutputPropertyNames };
    IList<string> currentOutputPropertyNames = new List<string>();
    IList<string> previousDataCategories = new List<string>() { func.OutputDataCategory ?? "" };
    IList<string> currentDataCategories = new List<string>();
    // When iterating the parameters of this specific function, use the mapping created for distinct parameters.
    foreach (var param in func.Parameters)
    {
        currentList = new List<string>(); //reset current list
        currentListNames = new List<string>();
        currentFormatStrings = new List<string>();
        currentDestinations = new List<string>();
        currentDisplayFolders = new List<string>();
        currentOutputProperties = new List<string>();
        currentOutputPropertyNames = new List<string>();
        currentDataCategories = new List<string>();
        // Retrieve the objects list for this parameter name from the map (prompting was done earlier)
        (IList<string> Values, string Type) paramObject;
        if (!parameterObjectsMap.TryGetValue(param.Name, out paramObject) || paramObject.Type == null || paramObject.Values.Count == 0)
        {
            Error(String.Format("No objects were selected earlier for parameter '{0}'.", param.Name));
            return;
        }
        for (int i = 0; i < previousList.Count; i++)
        {
            string s = previousList[i];
            string sName = previousListNames[i];
            string sFormatString = previousFormatStrings[i];
            string sDisplayFolder = previousDisplayFolders[i];
            string sDestination = previousDestinations[i];
            string sOutputProperties = previousOutputProperties[i];
            string sOutputPropertyNames = previousOutputPropertyNames[i];
            string sDataCategory = previousDataCategories[i];
            foreach (var o in paramObject.Values)
            {
                //extract original name and format string if the parameter is a measure
                string paramName = o;
                string paramFormatStringFull = "";
                string paramFormatStringRoot = "";
                string paramDisplayFolder = "";
                string paramTable = "";
                string paramProperties = "";
                string paramPropertyNames = "";
                string paramDataCategory = "";
                //prepare placeholder
                string paramNamePlaceholder = param.Name + "Name";
                string paramFormatStringRootPlaceholder = param.Name + "FormatStringRoot";
                string paramFormatStringFullPlaceholder = param.Name + "FormatStringFull";
                string paramDisplayFolderPlaceholder = param.Name + "DisplayFolder";
                string paramPropertiesPlaceholder = param.Name + "Properties";
                string paramPropertyNamesPlaceholder = param.Name + "PropertyNames";
                string paramDataCategoryPlaceholder = param.Name + "DataCategory";
                string paramTablePlaceholder = "";
                // Empty value means the parameter default is used: keep the comma (blank arg)
                bool useDefault = string.IsNullOrEmpty(o);
                if (useDefault)
                {
                    paramName = "";
                    paramTablePlaceholder = param.Name;
                }
                else if (paramObject.Type == "Measure")
                {
                    Measure m = Model.AllMeasures.FirstOrDefault(m => m.DaxObjectFullName == o);
                    paramName = m.Name;
                    paramFormatStringFull = m.FormatString;
                    paramDisplayFolder = m.DisplayFolder;
                    paramTable = m.Table.DaxObjectFullName;
                    paramProperties = m.GetAnnotation("Properties") ?? "";
                    paramPropertyNames = m.GetAnnotation("PropertyNames") ?? "";
                    paramDataCategory = m.DataCategory ?? "";
                    paramTablePlaceholder = param.Name + "Table";
                }
                else if (paramObject.Type == "Column")
                {
                    Column c = Model.AllColumns.FirstOrDefault(c => c.DaxObjectFullName == o);
                    paramName = c.Name;
                    paramFormatStringFull = c.FormatString;
                    paramDisplayFolder = c.DisplayFolder;
                    paramTable = c.Table.DaxObjectFullName;
                    paramProperties = c.GetAnnotation("Properties") ?? "";
                    paramPropertyNames = c.GetAnnotation("PropertyNames") ?? "";
                    paramDataCategory = c.DataCategory ?? "";
                    paramTablePlaceholder = param.Name + "Table";
                }
                else if (paramObject.Type == "Table")
                {
                    Table t = Model.Tables.FirstOrDefault(t => t.DaxObjectFullName == o);
                    paramName = t.Name;
                    paramFormatStringFull = "";
                    paramDisplayFolder = "";
                    paramTable = t.DaxObjectFullName;
                    paramProperties = t.GetAnnotation("Properties") ?? "";
                    paramPropertyNames = t.GetAnnotation("PropertyNames") ?? "";
                    paramTablePlaceholder = param.Name;
                }
                else if (paramObject.Type == "Calendar")
                {
                    // o is the quoted calendar reference ('Name'); strip quotes for the display name
                    paramName = o.Trim('\'');
                    paramTablePlaceholder = param.Name;
                }
                else
                {
                    // Scalar / Text: use the value itself (strip quotes) as the display name
                    paramName = o.Trim('"');
                    paramTablePlaceholder = param.Name;
                }
                if (paramFormatStringFull.Contains(";"))
                {
                    //keep the first part of the format string, strip it of any + sign
                    paramFormatStringRoot = paramFormatStringFull.Split(';')[0].Replace("+", "");
                }
                else
                {
                    paramFormatStringRoot = paramFormatStringFull;
                }
                currentList.Add(s + delimiter + o);
                currentListNames.Add(sName.Replace(paramNamePlaceholder, paramName));
                currentFormatStrings.Add(
                    sFormatString
                        .Replace(paramFormatStringFullPlaceholder, paramFormatStringFull)
                        .Replace(paramFormatStringRootPlaceholder, paramFormatStringRoot));
                currentDisplayFolders.Add(
                    sDisplayFolder
                        .Replace(paramNamePlaceholder, paramName)
                        .Replace(paramDisplayFolderPlaceholder, paramDisplayFolder));
                currentDestinations.Add(
                    sDestination.Replace(paramTablePlaceholder, paramTable));
                currentOutputProperties.Add(
                    sOutputProperties
                        .Replace(paramNamePlaceholder, paramName)
                        .Replace(paramPropertiesPlaceholder, paramProperties));
                currentOutputPropertyNames.Add(
                    sOutputPropertyNames
                        .Replace(paramPropertyNamesPlaceholder, paramPropertyNames));
                currentDataCategories.Add(
                    sDataCategory.Replace(paramDataCategoryPlaceholder, paramDataCategory));
            }
        }
        delimiter = ", ";
        previousList = currentList;
        previousListNames = currentListNames;
        previousDestinations = currentDestinations;
        previousDisplayFolders = currentDisplayFolders;
        previousFormatStrings = currentFormatStrings;
        previousOutputProperties = currentOutputProperties;
        previousOutputPropertyNames = currentOutputPropertyNames;
        previousDataCategories = currentDataCategories;
    }
    IList<Table> currentDestinationTables = new List<Table>();
    if (func.OutputType == "Measure" || func.OutputType == "Column")
    {
        for (int i = 0; i < currentDestinations.Count; i++)
        {
            //transform to actual tables, initialize if necessary
            Table destinationTable = Model.Tables.Where(
                t => t.DaxObjectFullName == currentDestinations[i])
                .FirstOrDefault();
            if (destinationTable == null)
            {
                if (cliMode)
                {
                    destinationTable = Fx.GetTableFromEnv(Model, "MF_DESTINATION_TABLE");
                }
#if !TECLI
                else
                {
                    destinationTable = SelectTable(preselect: null, label: $"Select destinatoin table for {func.OutputType} {currentListNames[i]}");
                }
#endif
                if (destinationTable == null) return;
            }
            currentDestinationTables.Add(destinationTable);
        }
    }
    if (func.OutputType == "Measure")
    {
        for (int i = 0; i < currentList.Count; i++)
        {
            //It normalizes a folder/display-folder string by collapsing repeated slashes, removing leading/trailing backslashes and trimming whitespace.
            string cleanCurrentDisplayFolder = Regex.Replace(currentDisplayFolders[i], @"[/]+", @"").Trim('\\').Trim();
            Measure measure = currentDestinationTables[i].AddMeasure(currentListNames[i], currentList[i] + ")");
            measure.FormatDax();
            measure.Description = String.Format("Measure created with {0} function. Check function for details.", func.Name);
            measure.DisplayFolder = cleanCurrentDisplayFolder;
            measure.FormatString = currentFormatStrings[i];
            // Set Properties annotation with the transformed outputProperties value
            if (!string.IsNullOrEmpty(currentOutputProperties[i]))
            {
                measure.SetAnnotation("Properties", currentOutputProperties[i]);
            }
            // Set PropertyNames annotation with the transformed outputPropertyNames value
            if (!string.IsNullOrEmpty(currentOutputPropertyNames[i]))
            {
                measure.SetAnnotation("PropertyNames", currentOutputPropertyNames[i]);
            }
            // Set data category unless Uncategorized (empty data category in the model)
            if (!string.IsNullOrEmpty(currentDataCategories[i]) && currentDataCategories[i] != "Uncategorized")
            {
                measure.DataCategory = currentDataCategories[i];
            }
            totalMeasuresCreated++;
        }
    }
    else if (func.OutputType == "Column")
    {
        for (int i = 0; i < currentList.Count; i++)
        {
            //It normalizes a folder/display-folder string by collapsing repeated slashes, removing leading/trailing backslashes and trimming whitespace.
            string cleanCurrentDisplayFolder = Regex.Replace(currentDisplayFolders[i], @"[/]+", @"").Trim('\\').Trim();
            Column column = currentDestinationTables[i].AddCalculatedColumn(currentListNames[i], currentList[i] + ")");
            //column.FormatDax();
            column.Description = String.Format("Column created with {0} function. Check function for details.", func.Name);
            column.DisplayFolder = cleanCurrentDisplayFolder;
            column.FormatString = currentFormatStrings[i];
            // Set Properties annotation with the transformed outputProperties value
            if (!string.IsNullOrEmpty(currentOutputProperties[i]))
            {
                column.SetAnnotation("Properties", currentOutputProperties[i]);
            }
            // Set PropertyNames annotation with the transformed outputPropertyNames value
            if (!string.IsNullOrEmpty(currentOutputPropertyNames[i]))
            {
                column.SetAnnotation("PropertyNames", currentOutputPropertyNames[i]);
            }
        }
    }
    else
    {
        Info("Not implemented yet for output types other than Measure.");
    }
}
// Display summary message
if (totalMeasuresCreated > 0 || totalColumnsCreated > 0)
{
    string message = "";
    if (totalMeasuresCreated > 0)
    {
        message += String.Format("{0} measure{1} created", totalMeasuresCreated, totalMeasuresCreated == 1 ? "" : "s");
    }
    if (totalColumnsCreated > 0)
    {
        if (totalMeasuresCreated > 0) message += " and ";
        message += String.Format("{0} column{1} created", totalColumnsCreated, totalColumnsCreated == 1 ? "" : "s");
    }
    Info(message + ".");
}

public static class Fx
{
#if !TECLI
    // In TE2 scripts run inside an "Hourglass" that sets Application.UseWaitCursor = true for the
    // whole application, so dialog boxes show a busy cursor (TE3 solves this with WaitFormVisible).
    // Turn the flag off and nudge Windows (WM_SETCURSOR) to refresh the cursor immediately,
    // otherwise it only updates on the next mouse message. Call this before showing any dialog.
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);
    public static void ResetWaitCursor()
    {
        if (!Application.UseWaitCursor) return;
        Application.UseWaitCursor = false;
        Cursor.Current = Cursors.Default;
        Form f = Form.ActiveForm;
        if (f != null && f.Handle != IntPtr.Zero)
            SendMessage(f.Handle, 0x20, f.Handle, (IntPtr)1); // WM_SETCURSOR
    }
    public static string GetNameFromUser(string Prompt, string Title, string DefaultResponse)
    {
        ResetWaitCursor();
        string response = Interaction.InputBox(Prompt, Title, DefaultResponse, 740, 400);
        if (response == null)
        {
            Error("No input provided.");
            return null;
        };
        return response;
    }
#endif
    public const string DefaultSentinel = "<<USE_DEFAULT>>";
#if !TECLI
    public static (IList<string> Values, string Type) SelectAnyObjects(Model model, string selectionType = null, string prompt1 = "select item type", string prompt2 = "select item(s)", string placeholderValue = "", string defaultValue = null, string description = null)
    {
        var returnEmpty = (Values: new List<string>(), Type: (string)null);
        if (prompt1.Contains("{0}"))
            prompt1 = string.Format(prompt1, placeholderValue ?? "");
        if(prompt2.Contains("{0}"))
            prompt2 = string.Format(prompt2, placeholderValue ?? "");
        if (selectionType == null)
        {
            IList<string> selectionTypeOptions = new List<string> { "Table", "Column", "Measure", "Calendar", "Scalar", "Text" };
            selectionType = ChooseString(selectionTypeOptions, label: prompt1, customWidth: 600);
        }
        if (selectionType == null) return returnEmpty;
        IList<string> selectedValues = new List<string>();
        switch (selectionType)
        {
            case "Table":
                selectedValues = SelectTableMultiple(model, label: prompt2, defaultValue: defaultValue, description: description);
                break;
            case "Column":
                selectedValues = SelectColumnMultiple(model, label: prompt2, defaultValue: defaultValue, description: description);
                break;
            case "Measure":
                selectedValues = SelectMeasureMultiple(model: model, label: prompt2, defaultValue: defaultValue, description: description);
                break;
            case "Calendar":
                selectedValues = SelectCalendarMultiple(model, label: prompt2, defaultValue: defaultValue, description: description);
                break;
            case "Scalar":
                IList<string> scalarList = new List<string>();
                scalarList.Add(GetNameFromUser(prompt2, "Scalar value", defaultValue ?? "0"));
                selectedValues = scalarList;
                break;
            case "Text":
                IList<string> textList = new List<string>();
                string textValue = GetNameFromUser(prompt2, "Text value", defaultValue ?? "");
                if (textValue == null) return returnEmpty;
                // Wrap in double quotes if not already a quoted DAX string literal
                if (!(textValue.StartsWith("\"") && textValue.EndsWith("\"")))
                    textValue = "\"" + textValue + "\"";
                textList.Add(textValue);
                selectedValues = textList;
                break;
            default:
                Error("Invalid selection type");
                return returnEmpty;
        }
        if (selectedValues == null || selectedValues.Count == 0) return returnEmpty;
        // Default-value option selected: return empty value so the argument is left blank (default used)
        if (selectedValues.Count == 1 && selectedValues[0] == DefaultSentinel)
            return (Values: new List<string>() { "" }, Type: selectionType);
        return (Values:selectedValues, Type:selectionType);
    }
    public static string ChooseString(IList<string> OptionList, string label = "Choose item", int customWidth = 400, int customHeight = 500)
    {
        return ChooseStringInternal(OptionList, MultiSelect: false, label: label, customWidth: customWidth, customHeight:customHeight) as string;
    }
    public static List<string> ChooseStringMultiple(IList<string> OptionList, string label = "Choose item(s)", int customWidth = 650, int customHeight = 550, string defaultValue = null, string description = null)
    {
        return ChooseStringInternal(OptionList, MultiSelect:true, label:label, customWidth: customWidth, customHeight: customHeight, defaultValue: defaultValue, description: description) as List<string>;
    }
    private static object ChooseStringInternal(IList<string> OptionList, bool MultiSelect, string label = "Choose item(s)", int customWidth = 400, int customHeight = 500, string defaultValue = null, string description = null)
    {
        Form form = new Form
        {
            Text =label,
            StartPosition = FormStartPosition.CenterScreen,
            Padding = new Padding(20)
        };
        // Keep track of selected items across filtering
        HashSet<string> selectedItemsSet = new HashSet<string>();
        bool isRestoringSelections = false;
        // Search panel at the top
        Panel searchPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 35,
            Padding = new Padding(0, 0, 0, 5)  // Add 5px bottom padding for spacing
        };
        Label searchLabel = new System.Windows.Forms.Label
        {
            Text = "Search:",
            AutoSize = true,
            Location = new System.Drawing.Point(0, 8),
            Width = 70
        };
        TextBox searchBox = new TextBox
        {
            Location = new System.Drawing.Point(75, 5),
            Width = customWidth - 115,
            Anchor = AnchorStyles.Left | AnchorStyles.Right
        };
        searchPanel.Controls.Add(searchLabel);
        searchPanel.Controls.Add(searchBox);
        // Optional info panel: parameter description and "use default" checkbox
        CheckBox useDefaultCheck = null;
        Panel infoPanel = null;
        int infoPanelHeight = 0;
        if (!string.IsNullOrEmpty(description) || !string.IsNullOrEmpty(defaultValue))
        {
            infoPanel = new Panel { Dock = DockStyle.Top, Padding = new Padding(0, 0, 0, 5) };
            int infoY = 0;
            if (!string.IsNullOrEmpty(description))
            {
                Label descLabel = new System.Windows.Forms.Label
                {
                    Text = description,
                    Location = new System.Drawing.Point(0, infoY),
                    Width = customWidth - 60,
                    Height = 36,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
                };
                infoPanel.Controls.Add(descLabel);
                infoY += 40;
                infoPanelHeight += 40;
            }
            if (!string.IsNullOrEmpty(defaultValue))
            {
                useDefaultCheck = new CheckBox
                {
                    Text = String.Format("Use default ({0})", defaultValue),
                    Location = new System.Drawing.Point(0, infoY),
                    AutoSize = true
                };
                infoPanel.Controls.Add(useDefaultCheck);
                infoPanelHeight += 30;
            }
            infoPanel.Height = infoPanelHeight;
        }
        // ListBox panel in the middle
        Panel listBoxPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 35 + infoPanelHeight, 0, 75)  // Top padding = info + searchPanel height, Bottom = buttonPanel height
        };
        ListBox listbox = new ListBox
        {
            Dock = DockStyle.Fill,
            SelectionMode = MultiSelect ? SelectionMode.MultiExtended : SelectionMode.One
        };
        listBoxPanel.Controls.Add(listbox);
        // Initial population
        listbox.Items.AddRange(OptionList.ToArray());
        if (!MultiSelect && OptionList.Count > 0)
            listbox.SelectedItem = OptionList[0];
        // Track manual selection changes
        listbox.SelectedIndexChanged += delegate
        {
            // Skip if we're programmatically restoring selections
            if (isRestoringSelections) return;
            // Get current items in listbox
            HashSet<string> currentItems = new HashSet<string>();
            foreach (object item in listbox.Items)
            {
                currentItems.Add(item.ToString());
            }
            // Remove only currently visible items that are NOT selected
            foreach (string item in currentItems)
            {
                if (!listbox.SelectedItems.Cast<object>().Any(selected => selected.ToString() == item))
                {
                    selectedItemsSet.Remove(item);
                }
            }
            // Add currently selected items
            foreach (object item in listbox.SelectedItems)
            {
                selectedItemsSet.Add(item.ToString());
            }
        };
        // Search/filter functionality
        searchBox.TextChanged += delegate
        {
            string searchText = searchBox.Text;
            // Filter the list
            var filteredList = OptionList
                .Where(item => item.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();
            // Set flag to prevent SelectedIndexChanged from firing during restoration
            isRestoringSelections = true;
            // Repopulate listbox
            listbox.Items.Clear();
            listbox.Items.AddRange(filteredList.ToArray());
            // Restore previous selections
            for (int i = 0; i < listbox.Items.Count; i++)
            {
                if (selectedItemsSet.Contains(listbox.Items[i].ToString()))
                {
                    listbox.SetSelected(i, true);
                }
            }
            // Re-enable selection tracking
            isRestoringSelections = false;
        };
        FlowLayoutPanel buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 75,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(10, 5, 10, 10)  // Add top padding for spacing from listbox
        };
        Button selectAllButton = new Button { Text = "Select All", Visible = MultiSelect , Height = 50, Width = 150};
        Button selectNoneButton = new Button { Text = "Select None", Visible = MultiSelect, Height = 50, Width = 150 };
        Button okButton = new Button { Text = "OK", DialogResult = DialogResult.OK, Height = 50, Width = 100 };
        Button cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Height = 50, Width = 100 };
        selectAllButton.Click += delegate
        {
            for (int i = 0; i < listbox.Items.Count; i++)
                listbox.SetSelected(i, true);
            // Update tracking set with all currently visible items
            foreach (object item in listbox.Items)
            {
                selectedItemsSet.Add(item.ToString());
            }
        };
        selectNoneButton.Click += delegate
        {
            for (int i = 0; i < listbox.Items.Count; i++)
                listbox.SetSelected(i, false);
            // Remove all currently visible items from tracking set
            foreach (object item in listbox.Items)
            {
                selectedItemsSet.Remove(item.ToString());
            }
        };
        buttonPanel.Controls.Add(selectAllButton);
        buttonPanel.Controls.Add(selectNoneButton);
        buttonPanel.Controls.Add(okButton);
        buttonPanel.Controls.Add(cancelButton);
        // Add controls in proper order for Dock: Bottom first, Top second, Fill last
        form.Controls.Add(buttonPanel);    // Bottom - add first
        form.Controls.Add(searchPanel);    // Top - add second
        if (infoPanel != null) form.Controls.Add(infoPanel);  // Top - sits above search panel
        form.Controls.Add(listBoxPanel);   // Fill - add last
        // When "use default" is checked, disable the list/search to make clear the default will be used
        if (useDefaultCheck != null)
        {
            useDefaultCheck.CheckedChanged += delegate
            {
                bool useDefault = useDefaultCheck.Checked;
                listbox.Enabled = !useDefault;
                searchBox.Enabled = !useDefault;
                selectAllButton.Enabled = !useDefault;
                selectNoneButton.Enabled = !useDefault;
            };
        }
        form.Width = customWidth;
        form.Height = customHeight + infoPanelHeight;
        ResetWaitCursor();
        DialogResult result = form.ShowDialog();
        if (result == DialogResult.Cancel)
        {
            Info("You Cancelled!");
            return null;
        }
        // Default checkbox checked: signal default usage to the caller
        if (useDefaultCheck != null && useDefaultCheck.Checked)
        {
            return MultiSelect ? (object)new List<string>() { DefaultSentinel } : (object)DefaultSentinel;
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
    public static IList<string> SelectMeasureMultiple(Model model, IEnumerable<Measure> measures = null, string label = "Select Measure(s)", string defaultValue = null, string description = null)
    {
        measures ??= model.AllMeasures;
        // Create display strings with format: TableName\DisplayFolder\[MeasureName]
        var measureDisplayMap = new Dictionary<string, string>();
        var displayList = new List<string>();
        foreach (var measure in measures.OrderBy(m => m.Table.Name).ThenBy(m => m.DisplayFolder).ThenBy(m => m.Name))
        {
            string displayString;
            if (string.IsNullOrEmpty(measure.DisplayFolder))
            {
                displayString = String.Format("{0}\\[{1}]", measure.Table.Name, measure.Name);
            }
            else
            {
                displayString = String.Format("{0}\\{1}\\[{2}]", measure.Table.Name, measure.DisplayFolder, measure.Name);
            }
            measureDisplayMap[displayString] = measure.DaxObjectFullName;
            displayList.Add(displayString);
        }
        // Show the display list to user
        var selectedDisplayStrings = ChooseStringMultiple(displayList, label: label, defaultValue: defaultValue, description: description);
        if (selectedDisplayStrings == null || selectedDisplayStrings.Count == 0)
            return new List<string>();
        // Pass through the default sentinel unchanged
        if (selectedDisplayStrings.Count == 1 && selectedDisplayStrings[0] == DefaultSentinel)
            return selectedDisplayStrings;
        // Map back to DaxObjectFullName
        var selectedMeasureNames = selectedDisplayStrings
            .Select(display => measureDisplayMap[display])
            .ToList();
        return selectedMeasureNames;
    }
    public static IList<string> SelectColumnMultiple(Model model, IEnumerable<Column> columns = null, string label = "Select Columns(s)", string defaultValue = null, string description = null)
    {
        columns ??= model.AllColumns;
        IList<string> columnNames = columns.Select(m => m.DaxObjectFullName).ToList();
        IList<string> selectedColumnNames = ChooseStringMultiple(columnNames, label: label, defaultValue: defaultValue, description: description);
        return selectedColumnNames;
    }
    // True when TE2 runs the script from the command line (-S / -SCRIPT). TE2 doesn't define TECLI and shows
    // Select*/Choose* dialogs modally, which hangs an unattended run. Macros call it in the #else branch of #if TECLI
    // (te is always CLI mode), so this method has no #if of its own.
    public static bool IsCliMode()
    {
        return Environment.GetCommandLineArgs().Any(a =>
            a.Equals("-S", StringComparison.OrdinalIgnoreCase) || a.Equals("-SCRIPT", StringComparison.OrdinalIgnoreCase));
    }
#endif
    // ENV helpers: used in the #if TECLI branch of macros so agents can pass inputs through environment variables.
    // Lists are separated by ';'. All return null (after an Error) when the variable is missing or a value can't be resolved,
    // except when a defaultValue is given, which is used if the variable is missing or empty.
    // Raw string value, e.g. a name for a new measure or table
    public static string GetStringFromEnv(string envVar, string defaultValue = null)
    {
        string value = Environment.GetEnvironmentVariable(envVar);
        if (!String.IsNullOrWhiteSpace(value)) return value;
        if (defaultValue != null) return defaultValue;
        Error("Environment variable " + envVar + " is not set.");
        return null;
    }
    // Several string values separated by ';'
    public static List<string> GetStringsFromEnv(string envVar)
    {
        string value = GetStringFromEnv(envVar);
        if (value == null) return null;
        return value.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
    }
    // One of the given options (answer to a ChooseString). Only a value taken from the variable is checked, not the default
    public static string GetOptionFromEnv(string envVar, IList<string> options, string defaultValue = null)
    {
        if (String.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(envVar))) return GetStringFromEnv(envVar, defaultValue);
        string value = GetStringFromEnv(envVar);
        if (options.Contains(value)) return value;
        Error("Environment variable " + envVar + " must be one of " + String.Join(", ", options) + ", not '" + value + "'.");
        return null;
    }
    // Resolves each ';'-separated reference against the candidates using the given name forms
    private static List<T> GetObjectsFromEnv<T>(string envVar, IEnumerable<T> candidates, Func<T, string[]> names, string kind)
    {
        List<string> references = GetStringsFromEnv(envVar);
        if (references == null) return null;
        var result = new List<T>();
        foreach (string reference in references)
        {
            T obj = candidates.FirstOrDefault(c => names(c).Contains(reference));
            if (obj == null) { Error(kind + " " + reference + " (from " + envVar + ") not found."); return null; }
            result.Add(obj);
        }
        return result;
    }
    // 'Sales'[Amount] or Sales[Amount]
    public static List<Column> GetColumnsFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.AllColumns, c => new[] { c.DaxObjectFullName, c.Table.Name + "[" + c.Name + "]" }, "Column");
    }
    // [Sales Amount], Sales Amount or 'Sales'[Sales Amount]
    public static List<Measure> GetMeasuresFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.AllMeasures, m => new[] { m.DaxObjectFullName, m.Name, m.Table.DaxObjectFullName + m.DaxObjectFullName }, "Measure");
    }
    // 'Sales' or Sales
    public static List<Table> GetTablesFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.Tables, t => new[] { t.DaxObjectFullName, t.Name }, "Table");
    }
    // Function name, e.g. Local.TimeIntel.YTD
    public static List<Function> GetFunctionsFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.Functions, f => new[] { f.Name }, "Function");
    }
    // 'Gregorian' or Gregorian
    public static List<Calendar> GetCalendarsFromEnv(Model model, string envVar)
    {
        return GetObjectsFromEnv(envVar, model.AllCalendars, c => new[] { "'" + c.Name + "'", c.Name }, "Calendar");
    }
    public static Table GetTableFromEnv(Model model, string envVar) { return Single(GetTablesFromEnv(model, envVar), envVar); }
    private static T Single<T>(List<T> list, string envVar) where T : class
    {
        if (list == null) return null;
        if (list.Count != 1) { Error("Environment variable " + envVar + " must contain exactly one reference."); return null; }
        return list[0];
    }
#if !TECLI
    public static IList<string> SelectTableMultiple(Model model, IEnumerable<Table> Tables = null, string label = "Select Tables(s)", int customWidth = 400, string defaultValue = null, string description = null)
    {
        Tables ??= model.Tables;
        IList<string> TableNames = Tables.Select(m => m.DaxObjectFullName).ToList();
        IList<string> selectedTableNames = ChooseStringMultiple(TableNames, label: label, customWidth: customWidth, defaultValue: defaultValue, description: description);
        return selectedTableNames;
    }
    public static IList<string> SelectCalendarMultiple(Model model, string label = "Select Calendar(s)", string defaultValue = null, string description = null)
    {
        // Build a display->reference map: display is the calendar name, reference is quoted 'Name'
        var calendarMap = new Dictionary<string, string>();
        var displayList = new List<string>();
        foreach (var calendar in model.AllCalendars.OrderBy(c => c.Name))
        {
            string display = calendar.Name;
            calendarMap[display] = "'" + calendar.Name + "'";
            displayList.Add(display);
        }
        var selectedDisplay = ChooseStringMultiple(displayList, label: label, defaultValue: defaultValue, description: description);
        if (selectedDisplay == null || selectedDisplay.Count == 0)
            return new List<string>();
        // Pass through the default sentinel unchanged
        if (selectedDisplay.Count == 1 && selectedDisplay[0] == DefaultSentinel)
            return selectedDisplay;
        return selectedDisplay.Select(d => calendarMap[d]).ToList();
    }
#endif
}

    

    public class FunctionExtended
    {
        public string Name { get; set; }
        public string Expression { get; set; }
        public string Description { get; set; }
        public string OutputFormatString { get; set; }
        public string OutputNameTemplate { get; set; }
        public string OutputType { get; set; }
        public string OutputDisplayFolder { get; set; }

        public string OutputDestination { get; set; }
        public string OutputProperties { get; set; }
        public string OutputPropertyNames { get; set; }
        public string OutputDataCategory { get; set; }
        public Function OriginalFunction { get; set; }

        // Data categories accepted by Power BI for measures (TOM values). "Uncategorized" means no data category.
        public static readonly IList<string> DataCategories = new List<string>
        {
            "Uncategorized", "Address", "Place", "City", "County", "StateOrProvince", "PostalCode",
            "Country", "Continent", "Latitude", "Longitude", "WebUrl", "ImageUrl", "Barcode"
        };
        public List<FunctionParameter> Parameters { get; set; }
        private static List<FunctionParameter> ExtractParametersFromExpression(string expression)
        {
            // Find the first set of parentheses before the "=>"
            int arrowIndex = expression.IndexOf("=>");
            if (arrowIndex == -1)
                return new List<FunctionParameter>();

            int openParenIndex = expression.LastIndexOf('(', arrowIndex);
            int closeParenIndex = expression.IndexOf(')', openParenIndex);
            if (openParenIndex == -1 || closeParenIndex == -1 || closeParenIndex > arrowIndex)
                return new List<FunctionParameter>();

            string paramSection = expression.Substring(openParenIndex + 1, closeParenIndex - openParenIndex - 1);
            var paramList = new List<FunctionParameter>();

            // Split parameters preserving per-line layout so that an end-of-line comment can be
            // captured as the parameter description (only when there is one parameter per line).
            var paramStrings = paramSection.IndexOf('\n') >= 0
                ? paramSection.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                : paramSection.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var rawParam in paramStrings)
            {
                var param = rawParam.Trim().TrimEnd(',');
                string description = null;

                // Extract end-of-line comment (// ...) as the parameter description
                int commentIndex = param.IndexOf("//");
                if (commentIndex >= 0)
                {
                    description = param.Substring(commentIndex + 2).Trim();
                    param = param.Substring(0, commentIndex);
                }

                // Extract default value if present (param: TYPE = defaultExpression)
                string defaultValue = null;
                int equalsIndex = param.IndexOf('=');
                if (equalsIndex >= 0)
                {
                    defaultValue = param.Substring(equalsIndex + 1).Trim().TrimEnd(',').Trim();
                    param = param.Substring(0, equalsIndex);
                }

                var trimmed = param.Trim().TrimEnd(',').Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;
                var nameParams = trimmed.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                //var parts = trimmed.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                var fp = new FunctionParameter();
                fp.DefaultValue = defaultValue;
                fp.Description = description;
                fp.Name = nameParams.Length > 0 ? nameParams[0].Trim() : param.Trim();
                fp.Type =
                    (nameParams.Length == 1) ? "ANYVAL" :
                    (nameParams.Length > 1) ?
                        (nameParams[1].IndexOf("calendarRef", StringComparison.OrdinalIgnoreCase) >= 0 ? "CALENDARREF" :
                         nameParams[1].IndexOf("measureRef", StringComparison.OrdinalIgnoreCase) >= 0 ? "MEASUREREF" :
                         nameParams[1].IndexOf("columnRef", StringComparison.OrdinalIgnoreCase) >= 0 ? "COLUMNREF" :
                         nameParams[1].IndexOf("tableRef", StringComparison.OrdinalIgnoreCase) >= 0 ? "TABLEREF" :
                         nameParams[1].IndexOf("anyVal", StringComparison.OrdinalIgnoreCase) >= 0 ? "ANYVAL" :
                         nameParams[1].IndexOf("Scalar", StringComparison.OrdinalIgnoreCase) >= 0 ? "SCALAR" :
                         nameParams[1].IndexOf("Table", StringComparison.OrdinalIgnoreCase) >= 0 ? "TABLE" :
                         nameParams[1].IndexOf("AnyRef", StringComparison.OrdinalIgnoreCase) >= 0 ? "ANYREF" :
                         "ANYVAL")
                    : "ANYVAL";

                fp.Subtype =
                    (fp.Type == "SCALAR" && nameParams.Length > 1) ?
                        (
                            nameParams[1].IndexOf("variant", StringComparison.OrdinalIgnoreCase) >= 0 ? "VARIANT" :
                            nameParams[1].IndexOf("int64", StringComparison.OrdinalIgnoreCase) >= 0 ? "INT64" :
                            nameParams[1].IndexOf("decimal", StringComparison.OrdinalIgnoreCase) >= 0 ? "DECIMAL" :
                            nameParams[1].IndexOf("double", StringComparison.OrdinalIgnoreCase) >= 0 ? "DOUBLE" :
                            nameParams[1].IndexOf("string", StringComparison.OrdinalIgnoreCase) >= 0 ? "STRING" :
                            nameParams[1].IndexOf("datetime", StringComparison.OrdinalIgnoreCase) >= 0 ? "DATETIME" :
                            nameParams[1].IndexOf("boolean", StringComparison.OrdinalIgnoreCase) >= 0 ? "BOOLEAN" :
                            nameParams[1].IndexOf("numeric", StringComparison.OrdinalIgnoreCase) >= 0 ? "NUMERIC" :
                            null
                        )
                    : null;

                // ParameterMode: check for VAL or EXPR (any casing) in the parameter string
                string paramMode = null;
                if (trimmed.IndexOf("VAL", StringComparison.OrdinalIgnoreCase) >= 0)
                    paramMode = "VAL";
                else if (trimmed.IndexOf("EXPR", StringComparison.OrdinalIgnoreCase) >= 0)
                    paramMode = "EXPR";
                else if (fp.Type == "ANYREF" || fp.Type.EndsWith("REF"))
                {
                    paramMode = "EXPR";
                }else
                    paramMode = "VAL";
                fp.ParameterMode = paramMode;

                paramList.Add(fp);
            }

            return paramList;
        }

        

        // Default values for the output annotations of a function, derived from its name and parameters
        private static (string NameTemplate, string FormatString, string DisplayFolder, string Destination, string OutputProperties, string OutputPropertyNames) GetMetadataDefaults(Function function, List<FunctionParameter> Parameters)
        {
            string nameTemplateDefault = "";
            string formatStringDefault = "";
            string displayFolderDefault = "";
            string functionGroupName = function.Name;
            
            var nameParts = function.Name.Split('.');
            if (nameParts.Length > 1)
            {
                functionGroupName = nameParts[1];
            }
            
            string functionNameShort = function.Name;
            string destinationDefault = ""; 

            if(function.Name.IndexOf(".") > 0)
            {
                functionNameShort = function.Name.Substring(function.Name.LastIndexOf(".") + 1);
            }

            string outputPropertiesDefault = "";
            string outputPropertyNamesDefault = "";
            
            if (Parameters.Count == 0) {
                nameTemplateDefault = function.Name;
                formatStringDefault = "";
                displayFolderDefault = "";
                destinationDefault = "";
                outputPropertiesDefault = functionNameShort;
                outputPropertyNamesDefault = functionNameShort;
            }
            else
            {
                nameTemplateDefault = string.Join(" ", Parameters.Select(p => p.Name + "Name")) + " " + functionNameShort;
                
                if(function.Name.IndexOf("Pct", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    formatStringDefault = "+0.0%;-0.0%;-";
                }
                else
                {
                    formatStringDefault = Parameters[0].Name + "FormatStringFull";
                }


                    
                displayFolderDefault = 
                    String.Format(
                        @"{0}DisplayFolder/{1}Name {2}", 
                        Parameters[0].Name, 
                        Parameters[0].Name,
                        functionGroupName);
                
                if (Parameters[0].Name.ToUpper().Contains("TABLE"))
                {
                    destinationDefault = Parameters[0].Name;
                }
                else if (Parameters[0].Name.ToUpper().Contains("MEASURE") || Parameters[0].Name.ToUpper().Contains("COLUMN"))
                {
                    destinationDefault = Parameters[0].Name + "Table";
                }
                else
                {
                    destinationDefault = "Custom";
                }
                
                outputPropertiesDefault = string.Join("|", Parameters.Select(p => p.Name + "Name")) + "|" + functionNameShort;
                outputPropertyNamesDefault = string.Join("|", Parameters.Select(p => p.Name)) + "|" + functionGroupName;

            };

            return (nameTemplateDefault, formatStringDefault, displayFolderDefault, destinationDefault, outputPropertiesDefault, outputPropertyNamesDefault);
        }

        // Non-interactive version of CreateFunctionExtended, used by skill-ready macros in CLI mode. Each output setting is read from
        // the environment variable <envPrefix>_<KEY>_<FUNCTION> (function name in upper case, other characters as '_'). When it is
        // not set, the function's annotation is used, or the default value if there is none. The values used are stored as annotations.
        public static FunctionExtended CreateFunctionExtendedFromEnv(Function function, string envPrefix)
        {
            List<FunctionParameter> Parameters = ExtractParametersFromExpression(function.Expression);
            var defaults = GetMetadataDefaults(function, Parameters);
            string envSuffix = "_" + new string(function.Name.ToUpper().Select(ch => char.IsLetterOrDigit(ch) ? ch : '_').ToArray());

            string Current(string annotation, string defaultValue)
            {
                string value = function.GetAnnotation(annotation);
                return string.IsNullOrEmpty(value) ? defaultValue : value;
            }

            string myOutputType = Fx.GetOptionFromEnv(envPrefix + "_OUTPUT_TYPE" + envSuffix,
                new List<string> { "Table", "Column", "Measure", "None" }, Current("outputType", "Measure"));
            if (myOutputType == null) return null;
            function.SetAnnotation("outputType", myOutputType);

            string myNameTemplate = Fx.GetStringFromEnv(envPrefix + "_NAME_TEMPLATE" + envSuffix, Current("nameTemplate", defaults.NameTemplate));
            function.SetAnnotation("nameTemplate", myNameTemplate);

            string myFormatString = Fx.GetStringFromEnv(envPrefix + "_FORMAT_STRING" + envSuffix, Current("formatString", defaults.FormatString));
            function.SetAnnotation("formatString", myFormatString);

            string myDisplayFolder = Fx.GetStringFromEnv(envPrefix + "_DISPLAY_FOLDER" + envSuffix, Current("displayFolder", defaults.DisplayFolder));
            function.SetAnnotation("displayFolder", myDisplayFolder);

            string myOutputDestination = function.GetAnnotation("outputDestination");
            if (myOutputType == "Table")
            {
                if (string.IsNullOrEmpty(myOutputDestination)) myOutputDestination = "Model";
            }
            else if (myOutputType == "Column" || myOutputType == "Measure")
            {
                myOutputDestination = Fx.GetStringFromEnv(envPrefix + "_DESTINATION" + envSuffix, Current("outputDestination", defaults.Destination));
                function.SetAnnotation("outputDestination", myOutputDestination);
            }

            string myOutputProperties = Fx.GetStringFromEnv(envPrefix + "_PROPERTIES" + envSuffix, Current("outputProperties", defaults.OutputProperties));
            function.SetAnnotation("outputProperties", myOutputProperties);

            string myOutputPropertyNames = Fx.GetStringFromEnv(envPrefix + "_PROPERTY_NAMES" + envSuffix, Current("outputPropertyNames", defaults.OutputPropertyNames));
            function.SetAnnotation("outputPropertyNames", myOutputPropertyNames);

            string myDataCategory = function.GetAnnotation("dataCategory");
            if (myOutputType == "Measure")
            {
                // same options as in interactive mode: fixed categories plus one placeholder per parameter that may reference a measure or column
                IList<string> dataCategoryOptions = new List<string>(DataCategories);
                foreach (var p in Parameters.Where(p =>
                    p.Type == "MEASUREREF" || p.Type == "COLUMNREF" || p.Type == "ANYREF"
                    || p.Name.ToUpper().Contains("MEASURE") || p.Name.ToUpper().Contains("COLUMN")))
                {
                    dataCategoryOptions.Add(p.Name + "DataCategory");
                }

                myDataCategory = Fx.GetOptionFromEnv(envPrefix + "_DATA_CATEGORY" + envSuffix, dataCategoryOptions, Current("dataCategory", "Uncategorized"));
                if (myDataCategory == null) return null;
                function.SetAnnotation("dataCategory", myDataCategory);
            }

            return new FunctionExtended
            {
                Name = function.Name,
                Expression = function.Expression,
                Description = function.Description,
                Parameters = Parameters,
                OutputFormatString = myFormatString,
                OutputNameTemplate = myNameTemplate,
                OutputType = myOutputType,
                OutputDisplayFolder = myDisplayFolder,
                OutputDestination = myOutputDestination,
                OutputProperties = myOutputProperties,
                OutputPropertyNames = myOutputPropertyNames,
                OutputDataCategory = myDataCategory,
                OriginalFunction = function
            };
        }
#if !TECLI

        public static FunctionExtended CreateFunctionExtended(Function function, bool completeMetadata = true)
        {

            FunctionExtended emptyFunction = null as FunctionExtended;
            List<FunctionParameter> Parameters =  ExtractParametersFromExpression (function.Expression);

            var defaults = GetMetadataDefaults(function, Parameters);
            string nameTemplateDefault = defaults.NameTemplate;
            string formatStringDefault = defaults.FormatString;
            string displayFolderDefault = defaults.DisplayFolder;
            string destinationDefault = defaults.Destination;
            string outputPropertiesDefault = defaults.OutputProperties;
            string outputPropertyNamesDefault = defaults.OutputPropertyNames;

            string myOutputType = function.GetAnnotation("outputType"); 
            string myNameTemplate = function.GetAnnotation("nameTemplate");
            string myFormatString = function.GetAnnotation("formatString");
            string myDisplayFolder = function.GetAnnotation("displayFolder");
            string myOutputDestination = function.GetAnnotation("outputDestination");
            string myOutputProperties = function.GetAnnotation("outputProperties");
            string myOutputPropertyNames = function.GetAnnotation("outputPropertyNames");
            string myDataCategory = function.GetAnnotation("dataCategory");

            if (completeMetadata)
            {
                if (string.IsNullOrEmpty(myOutputType))
                {
                    // the first option is preselected, so Measure (the usual output) is the default
                    IList<string> selectionTypeOptions = new List<string> { "Measure", "Column", "Table", "None" };
                    myOutputType =
                        Fx.ChooseString(
                            OptionList: selectionTypeOptions,
                            label: "Choose output type for function" + function.Name,
                            customWidth: 600);
                    if (string.IsNullOrEmpty(myOutputType)) return emptyFunction;
                    function.SetAnnotation("outputType", myOutputType);
                }

                if (string.IsNullOrEmpty(myNameTemplate))
                {
                    myNameTemplate = Fx.GetNameFromUser(Prompt: "Enter output name template for function " + function.Name, "Name Template", nameTemplateDefault);
                    if (string.IsNullOrEmpty(myNameTemplate)) return emptyFunction;
                    function.SetAnnotation("nameTemplate", myNameTemplate);
                }
                if (string.IsNullOrEmpty(myFormatString))
                {
                    myFormatString = Fx.GetNameFromUser(Prompt: "Enter output format string for function " + function.Name, "Format String", formatStringDefault);
                    if (string.IsNullOrEmpty(myFormatString)) return emptyFunction;
                    function.SetAnnotation("formatString", myFormatString);

                }
                if (string.IsNullOrEmpty(myDisplayFolder))
                {
                    myDisplayFolder =
                        Fx.GetNameFromUser(
                            Prompt: "Enter output display folder for function " + function.Name,
                            Title: "Display Folder",
                            DefaultResponse: displayFolderDefault);

                    if (string.IsNullOrEmpty(myDisplayFolder)) return emptyFunction;
                    function.SetAnnotation("displayFolder", myDisplayFolder);
                }

                if (string.IsNullOrEmpty(myOutputDestination))
                {
                    if (myOutputType == "Table")
                    {
                        myOutputDestination = "Model";
                    }
                    else if (myOutputType == "Column" || myOutputType == "Measure")
                    {
                        myOutputDestination =
                            Fx.GetNameFromUser(
                                Prompt: "Enter Destination template for " + function.Name,
                                Title: "Destination",
                                DefaultResponse: destinationDefault);

                        if (string.IsNullOrEmpty(myOutputDestination)) return emptyFunction;
                        function.SetAnnotation("outputDestination", myOutputDestination);
                    }
                }

                if (string.IsNullOrEmpty(myOutputProperties))
                {
                    myOutputProperties =
                        Fx.GetNameFromUser(
                            Prompt: "Enter output properties template for " + function.Name,
                            Title: "Output Properties",
                            DefaultResponse: outputPropertiesDefault);

                    if (string.IsNullOrEmpty(myOutputProperties)) return emptyFunction;
                    function.SetAnnotation("outputProperties", myOutputProperties);
                }

                if (string.IsNullOrEmpty(myOutputPropertyNames))
                {
                    myOutputPropertyNames =
                        Fx.GetNameFromUser(
                            Prompt: "Enter output property names for " + function.Name,
                            Title: "Output Property Names",
                            DefaultResponse: outputPropertyNamesDefault);

                    if (string.IsNullOrEmpty(myOutputPropertyNames)) return emptyFunction;
                    function.SetAnnotation("outputPropertyNames", myOutputPropertyNames);
                }

                if (string.IsNullOrEmpty(myDataCategory) && myOutputType == "Measure")
                {
                    // fixed categories plus one placeholder per parameter that may reference a measure or column,
                    // which resolves to the data category of the selected object. Uncategorized comes first, so it is preselected
                    IList<string> dataCategoryOptions = new List<string>(DataCategories);
                    foreach (var p in Parameters.Where(p =>
                        p.Type == "MEASUREREF" || p.Type == "COLUMNREF" || p.Type == "ANYREF"
                        || p.Name.ToUpper().Contains("MEASURE") || p.Name.ToUpper().Contains("COLUMN")))
                    {
                        dataCategoryOptions.Add(p.Name + "DataCategory");
                    }

                    myDataCategory =
                        Fx.ChooseString(
                            OptionList: dataCategoryOptions,
                            label: "Choose output data category for function " + function.Name,
                            customWidth: 600);

                    if (string.IsNullOrEmpty(myDataCategory)) return emptyFunction;
                    function.SetAnnotation("dataCategory", myDataCategory);
                }
            }

            var functionExtended = new FunctionExtended
            {
                Name = function.Name,
                Expression = function.Expression,
                Description = function.Description,
                Parameters = Parameters,
                OutputFormatString = myFormatString,
                OutputNameTemplate = myNameTemplate,
                OutputType = myOutputType,
                OutputDisplayFolder = myDisplayFolder,
                OutputDestination = myOutputDestination,
                OutputProperties = myOutputProperties,
                OutputPropertyNames = myOutputPropertyNames,
                OutputDataCategory = myDataCategory,
                OriginalFunction = function

            };

            return functionExtended;
        }
#endif

        
    }


    public class FunctionParameter
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Subtype { get; set; }
        public string ParameterMode { get; set; }
        public string DefaultValue { get; set; }
        public string Description { get; set; }
    }
