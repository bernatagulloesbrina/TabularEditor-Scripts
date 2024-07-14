using Microsoft.VisualBasic;

using System.Windows.Forms;
using System.IO;
// '2024-07-13 / B.Agullo / 
// Requirements:
// execute after running the Referential integrity check measures script (version  2024-07-13 or later)
// https://github.com/bernatagulloesbrina/TabularEditor-Scripts/blob/main/Advanced/One-Click%20Macros/Referential%20Integrity%20Check%20Measures.csx
// See blog posta here: 
// https://www.esbrina-ba.com/easy-management-of-referential-integrity/
// https://www.esbrina-ba.com/building-a-referential-integrity-report-page-with-a-c-script/


string pageName = "Referential Integrity";
int interObjectGap = 15; 
int totalCardX = interObjectGap;
int totalCardY = interObjectGap;
int totalCardWidth = 300;
int totalCardHeight = 150;
int detailCardFontSize = 14;
int detailCardX = interObjectGap;
int detailCardY = totalCardY + totalCardHeight + interObjectGap;
int detailCardWidth = 300;
int detailCardHeight = 200;
int tableHorizontalGap = 10;
int tablesPerRow = 3;
int tableHeight = 250;
int tableWidth = 300;
int backgroundTransparency = 0;
string backgroundColor = "#F0ECEC"; 

// do not modify below this line -- 
string annLabel = "ReferencialIntegrityMeasures";
string annValueTotal = "TotalUnmappedItems";
string annValueDetail = "DataProblems";
string annValueDataQualityMeasures = "DataQualityMeasure";
string annValueDataQualityTitles = "DataQualityTitle";
string annValueDataQualitySubtitles = "DataQualitySubitle";
string annValueFactColumn = "FactColumn";

/*uncomment in TE3 to avoid wating cursor infront of dialogs*/

ScriptHelper.WaitFormVisible = false;

bool waitCursor = Application.UseWaitCursor;
Application.UseWaitCursor = false;

DialogResult dialogResult = MessageBox.Show(text:"Did you save changes in PBI Desktop before running this macro?", caption:"Saved changes?", buttons:MessageBoxButtons.YesNo);

if(dialogResult != DialogResult.Yes){
    Info("Please save your changes first and then run this macro"); 
    return; 
};



string[] annotationArray = new string[4] { annValueTotal, annValueDetail, annValueDataQualityMeasures, annValueDataQualityTitles };
foreach (string annotation in annotationArray)
{
    if (!Model.AllMeasures.Any(m => m.GetAnnotation(annLabel).StartsWith(annotation)))
    {
        Error(String.Format("No measure found with annotation {0} starting with {1} ", annLabel, annotation));
        return;
    }
}
Measure totalMeasure = 
    Model.AllMeasures
        .Where(m => m.GetAnnotation(annLabel) == annValueTotal)
        .FirstOrDefault();
Measure detailMeasure = 
    Model.AllMeasures
        .Where(m => m.GetAnnotation(annLabel) == annValueDetail)
        .FirstOrDefault();
IList<Measure> dataQualityMeasures = 
    Model.AllMeasures
        .Where(m => m.GetAnnotation(annLabel) != null
            && m.GetAnnotation(annLabel).StartsWith(annValueDataQualityMeasures))
        .OrderBy(m => m.GetAnnotation(annLabel))
        .ToList<Measure>();

IList<Measure> dataQualityTitles = 
    Model.AllMeasures
        .Where(m => m.GetAnnotation(annLabel) != null
            && m.GetAnnotation(annLabel).StartsWith(annValueDataQualityTitles))
        .OrderBy(m => m.GetAnnotation(annLabel))
        .ToList<Measure>();


IList<Measure> dataQualitySubtitles =
    Model.AllMeasures
        .Where(m => m.GetAnnotation(annLabel) != null
            && m.GetAnnotation(annLabel).StartsWith(annValueDataQualitySubtitles))
        .OrderBy(m => m.GetAnnotation(annLabel))
        .ToList<Measure>();

IList<Column> factTableColumns =
    Model.AllColumns
        .Where(c => c.GetAnnotation(annLabel) != null
            && c.GetAnnotation(annLabel).StartsWith(annValueFactColumn))
        .OrderBy(c => c.GetAnnotation(annLabel))
        .ToList<Column>();
//now that we now number of tables we'll need, let's set up the page size. 
int tableCount = dataQualityMeasures.Count();
decimal rowsRaw = (decimal) tableCount / (decimal) tablesPerRow;
int rowsOfTables = (int)Math.Ceiling(rowsRaw);
int pageWidth = totalCardX + totalCardWidth + interObjectGap + (tableWidth + interObjectGap) * tablesPerRow ;
int totalTablesHeight = interObjectGap + (tableHeight + interObjectGap) * rowsOfTables;
int totalCardsHeight = detailCardY + detailCardHeight + interObjectGap;
int pageHeight = Math.Max(totalTablesHeight,totalCardsHeight); 

//adjust detail card height to fill the height if tables are taller
detailCardHeight = pageHeight - 3 * interObjectGap - totalCardHeight;

// Create an instance of the OpenFileDialog
OpenFileDialog openFileDialog = new OpenFileDialog();
openFileDialog.Title = "Please select definition.pbir file of the target report";
// Set filter options and filter index.
openFileDialog.Filter = "PBIR Files (*.pbir)|*.pbir";
openFileDialog.FilterIndex = 1;
// Call the ShowDialog method to show the dialog box.
DialogResult result = openFileDialog.ShowDialog();
// Process input if the user clicked OK.
if (result != DialogResult.OK)
{
    Error("You cancelled");
    return;
}
// Get the file name.
string pbirFilePath = openFileDialog.FileName;
string pageContentsTemplate = @"
{
    ""$schema"": ""https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.0.0/schema.json"",
    ""name"": ""{{newPageId}}"",
    ""displayName"": ""Referential Integrity"",
    ""displayOption"": ""FitToPage"",
    ""height"": {{pageHeight}},
    ""width"": {{pageWidth}},
    ""objects"": {
        ""background"": [
          {
            ""properties"": {
              ""transparency"": {
                ""expr"": {
                  ""Literal"": {
                    ""Value"": ""{{backgroundTransparency}}D""
                  }
                }
              },
              ""color"": {
                ""solid"": {
                  ""color"": {
                    ""expr"": {
                      ""Literal"": {
                        ""Value"": ""'{{backgroundColor}}'""
                      }
                    }
                  }
                }
              }
            }
          }
        ]
      }
}";
string totalCardContentsTemplate = @"
    {
        ""$schema"": ""https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.0.0/schema.json"",
        ""name"": ""{{newVisualId}}"",
        ""position"": {
            ""x"": {{totalCardX}},
            ""y"": {{totalCardY}},
            ""z"": {{zTabOrder}},
            ""width"": {{totalCardWidth}},
            ""height"": {{totalCardHeight}},
            ""tabOrder"": {{zTabOrder}}
          },
      ""visual"": {
        ""visualType"": ""card"",
        ""query"": {
          ""queryState"": {
            ""Values"": {
              ""projections"": [
                {
                  ""field"": {
                    ""Measure"": {
                      ""Expression"": {
                        ""SourceRef"": {
                          ""Entity"": ""{{totalMeasureTable}}""
                        }
                      },
                      ""Property"": ""{{totalMeasureName}}""
                    }
                  },
                  ""queryRef"": ""{{totalMeasureTable}}.{{totalMeasureName}}"",
                  ""nativeQueryRef"": ""{{totalMeasureName}}""
                }
              ]
            }
          },
          ""sortDefinition"": {
            ""sort"": [
              {
                ""field"": {
                  ""Measure"": {
                    ""Expression"": {
                      ""SourceRef"": {
                        ""Entity"": ""{{totalMeasureTable}}""
                      }
                    },
                    ""Property"": ""{{totalMeasureName}}""
                  }
                },
                ""direction"": ""Descending""
              }
            ],
            ""isDefaultSort"": true
          }
        },
        ""drillFilterOtherVisuals"": true
      }
    }";
string detailCardContentsTemplate = @"
    {
      ""$schema"": ""https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.0.0/schema.json"",
      ""name"": ""{{newVisualId}}"",
      ""position"": {
        ""x"": {{detailCardX}},
        ""y"": {{detailCardY}},
        ""z"": {{zTabOrder}},
        ""width"": {{detailCardWidth}},
        ""height"": {{detailCardHeight}},
        ""tabOrder"": {{zTabOrder}}
      },
      ""visual"": {
        ""visualType"": ""card"",
        ""query"": {
          ""queryState"": {
            ""Values"": {
              ""projections"": [
                {
                  ""field"": {
                    ""Measure"": {
                      ""Expression"": {
                        ""SourceRef"": {
                          ""Entity"": ""{{dataProblemsMeasureTable}}""
                        }
                      },
                      ""Property"": ""{{dataProblemsMeasureName}}""
                    }
                  },
                  ""queryRef"": ""{{dataProblemsMeasureTable}}.{{dataProblemsMeasureName}}"",
                  ""nativeQueryRef"": ""{{dataProblemsMeasureName}}""
                }
              ]
            }
          }
        },
        ""objects"": {
          ""labels"": [
            {
              ""properties"": {
                ""fontSize"": {
                  ""expr"": {
                    ""Literal"": {
                      ""Value"": ""{{detailCardFontSize}}D""
                    }
                  }
                }
              }
            }
          ]
        },
        ""drillFilterOtherVisuals"": true
      }
    }";
string tableContentsTemplate = @"
    {
      ""$schema"": ""https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.0.0/schema.json"",
      ""name"": ""{{newVisualId}}"",
      ""position"": {
        ""x"": {{tableX}},
        ""y"": {{tableY}},
        ""z"": {{zTabOrder}},
        ""width"": {{tableWidth}},
        ""height"": {{tableHeight}},
        ""tabOrder"": {{zTabOrder}}
      },
      ""visual"": {
        ""visualType"": ""tableEx"",
        ""query"": {
          ""queryState"": {
            ""Values"": {
              ""projections"": [
                {
                  ""field"": {
                    ""Column"": {
                      ""Expression"": {
                        ""SourceRef"": {
                          ""Entity"": ""{{factTableName}}""
                        }
                      },
                      ""Property"": ""{{factColumnName}}""
                    }
                  },
                  ""queryRef"": ""{{factTableName}}.{{factColumnName}}"",
                  ""nativeQueryRef"": ""{{factColumnName}}""
                },
                {
                  ""field"": {
                    ""Measure"": {
                      ""Expression"": {
                        ""SourceRef"": {
                          ""Entity"": ""{{dataQualityMeasureTable}}""
                        }
                      },
                      ""Property"": ""{{dataQualityMeasureName}}""
                    }
                  },
                  ""queryRef"": ""{{dataQualityMeasureTable}}.{{dataQualityMeasureName}}"",
                  ""nativeQueryRef"": ""{{dataQualityMeasureName}}""
                }
              ]
            }
          }
        },
        ""visualContainerObjects"": {
          ""title"": [
            {
              ""properties"": {
                ""show"": {
                  ""expr"": {
                    ""Literal"": {
                      ""Value"": ""true""
                    }
                  }
                },
                ""text"": {
                  ""expr"": {
                    ""Measure"": {
                      ""Expression"": {
                        ""SourceRef"": {
                          ""Entity"": ""{{dataQualityTitleMeasureTable}}""
                        }
                      },
                      ""Property"": ""{{dataQualityTitleMeasureName}}""
                    }
                  }
                }
              }
            }
          ],
          ""subTitle"": [
            {
              ""properties"": {
                ""show"": {
                  ""expr"": {
                    ""Literal"": {
                      ""Value"": ""true""
                    }
                  }
                },
                ""text"": {
                  ""expr"": {
                    ""Measure"": {
                      ""Expression"": {
                        ""SourceRef"": {
                          ""Entity"": ""{{dataQualitySubtitleMeasureTable}}""
                        }
                      },
                      ""Property"": ""{{dataQualitySubtitleMeasureName}}""
                    }
                  }
                }
              }
            }
          ]
        },
        ""drillFilterOtherVisuals"": true
      }
    }";
Dictionary<string, string> placeholders = new Dictionary<string, string>();
placeholders.Add("{{newPageId}}", "");
placeholders.Add("{{newVisualId}}", "");
placeholders.Add("{{pageName}}", pageName);
placeholders.Add("{{totalMeasureTable}}", totalMeasure.Table.Name);
placeholders.Add("{{totalMeasureName}}", totalMeasure.Name);
placeholders.Add("{{dataProblemsMeasureTable}}", detailMeasure.Table.Name);
placeholders.Add("{{dataProblemsMeasureName}}", detailMeasure.Name);
placeholders.Add("{{factTableName}}", "");  //factColumn.Table.Name);
placeholders.Add("{{factColumnName}}", "");  // factColumn.Name);
placeholders.Add("{{dataQualityMeasureTable}}", "");  // dataQualityMeasure.Table.Name);
placeholders.Add("{{dataQualityMeasureName}}", "");  //dataQualityMeasure.Name);
placeholders.Add("{{dataQualityTitleMeasureTable}}", "");  // dataQualityTitleMeasure.Table.Name);
placeholders.Add("{{dataQualityTitleMeasureName}}", "");  //dataQualityTitleMeasure.Name);
placeholders.Add("{{dataQualitySubtitleMeasureTable}}", "");  //dataQualitySubtitleMeasure.Table.Name);
placeholders.Add("{{dataQualitySubtitleMeasureName}}", "");  //dataQualitySubtitleMeasure.Name);
placeholders.Add("{{pageHeight}}", pageHeight.ToString());
placeholders.Add("{{pageWidth}}", pageWidth.ToString());
placeholders.Add("{{detailCardX}}", detailCardX.ToString());
placeholders.Add("{{detailCardY}}", detailCardY.ToString());
placeholders.Add("{{detailCardWidth}}", detailCardWidth.ToString());
placeholders.Add("{{detailCardHeight}}", detailCardHeight.ToString());
placeholders.Add("{{detailCardFontSize}}", detailCardFontSize.ToString());
placeholders.Add("{{totalCardX}}", totalCardX.ToString());
placeholders.Add("{{totalCardY}}", totalCardY.ToString());
placeholders.Add("{{totalCardWidth}}", totalCardWidth.ToString());
placeholders.Add("{{totalCardHeight}}", totalCardHeight.ToString());
placeholders.Add("{{zTabOrder}}", 0.ToString());
placeholders.Add("{{tableX}}", 0.ToString());
placeholders.Add("{{tableY}}", 0.ToString());
placeholders.Add("{{tableWidth}}", tableWidth.ToString());
placeholders.Add("{{tableHeight}}", tableHeight.ToString());
placeholders.Add("{{backgroundColor}}", backgroundColor);
placeholders.Add("{{backgroundTransparency}}", backgroundTransparency.ToString());

string pagesFolder = Fx.GetPagesFolder(pbirFilePath);
string newVisualId = "";
string tableContents = "";
int zTabOrder = -1000; 

//create new page
string newPageId = Guid.NewGuid().ToString();
placeholders["{{newPageId}}"] = newPageId;
zTabOrder = zTabOrder + 1000;
placeholders["{{zTabOrder}}"] = zTabOrder.ToString();
string pageContents = Fx.ReplacePlaceholders(pageContentsTemplate,placeholders);
string newPageFolder = Fx.AddNewPage(pageContents, pagesFolder, newPageId);

//create total card
newVisualId = Guid.NewGuid().ToString();
placeholders["{{newVisualId}}"] = newVisualId;
zTabOrder = zTabOrder + 1000;
placeholders["{{zTabOrder}}"] = zTabOrder.ToString();
string totalCardContents = Fx.ReplacePlaceholders(totalCardContentsTemplate,placeholders);
Fx.AddNewVisual(visualContents: totalCardContents, pageFolder: newPageFolder, newVisualId: newVisualId); 

//create detail card
newVisualId = Guid.NewGuid().ToString();
placeholders["{{newVisualId}}"] = newVisualId;
zTabOrder = zTabOrder + 1000;
placeholders["{{zTabOrder}}"] = zTabOrder.ToString();
string detailCardContents = Fx.ReplacePlaceholders(detailCardContentsTemplate, placeholders);
Fx.AddNewVisual(visualContents: detailCardContents, pageFolder: newPageFolder, newVisualId: newVisualId);

int currentRow = 1;
int currentColumn = 1;
int startX = totalCardX + totalCardWidth + interObjectGap; 
int startY = interObjectGap;

for(int i = 0; i < dataQualityMeasures.Count(); i++)
{
    //get references and calculate values
    Column factColumn = factTableColumns[i];
    Measure dataQualityMeasure = dataQualityMeasures[i];
    Measure dataQualityTitleMeasure = dataQualityTitles[i];
    Measure dataQualitySubtitleMeasure = dataQualitySubtitles[i];
    zTabOrder = zTabOrder + 1000;
    newVisualId = Guid.NewGuid().ToString();
    int tableX = startX + (currentColumn - 1) * (tableWidth + interObjectGap) ;
    int tableY = startY + (currentRow - 1) * (tableHeight + interObjectGap) ;
    //update the dictionary
    placeholders["{{newVisualId}}"] = newVisualId;
    placeholders["{{zTabOrder}}"] = zTabOrder.ToString();
    placeholders["{{factTableName}}"] =factColumn.Table.Name;
    placeholders["{{factColumnName}}"] = factColumn.Name;
    placeholders["{{dataQualityMeasureTable}}"] = dataQualityMeasure.Table.Name;
    placeholders["{{dataQualityMeasureName}}"] =dataQualityMeasure.Name;
    placeholders["{{dataQualityTitleMeasureTable}}"] = dataQualityTitleMeasure.Table.Name;
    placeholders["{{dataQualityTitleMeasureName}}"] =dataQualityTitleMeasure.Name;
    placeholders["{{dataQualitySubtitleMeasureTable}}"] =dataQualitySubtitleMeasure.Table.Name;
    placeholders["{{dataQualitySubtitleMeasureName}}"] =dataQualitySubtitleMeasure.Name;
    placeholders["{{tableX}}"] = tableX.ToString();
    placeholders["{{tableY}}"] = tableY.ToString();
    //fill the template
    tableContents = Fx.ReplacePlaceholders(tableContentsTemplate, placeholders);
    //create the folder & Json file
    Fx.AddNewVisual(visualContents: tableContents, pageFolder: newPageFolder, newVisualId: newVisualId);
    //update variables for the next table
    currentColumn = currentColumn + 1;
    if (currentColumn > tablesPerRow)
    {
        currentRow = currentRow + 1;
        currentColumn = 1;
    }
    
}

Info("New page added successfully. Close your PBIP project on Power BI desktop *without saving changes* and open again to see the new page with the button.");
//Uncomment in TE3
Application.UseWaitCursor = waitCursor;


public static class Fx
{
    public static string ReplacePlaceholders(string pageContentsTemplate, Dictionary<string, string> placeholders)
    {
        string pageContents = pageContentsTemplate; 
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
    public static string GetPagesFolder(string pbirFilePath)
    {
        FileInfo pbirFileInfo = new FileInfo(pbirFilePath);
        string pbirFolder = pbirFileInfo.Directory.FullName;
        string pagesFolder = Path.Combine(pbirFolder, "definition", "pages");
        return pagesFolder;
    }
    public static string AddNewPage(string pageContents, string pagesFolder, string newPageId)
    {

        string newPageFolder = Path.Combine(pagesFolder, newPageId);

        Directory.CreateDirectory(newPageFolder);

        string newPageFilePath = Path.Combine(newPageFolder, "page.json");
        File.WriteAllText(newPageFilePath, pageContents);

        string pagesFilePath = Path.Combine(pagesFolder, "pages.json");
        AddPageIdToPages(pagesFilePath, newPageId);

        return newPageFolder;
    }
    public static void AddNewVisual(string visualContents, string pageFolder, string newVisualId)
    {
        string visualsFolder = Path.Combine(pageFolder, "visuals");

        //maybe created earlier
        if (!Directory.Exists(visualsFolder))
        {
            Directory.CreateDirectory(visualsFolder);
        }

        string newVisualFolder = Path.Combine(visualsFolder, newVisualId); 

        Directory.CreateDirectory(newVisualFolder);

        string newVisualFilePath = Path.Combine(newVisualFolder, "visual.json");
        File.WriteAllText(newVisualFilePath, visualContents);

    }
    public static Table CreateCalcTable(Model model, string tableName, string tableExpression)
    {
        if (!model.Tables.Any(t => t.Name == tableName))
        {
            return model.AddCalculatedTable(tableName, tableExpression);
        }
        else
        {
            return model.Tables.Where(t => t.Name == tableName).First();
        }
    }
    public static string GetNameFromUser(string Prompt, string Title, string DefaultResponse)
    {
        string response = Interaction.InputBox(Prompt, Title, DefaultResponse, 740, 400);
        return response;
    }
    public static string ChooseString(IList<string> OptionList)
    {
        Func<IList<string>, string, string> SelectString = (IList<string> options, string title) =>
        {
            var form = new Form();
            form.Text = title;
            var buttonPanel = new Panel();
            buttonPanel.Dock = DockStyle.Bottom;
            buttonPanel.Height = 30;
            var okButton = new Button() { DialogResult = DialogResult.OK, Text = "OK" };
            var cancelButton = new Button() { DialogResult = DialogResult.Cancel, Text = "Cancel", Left = 80 };
            var listbox = new ListBox();
            listbox.Dock = DockStyle.Fill;
            listbox.Items.AddRange(options.ToArray());
            listbox.SelectedItem = options[0];
            form.Controls.Add(listbox);
            form.Controls.Add(buttonPanel);
            buttonPanel.Controls.Add(okButton);
            buttonPanel.Controls.Add(cancelButton);
            var result = form.ShowDialog();
            if (result == DialogResult.Cancel) return null;
            return listbox.SelectedItem.ToString();
        };
        //let the user select the name of the macro to copy
        String select = SelectString(OptionList, "Choose a macro");
        //check that indeed one macro was selected
        if (select == null)
        {
            Info("You cancelled!");
        }
        return select;
    }
    public static IEnumerable<Table> GetDateTables(Model model)
    {
        IEnumerable<Table> dateTables = null as IEnumerable<Table>;
        if (model.Tables.Any(t => t.DataCategory == "Time" && t.Columns.Any(c => c.IsKey == true)))
        {
            dateTables = model.Tables.Where(t => t.DataCategory == "Time" && t.Columns.Any(c => c.IsKey == true && c.DataType == DataType.DateTime));
        }
        else
        {
            Error("No date table detected in the model. Please mark your date table(s) as date table");
        }
        return dateTables;
    }
    public static Table GetTablesWithAnnotation(IEnumerable<Table> tables, string annotationLabel, string annotationValue)
    {
        Func<Table, bool> lambda = t => t.GetAnnotation(annotationLabel) == annotationValue;
        IEnumerable<Table> matchTables = GetFilteredTables(tables, lambda);
        if (matchTables == null)
        {
            return null;
        }
        else
        {
            return matchTables.First();
        }
    }
    public static IEnumerable<Table> GetFilteredTables(IEnumerable<Table> tables, Func<Table, bool> lambda)
    {
        if (tables.Any(t => lambda(t)))
        {
            return tables.Where(t => lambda(t));
        }
        else
        {
            return null as IEnumerable<Table>;
        }
    }
    public static IEnumerable<Column> GetFilteredColumns(IEnumerable<Column> columns, Func<Column, bool> lambda, bool returnAllIfNoneFound = true)
    {
        if (columns.Any(c => lambda(c)))
        {
            return columns.Where(c => lambda(c));
        }
        else
        {
            if (returnAllIfNoneFound)
            {
                return columns;
            }
            else
            {
                return null as IEnumerable<Column>;
            }
        }
    }
    public static Table SelectTableExt(Model model, string possibleName = null, string annotationName = null, string annotationValue = null, 
        Func<Table,bool>  lambdaExpression = null, string label = "Select Table", bool skipDialogIfSingleMatch = true, bool showOnlyMatchingTables = true)
    {
        if (lambdaExpression == null)
        {
            if (possibleName != null) { 
                lambdaExpression = (t) => t.Name == possibleName;
            } else if(annotationName!= null && annotationValue != null)
            {
                lambdaExpression = (t) => t.GetAnnotation(annotationName) == annotationValue;
            }
        }
        IEnumerable<Table> tables = model.Tables.Where(lambdaExpression);
        //none found, let the user choose from all tables
        if (tables.Count() == 0)
        {
            return SelectTable(tables: model.Tables, label: label);
        }
        else if (tables.Count() == 1 && !skipDialogIfSingleMatch)
        {
            return SelectTable(tables: model.Tables, preselect: tables.First(), label: label);
        }
        else if (tables.Count() == 1 && skipDialogIfSingleMatch)
        {
            return tables.First();
        } 
        else if (tables.Count() > 1 && showOnlyMatchingTables)
        {
            return SelectTable(tables: tables, preselect: tables.First(), label: label);
        }
        else if (tables.Count() > 1 && !showOnlyMatchingTables)
        {
            return SelectTable(tables: model.Tables, preselect: tables.First(), label: label);
        } else
        {
            Error(@"Unexpected logic in ""SelectTableExt""");
            return null;
        }
    }
    //add other methods always as "public static" followed by the data type they will return or void if they do not return anything.

    private static void AddPageIdToPages(string pagesFilePath, string pageId)
    {
        string pagesFileContents = File.ReadAllText(pagesFilePath);
        PagesDto pagesDto = JsonConvert.DeserializeObject<PagesDto>(pagesFileContents);
        if(pagesDto.pageOrder == null)
        {
            pagesDto.pageOrder = new List<string>();
        }
        
        if (!pagesDto.pageOrder.Contains(pageId)) { 

            pagesDto.pageOrder.Add(pageId);
            string resultFile = JsonConvert.SerializeObject(pagesDto, Formatting.Indented);
            File.WriteAllText(pagesFilePath, resultFile);
        }
    }
}

public class PagesDto
{
    [JsonProperty("$schema")]
    public string schema { get; set; }
    public List<string> pageOrder { get; set; }
    public string activePageName { get; set; }
}