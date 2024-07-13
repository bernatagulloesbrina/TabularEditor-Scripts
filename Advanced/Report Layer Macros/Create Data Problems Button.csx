using System.Windows.Forms;
using System.IO;
// '2024-07-10 / B.Agullo / 
// Instructions:
// execute after running latest version of dataProblemsButtomMeasureCreation macro
// See https://www.esbrina-ba.com/c-scripting-the-report-layer-with-tabular-editor/ for detail


/*uncomment in TE3 to avoid wating cursor infront of dialogs*/

//ScriptHelper.WaitFormVisible = false;
//
//bool waitCursor = Application.UseWaitCursor;
//Application.UseWaitCursor = false;

DialogResult dialogResult = MessageBox.Show(text:"Did you save your model changes before running this macro?", caption:"Saved changes?", buttons:MessageBoxButtons.YesNo);

if(dialogResult != DialogResult.Yes){
    Info("Please save your changes first and then run this macro"); 
    return; 
};

string annotationLabel = "DataProblemsMeasures";
string annotationValueNavigation = "ButtonNavigationMeasure";
string annotationValueText = "ButtonTextMeasure";
string annotationValueBackground = "ButtonBackgroundMeasure";
string[] annotationArray = new string[3] { annotationValueNavigation, annotationValueText, annotationValueBackground };
foreach(string annotation in annotationArray)
{
    if(!Model.AllMeasures.Any(m => m.GetAnnotation(annotationLabel) == annotation))
    {
        Error(String.Format("No measure found with annotation {0} = {1} ", annotationLabel, annotationValueNavigation));
        return;
    }
}


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
    


string newPageId = Guid.NewGuid().ToString();

string newVisualId = Guid.NewGuid().ToString();




Measure navigationMeasure = 
    Model.AllMeasures
        .Where(m => m.GetAnnotation(annotationLabel) == annotationValueNavigation)
        .FirstOrDefault();
Measure textMeasure = 
    Model.AllMeasures
        .Where(m => m.GetAnnotation(annotationLabel) == annotationValueText)
        .FirstOrDefault();
Measure backgroundMeasure = 
    Model.AllMeasures
        .Where(m => m.GetAnnotation(annotationLabel) == annotationValueBackground)
        .FirstOrDefault();



string newPageContent =@"
{
  ""$schema"": ""https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.0.0/schema.json"",
  ""name"": ""{{newPageId}}"",
  ""displayName"": ""Problems Button"",
  ""displayOption"": ""FitToPage"",
  ""height"": 720,
  ""width"": 1280
}";
 


string newVisualContent = @"{
	""$schema"": ""https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.0.0/schema.json"",
	""name"": ""{{newVisualId}}"",
	""position"": {
		""x"": 510.44776119402987,
		""y"": 256.1194029850746,
		""z"": 0,
		""width"": 188.0597014925373,
		""height"": 50.14925373134328
	},
	""visual"": {
		""visualType"": ""actionButton"",
		""objects"": {
			""icon"": [
				{
					""properties"": {
						""shapeType"": {
							""expr"": {
								""Literal"": {
									""Value"": ""'blank'""
								}
							}
						}
					},
					""selector"": {
						""id"": ""default""
					}
				},
				{
					""properties"": {
						""show"": {
							""expr"": {
								""Literal"": {
									""Value"": ""false""
								}
							}
						}
					}
				}
			],
			""outline"": [
				{
					""properties"": {
						""show"": {
							""expr"": {
								""Literal"": {
									""Value"": ""false""
								}
							}
						}
					}
				}
			],
			""text"": [
				{
					""properties"": {
						""show"": {
							""expr"": {
								""Literal"": {
									""Value"": ""true""
								}
							}
						}
					}
				},
				{
					""properties"": {
						""text"": {
							""expr"": {
								""Measure"": {
									""Expression"": {
										""SourceRef"": {
											""Entity"": ""{{textMeasureTable}}""
										}
									},
									""Property"": ""{{textMeasureName}}""
								}
							}
						},
						""bold"": {
							""expr"": {
								""Literal"": {
									""Value"": ""true""
								}
							}
						},
						""fontColor"": {
							""solid"": {
								""color"": {
									""expr"": {
										""ThemeDataColor"": {
											""ColorId"": 0,
											""Percent"": 0
										}
									}
								}
							}
						}
					},
					""selector"": {
						""id"": ""default""
					}
				}
			],
			""fill"": [
				{
					""properties"": {
						""show"": {
							""expr"": {
								""Literal"": {
									""Value"": ""true""
								}
							}
						}
					}
				},
				{
					""properties"": {
						""fillColor"": {
							""solid"": {
								""color"": {
									""expr"": {
										""Measure"": {
											""Expression"": {
												""SourceRef"": {
													""Entity"": ""{{backgroundMeasureTable}}""
												}
											},
											""Property"": ""{{backgroundMeasureName}}""
										}
									}
								}
							}
						},
						""transparency"": {
							""expr"": {
								""Literal"": {
									""Value"": ""0D""
								}
							}
						}
					},
					""selector"": {
						""id"": ""default""
					}
				}
			]
		},
		""visualContainerObjects"": {
			""visualLink"": [
				{
					""properties"": {
						""show"": {
							""expr"": {
								""Literal"": {
									""Value"": ""true""
								}
							}
						},
						""type"": {
							""expr"": {
								""Literal"": {
									""Value"": ""'PageNavigation'""
								}
							}
						},
						""navigationSection"": {
							""expr"": {
								""Measure"": {
									""Expression"": {
										""SourceRef"": {
											""Entity"": ""{{navigationMeasureTable}}""
										}
									},
									""Property"": ""{{navigationMeasureName}}""
								}
							}
						}
					}
				}
			]
		},
		""drillFilterOtherVisuals"": true
	},
	""howCreated"": ""InsertVisualButton""
}";


Dictionary<string,string> placeholders = new Dictionary<string,string>();
placeholders.Add("{{newPageId}}", newPageId);
placeholders.Add("{{newVisualId}}", newVisualId);
placeholders.Add("{{textMeasureTable}}",textMeasure.Table.Name);
placeholders.Add("{{textMeasureName}}",textMeasure.Name);
placeholders.Add("{{backgroundMeasureTable}}",backgroundMeasure.Table.Name);
placeholders.Add("{{backgroundMeasureName}}",backgroundMeasure.Name);
placeholders.Add("{{navigationMeasureTable}}",navigationMeasure.Table.Name);
placeholders.Add("{{navigationMeasureName}}",navigationMeasure.Name);


newPageContent = ReportManager.ReplacePlaceholders(newPageContent,placeholders);
newVisualContent = ReportManager.ReplacePlaceholders(newVisualContent, placeholders);
ReportManager.AddNewPage(newPageContent, newVisualContent,pbirFilePath,newPageId,newVisualId);



Info("New page added successfully. Close your PBIP project on Power BI desktop *without saving changes* and open again to see the new page with the button.");
//Application.UseWaitCursor = waitCursor;

public static class ReportManager
{

    public static string ReplacePlaceholders(string jsonContents, Dictionary<string, string> placeholders)
    {
        foreach(string placeholder in placeholders.Keys)
        {
            string valueToReplace = placeholders[placeholder];

            jsonContents = jsonContents.Replace(placeholder, valueToReplace);

        }

        return jsonContents;
    }
    
    
    public static void AddNewPage(string pageContents, string visualContents, string pbirFilePath, string newPageId, string newVisualId)
    {
        
        FileInfo pbirFileInfo = new FileInfo(pbirFilePath);

        string pbirFolder = pbirFileInfo.Directory.FullName;
        string pagesFolder = Path.Combine(pbirFolder, "definition", "pages");
        string pagesFilePath = Path.Combine(pagesFolder, "pages.json");

        string newPageFolder = Path.Combine(pagesFolder, newPageId);

        Directory.CreateDirectory(newPageFolder);

        string newPageFilePath = Path.Combine(newPageFolder, "page.json");
        File.WriteAllText(newPageFilePath, pageContents);

        string visualsFolder = Path.Combine(newPageFolder,"visuals");
        Directory.CreateDirectory(visualsFolder);

        string newVisualFolder = Path.Combine(visualsFolder,newVisualId);
        Directory.CreateDirectory(newVisualFolder); 

        string newVisualFilePath = Path.Combine(newVisualFolder,"visual.json"); 
        File.WriteAllText(newVisualFilePath,visualContents); 

        AddPageIdToPages(pagesFilePath, newPageId);
    }

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