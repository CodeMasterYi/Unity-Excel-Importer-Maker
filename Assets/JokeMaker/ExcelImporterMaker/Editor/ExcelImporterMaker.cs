#pragma warning disable 0219

using UnityEngine;
using UnityEditor;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Text;

public class ExcelImporterMaker : EditorWindow
{
    private Vector2 scrollPos = Vector2.zero;
    private string filePath = string.Empty;
    private bool separateSheet = false;
    private readonly List<ExcelRowParameter> typeList = new List<ExcelRowParameter>();
    private readonly List<ExcelSheetParameter> sheetList = new List<ExcelSheetParameter>();
    private string className = string.Empty;
    private string fileName = string.Empty;
    private static string editorPrefKeyPrefix = "JokeMaker.excel-importer-maker.";

    private void OnGUI()
    {
        GUILayout.Label("Making Importer", EditorStyles.boldLabel);
        className = EditorGUILayout.TextField("Class Name", className);
        separateSheet = EditorGUILayout.ToggleLeft("Separate Sheet", separateSheet);

        EditorPrefs.SetBool($"{editorPrefKeyPrefix}{fileName}.separateSheet", separateSheet);

        if (GUILayout.Button("Generate"))
        {
            EditorPrefs.SetString($"{editorPrefKeyPrefix}{fileName}.className", className);
            ExportEntity();
            ExportImporter();

            AssetDatabase.ImportAsset(filePath);
            AssetDatabase.Refresh(ImportAssetOptions.ForceUpdate);
            Close();
        }

        EditorGUILayout.LabelField("Sheet Settings");
        EditorGUILayout.BeginVertical("box");
        foreach (var sheet in sheetList)
        {
            sheet.isEnable = EditorGUILayout.ToggleLeft(sheet.sheetName, sheet.isEnable);
            EditorPrefs.SetBool(editorPrefKeyPrefix + fileName + ".sheet." + sheet.sheetName, sheet.isEnable);
        }
        EditorGUILayout.EndVertical();

        // selecting parameters
        EditorGUILayout.LabelField("Field Settings");
        scrollPos = EditorGUILayout.BeginScrollView(scrollPos);
        EditorGUILayout.BeginVertical("box");
        var lastCellName = string.Empty;
        foreach (var cell in typeList)
        {
            if (cell.isArray && lastCellName != null && cell.name.Equals(lastCellName))
            {
                continue;
            }

            GUILayout.BeginHorizontal();
            var cellLabelName = cell.name;
            if (cell.isArray)
            {
                cellLabelName += " [Array]";
            }
            cell.isEnable = EditorGUILayout.ToggleLeft(cellLabelName, cell.isEnable);
            cell.type = (ValueType)EditorGUILayout.EnumPopup(cell.type, GUILayout.MaxWidth(100));
            GUILayout.EndHorizontal();
            lastCellName = cell.name;
        }
        EditorGUILayout.EndVertical();
        EditorGUILayout.EndScrollView();
    }

    private enum ValueType
    {
        BOOL,
        STRING,
        INT,
		FLOAT,
		DOUBLE,
    }

    [MenuItem("Assets/XLS Import Settings...")]
    private static void ExportExcelToAsset()
    {
        foreach (var obj in Selection.objects)
        {
            var window = CreateInstance<ExcelImporterMaker>();
            window.filePath = AssetDatabase.GetAssetPath(obj);
            window.fileName = Path.GetFileNameWithoutExtension(window.filePath);
		
			using (var stream = File.Open (window.filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var ext = Path.GetExtension(window.filePath);
                var book = ext == ".xls" ? (IWorkbook) new HSSFWorkbook(stream) : new XSSFWorkbook(stream);

                for (var i = 0; i < book.NumberOfSheets; ++i)
                {
                    var s = book.GetSheetAt(i);
                    var sht = new ExcelSheetParameter {sheetName = s.SheetName};
                    sht.isEnable = EditorPrefs.GetBool(editorPrefKeyPrefix + window.fileName + ".sheet." + sht.sheetName, true);
                    window.sheetList.Add(sht);
                }
			
                var sheet = book.GetSheetAt(0);
                window.className = EditorPrefs.GetString(editorPrefKeyPrefix + window.fileName + ".className", "Entity_" + sheet.SheetName);
                window.separateSheet = EditorPrefs.GetBool(editorPrefKeyPrefix + window.fileName + ".separateSheet");

                var titleRow = sheet.GetRow(0);
                var dataRow = sheet.GetRow(1);
                for (int i=0; i < titleRow.LastCellNum; i++)
                {
                    ExcelRowParameter lastParser = null;
                    ExcelRowParameter parser = new ExcelRowParameter();
                    parser.name = titleRow.GetCell(i).StringCellValue;
                    parser.isArray = parser.name.Contains("[]");
                    if (parser.isArray)
                    {
                        parser.name = parser.name.Remove(parser.name.LastIndexOf("[]"));
                    }

                    ICell cell = dataRow.GetCell(i);

                    // array support
                    if (window.typeList.Count > 0)
                    {
                        lastParser = window.typeList [window.typeList.Count - 1];
                        if (lastParser.isArray && parser.isArray && lastParser.name.Equals(parser.name))
                        {
                            // trailing array items must be the same as the top type
                            parser.isEnable = lastParser.isEnable;
                            parser.type = lastParser.type;
                            lastParser.nextArrayItem = parser;
                            window.typeList.Add(parser);
                            continue;
                        }
                    }
				
                    if (cell.CellType != CellType.Unknown && cell.CellType != CellType.Blank)
                    {
                        parser.isEnable = true;

                        try
                        {
                            if (EditorPrefs.HasKey(editorPrefKeyPrefix + window.fileName + ".type." + parser.name))
                            {
                                parser.type = (ValueType)EditorPrefs.GetInt(editorPrefKeyPrefix + window.fileName + ".type." + parser.name);
                            } else
                            {
                                string sampling = cell.StringCellValue;
                                parser.type = ValueType.STRING;
                            }
                        } catch
                        {
                        }
                        try
                        {
                            if (EditorPrefs.HasKey(editorPrefKeyPrefix + window.fileName + ".type." + parser.name))
                            {
                                parser.type = (ValueType)EditorPrefs.GetInt(editorPrefKeyPrefix + window.fileName + ".type." + parser.name);
                            } else
                            {
                                double sampling = cell.NumericCellValue;
                                parser.type = ValueType.DOUBLE;
                            }
                        } catch
                        {
                        }
                        try
                        {
                            if (EditorPrefs.HasKey(editorPrefKeyPrefix + window.fileName + ".type." + parser.name))
                            {
                                parser.type = (ValueType)EditorPrefs.GetInt(editorPrefKeyPrefix + window.fileName + ".type." + parser.name);
                            } else
                            {
                                bool sampling = cell.BooleanCellValue;
                                parser.type = ValueType.BOOL;
                            }
                        } catch
                        {
                        }
                    }
				
                    window.typeList.Add(parser);
                }
			
                window.Show();
            }
        }
    }
	
    void ExportEntity()
    {
        string templateFilePath = (separateSheet) ? "Assets/Terasurware/Editor/EntityTemplate2.txt" : "Assets/Terasurware/Editor/EntityTemplate.txt";
        string entittyTemplate = File.ReadAllText(templateFilePath);
        entittyTemplate = entittyTemplate.Replace("\r\n", "\n").Replace("\n", System.Environment.NewLine);
        StringBuilder builder = new StringBuilder();
        bool isInbetweenArray = false;
        foreach (ExcelRowParameter row in typeList)
        {
            if (row.isEnable)
            {
                if (!row.isArray)
                {
                    builder.AppendLine();
                    builder.AppendFormat("		public {0} {1};", row.type.ToString().ToLower(), row.name);
                } else
                {
                    if (!isInbetweenArray)
                    {
                        builder.AppendLine();
                        builder.AppendFormat("		public {0}[] {1};", row.type.ToString().ToLower(), row.name);
                    } 
                    isInbetweenArray = (row.nextArrayItem != null);
                }
            }
        }
		
        entittyTemplate = entittyTemplate.Replace("$Types$", builder.ToString());
        entittyTemplate = entittyTemplate.Replace("$ExcelData$", className);
		
        Directory.CreateDirectory("Assets/Terasurware/Classes/");
        File.WriteAllText("Assets/Terasurware/Classes/" + className + ".cs", entittyTemplate);
    }
	
    void ExportImporter()
    {
        string templateFilePath = (separateSheet) ? "Assets/Terasurware/Editor/ExportTemplate2.txt" : "Assets/Terasurware/Editor/ExportTemplate.txt";

        string importerTemplate = File.ReadAllText(templateFilePath);
		
        StringBuilder builder = new StringBuilder();
        StringBuilder sheetListbuilder = new StringBuilder();
        int rowCount = 0;
        string tab = "					";
        bool isInbetweenArray = false;

        //public string[] sheetNames = {"hoge", "fuga"};
        //$SheetList$
        foreach (ExcelSheetParameter sht in sheetList)
        {
            if (sht.isEnable)
            {
                sheetListbuilder.Append("\"" + sht.sheetName + "\",");
            }
            /*
            if (sht != sheetList [sheetList.Count - 1])
            {
                sheetListbuilder.Append(",");
            }
            */
        }
		
        foreach (ExcelRowParameter row in typeList)
        {
            if (row.isEnable)
            {
                if (!row.isArray)
                {
                    builder.AppendLine();
                    switch (row.type)
                    {
                        case ValueType.BOOL:
                            builder.AppendFormat(tab + "cell = row.GetCell({1}); p.{0} = (cell == null ? false : cell.BooleanCellValue);", row.name, rowCount);
                            break;
                        case ValueType.DOUBLE:
                            builder.AppendFormat(tab + "cell = row.GetCell({1}); p.{0} = (cell == null ? 0.0 : cell.NumericCellValue);", row.name, rowCount);
                            break;
                        case ValueType.INT:
                            builder.AppendFormat(tab + "cell = row.GetCell({1}); p.{0} = (int)(cell == null ? 0 : cell.NumericCellValue);", row.name, rowCount);
                            break;
						case ValueType.FLOAT:
							builder.AppendFormat(tab + "cell = row.GetCell({1}); p.{0} = (float)(cell == null ? 0 : cell.NumericCellValue);", row.name, rowCount);
							break;
						case ValueType.STRING:
                            builder.AppendFormat(tab + "cell = row.GetCell({1}); p.{0} = (cell == null ? \"\" : cell.StringCellValue);", row.name, rowCount);
                            break;
					}
                } else
                {
                    // only the head of array should generate code

                    if (!isInbetweenArray)
                    {
                        int arrayLength = 0;
                        for (ExcelRowParameter r = row; r != null; r = r.nextArrayItem, ++arrayLength)
                        {
                        }

                        builder.AppendLine();
                        switch (row.type)
                        {
                            case ValueType.BOOL:
                                builder.AppendFormat(tab + "p.{0} = new bool[{1}];", row.name, arrayLength);
                                break;
                            case ValueType.DOUBLE:
                                builder.AppendFormat(tab + "p.{0} = new double[{1}];", row.name, arrayLength);
                                break;
                            case ValueType.INT:
                                builder.AppendFormat(tab + "p.{0} = new int[{1}];", row.name, arrayLength);
                                break;
							case ValueType.FLOAT:
								builder.AppendFormat(tab + "p.{0} = new float[{1}];", row.name, arrayLength);
								break;
                            case ValueType.STRING:
                                builder.AppendFormat(tab + "p.{0} = new string[{1}];", row.name, arrayLength);
                                break;
                        }
						
                        for (int i = 0; i < arrayLength; ++i)
                        {
                            builder.AppendLine();
                            switch (row.type)
                            {
                                case ValueType.BOOL:
                                    builder.AppendFormat(tab + "cell = row.GetCell({1}); p.{0}[{2}] = (cell == null ? false : cell.BooleanCellValue);", row.name, rowCount + i, i);
                                    break;
                                case ValueType.DOUBLE:
                                    builder.AppendFormat(tab + "cell = row.GetCell({1}); p.{0}[{2}] = (cell == null ? 0.0 : cell.NumericCellValue);", row.name, rowCount + i, i);
                                    break;
                                case ValueType.INT:
                                    builder.AppendFormat(tab + "cell = row.GetCell({1}); p.{0}[{2}] = (int)(cell == null ? 0 : cell.NumericCellValue);", row.name, rowCount + i, i);
									break;
								case ValueType.FLOAT:
									builder.AppendFormat(tab + "cell = row.GetCell({1}); p.{0}[{2}] = (float)(cell == null ? 0.0 : cell.NumericCellValue);", row.name, rowCount + i, i);
									break;
                                case ValueType.STRING:
                                    builder.AppendFormat(tab + "cell = row.GetCell({1}); p.{0}[{2}] = (cell == null ? \"\" : cell.StringCellValue);", row.name, rowCount + i, i);
                                    break;
                            }
                        }
                    }
                    isInbetweenArray = (row.nextArrayItem != null);
                }
            }
            rowCount += 1;
        }

        importerTemplate = importerTemplate.Replace("$IMPORT_PATH$", filePath);
        importerTemplate = importerTemplate.Replace("$ExportAssetDirectry$", Path.GetDirectoryName(filePath));
        importerTemplate = importerTemplate.Replace("$EXPORT_PATH$", Path.ChangeExtension(filePath, ".asset"));
        importerTemplate = importerTemplate.Replace("$ExcelData$", className);
        importerTemplate = importerTemplate.Replace("$SheetList$", sheetListbuilder.ToString());
        importerTemplate = importerTemplate.Replace("$EXPORT_DATA$", builder.ToString());
        importerTemplate = importerTemplate.Replace("$ExportTemplate$", fileName + "_importer");
			
        Directory.CreateDirectory("Assets/Terasurware/Classes/Editor/");
        File.WriteAllText("Assets/Terasurware/Classes/Editor/" + fileName + "_importer.cs", importerTemplate);
    }
	
    private class ExcelSheetParameter
    {
        public string sheetName;
        public bool isEnable;
    }

    private class ExcelRowParameter
    {
        public ValueType type;
        public string name;
        public bool isEnable;
        public bool isArray;
        public ExcelRowParameter nextArrayItem;
    }
}
