#pragma warning disable 0219

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using UnityEngine;
using UnityEditor;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Object = UnityEngine.Object;

namespace JokeMaker.Editor
{
    public class ExcelImporterGenerator : EditorWindow
    {
        private Vector2 scrollPos = Vector2.zero;
        private readonly List<Object> excels = new List<Object>();
        private readonly List<string> excelFileExtensions = new List<string> {".xls", ".xlsx", ".xlsm"};
        private int excelIndex = -1;
        private const string entityNamespace = "JokeMaker.Entity";
        private const string importerNamespace = "JokeMaker.Importer";
        private const string defaultAssetOutputFolder = "Assets/JokeMaker/ExcelImporterGenerator/Output";
        private const string defaultClassesFolder = "Assets/JokeMaker/ExcelImporterGenerator/Classes";
        private const string defaultImporterFolder = "Assets/JokeMaker/ExcelImporterGenerator/Classes/Editor";
        private const string settingsFile = "Assets/JokeMaker/ExcelImporterGenerator/Editor/GeneratorSettings.asset";

        private GeneratorSettings settings;
        private GeneratorSettings Settings
        {
            get
            {
                if (settings != null) return settings;
                settings = AssetDatabase.LoadAssetAtPath<GeneratorSettings>(settingsFile);
                if (settings != null) return settings;
                settings = CreateInstance<GeneratorSettings>();
                settings.EntityNamespace = entityNamespace;
                settings.ImporterNamespace = importerNamespace;
                settings.AssetOutputFolder = defaultAssetOutputFolder;
                settings.EntitiesFolder = defaultClassesFolder;
                settings.ImportersFolder = $"{defaultClassesFolder}/Editor";
                settings.ExcelFiles = new List<string>();
                settings.hideFlags = HideFlags.NotEditable;
                AssetDatabase.CreateAsset(settings, settingsFile);
                settings = AssetDatabase.LoadAssetAtPath<GeneratorSettings>(settingsFile);
                return settings;
            }
        }

        private void OnGUI()
        {
            // Generator Settings
            EditorGUILayout.LabelField("Generator Settings");
            EditorGUILayout.BeginVertical("box");
            Settings.EntityNamespace = EditorGUILayout.TextField($"{nameof(Settings.EntityNamespace)}", Settings.EntityNamespace);
            Settings.ImporterNamespace = EditorGUILayout.TextField($"{nameof(Settings.ImporterNamespace)}", Settings.ImporterNamespace);
            var dataPath = Application.dataPath.Replace('\\', '/').TrimEnd('/');
            //- AssetOutputFolder
            EditorGUILayout.BeginHorizontal();
            EditorGUI.BeginDisabledGroup(true);
            EditorGUILayout.TextField($"{nameof(Settings.AssetOutputFolder)}", Settings.AssetOutputFolder);
            EditorGUI.EndDisabledGroup();
            if (GUILayout.Button("...", GUILayout.Width(25)))
            {
                var path = Path.Combine(dataPath, Settings.AssetOutputFolder.Substring(7));
                path = EditorUtility.OpenFolderPanel("Asset Output Folder", path, null);
                if (path != null)
                {
                    path = path.Replace('\\', '/');
                    if (path.StartsWith(dataPath))
                    {
                        Settings.AssetOutputFolder = "Assets/" + path.Substring(dataPath.Length + 1);
                    }
                }
            }
            EditorGUILayout.EndHorizontal();
            //- EntitiesFolder
            EditorGUILayout.BeginHorizontal();
            EditorGUI.BeginDisabledGroup(true);
            EditorGUILayout.TextField($"{nameof(Settings.EntitiesFolder)}", Settings.EntitiesFolder);
            EditorGUI.EndDisabledGroup();
            if (GUILayout.Button("...", GUILayout.Width(25)))
            {
                var path = Path.Combine(dataPath, Settings.EntitiesFolder.Substring(7));
                path = EditorUtility.OpenFolderPanel("Entities Folder", path, null);
                if (path != null)
                {
                    path = path.Replace('\\', '/');
                    if (path.StartsWith(dataPath))
                    {
                        Settings.AssetOutputFolder = "Assets/" + path.Substring(dataPath.Length + 1);
                    }
                }
            }
            EditorGUILayout.EndHorizontal();
            //- ImportersFolder
            EditorGUILayout.BeginHorizontal();
            EditorGUI.BeginDisabledGroup(true);
            EditorGUILayout.TextField($"{nameof(Settings.ImportersFolder)}", Settings.ImportersFolder);
            EditorGUI.EndDisabledGroup();
            if (GUILayout.Button("...", GUILayout.Width(25)))
            {
                var path = Path.Combine(dataPath, Settings.ImportersFolder.Substring(7));
                path = EditorUtility.OpenFolderPanel("Importers Folder", path, null);
                if (path != null && (path.Contains("/Editor/") || path.EndsWith("/Editor")))
                {
                    path = path.Replace('\\', '/');
                    if (path.StartsWith(dataPath))
                    {
                        Settings.AssetOutputFolder = "Assets/" + path.Substring(dataPath.Length + 1);
                    }
                }
            }
            EditorGUILayout.EndHorizontal();
            if (GUILayout.Button("Save Settings"))
            {
                Settings.ExcelFiles.Clear();
                foreach (var excelObj in excels)
                {
                    if (excelObj == null) continue;
                    Settings.ExcelFiles.Add(AssetDatabase.GetAssetPath(excelObj));
                }
                EditorUtility.SetDirty(Settings);
                AssetDatabase.SaveAssets();
                AssetDatabase.Refresh();
            }
            EditorGUILayout.EndVertical();

            // Excel List
            if (excels.Count == 0 && Settings.ExcelFiles.Count > 0)
            {
                foreach (var excelFile in Settings.ExcelFiles)
                {
                    excels.Add(AssetDatabase.LoadAssetAtPath<Object>(excelFile));
                }
            }

            EditorGUILayout.Space();
            EditorGUILayout.LabelField($"Excel List[{excels.Count}]");
            EditorGUILayout.BeginVertical("box");
            if (GUILayout.Button("+"))
            {
                excels.Add(null);
            }

            if (excels.Count > 0 && GUILayout.Button("~"))
            {
                for (var i = excels.Count - 1; i >= 0; i--)
                {
                    if (excels[i] == null) excels.RemoveAt(i);
                }
            }

            if (excels.Count > 0)
            {
                var height = excels.Count * 20f + 2;
                const float maxHeight = 200f;
                var layoutOpt = height > maxHeight ? GUILayout.MaxHeight(maxHeight) : GUILayout.Height(height);
                scrollPos = EditorGUILayout.BeginScrollView(scrollPos, layoutOpt);
                for (var i = 0; i < excels.Count; i++)
                {
                    EditorGUILayout.BeginHorizontal();
                    excels[i] = EditorGUILayout.ObjectField(excels[i], typeof(Object), false);
                    var assetPath = AssetDatabase.GetAssetPath(excels[i]);
                    var ext = Path.GetExtension(assetPath);
                    if (!excelFileExtensions.Contains(ext)) excels[i] = null;
                    if (GUILayout.Button("-", GUILayout.MaxWidth(25)))
                    {
                        excels.RemoveAt(i);
                        if (excelIndex == i) excelIndex = -1;
                        break;
                    }

                    if (excelIndex == i && excels[i] == null) excelIndex = -1;
                    EditorGUI.BeginDisabledGroup(excels[i] == null);
                    var toggleOn = GUILayout.Toggle(excelIndex == i, string.Empty, GUILayout.MaxWidth(25));
                    EditorGUI.EndDisabledGroup();
                    if (excels[i] != null) excelIndex = toggleOn ? i : -1;
                    EditorGUILayout.EndHorizontal();
                }

                EditorGUILayout.EndScrollView();
            }

            EditorGUILayout.EndVertical();

            // Chosen Sheet's Info List
            if (excelIndex < 0 || excelIndex >= excels.Count) return;
            var obj = excels[excelIndex];
            var filePath = AssetDatabase.GetAssetPath(obj);
            var sheetList = GetSheetInfoList(filePath);

            if (sheetList.Count == 0)
            {
                EditorGUILayout.HelpBox("There is no sheet!", MessageType.Warning, true);
                return;
            }

            EditorGUILayout.Space();
            EditorGUILayout.LabelField("Including Sheets");
            EditorGUILayout.BeginVertical("box");
            foreach (var sheet in sheetList)
            {
                var sheetName = sheet.Name;
                if (!string.IsNullOrEmpty(sheet.SubName)) sheetName += $"_{sheet.SubName}";
                EditorGUI.BeginDisabledGroup(true);
                EditorGUILayout.ToggleLeft(sheetName, sheet.Enabled);
                EditorGUI.EndDisabledGroup();
            }

            EditorGUILayout.EndVertical();

            // Sheet Group
            if (sheetList.Count == 0) return;
            var groupList = GetSheetGroupList(filePath, sheetList);

            if (groupList.Count == 0)
            {
                EditorGUILayout.HelpBox("There is no valid sheet group!", MessageType.Warning, true);
                return;
            }

            EditorGUILayout.Space();
            EditorGUILayout.LabelField("Sheet Groups");
            foreach (var sheetGroup in groupList)
            {
                EditorGUILayout.BeginVertical("box");
                EditorGUILayout.LabelField("Group Name:", sheetGroup.Name);
                if (!sheetGroup.Valid)
                {
                    EditorGUILayout.HelpBox("Invalid Group!", MessageType.Error);
                }
                else
                {
                    if (sheetGroup.MultiParts)
                    {
                        EditorGUILayout.BeginVertical("box");
                        EditorGUI.BeginDisabledGroup(true);
                        foreach (var sheetInfo in sheetGroup.Sheets)
                        {
                            EditorGUILayout.ToggleLeft(sheetInfo.SubName, sheetInfo.Enabled);
                        }

                        EditorGUI.EndDisabledGroup();
                        EditorGUILayout.EndVertical();
                    }

                    var colInfos = sheetGroup.Sheets[0].ColumnInfos;
                    if (colInfos.Count == 0)
                    {
                        EditorGUILayout.HelpBox("No Fields!", MessageType.Error);
                    }
                    else
                    {
                        EditorGUILayout.BeginVertical("box");
                        EditorGUILayout.BeginHorizontal();
                        EditorGUILayout.LabelField("FieldName");
                        EditorGUILayout.LabelField("IsArray", GUILayout.MaxWidth(80));
                        EditorGUILayout.LabelField("ValueType", GUILayout.MaxWidth(100));
                        EditorGUILayout.EndHorizontal();
                        foreach (var colInfo in colInfos)
                        {
                            EditorGUILayout.BeginHorizontal();
                            EditorGUILayout.ToggleLeft(colInfo.Name, colInfo.Enabled);
                            EditorGUI.BeginDisabledGroup(true);
                            EditorGUILayout.ToggleLeft("", colInfo.IsArray, GUILayout.MaxWidth(80));
                            EditorGUILayout.EnumPopup(colInfo.ValType, GUILayout.MaxWidth(100));
                            EditorGUI.EndDisabledGroup();
                            EditorGUILayout.EndHorizontal();
                        }

                        EditorGUILayout.EndVertical();
                    }
                }

                EditorGUILayout.EndVertical();
            }

            if (GUILayout.Button("Generate"))
            {
                foreach (var sheetGroup in groupList)
                {
                    ExportEntity(sheetGroup);
                }
                ExportImporter(groupList);
            }
        }

        private enum ValueType
        {
            STRING,
            BOOL,
            INT,
            LONG,
            FLOAT,
            DOUBLE,
            UNKNOWN
        }

        [MenuItem("Window/XLS Import Settings...")]
        private static void ExportExcelToAssetX()
        {
            var window = GetWindow<ExcelImporterGenerator>();
            window.Show();
        }

        private IList<ExcelSheetInfo> GetSheetInfoList(string filePath)
        {
            var sheetList = new List<ExcelSheetInfo>();
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var ext = Path.GetExtension(filePath);
                var book = ext == ".xls" ? (IWorkbook) new HSSFWorkbook(stream) : new XSSFWorkbook(stream);

                for (var i = 0; i < book.NumberOfSheets; ++i)
                {
                    var sheet = book.GetSheetAt(i);
                    var sheetInfo = new ExcelSheetInfo();
                    var sheetName = sheet.SheetName.Trim().Replace(" ", "").Replace("\t", "");
                    var parts = sheetName.Split(new[] {'_'}, StringSplitOptions.RemoveEmptyEntries);
                    sheetInfo.Name = parts[0];
                    sheetInfo.SubName = parts.Length > 1 ? parts[1] : null;
                    sheetList.Add(sheetInfo);
                    if (!sheetInfo.NameValid || !sheetInfo.SubNameValid)
                    {
                        Debug.LogError($"Sheet Name is invalid! [{sheetName}]\neg. Name_Sub, Name, -Name_Sub, -Name");
                    }

                    if (!sheetInfo.Enabled) continue;

                    var titleRow = sheet.GetRow(0); // 1st row [FieldName]
                    var typeRow = sheet.GetRow(1); // 2nd row [TypeSymbol]
                    for (var j = 0; j < titleRow.LastCellNum; j++)
                    {
                        var colInfo = new ExcelColumnInfo {Index = j};

                        var fieldName = titleRow.GetCell(j).StringCellValue;
                        fieldName = fieldName.Trim().Replace(" ", "").Replace("\t", "");
                        colInfo.Name = fieldName;
                        var typeStr = typeRow.GetCell(j).StringCellValue;
                        (colInfo.ValType, colInfo.IsArray, colInfo.ArraySep) = ParseValueType(typeStr);
                        sheetInfo.ColumnInfos.Add(colInfo);
                    }
                }
            }

            return sheetList;
        }

        private IList<ExcelSheetGroup> GetSheetGroupList(string filePath, IList<ExcelSheetInfo> sheetList)
        {
            var dictSheetGroup = new Dictionary<string, ExcelSheetGroup>();
            foreach (var sheetInfo in sheetList)
            {
                if (!sheetInfo.Enabled) continue;
                if (!dictSheetGroup.ContainsKey(sheetInfo.Name))
                {
                    dictSheetGroup[sheetInfo.Name] = new ExcelSheetGroup {Name = sheetInfo.Name};
                }

                var group = dictSheetGroup[sheetInfo.Name];
                if (group.Sheets.Count == 0)
                {
                    group.Sheets.Add(sheetInfo);
                    group.MultiParts = !string.IsNullOrEmpty(sheetInfo.SubName);
                    group.ExcelFilePath = filePath;
                }
                else
                {
                    if (!group.MultiParts)
                    {
                        Debug.LogError(
                            $"{Path.GetFileName(filePath)} - {sheetInfo.Name}_{sheetInfo.SubName} [Already exists an Single-part sheet group!]");
                        continue;
                    }

                    if (!string.IsNullOrEmpty(sheetInfo.SubName))
                    {
                        group.Sheets.Add(sheetInfo);
                    }
                    else
                    {
                        Debug.LogError(
                            $"{Path.GetFileName(filePath)} - {sheetInfo.Name} [Already exists an Multi-part sheet group!]");
                        continue;
                    }
                }
            }

            return new List<ExcelSheetGroup>(dictSheetGroup.Values);
        }

        private const string entityTemplateFile = "Assets/JokeMaker/ExcelImporterGenerator/Editor/Templates/EntityTemplate.txt";
        private void ExportEntity(ExcelSheetGroup group)
        {
            var entityTemplate = File.ReadAllText(entityTemplateFile);
            var content = entityTemplate.Replace("\r\n", "\n");
            content = content.TrimEnd() + '\n';
            var fieldBuilder = new StringBuilder();
            AppendLine(fieldBuilder);
            var sheet = group.Sheets[0];
            foreach (var colInfo in sheet.ColumnInfos)
            {
                if (!colInfo.Enabled) continue;
                var typeStr = colInfo.ValType.ToString().ToLowerInvariant();
                var fieldName = colInfo.Name;
                if (colInfo.IsArray) typeStr += "[]";
                AppendLine(fieldBuilder, $"public {typeStr} {fieldName};", 3);
            }

            var fieldsStr = fieldBuilder.ToString().Replace("\r\n", "\n").TrimEnd();
            content = content.Replace("$Fields$", fieldsStr);
            content = content.Replace("$ExcelData$", $"Entity{group.Name}");
            content = content.Replace("$Namespace$", entityNamespace);

            Directory.CreateDirectory(defaultClassesFolder);
            File.WriteAllText($"{defaultClassesFolder}/Entity{group.Name}.cs", content);
        }

        private const string importerTemplateFile = "Assets/JokeMaker/ExcelImporterGenerator/Editor/Templates/ImporterTemplate.txt";
        private void ExportImporter(IList<ExcelSheetGroup> groupList)
        {
            var importerTemplate = File.ReadAllText(importerTemplateFile);
            importerTemplate = importerTemplate.Replace("\r\n", "\n");
            var match1 = Regex.Match(importerTemplate, @"##ExportFunction1([\s\S]+?)##", RegexOptions.Multiline);
            var match2 = Regex.Match(importerTemplate, @"##ExportFunction2([\s\S]+?)##", RegexOptions.Multiline);
            importerTemplate = Regex.Replace(importerTemplate, @"##ExportFunction1[\s\S]+?##", "", RegexOptions.Multiline);
            importerTemplate = Regex.Replace(importerTemplate, @"##ExportFunction2[\s\S]+?##", "", RegexOptions.Multiline);
            importerTemplate = importerTemplate.TrimEnd() + '\n';
            var exportFunction1Template = match1.Groups[1].Value;
            var exportFunction2Template = match2.Groups[1].Value;
            
            var exportFunctionsBuilder = new StringBuilder();
            AppendLine(exportFunctionsBuilder);
            var exportFunctionCallsBuilder = new StringBuilder();
            AppendLine(exportFunctionCallsBuilder);
            AppendLine(exportFunctionCallsBuilder);
            var exportFieldsBuilder = new StringBuilder();

            foreach (var sheetGroup in groupList)
            {
                var functionContent = sheetGroup.MultiParts ? exportFunction2Template : exportFunction1Template;
                var functionName = $"ExportEntity{sheetGroup.Name}";
                if (sheetGroup.MultiParts)
                {
                    foreach (var sheet in sheetGroup.Sheets)
                    {
                        AppendLine(exportFunctionCallsBuilder, $"{functionName}(book, \"{sheet.SubName}\");", 5);
                    }
                }
                else
                {
                    AppendLine(exportFunctionCallsBuilder, $"{functionName}(book);", 5);
                }
                exportFieldsBuilder.Clear();
                AppendLine(exportFieldsBuilder);
                foreach (var columnInfo in sheetGroup.Sheets[0].ColumnInfos)
                {
                    if (!columnInfo.Enabled) continue;
                    var i = columnInfo.Index;
                    var sep = columnInfo.ArraySep;
                    var n = columnInfo.Name;
                    switch (columnInfo.ValType)
                    {
                        case ValueType.STRING:
                            AppendLine(exportFieldsBuilder, $"cell = row.GetCell({i});", 2);
                            AppendLine(exportFieldsBuilder,
                                columnInfo.IsArray
                                    ? $"p.{n} = cell.ToStringArray('{sep}', out var val{i}) ? val{i} : default;"
                                    : $"p.{n} = cell.ToString(out var val{i}) ? val{i} : default;", 2);
                            break;
                        case ValueType.BOOL:
                            AppendLine(exportFieldsBuilder, $"cell = row.GetCell({i});", 2);
                            AppendLine(exportFieldsBuilder,
                                columnInfo.IsArray
                                    ? $"p.{n} = cell.ToBoolArray('{sep}', out var val{i}) ? val{i} : default;"
                                    : $"p.{n} = cell.ToBool(out var val{i}) ? val{i} : default;", 2);
                            break;
                        case ValueType.INT:
                            AppendLine(exportFieldsBuilder, $"cell = row.GetCell({i});", 2);
                            AppendLine(exportFieldsBuilder,
                                columnInfo.IsArray
                                    ? $"p.{n} = cell.ToIntArray('{sep}', out var val{i}) ? val{i} : default;"
                                    : $"p.{n} = cell.ToInt(out var val{i}) ? val{i} : default;", 2);
                            break;
                        case ValueType.LONG:
                            AppendLine(exportFieldsBuilder, $"cell = row.GetCell({i});", 2);
                            AppendLine(exportFieldsBuilder,
                                columnInfo.IsArray
                                    ? $"p.{n} = cell.ToLongArray('{sep}', out var val{i}) ? val{i} : default;"
                                    : $"p.{n} = cell.ToLong(out var val{i}) ? val{i} : default;", 2);
                            break;
                        case ValueType.FLOAT:
                            AppendLine(exportFieldsBuilder, $"cell = row.GetCell({i});", 2);
                            AppendLine(exportFieldsBuilder,
                                columnInfo.IsArray
                                    ? $"p.{n} = cell.ToFloatArray('{sep}', out var val{i}) ? val{i} : default;"
                                    : $"p.{n} = cell.ToFloat(out var val{i}) ? val{i} : default;", 2);
                            break;
                        case ValueType.DOUBLE:
                            AppendLine(exportFieldsBuilder, $"cell = row.GetCell({i});", 2);
                            AppendLine(exportFieldsBuilder,
                                columnInfo.IsArray
                                    ? $"p.{n} = cell.ToDoubleArray('{sep}', out var val{i}) ? val{i} : default;"
                                    : $"p.{n} = cell.ToDouble(out var val{i}) ? val{i} : default;", 2);
                            break;
                        default:
                            throw new ArgumentOutOfRangeException();
                    }
                }

                functionContent = functionContent.Replace("$MainSheetName$", sheetGroup.Name);
                functionContent = functionContent.Replace("$EXPORT_DATA$", exportFieldsBuilder.ToString());
                functionContent = functionContent.Replace("\r\n", "\n").TrimEnd();
                AppendLine(exportFunctionsBuilder, functionContent);
            }

            var ss = new StringReader(exportFunctionsBuilder.ToString());
            exportFunctionsBuilder.Clear();
            string line = null;
            while ((line = ss.ReadLine()) != null)
            {
                line = line.TrimEnd();
                if (string.IsNullOrEmpty(line)) AppendLine(exportFunctionsBuilder);
                else
                {
                    AppendLine(exportFunctionsBuilder, line, 2);
                }
            }

            var mainContent = importerTemplate;
            var filePath = groupList[0].ExcelFilePath;
            var fileName = Path.GetFileNameWithoutExtension(groupList[0].ExcelFilePath);
            mainContent = mainContent.Replace("$EntityNamespace$", entityNamespace);
            mainContent = mainContent.Replace("$Namespace$", importerNamespace);
            mainContent = mainContent.Replace("$ExcelName$", fileName);
            mainContent = mainContent.Replace("$IMPORT_PATH$", filePath);
            mainContent = mainContent.Replace("$EXPORT_FOLDER$", defaultAssetOutputFolder);
            mainContent = mainContent.Replace("$ExportFunctionCalls$", exportFunctionCallsBuilder.ToString().TrimEnd());
            mainContent = mainContent.Replace("$ExportFunctions$", exportFunctionsBuilder.ToString().TrimEnd());

            Directory.CreateDirectory(defaultImporterFolder);
            File.WriteAllText($"{defaultImporterFolder}/ExcelImporter_{fileName}.cs", mainContent);
        }

        private static void AppendLine(StringBuilder sb, string msg = "", uint indentLevel = 0)
        {
            if (indentLevel == 0) sb.AppendLine(msg);
            else
            {
                const string indentUnit = "    ";
                var msgBuilder = new StringBuilder();
                for (var i = 0; i < indentLevel; ++i)
                {
                    msgBuilder.Append(indentUnit);
                }
                msgBuilder.Append(msg);

                sb.AppendLine(msgBuilder.ToString());
            }
        }

        private class ExcelSheetGroup
        {
            public string Name;
            public bool MultiParts;
            public readonly List<ExcelSheetInfo> Sheets = new List<ExcelSheetInfo>();
            public string ExcelFilePath;

            public bool Valid
            {
                get
                {
                    if (string.IsNullOrEmpty(Name)) return false;
                    if (Sheets.Count == 0) return false;
                    var sheet0 = Sheets[0];
                    if (!MultiParts)
                    {
                        return Sheets.Count == 1
                               && Name == sheet0.Name
                               && string.IsNullOrEmpty(sheet0.SubName)
                               && sheet0.Enabled;
                    }

                    foreach (var sheet in Sheets)
                    {
                        var valid = sheet.Name == Name
                                    && !string.IsNullOrEmpty(sheet.SubName)
                                    && sheet.Enabled;
                        if (!valid) return false;
                        if (sheet.ColumnInfos.Count != sheet0.ColumnInfos.Count) return false;
                        for (var i = 0; i < sheet.ColumnInfos.Count; ++i)
                        {
                            var colInfo = sheet.ColumnInfos[i];
                            if (!colInfo.Equals(sheet0.ColumnInfos[i])) return false;
                        }
                    }

                    return true;
                }
            }
        }

        private class ExcelSheetInfo
        {
            public string Name;
            public string SubName;
            public readonly List<ExcelColumnInfo> ColumnInfos = new List<ExcelColumnInfo>();

            public bool Enabled => NameValid && SubNameValid && !Name.StartsWith("-");
            public bool NameValid => Regex.IsMatch(Name, @"^-{0,1}([A-Z][a-z]*?)+?$");

            public bool SubNameValid =>
                string.IsNullOrEmpty(SubName) || Regex.IsMatch(SubName, @"^([A-Z][a-z]*?)+?[0-9]{0,3}$");
        }

        private class ExcelColumnInfo
        {
            public string Name;
            public ValueType ValType;
            public bool IsArray;
            public int Index;
            public char ArraySep;

            public bool Enabled => Valid && !Name.StartsWith("*");
            public bool Valid => ValType != ValueType.UNKNOWN && Regex.IsMatch(Name, @"^([A-Z][a-z]*?)+?[0-9]{0,1}$");

            public override bool Equals(object obj)
            {
                if (!(obj is ExcelColumnInfo colInfo)) return false;
                return Name == colInfo.Name && ValType == colInfo.ValType && IsArray == colInfo.IsArray;
            }
        }

        private static (ValueType, bool, char) ParseValueType(string typeStr)
        {
            var isArray = false;
            var arrSep = '#';
            if (typeStr.Contains("[]"))
            {
                var parts = typeStr.Split(new[] {"[]"}, StringSplitOptions.RemoveEmptyEntries);
                typeStr = parts[0];
                isArray = true;
                if (parts.Length > 1) arrSep = parts[1][0];
            }

            if (!Enum.TryParse(typeStr, out ValueType valType))
            {
                valType = ValueType.UNKNOWN;
            }

            return (valType, isArray, arrSep);
        }
    }
}
