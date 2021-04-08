﻿#pragma warning disable 0219

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
        private string entityNamespace = "JokeMaker.Entity";
        private string importerNamespace = "JokeMaker.Importer";
        private string defaultAssetOutputFolder = "Assets/JokeMaker/ExcelImporterGenerator/Output";
        private string defaultClassFolder = "Assets/JokeMaker/ExcelImporterGenerator/Classes";
        private string defaultImporterFolder = "Assets/JokeMaker/ExcelImporterGenerator/Classes/Editor";

        private void OnGUI()
        {
            // Excel List
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
            EditorGUILayout.LabelField("Sheet Settings");
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
                        var colInfo = new ExcelColumnInfo {Index = j + 1};
                        var fieldName = titleRow.GetCell(j).StringCellValue;
                        fieldName = fieldName.Trim().Replace(" ", "").Replace("\t", "");
                        colInfo.Name = fieldName;
                        var typeStr = typeRow.GetCell(j).StringCellValue;
                        (colInfo.ValType, colInfo.IsArray) = ParseValueType(typeStr);
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

        private const string entityTemplateFile = "Assets/JokeMaker/ExcelImporterGenerator/Editor/EntityTemplate.txt";
        private void ExportEntity(ExcelSheetGroup group)
        {
            var entityTemplate = File.ReadAllText(entityTemplateFile);
            var content = entityTemplate.Replace("\r\n", "\n");
            content = content.TrimEnd('\n') + '\n';
            var fieldBuilder = new StringBuilder();
            fieldBuilder.AppendLine();
            var sheet = group.Sheets[0];
            foreach (var colInfo in sheet.ColumnInfos)
            {
                if (!colInfo.Enabled) continue;
                var typeStr = colInfo.ValType.ToString().ToLowerInvariant();
                var fieldName = colInfo.Name;
                fieldBuilder.AppendLine(colInfo.IsArray
                    ? $"            public {typeStr}[] {fieldName};"
                    : $"            public {typeStr} {fieldName};");
            }

            var fieldsStr = fieldBuilder.ToString().Replace("\r\n", "\n").TrimEnd('\n');
            content = content.Replace("$Fields$", fieldsStr);
            content = content.Replace("$ExcelData$", $"Entity{group.Name}");
            content = content.Replace("$Namespace$", entityNamespace);

            Directory.CreateDirectory(defaultClassFolder);
            File.WriteAllText($"{defaultClassFolder}/Entity_{group.Name}.cs", content);
        }

        private const string importerTemplateFile = "Assets/JokeMaker/ExcelImporterGenerator/Editor/ImporterTemplate.txt";
        private void ExportImporter(IList<ExcelSheetGroup> groupList)
        {
            var importerTemplate = File.ReadAllText(importerTemplateFile);
            importerTemplate = importerTemplate.Replace("\r\n", "\n");
            var match1 = Regex.Match(importerTemplate, @"##ExportFunction1([\s\S]+?)##", RegexOptions.Multiline);
            var match2 = Regex.Match(importerTemplate, @"##ExportFunction2([\s\S]+?)##", RegexOptions.Multiline);
            importerTemplate = Regex.Replace(importerTemplate, @"##ExportFunction1[\s\S]+?##", "", RegexOptions.Multiline);
            importerTemplate = Regex.Replace(importerTemplate, @"##ExportFunction2[\s\S]+?##", "", RegexOptions.Multiline);
            importerTemplate = importerTemplate.TrimEnd('\n') + '\n';
            var exportFunction1Template = match1.Groups[1].Value;
            var exportFunction2Template = match2.Groups[1].Value;
            
            var exportFunctionsBuilder = new StringBuilder();
            exportFunctionsBuilder.AppendLine();
            var exportFunctionCallsBuilder = new StringBuilder();
            exportFunctionCallsBuilder.AppendLine();
            var exportFieldsBuilder = new StringBuilder();

            foreach (var sheetGroup in groupList)
            {
                var functionContent = sheetGroup.MultiParts ? exportFunction2Template : exportFunction1Template;
                var functionName = $"ExportEntity{sheetGroup.Name}";
                if (sheetGroup.MultiParts)
                {
                    foreach (var sheet in sheetGroup.Sheets)
                    {
                        exportFunctionCallsBuilder.AppendLine($"{functionName}(book, \"{sheet.SubName}\");");
                    }
                }
                else
                {
                    exportFunctionCallsBuilder.AppendLine($"{functionName}(book);");
                }
                exportFieldsBuilder.Clear();
                exportFieldsBuilder.AppendLine();
                foreach (var columnInfo in sheetGroup.Sheets[0].ColumnInfos)
                {
                    if (!columnInfo.Enabled) continue;
                    switch (columnInfo.ValType)
                    {
                        case ValueType.STRING:
                            exportFieldsBuilder.AppendLine(
                                $"cell = row.GetCell({columnInfo.Index}); p.{columnInfo.Name} = (cell == null ? string.Empty : cell.StringCellValue);");
                            break;
                        case ValueType.BOOL:
                            exportFieldsBuilder.AppendLine(
                                $"cell = row.GetCell({columnInfo.Index}); p.{columnInfo.Name} = (cell == null ? false : cell.BooleanCellValue);");
                            break;
                        case ValueType.INT:
                            exportFieldsBuilder.AppendLine(
                                $"cell = row.GetCell({columnInfo.Index}); p.{columnInfo.Name} = (cell == null ? 0 : (int) cell.NumericCellValue);");
                            break;
                        case ValueType.LONG:
                            exportFieldsBuilder.AppendLine(
                                $"cell = row.GetCell({columnInfo.Index}); p.{columnInfo.Name} = (cell == null ? 0L : (long) cell.NumericCellValue);");
                            break;
                        case ValueType.FLOAT:
                            exportFieldsBuilder.AppendLine(
                                $"cell = row.GetCell({columnInfo.Index}); p.{columnInfo.Name} = (cell == null ? 0f : (float) cell.NumericCellValue);");
                            break;
                        case ValueType.DOUBLE:
                            exportFieldsBuilder.AppendLine(
                                $"cell = row.GetCell({columnInfo.Index}); p.{columnInfo.Name} = (cell == null ? 0.0 : cell.NumericCellValue);");
                            break;
                        case ValueType.UNKNOWN:
                            break;
                        default:
                            throw new ArgumentOutOfRangeException();
                    }
                }

                functionContent = functionContent.Replace("$MainSheetName$", sheetGroup.Name);
                functionContent = functionContent.Replace("$EXPORT_DATA$", exportFieldsBuilder.ToString());
                exportFunctionsBuilder.AppendLine(functionContent);
            }

            var mainContent = importerTemplate;
            var filePath = groupList[0].ExcelFilePath;
            var fileName = Path.GetFileNameWithoutExtension(groupList[0].ExcelFilePath);
            mainContent = mainContent.Replace("$EntityNamespace$", entityNamespace);
            mainContent = mainContent.Replace("$Namespace$", importerNamespace);
            mainContent = mainContent.Replace("$ExcelName$", fileName);
            mainContent = mainContent.Replace("$IMPORT_PATH$", filePath);
            mainContent = mainContent.Replace("$EXPORT_FOLDER$", defaultAssetOutputFolder);
            mainContent = mainContent.Replace("$ExportFunctionCalls$", exportFunctionCallsBuilder.ToString());
            mainContent = mainContent.Replace("$ExportFunctions$", exportFunctionsBuilder.ToString());

            Directory.CreateDirectory(defaultImporterFolder);
            File.WriteAllText($"{defaultImporterFolder}/ExcelImporter_{fileName}.cs", mainContent);
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

            public bool Enabled => Valid && !Name.StartsWith("*");
            public bool Valid => ValType != ValueType.UNKNOWN && Regex.IsMatch(Name, @"^([A-Z][a-z]*?)+?[0-9]{0,1}$");

            public override bool Equals(object obj)
            {
                if (!(obj is ExcelColumnInfo colInfo)) return false;
                return Name == colInfo.Name && ValType == colInfo.ValType && IsArray == colInfo.IsArray;
            }
        }

        private static (ValueType, bool) ParseValueType(string typeStr)
        {
            var isArray = false;
            if (typeStr.EndsWith("[]"))
            {
                isArray = true;
                typeStr = typeStr.Substring(0, typeStr.Length - 2);
            }

            if (!Enum.TryParse(typeStr, out ValueType valType))
            {
                valType = ValueType.UNKNOWN;
            }

            return (valType, isArray);
        }
    }
}
