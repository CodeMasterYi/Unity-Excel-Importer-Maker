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
                if (!string.IsNullOrEmpty(sheet.SubName)) sheetName += $"-{sheet.SubName}";
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

//        GUILayout.Label("Making Importer", EditorStyles.boldLabel);
//        className = EditorGUILayout.TextField("Class Name", className);
//
//        EditorGUILayout.LabelField("Sheet Settings");
//        EditorGUILayout.BeginVertical("box");
//        foreach (var sheet in sheetList)
//        {
//            EditorGUI.BeginDisabledGroup(true);
//            EditorGUILayout.ToggleLeft(sheet.Name.TrimStart('_'), sheet.Enabled);
//            EditorGUI.EndDisabledGroup();
//        }
//        EditorGUILayout.EndVertical();
//
//        EditorGUILayout.LabelField("Field Settings");
//        scrollPos = EditorGUILayout.BeginScrollView(scrollPos);
//        EditorGUILayout.BeginVertical("box");
//        var lastCellName = string.Empty;
//        foreach (var cell in typeList)
//        {
//            if (cell.IsArray && lastCellName != null && cell.Name.Equals(lastCellName))
//            {
//                continue;
//            }
//
//            GUILayout.BeginHorizontal();
//            var cellLabelName = cell.Name;
//            if (cell.IsArray)
//            {
//                cellLabelName += " [Array]";
//            }
//
//            cell.Enabled = EditorGUILayout.ToggleLeft(cellLabelName, cell.Enabled);
//            cell.ValType = (ValueType) EditorGUILayout.EnumPopup(cell.ValType, GUILayout.MaxWidth(100));
//            GUILayout.EndHorizontal();
//            lastCellName = cell.Name;
//        }
//        EditorGUILayout.EndVertical();
//        EditorGUILayout.EndScrollView();
//
//        if (GUILayout.Button("Generate"))
//        {
//            EditorPrefs.SetString($"{editorPrefKeyPrefix}{fileName}.className", className);
//            ExportEntity();
//            ExportImporter();
//
//            AssetDatabase.ImportAsset(filePath);
//            AssetDatabase.Refresh(ImportAssetOptions.ForceUpdate);
//            Close();
//        }
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

        [MenuItem("Assets/XLS Import Settings...")]
        private static void ExportExcelToAsset()
        {
            var sheetList = new List<ExcelSheetInfo>();
            foreach (var obj in Selection.objects)
            {

            }
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
                    var parts = sheetName.Split(new[] {'-'}, StringSplitOptions.RemoveEmptyEntries);
                    sheetInfo.Name = parts[0];
                    sheetInfo.SubName = parts.Length > 1 ? parts[1] : null;
                    sheetList.Add(sheetInfo);
                    if (!sheetInfo.NameValid || !sheetInfo.SubNameValid)
                    {
                        Debug.LogError($"Sheet Name is invalid! [{sheetName}]\neg. Name-Sub, Name, _Name-Sub, _Name");
                    }

                    if (!sheetInfo.Enabled) continue;

                    var titleRow = sheet.GetRow(0); // 1st row [FieldName]
                    var typeRow = sheet.GetRow(1); // 2nd row [TypeSymbol]
                    for (var j = 0; j < titleRow.LastCellNum; j++)
                    {
                        var colInfo = new ExcelColumnInfo();
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

//    private void ExportEntity()
//    {
//        var templateFilePath = (separateSheet)
//            ? "Assets/JokeMaker/ExcelImporterGenerator/Editor/EntityTemplate2.txt"
//            : "Assets/JokeMaker/ExcelImporterGenerator/Editor/EntityTemplate.txt";
//        var entityTemplate = File.ReadAllText(templateFilePath);
//        entityTemplate = entityTemplate.Replace("\r\n", "\n");
//        if (!entityTemplate.EndsWith("\n")) entityTemplate += '\n';
//        var builder = new StringBuilder();
//        bool isInbetweenArray = false;
//        foreach (var row in typeList)
//        {
//            if (row.Enabled)
//            {
//                if (!row.IsArray)
//                {
//                    builder.AppendLine();
//                    builder.AppendFormat("		public {0} {1};", row.ValType.ToString().ToLower(), row.Name);
//                }
//                else
//                {
//                    if (!isInbetweenArray)
//                    {
//                        builder.AppendLine();
//                        builder.AppendFormat("        public {0}[] {1};", row.ValType.ToString().ToLower(), row.Name);
//                    }
//
//                    isInbetweenArray = (row.NextArrayItem != null);
//                }
//            }
//        }
//
//        entityTemplate = entityTemplate.Replace("$Types$", builder.ToString());
//        entityTemplate = entityTemplate.Replace("$ExcelData$", className);
//
//        Directory.CreateDirectory("Assets/JokeMaker/ExcelImporterGenerator/Classes/");
//        File.WriteAllText("Assets/JokeMaker/ExcelImporterGenerator/Classes/" + className + ".cs", entityTemplate);
//    }
//
//    void ExportImporter()
//    {
//        string templateFilePath = (separateSheet)
//            ? "Assets/JokeMaker/ExcelImporterGenerator/Editor/ExportTemplate2.txt"
//            : "Assets/JokeMaker/ExcelImporterGenerator/Editor/ExportTemplate.txt";
//
//        string importerTemplate = File.ReadAllText(templateFilePath);
//
//        StringBuilder builder = new StringBuilder();
//        StringBuilder sheetListbuilder = new StringBuilder();
//        int rowCount = 0;
//        string indent = "                    ";
//        bool isInbetweenArray = false;
//
//        //public string[] sheetNames = {"hoge", "fuga"};
//        //$SheetList$
//        foreach (ExcelSheetInfo sht in sheetList)
//        {
//            if (sht.Enabled)
//            {
//                sheetListbuilder.Append("\"" + sht.Name + "\",");
//            }
//
//            /*
//            if (sht != sheetList [sheetList.Count - 1])
//            {
//                sheetListbuilder.Append(",");
//            }
//            */
//        }
//
//        foreach (ExcelColumnInfo row in typeList)
//        {
//            if (row.Enabled)
//            {
//                if (!row.IsArray)
//                {
//                    builder.AppendLine();
//                    switch (row.ValType)
//                    {
//                        case ValueType.STRING:
//                            builder.AppendFormat(
//                                indent +
//                                "cell = row.GetCell({1}); p.{0} = (cell == null ? string.Empty : cell.StringCellValue);",
//                                row.Name, rowCount);
//                            break;
//                        case ValueType.BOOL:
//                            builder.AppendFormat(
//                                indent +
//                                "cell = row.GetCell({1}); p.{0} = (cell == null ? false : cell.BooleanCellValue);",
//                                row.Name, rowCount);
//                            break;
//                        case ValueType.INT:
//                            builder.AppendFormat(
//                                indent +
//                                "cell = row.GetCell({1}); p.{0} = (cell == null ? 0 : (int) cell.NumericCellValue);",
//                                row.Name, rowCount);
//                            break;
//                        case ValueType.LONG:
//                            builder.AppendFormat(
//                                indent +
//                                "cell = row.GetCell({1}); p.{0} = (cell == null ? 0L : (long) cell.NumericCellValue);",
//                                row.Name, rowCount);
//                            break;
//                        case ValueType.FLOAT:
//                            builder.AppendFormat(
//                                indent +
//                                "cell = row.GetCell({1}); p.{0} = (cell == null ? 0f : (float) cell.NumericCellValue);",
//                                row.Name, rowCount);
//                            break;
//                        case ValueType.DOUBLE:
//                            builder.AppendFormat(
//                                indent +
//                                "cell = row.GetCell({1}); p.{0} = (cell == null ? 0.0 : cell.NumericCellValue);",
//                                row.Name, rowCount);
//                            break;
//                    }
//                }
//                else
//                {
//                    // only the head of array should generate code
//
//                    if (!isInbetweenArray)
//                    {
//                        int arrayLength = 0;
//                        for (ExcelColumnInfo r = row; r != null; r = r.NextArrayItem, ++arrayLength)
//                        {
//                        }
//
//                        builder.AppendLine();
//                        switch (row.ValType)
//                        {
//                            case ValueType.STRING:
//                                builder.AppendFormat(indent + "p.{0} = new string[{1}];", row.Name, arrayLength);
//                                break;
//                            case ValueType.BOOL:
//                                builder.AppendFormat(indent + "p.{0} = new bool[{1}];", row.Name, arrayLength);
//                                break;
//                            case ValueType.INT:
//                                builder.AppendFormat(indent + "p.{0} = new int[{1}];", row.Name, arrayLength);
//                                break;
//                            case ValueType.LONG:
//                                builder.AppendFormat(indent + "p.{0} = new long[{1}];", row.Name, arrayLength);
//                                break;
//                            case ValueType.FLOAT:
//                                builder.AppendFormat(indent + "p.{0} = new float[{1}];", row.Name, arrayLength);
//                                break;
//                            case ValueType.DOUBLE:
//                                builder.AppendFormat(indent + "p.{0} = new double[{1}];", row.Name, arrayLength);
//                                break;
//                        }
//
//                        for (var i = 0; i < arrayLength; ++i)
//                        {
//                            builder.AppendLine();
//                            switch (row.ValType)
//                            {
//                                case ValueType.STRING:
//                                    builder.AppendFormat(
//                                        indent +
//                                        "cell = row.GetCell({1}); p.{0}[{2}] = (cell == null ? string.Empty : cell.StringCellValue);",
//                                        row.Name, rowCount + i, i);
//                                    break;
//                                case ValueType.BOOL:
//                                    builder.AppendFormat(
//                                        indent +
//                                        "cell = row.GetCell({1}); p.{0}[{2}] = (cell == null ? false : cell.BooleanCellValue);",
//                                        row.Name, rowCount + i, i);
//                                    break;
//                                case ValueType.INT:
//                                    builder.AppendFormat(
//                                        indent +
//                                        "cell = row.GetCell({1}); p.{0}[{2}] = (cell == null ? 0 : (int) cell.NumericCellValue);",
//                                        row.Name, rowCount + i, i);
//                                    break;
//                                case ValueType.LONG:
//                                    builder.AppendFormat(
//                                        indent +
//                                        "cell = row.GetCell({1}); p.{0}[{2}] = (cell == null ? 0L : (long) cell.NumericCellValue);",
//                                        row.Name, rowCount + i, i);
//                                    break;
//                                case ValueType.FLOAT:
//                                    builder.AppendFormat(
//                                        indent +
//                                        "cell = row.GetCell({1}); p.{0}[{2}] = (cell == null ? 0f : (float) cell.NumericCellValue);",
//                                        row.Name, rowCount + i, i);
//                                    break;
//                                case ValueType.DOUBLE:
//                                    builder.AppendFormat(
//                                        indent +
//                                        "cell = row.GetCell({1}); p.{0}[{2}] = (cell == null ? 0.0 : cell.NumericCellValue);",
//                                        row.Name, rowCount + i, i);
//                                    break;
//                            }
//                        }
//                    }
//
//                    isInbetweenArray = (row.NextArrayItem != null);
//                }
//            }
//
//            rowCount += 1;
//        }
//
//        importerTemplate = importerTemplate.Replace("$IMPORT_PATH$", filePath);
//        importerTemplate = importerTemplate.Replace("$ExportAssetDirectry$", Path.GetDirectoryName(filePath));
//        importerTemplate = importerTemplate.Replace("$EXPORT_PATH$", Path.ChangeExtension(filePath, ".asset"));
//        importerTemplate = importerTemplate.Replace("$ExcelData$", className);
//        importerTemplate = importerTemplate.Replace("$SheetList$", sheetListbuilder.ToString());
//        importerTemplate = importerTemplate.Replace("$EXPORT_DATA$", builder.ToString());
//        importerTemplate = importerTemplate.Replace("$ExportTemplate$", fileName + "_importer");
//
//        Directory.CreateDirectory("Assets/JokeMaker/ExcelImporterGenerator/Classes/Editor/");
//        File.WriteAllText("Assets/JokeMaker/ExcelImporterGenerator/Classes/Editor/" + fileName + "_importer.cs",
//            importerTemplate);
//    }

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

            public bool Enabled => NameValid && SubNameValid && !Name.StartsWith("_");
            public bool NameValid => Regex.IsMatch(Name, @"^_{0,1}([A-Z][a-z]*?)+?$");

            public bool SubNameValid =>
                string.IsNullOrEmpty(SubName) || Regex.IsMatch(SubName, @"^([A-Z][a-z]*?)+?[0-9]{0,3}$");
        }

        private class ExcelColumnInfo
        {
            public string Name;
            public ValueType ValType;
            public bool IsArray;

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
