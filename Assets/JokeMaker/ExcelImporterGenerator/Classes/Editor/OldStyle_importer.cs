using System.Collections;
using System.IO;
using System.Xml.Serialization;
using UnityEditor;
using UnityEngine;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

public class OldStyle_importer : AssetPostprocessor {
    private static readonly string filePath = "Assets/ExcelData/OldStyle.xlsx";
    private static readonly string exportPath = "Assets/ExcelData/OldStyle.asset";
    private static readonly string[] sheetNames = { "Item", };
    
    static void OnPostprocessAllAssets (string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
    {
        foreach (string asset in importedAssets) {
            if (!filePath.Equals (asset))
                continue;
                
            Entity_Item data = (Entity_Item)AssetDatabase.LoadAssetAtPath (exportPath, typeof(Entity_Item));
            if (data == null) {
                data = ScriptableObject.CreateInstance<Entity_Item> ();
                AssetDatabase.CreateAsset ((ScriptableObject)data, exportPath);
                data.hideFlags = HideFlags.NotEditable;
            }
            
            data.sheets.Clear ();
            using (FileStream stream = File.Open (filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                IWorkbook book = null;
                if (Path.GetExtension (filePath) == ".xls") {
                    book = new HSSFWorkbook(stream);
                } else {
                    book = new XSSFWorkbook(stream);
                }
                
                foreach(string sheetName in sheetNames) {
                    ISheet sheet = book.GetSheet(sheetName);
                    if( sheet == null ) {
                        Debug.LogError("[QuestData] sheet not found:" + sheetName);
                        continue;
                    }

                    Entity_Item.Sheet s = new Entity_Item.Sheet ();
                    s.name = sheetName;
                
                    for (int i=1; i<= sheet.LastRowNum; i++) {
                        IRow row = sheet.GetRow (i);
                        ICell cell = null;
                        
                        Entity_Item.Param p = new Entity_Item.Param ();
                        
                    cell = row.GetCell(0); p.ID = (cell == null ? 0 : (int) cell.NumericCellValue);
                    cell = row.GetCell(1); p.string_data = (cell == null ? string.Empty : cell.StringCellValue);
                    cell = row.GetCell(2); p.int_data = (cell == null ? 0 : (int) cell.NumericCellValue);
                    cell = row.GetCell(3); p.double_data = (cell == null ? 0.0 : cell.NumericCellValue);
                    cell = row.GetCell(4); p.bool_data = (cell == null ? false : cell.BooleanCellValue);
                    cell = row.GetCell(5); p.math_1 = (cell == null ? 0f : (float) cell.NumericCellValue);
                    p.array = new long[2];
                    cell = row.GetCell(6); p.array[0] = (cell == null ? 0L : (long) cell.NumericCellValue);
                    cell = row.GetCell(7); p.array[1] = (cell == null ? 0L : (long) cell.NumericCellValue);
                        s.list.Add (p);
                    }
                    data.sheets.Add(s);
                }
            }

            ScriptableObject obj = AssetDatabase.LoadAssetAtPath (exportPath, typeof(ScriptableObject)) as ScriptableObject;
            EditorUtility.SetDirty (obj);
        }
    }
}
