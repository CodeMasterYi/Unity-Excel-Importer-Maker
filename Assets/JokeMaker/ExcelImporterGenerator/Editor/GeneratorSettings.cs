using System.Collections.Generic;
using UnityEngine;

public class GeneratorSettings : ScriptableObject
{
    public string EntityNamespace;
    public string ImporterNamespace;
    public string AssetOutputFolder;
    public string EntitiesFolder;
    public string ImportersFolder;
    public List<string> ExcelFiles;
}
