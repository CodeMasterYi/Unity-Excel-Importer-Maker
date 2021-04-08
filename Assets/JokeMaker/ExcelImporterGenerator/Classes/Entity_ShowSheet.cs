using System;
using System.Collections.Generic;
using UnityEngine;

namespace JokeMaker.Entity
{
    public class Entity_ShowSheet : ScriptableObject
    {
        public string ExcelName;
        public string SheetName;
        public List<Data> param = new List<Data>();

        [Serializable]
        public class Data
        {
            public int ID;
            public string Desc;
            public double[] Num;
            public string[] XXX;
        }
    }
}
