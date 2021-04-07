using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class Entity_Item : ScriptableObject
{
    public List<Sheet> sheets = new List<Sheet>();

    [Serializable]
    public class Sheet
    {
        public string name = string.Empty;
        public List<Param> list = new List<Param>();
    }

    [Serializable]
    public class Param
    {
        
		public int ID;
		public string string_data;
		public int int_data;
		public double double_data;
		public bool bool_data;
		public float math_1;
		public long[] array;
    }
}
