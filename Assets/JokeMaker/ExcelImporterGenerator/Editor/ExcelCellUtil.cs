using System;
using NPOI.SS.UserModel;

namespace JokeMaker.Editor
{
    public static class ExcelCellUtil
    {
        public static bool ToString(this ICell cell, out string val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            val = cell.StringCellValue;
            return true;
        }

        public static bool ToStringArray(this ICell cell, char arrSep, out string[] val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            var s = cell.StringCellValue;
            var parts = s.Split(new[] {arrSep}, StringSplitOptions.RemoveEmptyEntries);
            val = parts;
            return true;
        }

        public static bool ToBool(this ICell cell, out bool val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            val = cell.BooleanCellValue;
            return true;
        }

        public static bool ToBoolArray(this ICell cell, char arrSep, out bool[] val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            var s = cell.StringCellValue;
            var parts = s.Split(new[] {arrSep}, StringSplitOptions.RemoveEmptyEntries);
            val = new bool[parts.Length];
            for (var i = 0; i < parts.Length; ++i)
            {
                if (bool.TryParse(parts[i], out var v))
                {
                    val[i] = v;
                }
                else
                {
                    val = default;
                    return false;
                }
            }

            return true;
        }

        public static bool ToInt(this ICell cell, out int val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            val = (int) cell.NumericCellValue;
            return true;
        }

        public static bool ToIntArray(this ICell cell, char arrSep, out int[] val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            var s = cell.StringCellValue;
            var parts = s.Split(new[] {arrSep}, StringSplitOptions.RemoveEmptyEntries);
            val = new int[parts.Length];
            for (var i = 0; i < parts.Length; ++i)
            {
                if (int.TryParse(parts[i], out var v))
                {
                    val[i] = v;
                }
                else
                {
                    val = default;
                    return false;
                }
            }

            return true;
        }

        public static bool ToLong(this ICell cell, out long val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            val = (long) cell.NumericCellValue;
            return true;
        }

        public static bool ToLongArray(this ICell cell, char arrSep, out long[] val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            var s = cell.StringCellValue;
            var parts = s.Split(new[] {arrSep}, StringSplitOptions.RemoveEmptyEntries);
            val = new long[parts.Length];
            for (var i = 0; i < parts.Length; ++i)
            {
                if (long.TryParse(parts[i], out var v))
                {
                    val[i] = v;
                }
                else
                {
                    val = default;
                    return false;
                }
            }

            return true;
        }

        public static bool ToFloat(this ICell cell, out float val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            val = (float) cell.NumericCellValue;
            return true;
        }

        public static bool ToFloatArray(this ICell cell, char arrSep, out float[] val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            var s = cell.StringCellValue;
            var parts = s.Split(new[] {arrSep}, StringSplitOptions.RemoveEmptyEntries);
            val = new float[parts.Length];
            for (var i = 0; i < parts.Length; ++i)
            {
                if (float.TryParse(parts[i], out var v))
                {
                    val[i] = v;
                }
                else
                {
                    val = default;
                    return false;
                }
            }

            return true;
        }

        public static bool ToDouble(this ICell cell, out double val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            val = cell.NumericCellValue;
            return true;
        }

        public static bool ToDoubleArray(this ICell cell, char arrSep, out double[] val)
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            var s = cell.StringCellValue;
            var parts = s.Split(new[] {arrSep}, StringSplitOptions.RemoveEmptyEntries);
            val = new double[parts.Length];
            for (var i = 0; i < parts.Length; ++i)
            {
                if (double.TryParse(parts[i], out var v))
                {
                    val[i] = v;
                }
                else
                {
                    val = default;
                    return false;
                }
            }

            return true;
        }

        public static bool ToEnum<T>(this ICell cell, out T val) where T : struct, Enum
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            var s = cell.StringCellValue;
            if (Enum.TryParse<T>(s, out var v))
            {
                val = v;
                return true;
            }
            val = default;
            return false;
        }

        public static bool ToEnumArray<T>(this ICell cell, char arrSep, out T[] val) where T : struct, Enum
        {
            if (cell == null)
            {
                val = default;
                return false;
            }

            var s = cell.StringCellValue;
            var parts = s.Split(new[] {arrSep}, StringSplitOptions.RemoveEmptyEntries);
            val = new T[parts.Length];
            for (var i = 0; i < parts.Length; ++i)
            {
                if (Enum.TryParse<T>(parts[i], out var v))
                {
                    val[i] = v;
                }
                else
                {
                    val = default;
                    return false;
                }
            }
            return true;
        }

//        public static bool ToCustomClass<T>(this ICell cell, out T val) where T : class, ICellConvert<T>, new()
//        {
//            if (cell == null)
//            {
//                val = default;
//                return false;
//            }
//
//            var s = cell.StringCellValue;
//            ICellConvert<T> t = new T();
//            if (t.TryConvert(s, out var v))
//            {
//                val = v;
//                return true;
//            }
//            val = default;
//            return false;
//        }
//
//        public static bool ToCustomClassArray<T>(this ICell cell, char arrSep, out T[] val) where T : class, ICellConvert<T>, new()
//        {
//            if (cell == null)
//            {
//                val = default;
//                return false;
//            }
//
//            var s = cell.StringCellValue;
//            var parts = s.Split(new[] {arrSep}, StringSplitOptions.RemoveEmptyEntries);
//            val = new T[parts.Length];
//            for (var i = 0; i < parts.Length; ++i)
//            {
//                ICellConvert<T> t = new T();
//                if (t.TryConvert(parts[i], out var v))
//                {
//                    val[i] = v;
//                }
//                else
//                {
//                    val = null;
//                    return false;
//                }
//            }
//            return true;
//        }
//
//        public interface ICellConvert<T> where T : ICellConvert<T>
//        {
//            bool TryConvert(string v, out T val);
//        }
//
//        [Serializable]
//        public abstract class CustomCellType : ICellConvert<CustomCellType>
//        {
//            [SerializeField]
//            private string Name;
//
//            public bool TryConvert(string v, out CustomCellType val)
//            {
//                Name = v;
//                val = this;
//                return true;
//            }
//        }
    }
}
