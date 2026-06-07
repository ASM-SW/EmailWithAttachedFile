// Copyright © 2016-2019  ASM-SW
//asmeyers@outlook.com  https://github.com/asm-sw

using System.Reflection;
using System.Text;

namespace EmailWithAttachedFile
{
    public static class Extensions
    {
        public static string ToCSV<T>(this IEnumerable<T> items)
        {
            StringBuilder result = new();
            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            // Header
            for (int i = 0; i < props.Length; i++)
            {
                result.AppendFormat("\"{0}\"", props[i].Name);
                result.Append(i == props.Length - 1 ? "\n" : ",");
            }

            // Data
            foreach (T item in items)
            {
                for (int i = 0; i < props.Length; i++)
                {
                    result.AppendFormat("\"{0}\"", props[i].GetValue(item, null));
                    result.Append(i == props.Length - 1 ? "\n" : ",");
                }
            }

            return result.ToString();
        }

    }
}
