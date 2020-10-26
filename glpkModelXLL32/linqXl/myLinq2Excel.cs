using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace linqXl
{
    public static class myLinq2Excel
    {
        public class maplet
        {
            public PropertyInfo property;
            public Type convertType;
            public int columnIndex;
            public maplet(PropertyInfo p, int i)
            {
                property = p;
                columnIndex = i;
                convertType = Nullable.GetUnderlyingType(property.PropertyType);
                if (convertType == null) convertType = property.PropertyType;
            }
        }

        public static List<T> WorksheetList<T>(this Excel.Workbook wb, string worksheetName, string listName)
        {
            return myLinq2Excel.WorksheetList<T>(wb.Worksheets[worksheetName], listName);
        }

        public static List<T> WorksheetList<T>(Excel.Worksheet sheet, string listName)
        {
            var table = sheet.ListObjects[listName];
            var columns =
                from index in Enumerable.Range(1, table.HeaderRowRange.Columns.Count)
                select new { columnIndex = index, columnName = table.HeaderRowRange[1, index].Value };
            var map = (
                            from property in typeof(T).GetProperties()
                            where property.CanWrite
                            from column in columns
                            where property.Name == column.columnName
                            select new maplet(property, column.columnIndex)
                        ).ToList();
            var ls = (
                            from r in Enumerable.Range(1, table.ListRows.Count)
                            select (T)createInstance<T>(map, table.ListRows[r].Range.Value2)
                        ).ToList();
            return ls;
        }

        public static T createInstance<T>(List<maplet> map, object[,] v)
        {
            T instance = Activator.CreateInstance<T>();
            foreach (maplet m in map)
            {
                if (v[1, m.columnIndex] != null)
                {
                    m.property.SetValue(instance, Convert.ChangeType(v[1, m.columnIndex], m.convertType));
                }
                else
                {
                    if (m.convertType.Name == "String") m.property.SetValue(instance, "");
                }
            }
            return instance;
        }
    }
}
