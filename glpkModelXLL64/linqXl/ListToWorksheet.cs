using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;

namespace linqXl
{
    public static class ListToWorksheet
    {
        static Excel.Worksheet ws;
        static Excel.Range anchor;
        static Excel.ListObject lo;
        private static bool newSheet;
        private static bool newList;

        public static Excel.Worksheet write<T>(this List<T> array, Excel.Workbook wb, string sheetName, string listObjectName, bool append = false)
        {
            try                                                                 // get the worksheet or create it
            {
                ws = wb.Worksheets[sheetName];
                newSheet = false;
            }
            catch
            {
                ws = wb.Worksheets.Add(Type.Missing, wb.Sheets[wb.Sheets.Count]);
                ws.Name = sheetName;
                newSheet = true;
            }

            try                                                                 // get the list object or create it
            {
                lo = ws.ListObjects[listObjectName];
                anchor = lo.HeaderRowRange.Cells[1, 1];
                newList = false;
            }
            catch
            {
                anchor = ws.UsedRange;
                anchor = anchor.Cells[1, anchor.Columns.Count].EntireColumn;
                anchor = anchor.Cells[4, 2];
                var headers = array.getHeaders();
                anchor = anchor.Resize[Type.Missing, headers.GetLength(1)];
                anchor.Value2 = headers;
                lo = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, anchor, null, Excel.XlYesNoGuess.xlYes);
                lo.Name = listObjectName;
                lo.ShowTableStyleRowStripes = false;
                lo.TableStyle = "TableStyleMedium3";
                wb.Application.ActiveWindow.DisplayGridlines = false;
                newList = true;
            }

            if (!append)
            {
                try                                                                 // clear the list object
                {
                    if (lo.DataBodyRange != null) lo.DataBodyRange.Clear();
                    lo.Resize(lo.HeaderRowRange.Resize[2]);
                }
                catch { /* Ignored */ }
            }

            try                                                                 // copy the data to the list object
            {
                dynamic[,] data = array.getDataTable();
                int r1 = lo.ListRows.Count + 1;
                int r2 = r1 + data.GetLength(0);
                int c2 = lo.ListColumns.Count;
                Excel.Range appendRange = ws.Range[lo.Range.Cells[r1 + 1, 1], lo.Range.Cells[r2, c2]];
                appendRange.Value2 = data;
                if (lo.DataBodyRange != null) lo.DataBodyRange.WrapText = false;
            }
            catch { /* Ignored */ }

            try
            {
                if (newSheet) ws.Columns.AutoFit();
            }
            catch (Exception e) { /* Ignored */ }

            try
            {
                if (newList)
                {
                    lo.ShowTableStyleRowStripes = true;
                    lo.TableStyle = "TableStyleLight4";
                    lo.Range.Font.Color = -65536;
                }
            }
            catch (Exception e) { /* Ignored */ }

            return ws;
        }

        public static dynamic[,] getHeaders<T>(this List<T> arrays)
        {
            var properties = typeof(T).GetProperties().ToArray();
            dynamic[,] header = new dynamic[1, properties.Length];
            for (int j = 0; j < properties.Length; j++) header[0, j] = properties[j].Name;
            return header;
        }

        public static dynamic[,] getDataTable<T>(this List<T> arrays)
        {
            var properties = typeof(T).GetProperties().ToArray();
            dynamic[,] data = new dynamic[arrays.Count, properties.Length];
            for (int j = 0; j < properties.Length; j++)
            {
                for (int i = 0; i < arrays.Count; i++) data[i, j] = properties[j].GetValue(arrays[i], null);
            }
            return data;
        }
    }
}
