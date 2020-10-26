using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using linqXl;

namespace glpkXl
{
    // element  name    subscript expression  attributes worksheet   values indexes
    public class modelLine
    {
        public string _statement;

        public string _worksheet;
        public string _values;
        public string _indexRows;
        public string _indexCols;

        public string statement { get { return _statement; } set { if (value == null) _statement = ""; else _statement = value.Trim(); } }

        public string worksheet { get { return _worksheet; } set { if (value == null) _worksheet = ""; else _worksheet = value.Trim(); } }
        public string values { get { return _values; } set { if (value == null) _values = ""; else _values = value.Trim(); } }
        public string indexRows { get { return _indexRows; } set { if (value == null) _indexRows = ""; else _indexRows = value.Trim(); } }
        public string indexCols { get { return _indexCols; } set { if (value == null) _indexCols = ""; else _indexCols = value.Trim(); } }

        public string element()
        {
            string [] split = statement.Split(new char[] { ' ', ';' }, StringSplitOptions.RemoveEmptyEntries);
            if (split.Length > 0) return split[0].ToLower();
            else return "";
        }

        public string name()
        {
            string[] split = statement.Split(new char[] { ' ',';' }, StringSplitOptions.RemoveEmptyEntries);
            if (split.Length > 1) return split[1];
            else return "";
        }
    }

    public static class modelSheet
    {
        public static List<modelLine> modelStatements(this Workbook wb)
        {
            try
            {
                List<modelLine> modelStatements = wb.WorksheetList<modelLine>("model", "modelTable");
                return modelStatements;
            }
            catch (Exception)
            {
                return null;
            }
        }
        public static bool hasGlpkModel(this Workbook wb) => wb.modelStatements() != null;

        public static void createModelsSheet(this Workbook wb)
        {
            DialogResult yn = MessageBox.Show("Do you want to create a model sheet?", "Model sheet missing.", MessageBoxButtons.YesNo);
            if (yn != DialogResult.Yes) return;
            try
            {
                List<modelLine> statements = new List<modelLine>();
                modelLine statement = new modelLine() {};
                statement.statement     = "param data { c in commodities, { 'price', 'weight'} union nutriants};";
                statement.worksheet     = "data";
                statement.values        = "data[[price]:[ascorbicAcid]]";
                statement._indexRows    = "data[commodity]";
                statement._indexCols    = "data[[#Headers],[price]:[ascorbicAcid]]";
                statements.Add(statement);
                statement = new modelLine();
                for (int i = 0; i < 40; i++) statements.Add(statement);
                Worksheet ws = statements.write(wb, "model", "modelTable");
            }
            catch { /* Ignored */ }
        }
    }
}
