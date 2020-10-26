using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using linqXl;

namespace glpkXl
{
    public static class glpkxlSolver
    {
        public static string tempFolder;
        public static string scenario;
        public static Workbook _wb;

        public static int solve(Workbook wb)
        {
            _wb = wb;
            tempFolder = Path.Combine(Path.GetTempPath(), "glpkxlSolver");
            Directory.CreateDirectory(tempFolder);
            int retStatus = -1;
            glpkxlMessages.initialize();
            glpkxlMessages.log(wb, "glpkxlSolver started\n");
            glpkxlMessages.log(wb, DateTime.Now.ToString("yyyy MM dd  hh:mm:ss"));
            try
            {
                createMOD(wb);
                createDAT(wb);
                glpkxlMessages.log(wb, "Solve problem");
                glpkSolver solver = new glpkSolver();
                retStatus = solver.solve(modFullName(), datFullName(),lpFullName());
                if (retStatus == 0)
                {
                    glpkxlMessages.log(wb, "Problem solved");
                    solver.status.write(wb, "status", "glpkStatus");
                    List<modelLine> statements = modelSheet.modelStatements(wb);
                    foreach (modelLine s in statements)
                    {
                        if (s.element().ToLower() == "var" && s.worksheet != "" && s.values != "")
                        {
                            glpkxlMessages.log(wb, "Write " + s.worksheet + " / " + s.values);
                            var list = solver.columns.Where(c => c.name == s.name()).ToList();
                            list.ForEach(t => t.scenario = scenario);
                            list.write(wb, s.worksheet, s.values);
                        }
                        if (s.element().ToLower() == "variables" && s.worksheet != "" && s.values != "")
                        {
                            glpkxlMessages.log(wb, "Write " + s.worksheet + " / " + s.values);
                            var list = solver.columns.ToList();
                            list.ForEach(t => t.scenario = scenario);
                            list.write(wb, s.worksheet, s.values);
                        }
                        if (s.element().ToLower() == "s.t." && s.worksheet != "" && s.values != "")
                        {
                            glpkxlMessages.log(wb, "Write " + s.worksheet + " / " + s.values);
                            var list = solver.rows.Where(c => c.name == s.name()).ToList();
                            list.ForEach(t => t.scenario = scenario);
                            list.write(wb, s.worksheet, s.values);
                        }
                        if (s.element().ToLower() == "constraints" && s.worksheet != "" && s.values != "")
                        {
                            glpkxlMessages.log(wb, "Write " + s.worksheet + " / " + s.values);
                            var list = solver.rows.ToList();
                            list.ForEach(t => t.scenario = scenario);
                            list.write(wb, s.worksheet, s.values);
                        }
                    }
                    glpkxlMessages.log(wb, "Elapsed: " + stopWatch.secondsNow("G3") + " s");
                    glpkxlMessages.log(wb, "Solve successful");
                    bool refreshAutom = wb.Names.Item("refreshAutom").RefersTo.Contains("TRUE");
                    if (refreshAutom) wb.RefreshAll();
                    if (refreshAutom) glpkxlMessages.log(wb, "Workbook refreshed");
                }
                else
                {
                    glpkxlMessages.log(wb, "retStatus = " + retStatus);
                    glpkxlMessages.log(wb, "Solve failed");
                }
            }
            catch (Exception e) { glpkxlMessages.log(e); }
            glpkxlMessages.replace(tempFolder + "\\", "");
            glpkxlMessages.messages.write(wb, "log", "log");
            return retStatus;
        }

        public static string xlsName()     { return _wb.Name; }
        public static string modFullName() { return tempFolder + "\\" + xlsName() + ".mod"; }
        public static string datFullName() { return tempFolder + "\\" + xlsName() + ".dat"; }
        public static string lpFullName()  { return tempFolder + "\\" + xlsName() + ".lp"; }

        private static void createMOD(Workbook wb)
        {
            glpkxlMessages.log(wb, "create mod file");
            List<modelLine> statements = wb.modelStatements();
            string day = DateTime.Now.ToLongDateString();
            string tim = DateTime.Now.ToLongTimeString();
            string model = $"/* Created by glpkXl, {day} {tim} */\n\n";
            foreach (modelLine s in statements)
            {
                string elem = s.element();
                if      (elem == "scenario")        { model = model + "#" + s.statement + "\n"; scenario = s.name(); }
                else if (elem == "variables")       { model = model + "#" + s.statement + "\n"; }
                else if (elem == "constraints")     { model = model + "#" + s.statement + "\n"; }
                else if (elem == "set")             { model = model + s.statement + "\n"; }
                else if (elem == "param")           { model = model + s.statement + "\n"; }
                else if (elem == "var")             { model = model + s.statement + "\n"; }
                else if (elem == "s.t.")            { model = model + s.statement + "\n"; }
                else if (elem == "subject to")      { model = model + s.statement + "\n"; }
                else if (elem == "minimize")        { model = model + s.statement + "\n"; }
                else if (elem == "maximize")        { model = model + s.statement + "\n"; }
                else if (elem == "solve")           { model = model + s.statement + "\n"; }
                else if (elem == "check")           { model = model + s.statement + "\n"; }
                else if (elem == "display")         { model = model + s.statement + "\n"; }
                else if (elem == "printf")          { model = model + s.statement + "\n"; }
                else if (elem == "for")             { model = model + s.statement + "\n"; }
                else if (elem == "table")           { model = model + s.statement + "\n"; }
                else                                { model = model + s.statement + "\n"; }
           }
            model = model + "end;\n";
            var file = new StreamWriter(modFullName());
            file.Write(model);
            file.Close();
        }

        private static void createDAT(Workbook wb)
        {
            glpkxlMessages.log(wb, "create dat file");
            List<modelLine> statements = wb.modelStatements();
            StringBuilder data = new StringBuilder();
            string day = DateTime.Now.ToLongDateString();
            string tim = DateTime.Now.ToLongTimeString();
            data.Append($"/* Created by glpkXl, {day} {tim} */\n\n");
            data.Append("\n");
            data.Append("data;\n");
            data.Append("\n");
            foreach (modelLine s in statements)
            {
                if (s.element().ToLower() == "set")       dataItem(data, s);
                if (s.element().ToLower() == "param")     dataItem(data, s);
            }
            data.Append("end;\n");
            var file = new StreamWriter(datFullName());
            file.Write(data);
            file.Close();
        }

        private static void dataItem(StringBuilder data, modelLine s)
        {
            if (s.values == "") return;
            try
            {
                dynamic values = null;
                dynamic indexRowValues = null;
                dynamic indexColValues = null;
                dynamic obj;
                if (s.worksheet != "") obj = _wb.Worksheets[s.worksheet];
                else obj = _wb.Application; 
                if (s.values    != "") values           = obj.Range[s.values.Trim()].Value2;
                if (s.indexRows != "") indexRowValues   = obj.Range[s.indexRows].Value2;
                if (s.indexCols != "") indexColValues   = obj.Range[s.indexCols].Value2;
                dataItem(data, s.element(), s.name(), values, indexRowValues, indexColValues);
            }
            catch { /* Ignored */ }
        }

        private static void dataItem(StringBuilder data, string element, string name, dynamic[,] values, dynamic[,] indexRowValues, dynamic[,] indexColValues)
        {
            try
            {
                data.Append($"{element} {name} :=\n");
                if (indexRowValues != null && indexColValues != null) dataItemRowColVal(data, values, indexRowValues, indexColValues);
                if (indexRowValues != null && indexColValues == null) dataItemRowVal(data, values, indexRowValues);
                if (indexRowValues == null && indexColValues != null) dataItemColVal(data, values, indexColValues);
                if (indexRowValues == null && indexColValues == null) dataItemVal(data, values);
                data.Append(";\n\n");
            }
            catch { /* Ignored */ }
        }

        private static void dataItemRowColVal(StringBuilder data, dynamic[,] values, dynamic[,] indexRowValues, dynamic[,] indexColValues)
        {
            try
            {
                for (int r = 1; r <= values.GetLength(0); r++)
                {
                    for (int c = 1; c <= values.GetLength(1); c++)
                    {
                        string keys = "";
                        for (int ci = 1; ci <= indexRowValues.GetLength(1); ci++) keys = keys + indexRowValues[r, ci] + ",";
                        for (int ri = 1; ri <= indexColValues.GetLength(0); ri++) keys = keys + indexColValues[ri, c] + ",";
                        keys = keys.Trim(',');
                        keys = "[" + keys + "]";
                        bool keysOk = keys.Split('[', ',', ']', ' ').Select(k => k.Length).Max() > 0;
                        if (keysOk) data.Append(keys + values[r, c] + " ");
                    }
                    data.Append("\n");
                }
            }
            catch { /* Ignored */ }
        }

        private static void dataItemRowVal(StringBuilder data, dynamic[,] values, dynamic[,] indexRowValues)
        {
            try
            {
                for (int r = 1; r <= values.GetLength(0); r++)
                {
                    for (int c = 1; c <= values.GetLength(1); c++)
                    {
                        string keys = "";
                        for (int ci = 1; ci <= indexRowValues.GetLength(1); ci++) keys = keys + indexRowValues[r, ci] + ",";
                        keys = keys.Trim(',');
                        keys = "[" + keys + "]";
                        bool keysOk = keys.Split('[', ',', ']', ' ').Select(k => k.Length).Max() > 0;
                        if (keysOk) data.Append(keys + values[r, c] + " ");
                    }
                    data.Append("\n");
                }
            }
            catch { /* Ignored */ }
        }

        private static void dataItemColVal(StringBuilder data, dynamic[,] values, dynamic[,] indexColValues)
        {
            try
            {
                for (int r = 1; r <= values.GetLength(0); r++)
                {
                    for (int c = 1; c <= values.GetLength(1); c++)
                    {
                        string keys = "";
                        for (int ri = 1; ri <= indexColValues.GetLength(0); ri++) keys = keys + indexColValues[ri, c] + ",";
                        keys = keys.Trim(',');
                        keys = "[" + keys + "]";
                        bool keysOk = keys.Split('[', ',', ']', ' ').Select(k => k.Length).Max() > 0;
                        if (keysOk) data.Append(keys + values[r, c] + " ");
                    }
                    data.Append("\n");
                }
                data.Append(";\n\n");
            }
            catch { /* Ignored */ }
        }

        private static void dataItemVal(StringBuilder data, dynamic[,] values)
        {
            try
            {
                for (int r = 1; r <= values.GetLength(0); r++)
                {
                    string vals = "";
                    for (int c = 1; c <= values.GetLength(1); c++) vals = vals + values[r, c] + " ";
                    bool valsOk = vals.Split(' ').Select(k => k.Length).Max() > 0;
                    if (valsOk) data.Append(vals + "\n");
                }
            }
            catch { /* Ignored */ }
        }
    }
}
