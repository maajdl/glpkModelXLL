using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using linqXl;
using glpkXl;

namespace glpkModelXLL
{
    public class scenario
    {
        private string _wbPath;
        private string _ws;
        private string _list;
        private string _dest;

        public string workbook      { get { return _wbPath; } set { if (value == null) _wbPath = ""; else _wbPath = value.Trim(); } }
        public string worksheet     { get { return _ws; } set { if (value == null) _ws = ""; else _ws = value.Trim(); } }
        public string list { get    { return _list; } set { if (value == null) _list = ""; else _list = value.Trim(); } }
        public string destination   { get { return _dest; } set { if (value == null) _dest = ""; else _dest = value.Trim(); } }
    }

    public static class scenariosSheet
    {
        private static List<scenario> scenarios(this Workbook wb)
        {
            try
            {
                List<scenario> scenarios = wb.WorksheetList<scenario>("scenarios", "scenariosTable");
                scenarios = scenarios.Where(s => s.workbook != "").Where(s => s.worksheet != "").Where(s => s.list != "").Where(s => s.destination != "").ToList();
                return scenarios;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool hasScenarios(this Workbook wb) => wb.scenarios() != null;

        public static void refreshScenarios(this Workbook wb)
        {
            glpkExcelAddIn.disableSheetChange();
                if ( wb.hasScenarios())
                {
                    var sceDest = wb.scenarios().GroupBy(s => s.destination);
                    foreach (var s in sceDest) refreshScenarios(wb, s.ToList(), s.Key);
                }
                if (!wb.hasScenarios())
                {
                    createScenariosSheet(wb);
                }
            glpkExcelAddIn.enableSheetChange();
        }

        private static void refreshScenarios(this Workbook wb, List<scenario> scenarios, string destination)
        {
            List<glpkColumn> columns = new List<glpkColumn>();
            try
            {
                foreach (scenario s in scenarios)
                {
                    try
                    {
                        Workbook wb1 = glpkExcelAddIn.Application.Workbooks.Open(s.workbook);
                        List<glpkColumn> columns1 = wb1.WorksheetList<glpkColumn>(s.worksheet, s.list);
                        if (wb1!=wb) wb1.Close(false);
                        columns = columns.Union(columns1).ToList();
                    }
                    catch { /* Ignored */ }
                }
                wb.Activate();
                if (columns.Count > 0) columns.write(wb, destination, destination);
            }
            catch { /* Ignored */ }
        }

        private static void createScenariosSheet(this Workbook wb)
        {
            DialogResult yn = MessageBox.Show("Do you want to create a scenario sheet?", "Scenario sheet missing.", MessageBoxButtons.YesNo);
            if (yn != DialogResult.Yes) return;
            try
            {
                List<scenario> scenarios = new List<scenario>();
                //scenario scenario = new scenario() { workbook = @"C:\folder\aa\scenario1.xls", worksheet = @"worksheet", list = "list", destination = "scenarioList" };
                scenario scenario = new scenario() { workbook = wb.FullName, worksheet = @"worksheet", list = "list", destination = "scenarioList" };
                scenarios.Add(scenario);
                scenario = new scenario();
                for (int i = 0; i < 40; i++) scenarios.Add(scenario);
                Worksheet ws = scenarios.write(wb, "scenarios", "scenariosTable");
            }
            catch { /* Ignored */ }
        }
    }
}
