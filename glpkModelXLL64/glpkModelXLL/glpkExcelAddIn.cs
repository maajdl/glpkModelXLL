using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using glpkXl;

namespace glpkModelXLL
{
    // ReSharper disable once ClassNeverInstantiated.Global
    public class glpkExcelAddIn : IExcelAddIn
    {
        public static Application Application => (Application)ExcelDnaUtil.Application;
        public static Workbook ActiveWorkbook => Application.ActiveWorkbook;

        public void AutoOpen()  { enableSheetChange(); }
        public void AutoClose() { disableSheetChange(); }

        public static void enableSheetChange() { Application.SheetChange += Wb_SheetChange; }
        public static void disableSheetChange() { Application.SheetChange -= Wb_SheetChange; }

        static void Wb_SheetChange(object Sh, Range Target)
        {
            Worksheet ws = (Worksheet)Sh;
            Workbook wb = ws.Parent;
            bool solveAutom = false;
            try { solveAutom = wb.Names.Item("solveAutom").RefersTo.Contains("TRUE");}
            catch { }
            if (solveAutom) solve();
        }

        public static void solve()
        {
            var ocult = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            try
            {
                disableSheetChange();
                Workbook wb = ActiveWorkbook;
                if (wb.hasGlpkModel())
                {
                    glpkxlSolver.solve(wb);
                }
                if (!wb.hasGlpkModel())
                {
                    wb.createModelsSheet();
                }
                enableSheetChange();
            }
            catch
            { /* Ignored */ }
            System.Threading.Thread.CurrentThread.CurrentCulture = ocult;
        }

        public static void refresh()
        {
            disableSheetChange();
            ActiveWorkbook.RefreshAll();
            enableSheetChange();
        }
    }
}
