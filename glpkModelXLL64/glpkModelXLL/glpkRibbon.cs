using ExcelDna.Integration.CustomUI;
using System.Diagnostics;
using System.Runtime.InteropServices;
using glpkXl;

// ReSharper disable UnusedMember.Global

namespace glpkModelXLL
{
    [ComVisible(true)]
    public class glpkRibbon : ExcelRibbon
    {
        public string   runGroupLabel(IRibbonControl group)
        {
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            return glpkExcelAddIn.ActiveWorkbook.Name;
        }

        public void     runBtn_Click(IRibbonControl ribbon)
        {
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            glpkExcelAddIn.solve();
        }

        public void     refreshBtn_Click(IRibbonControl ribbon)
        {
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            glpkExcelAddIn.refresh();
        }

        public void     modBtn_Click(IRibbonControl ribbon)
        {
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            Process.Start(glpkxlSolver.modFullName());
        }

        public void     datBtn_Click(IRibbonControl ribbon)
        {
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            Process.Start(glpkxlSolver.datFullName());
        }

        public void     lpBtn_Click(IRibbonControl  ribbon)
        {
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            Process.Start(glpkxlSolver.lpFullName());
        }

        public bool autom_getPressed(IRibbonControl checkBox)
        {
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            try { return glpkExcelAddIn.ActiveWorkbook.Names.Item(checkBox.Id).RefersTo.Contains("TRUE"); }
            catch { return false; }
        }

        public void     autom_Click(IRibbonControl checkBox, bool pressed)
        {
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            glpkExcelAddIn.ActiveWorkbook.Names.Add(checkBox.Id, pressed);
        }

        public void     refreshScenarios_Click(IRibbonControl ribbon)
        {
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            glpkExcelAddIn.ActiveWorkbook.refreshScenarios();
        }
    }
}
