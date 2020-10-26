using org.gnu.glpk;
using System.Collections.Generic;
// ReSharper disable AutoPropertyCanBeMadeGetOnly.Global
// ReSharper disable UnusedAutoPropertyAccessor.Global

namespace glpkXl
{
    public class glpkRow
    {
        public static string[] row_type = { "", "GLP_FR", "GLP_LO", "GLP_UP", "GLP_DB", "GLP_FX" };
        public static string[] row_stat = { "", "GLP_BS", "GLP_NL", "GLP_NU", "GLP_NF", "GLP_NS" };
        public static string[] row_kind = { "", "GLP_CV", "GLP_IV", "GLP_BV" };

        public string   scenario{ get; set; }
        public string   name    { get; set; }
        public string   index1  { get; set; }
        public string   index2  { get; set; }
        public string   index3  { get; set; }
        public string   index4  { get; set; }
        public string   index5  { get; set; }
        public string   index6  { get; set; }
        public string   index7  { get; set; }
        public string   index8  { get; set; }
        public string   index9  { get; set; }
        public string   type    { get; set; }
        public double   lb      { get; set; }
        public double   ub      { get; set; }
        public string   stat    { get; set; }
        public double   prim    { get; set; }
        public double   dual    { get; set; }
        public string   kind    { get; set; }
        public double   val     { get; set; }

        public glpkRow(glp_prob lp, int r)
        {
            char[] toTrim = { ' ', '\'' };
            string row_name = GLPK.glp_get_row_name(lp, r);
            string[] parsed = row_name.Split('[', ',', ']');
            name = parsed[0];
            try
            {
                index1 = parsed[1].Trim(toTrim);
                index2 = parsed[2].Trim(toTrim);
                index3 = parsed[3].Trim(toTrim);
                index4 = parsed[4].Trim(toTrim);
                index5 = parsed[5].Trim(toTrim);
                index6 = parsed[6].Trim(toTrim);
                index7 = parsed[7].Trim(toTrim);
                index8 = parsed[8].Trim(toTrim);
                index9 = parsed[9].Trim(toTrim);
            }
            catch { /* Ignored */ /* will stop when nothing left to parse */ }
            type    = row_type[GLPK.glp_get_row_type(lp, r)];
            lb      = GLPK.glp_get_row_lb(lp, r); lb = System.Math.Max(lb, -1e30);
            ub      = GLPK.glp_get_row_ub(lp, r); ub = System.Math.Min(ub,  1e30);
            stat    = row_stat[GLPK.glp_get_row_stat(lp, r)];
            prim    = GLPK.glp_get_row_prim(lp, r);
            dual    = GLPK.glp_get_row_dual(lp, r);
            kind    = row_kind[1];
            val     = GLPK.glp_mip_row_val(lp, r);
        }
    }

    public class glpkRows : List<glpkRow>
    {
        public glpkRows(glp_prob lp)
        {
            int nrow = GLPK.glp_get_num_rows(lp);
            for (int r = 1; r <= nrow; r++)
            {
                Add(new glpkRow(lp, r));
            }
        }
    }
}
