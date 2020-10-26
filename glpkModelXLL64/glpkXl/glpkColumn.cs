using org.gnu.glpk;
using System.Collections.Generic;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable AutoPropertyCanBeMadeGetOnly.Global

namespace glpkXl
{
    public class glpkColumn
    {
        private static string[] col_type = { "", "GLP_FR", "GLP_LO", "GLP_UP", "GLP_DB", "GLP_FX" };
        private static string[] col_stat = { "", "GLP_BS", "GLP_NL", "GLP_NU", "GLP_NF", "GLP_NS" };
        private static string[] col_kind = { "", "GLP_CV", "GLP_IV", "GLP_BV" };

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

        // ReSharper disable once UnusedMember.Global
        // needed for reflexion
        public glpkColumn() { }

        public glpkColumn(glp_prob lp, int c)
        {
            char[] toTrim = { ' ', '\''};
            string col_name = GLPK.glp_get_col_name(lp, c);
            string[] parsed = col_name.Split('[', ',', ']');
            name    = parsed[0];
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
            type    = col_type[GLPK.glp_get_col_type(lp, c)];
            lb      = GLPK.glp_get_col_lb(lp, c);   lb = System.Math.Max(lb, -1e30);
            ub      = GLPK.glp_get_col_ub(lp, c);   ub = System.Math.Min(ub,  1e30);
            stat    = col_stat[GLPK.glp_get_col_stat(lp, c)];
            prim    = GLPK.glp_get_col_prim(lp, c);
            dual    = GLPK.glp_get_col_dual(lp, c);
            kind    = col_kind[GLPK.glp_get_col_kind(lp, c)];
            val     = GLPK.glp_mip_col_val(lp, c);
        }
    }

    public class glpkColumns: List<glpkColumn>
    {
        public glpkColumns(glp_prob lp)
        {
            int ncol = GLPK.glp_get_num_cols(lp);
            for (int c=1; c<=ncol;c++)
            {
                Add(new glpkColumn(lp, c));
            }
        }
    }
}
