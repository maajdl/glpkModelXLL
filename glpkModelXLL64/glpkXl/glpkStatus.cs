using org.gnu.glpk;
using System.Collections.Generic;
// ReSharper disable UnusedAutoPropertyAccessor.Global


namespace glpkXl
{
    public class glpkStatus
    {
        public static string[] status = { "", "GLP_UNDEF", "GLP_FEAS", "GLP_INFEAS", "GLP_NOFEAS", "GLP_OPT" };
        public string name { get; set; }
        public string value { get; set; }

        // ReSharper disable once EmptyConstructor
        // empty constructor needed for reflexion
        public glpkStatus()
        {
        }
        public static glpkStatus asString(string name, int value)
        {
            glpkStatus gs = new glpkStatus();
            gs.name = name;
            gs.value = status[value];
            return gs;
        }
        public static glpkStatus asNumber(string name, int value)
        {
            glpkStatus gs = new glpkStatus();
            gs.name = name;
            gs.value = value.ToString();
            return gs;
        }
        public static glpkStatus asNumber(string name, double value)
        {
            glpkStatus gs = new glpkStatus();
            gs.name = name;
            gs.value = value.ToString();
            return gs;
        }
    }

    public class glpkStatusList : List<glpkStatus>
    {
        public glpkStatusList(glp_prob lp)
        {
            Add(glpkStatus.asString("status_simplex", GLPK.glp_get_status(lp)));
            Add(glpkStatus.asString("status_simplex_prim", GLPK.glp_get_prim_stat(lp)));
            Add(glpkStatus.asString("status_simplex_dual", GLPK.glp_get_dual_stat(lp)));
            Add(glpkStatus.asString("status_ipt", GLPK.glp_ipt_status(lp)));
            Add(glpkStatus.asString("status_mip", GLPK.glp_mip_status(lp)));

            Add(glpkStatus.asString(" ", 0));

            Add(glpkStatus.asNumber("num_rows"  , GLPK.glp_get_num_rows(lp)));
            Add(glpkStatus.asNumber("num_cols", GLPK.glp_get_num_cols(lp)));
            Add(glpkStatus.asNumber("num_nz", GLPK.glp_get_num_nz(lp)));
            Add(glpkStatus.asNumber("num_int", GLPK.glp_get_num_int(lp)));
            Add(glpkStatus.asNumber("num_bin", GLPK.glp_get_num_bin(lp)));

            Add(glpkStatus.asString(" ", 0));

            Add(glpkStatus.asNumber("obj_val_simplex", GLPK.glp_get_obj_val(lp)));
            Add(glpkStatus.asNumber("obj_val_ipt", GLPK.glp_ipt_obj_val(lp)));
            Add(glpkStatus.asNumber("obj_val_mip", GLPK.glp_mip_obj_val(lp)));
        }
    }
}
