using org.gnu.glpk;
using System;
// ReSharper disable CollectionNeverUpdated.Global

namespace glpkXl
{
    public class glpkSolver : IGlpkTerminalListener
    {
        public glpkColumns      columns;
        public glpkRows         rows;
        public glpkStatusList   status;

        public int solve(String modelFile, String dataFile, String lpFile)
        {
            glpkxlMessages.log("GLPK version = " + GLPK.glp_version());
            GLPK.glp_cli_set_numeric_locale("C");
            GlpkTerminal.addListener(this);
            glp_prob lp = GLPK.glp_create_prob();
            glp_tran tran = GLPK.glp_mpl_alloc_wksp();

            int retStatus = -1;
            try
            {
                if (GLPK.glp_mpl_read_model(tran, modelFile, 0) != 0)   throw new ApplicationException("Model file not valid: ");
                if (GLPK.glp_mpl_read_data(tran, dataFile) != 0)        throw new ApplicationException("Data file not valid: ");
                if (GLPK.glp_mpl_generate(tran, null)       != 0)       throw new ApplicationException("Cannot generate model: ");
                GLPK.glp_mpl_build_prob(tran, lp);

                glp_iocp iocp = new glp_iocp();
                GLPK.glp_init_iocp(iocp);
                iocp.presolve = GLPK.GLP_ON;

                retStatus = GLPK.glp_intopt(lp, iocp);
                if (retStatus == 0)
                {
                    GLPK.glp_mpl_postsolve(tran, lp, GLPK.GLP_MIP);
                    status = new glpkStatusList(lp);
                    columns = new glpkColumns(lp);
                    rows = new glpkRows(lp);
                }
                GLPK.glp_write_lp(lp, null, lpFile);
            }
            catch (Exception e)
            {
                glpkxlMessages.log(e.Message);
            }

            GLPK.glp_mpl_free_wksp(tran);
            GLPK.glp_delete_prob(lp); 
            GlpkTerminal.removeListener(this);
            return retStatus;
        }

        string outputBuffer = "";
        public bool output(string str)
        {
            outputBuffer = outputBuffer + str;
            if (outputBuffer.EndsWith("\n"))
            {
                outputBuffer = outputBuffer.Substring(0, outputBuffer.Length - 1);
                glpkxlMessages.log(outputBuffer);
                outputBuffer = "";
            }
            return false;
        }
    }
}