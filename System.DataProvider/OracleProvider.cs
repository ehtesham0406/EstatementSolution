using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

namespace System.Common
{
    public class OracleProvider
    {
        private OracleConnection OracleCn;
        private OracleCommand cmd;
        private OracleDataAdapter OracleDa;
        private OracleTransaction objTrans;

        private string ConStr = null;

        public OracleProvider(string _ConStr)
        {
            ConStr = _ConStr;
        }
                
        public string RunQuery(string sqlString)
        {
            try
            {
                OracleCn = new OracleConnection(ConStr);
                OracleCn.Open();
                
                objTrans = OracleCn.BeginTransaction();
                cmd = new OracleCommand(sqlString, OracleCn);
                cmd.Transaction = objTrans;

                cmd.ExecuteNonQuery();
                objTrans.Commit();

                return "Success";
            }
            catch (Exception ex)
            {
                objTrans.Rollback();
                return "Error: "+ex.Message;
            }
            finally
            {
                OracleCn.Close();
                OracleCn.Dispose();
                cmd.Dispose();
                objTrans.Dispose();
            }
        }

        public System.Data.DataSet ReturnData(string sqlString, ref string replymsg)
        {
            System.Data.DataSet dt = new System.Data.DataSet();
            try
            {
                OracleCn = new OracleConnection(ConStr);
                OracleCn.Open();

                cmd = new OracleCommand(sqlString, OracleCn);
                OracleDa = new OracleDataAdapter(cmd);
                OracleDa.Fill(dt);
                replymsg = "Success";
                return dt;
            }
            catch (Exception ex)
            {
                replymsg = "Error: " + ex.Message;
                if (OracleCn.State == System.Data.ConnectionState.Open)
                {
                    OracleCn.Dispose();
                    cmd.Dispose();
                    OracleDa.Dispose();
                }
                return dt = null;
            }
            finally
            {
                if (OracleCn.State == System.Data.ConnectionState.Open)
                {
                    OracleCn.Close();
                    OracleCn.Dispose();
                    cmd.Dispose();
                    OracleDa.Dispose();
                }
            }

        }
    }
}
