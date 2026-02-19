using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;

namespace System.Common
{
    public class OledbProvider
    {
        private OleDbConnection OleDbCn;
        private OleDbCommand cmd;
        private OleDbDataAdapter OracleDa;

        private string ConStr = null;

        public OledbProvider(string _ConStr)
        {
            ConStr = _ConStr;
        }
                
        public string RunQuery(string sqlString)
        {
            try
            {
                OleDbCn = new OleDbConnection(ConStr);
                OleDbCn.Open();
                
                cmd = new OleDbCommand(sqlString, OleDbCn);
                cmd.ExecuteNonQuery();

                return "Success";
            }
            catch (Exception ex)
            {
                return "Error: "+ex.Message;
            }
            finally
            {
                OleDbCn.Close();
                OleDbCn.Dispose();
                cmd.Dispose();
            }
        }

        public System.Data.DataSet ReturnData(string sqlString, ref string replymsg)
        {
            System.Data.DataSet dt = new System.Data.DataSet();
            try
            {
                OleDbCn = new OleDbConnection(ConStr);
                OleDbCn.Open();

                cmd = new OleDbCommand(sqlString, OleDbCn);
                OracleDa = new OleDbDataAdapter(cmd);
                OracleDa.Fill(dt);
                replymsg = "Success";
                return dt;
            }
            catch (Exception ex)
            {
                replymsg = "Error: " + ex.Message;
                if (OleDbCn.State == System.Data.ConnectionState.Open)
                {
                    OleDbCn.Dispose();
                    cmd.Dispose();
                    OracleDa.Dispose();
                }
                return dt = null;
            }
            finally
            {
                if (OleDbCn.State == System.Data.ConnectionState.Open)
                {
                    OleDbCn.Close();
                    OleDbCn.Dispose();
                    cmd.Dispose();
                    OracleDa.Dispose();
                }
            }

        }
    }
}
