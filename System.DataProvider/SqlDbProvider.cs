using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Net;
//
namespace System.Common
{
    public class SqlDbProvider
    {
        private SqlConnection OleDbCn;
        private SqlCommand cmd;
        private SqlDataAdapter OracleDa;

        private string ConStr = null;

        public SqlDbProvider(string _ConStr)
        { // use for MDBL only
           // ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
            //ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
            ConStr = _ConStr;
        }
                
        public string RunQuery(string sqlString)
        {
            try
            {
                OleDbCn = new SqlConnection(ConStr);
                OleDbCn.Open();

                cmd = new SqlCommand(sqlString, OleDbCn);
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
                OleDbCn = new SqlConnection(ConStr);
                OleDbCn.Open();

                cmd = new SqlCommand(sqlString, OleDbCn);
                OracleDa = new SqlDataAdapter(cmd);
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
