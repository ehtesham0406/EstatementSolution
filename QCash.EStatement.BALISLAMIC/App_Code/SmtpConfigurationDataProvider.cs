using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Connection;
using Common;
using System.Data;

namespace StatementGenerator.App_Code
{
    public class SmtpConfigurationDataProvider : ISmtpConfiguration
    {
        public static SmtpConfigurationDataProvider Instance()
        {
            return new SmtpConfigurationDataProvider();
        }

        #region ISmtpConfiguration Members

        public string SaveSmtpConfiguration(SmtpConfigurationInfo objSmtpConfig)
        {
            ConnectionStringBuilder objConStr = new ConnectionStringBuilder(1);
            SPExecute objSqlPro = new SPExecute(objConStr.ConnectionString_DBConfig);
            int reply = objSqlPro.ExecuteNonQuery("sp_AddSmtpConfiguration", objSmtpConfig.FIID, objSmtpConfig.Smtp_Server, objSmtpConfig.Smtp_Port,
                objSmtpConfig.EnableSSL, objSmtpConfig.From_Address, objSmtpConfig.From_User, objSmtpConfig.From_Password, objSmtpConfig.Status);
            //if (reply > 0)
                return "Success";
        }

        public SmtpConfigurationList GetSmtpConfiguration(string Fid, int status)
        {
            ConnectionStringBuilder objConStr = new ConnectionStringBuilder(1);
            SPExecute objSqlPro = new SPExecute(objConStr.ConnectionString_DBConfig);

            DataSet ds = objSqlPro.ExecuteDataset("sp_GetSmtpConfiguration", Fid);

            if(ds!=null)
                if (ds.Tables.Count > 0)
                {
                    SmtpConfigurationList objSmtpList = new SmtpConfigurationList();
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            SmtpConfigurationInfo objSmtpInfo = new SmtpConfigurationInfo();
                            objSmtpInfo.FIID = ds.Tables[0].Rows[i]["FIID"].ToString();
                            objSmtpInfo.Smtp_Server = ds.Tables[0].Rows[i]["Smtp_Server"].ToString();
                            objSmtpInfo.Smtp_Port = Convert.ToInt32(ds.Tables[0].Rows[i]["Smtp_Port"].ToString());
                            objSmtpInfo.EnableSSL = Convert.ToInt32(ds.Tables[0].Rows[i]["EnableSSL"]);
                            objSmtpInfo.From_Address = ds.Tables[0].Rows[i]["From_Address"].ToString();
                            objSmtpInfo.From_User = ds.Tables[0].Rows[i]["From_User"].ToString();
                            objSmtpInfo.From_Password = ds.Tables[0].Rows[i]["From_Password"].ToString();
                            objSmtpInfo.Status = Convert.ToInt32(ds.Tables[0].Rows[i]["Status"].ToString());
                            objSmtpList.Add(objSmtpInfo);
                        }
                    }
                    return objSmtpList;

                }
                return null;
        }

        public string UpdateSmtpConfiguration(SmtpConfigurationInfo objSmtpConfig)
        {
            ConnectionStringBuilder objConStr = new ConnectionStringBuilder(1);
            SPExecute objSqlPro = new SPExecute(objConStr.ConnectionString_DBConfig);
            int reply = objSqlPro.ExecuteNonQuery("sp_UpdateSmtpConfiguration", objSmtpConfig.FIID, objSmtpConfig.Smtp_Server, objSmtpConfig.Smtp_Port,
                objSmtpConfig.EnableSSL, objSmtpConfig.From_Address, objSmtpConfig.From_User, objSmtpConfig.From_Password, objSmtpConfig.Status);
            //if (reply > 0)
            return "Success";
        }

        #endregion
    }
}
