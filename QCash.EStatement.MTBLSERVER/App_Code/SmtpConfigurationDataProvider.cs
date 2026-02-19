using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Connection;
using Common;
using System.Data;
//using Oracle.DataAccess.Client;
using ITOSystem.DatabaseEngine.AccessPoint;
using UUL.UDBH.SQLData;
using System.Data.OracleClient;
using FlexiStar.Utilities;

namespace StatementGenerator.App_Code
{
    public class SmtpConfigurationDataProvider : ISmtpConfiguration
    {

        #region Variable & Instance Declaration
        private DataHandler dbHelper = new DataHandler();
        private DataHandler dataHandler = new DataHandler();
        private DataHandler dbHandel = new DataHandler();
        #endregion

        public static SmtpConfigurationDataProvider Instance()
        {
            return new SmtpConfigurationDataProvider();
        }

        #region ISmtpConfiguration Members

     /*   public string SaveSmtpConfiguration(SmtpConfigurationInfo objSmtpConfig)
        {
            ConnectionStringBuilder objConStr = new ConnectionStringBuilder(1);
            SPExecute objSqlPro = new SPExecute(objConStr.ConnectionString_DBConfig);
            int reply = objSqlPro.ExecuteNonQuery("sp_AddSmtpConfiguration", objSmtpConfig.FIID, objSmtpConfig.Smtp_Server, objSmtpConfig.Smtp_Port,
                objSmtpConfig.EnableSSL, objSmtpConfig.From_Address, objSmtpConfig.From_User, objSmtpConfig.From_Password, objSmtpConfig.Status);
            //if (reply > 0)
                return "Success";
        }  */


        public string SaveSmtpConfiguration(SmtpConfigurationInfo objSmtpConfig)
        {
            //SmtpConfigurationInfo objSmtpConfig=new SmtpConfigurationInfo();
            int sq_Type = 0;
            // sq_Type = (executeType == ExecuteType.INSERT ? 1 : 2);
            OracleCommand cmd = new OracleCommand();
            //cmd.Parameters.Add(new OracleParameter("pExecuteType", 2));
            cmd.Parameters.Add(new OracleParameter("pfid", objSmtpConfig.FIID));
            cmd.Parameters.Add(new OracleParameter("pSmtp_Server", objSmtpConfig.Smtp_Server));
            cmd.Parameters.Add(new OracleParameter("pSmtp_Port", objSmtpConfig.Smtp_Port));
            cmd.Parameters.Add(new OracleParameter("pEnableSSL", objSmtpConfig.EnableSSL));
            cmd.Parameters.Add(new OracleParameter("pFrom_Address", objSmtpConfig.From_Address));
            cmd.Parameters.Add(new OracleParameter("pFrom_User", objSmtpConfig.From_User));
            cmd.Parameters.Add(new OracleParameter("pFrom_Password", objSmtpConfig.From_Password));
            cmd.Parameters.Add(new OracleParameter("pStatus", objSmtpConfig.Status));


            try
            {
                sq_Type = dataHandler.DBOracleManupulation.SPExecute("sp_AddSmtpConfiguration", cmd);
            }

            catch (Exception ex)
            {
                throw ex;

            }

            if (sq_Type != 0)
            {
                return "Success";
            }
            else
            {
                return "unSuccessfull";
            }
        } 




     /*   public SmtpConfigurationList GetSmtpConfiguration(string Fid, int status)
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
        } */




        public SmtpConfigurationList GetSmtpConfiguration(string Fid, int status)
        {
            DataTable dt = new DataTable();
            

                    OracleCommand cmd = new OracleCommand();
                    OracleParameter pOUTTABLE = new OracleParameter("pOUTTABLE", OracleType.Cursor);
                    pOUTTABLE.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(new OracleParameter("fid", OracleType.VarChar)).Value = Fid;
                   
                    cmd.Parameters.Add(pOUTTABLE);
                    dt = dataHandler.DBOracleManupulation.RecordSet(" sp_GetSmtpConfiguration", cmd);
                
                 
                        SmtpConfigurationList objSmtpList = new SmtpConfigurationList();
                        for (int rowCounter = 0; rowCounter < dt.Rows.Count;rowCounter++ )
                    
                        {
                            try
                            {
                                SmtpConfigurationInfo objSmtpInfo = new SmtpConfigurationInfo();
                                objSmtpInfo.FIID = dt.Rows[rowCounter]["FIID"].ToString();
                                objSmtpInfo.Smtp_Server = dt.Rows[rowCounter]["Smtp_Server"].ToString();

                                if (!string.IsNullOrEmpty(dt.Rows[rowCounter]["Smtp_Port"].ToString()))
                                {
                                    objSmtpInfo.Smtp_Port = Convert.ToInt32(dt.Rows[rowCounter]["Smtp_Port"].ToString());
                                }
                                if (!string.IsNullOrEmpty(dt.Rows[rowCounter]["EnableSSL"].ToString()))
                                {
                                    objSmtpInfo.EnableSSL = Convert.ToInt32(dt.Rows[rowCounter]["EnableSSL"].ToString());
                                }

                              
                              
                                objSmtpInfo.From_Address = dt.Rows[rowCounter]["From_Address"].ToString();
                                objSmtpInfo.From_User = dt.Rows[rowCounter]["From_User"].ToString();
                                objSmtpInfo.From_Password = dt.Rows[rowCounter]["From_Password"].ToString();

                                if (!string.IsNullOrEmpty(dt.Rows[rowCounter]["Status"].ToString()))
                                {
                                    objSmtpInfo.Status = Convert.ToInt32(dt.Rows[rowCounter]["Status"].ToString());
                                }

                                objSmtpList.Add(objSmtpInfo);
                            }

                            catch(Exception ex)
                            {
                                throw ex;
                            }
                        }
                    
                    return objSmtpList;

                
        }




    /*    public string UpdateSmtpConfiguration(SmtpConfigurationInfo objSmtpConfig)
        {
            ConnectionStringBuilder objConStr = new ConnectionStringBuilder(1);
            SPExecute objSqlPro = new SPExecute(objConStr.ConnectionString_DBConfig);
            int reply = objSqlPro.ExecuteNonQuery("sp_UpdateSmtpConfiguration", objSmtpConfig.FIID, objSmtpConfig.Smtp_Server, objSmtpConfig.Smtp_Port,
                objSmtpConfig.EnableSSL, objSmtpConfig.From_Address, objSmtpConfig.From_User, objSmtpConfig.From_Password, objSmtpConfig.Status);
            //if (reply > 0)
            return "Success";
        }  */

        public string UpdateSmtpConfiguration(SmtpConfigurationInfo objSmtpConfig)
        {
            //SmtpConfigurationInfo objSmtpConfig=new SmtpConfigurationInfo();
            int sq_Type = 0;
           // sq_Type = (executeType == ExecuteType.INSERT ? 1 : 2);
            OracleCommand cmd = new OracleCommand();
            cmd.Parameters.Add(new OracleParameter("pExecuteType", 2));
            cmd.Parameters.Add(new OracleParameter("pfid", objSmtpConfig.FIID));
            cmd.Parameters.Add(new OracleParameter("pSmtp_Server", objSmtpConfig.Smtp_Server));
            cmd.Parameters.Add(new OracleParameter("pSmtp_Port", objSmtpConfig.Smtp_Port));
            cmd.Parameters.Add(new OracleParameter("pEnableSSL", objSmtpConfig.EnableSSL));
            cmd.Parameters.Add(new OracleParameter("pFrom_Address", objSmtpConfig.From_Address));
            cmd.Parameters.Add(new OracleParameter("pFrom_User", objSmtpConfig.From_User));
            cmd.Parameters.Add(new OracleParameter("pFrom_Password", objSmtpConfig.From_Password));
            cmd.Parameters.Add(new OracleParameter("pStatus", objSmtpConfig.Status));


            try
            {
                sq_Type = dataHandler.DBOracleManupulation.SPExecute("sp_UpdateSmtpConfiguration", cmd);
            }

            catch( Exception ex)

            {
                throw ex;
            
            }

            if (sq_Type != 0)
            {
                return "Success";
            }
            else
            {
                return "unSuccessfull";
            }
        } 

        #endregion
    }
}
