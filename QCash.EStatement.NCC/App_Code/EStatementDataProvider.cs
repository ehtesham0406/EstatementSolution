using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Connection;
using System.Common;
using System.Data;
using Common;

namespace StatementGenerator.App_Code
{
    public class EStatementDataProvider : IEStatement
    {
        private ConnectionStringBuilder ConStr = null;
        private SqlDbProvider objProvider = null;

        private EStatementDataProvider() 
        {
 
        }

        public static EStatementDataProvider Instance() 
        { 
            return new EStatementDataProvider(); 
        }

        #region IEStatement Members

        public EStatementList GetAllEStatements(string bankcode, string stdate, string status, ref string reply)
        {
            EStatementList objEstList = null;
           ConStr = new ConnectionStringBuilder(1);
           objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
           string s = "";
           DataSet ds = new DataSet();
            /////////////////***************Previous*********************///////////////////
            //if(status=="1")
            //    ds = objProvider.ReturnData("(select * from email_notification where bank_code='" + bankcode + "' and status='" + status + "' and stmdate='" + stdate + "') "+
            //" union (select * from email_notification where bank_code='" + bankcode + "' and status='" + status + "' and stmdate='" + stdate + "')", ref reply);
            //else if (status == "2")
            //    ds = objProvider.ReturnData("(select * from email_notification where bank_code='" + bankcode + "' and stmdate='" + stdate + "')"+
            //       " union (select * from email_notification_arc where bank_code='" + bankcode + "' and stmdate='" + stdate + "')", ref reply);
            //else if (status == "3")
            //    ds = objProvider.ReturnData("select * from email_notification where bank_code='" + bankcode + "' and stmdate='" + stdate + "'", ref reply);
           ///////////////////*******************************************/////////////////////

           //With Date Format
           //if (status == "1")
           //{
           //    s = "(select * from email_notification where bank_code='" + bankcode + "' and status='" + status + "' and convert(date,stmdate,103)=convert(date,'" + stdate + "',103)) " +
           //        " union (select * from email_notification where bank_code='" + bankcode + "' and status='" + status + "' and convert(date,stmdate,103)=convert(date,'" + stdate + "',103))";
           //    ds = objProvider.ReturnData(s, ref reply);
           //}
           //else if (status == "2")
           //    ds = objProvider.ReturnData("(select * from email_notification where bank_code='" + bankcode + "' and convert(date,stmdate,103)=convert(date,'" + stdate + "',103))" +
           //                                " union (select * from email_notification_arc where bank_code='" + bankcode + "' and convert(date,stmdate,103)=convert(date,'" + stdate + "',103))", ref reply);
           //else if (status == "3")
           //    ds = objProvider.ReturnData("select * from email_notification where bank_code='" + bankcode + "' and convert(date,stmdate,103)=convert(date,'" + stdate + "',103)", ref reply);

           //// Without Date Format
           if (status == "1")
           {
               s = "(select * from email_notification where bank_code='" + bankcode + "' and status='" + status + "' and stmdate='" + stdate + "') " +
                   " union (select * from email_notification where bank_code='" + bankcode + "' and status='" + status + "' and stmdate='" + stdate + "')";
               ds = objProvider.ReturnData(s, ref reply);
           }
           else if (status == "2")
               ds = objProvider.ReturnData("(select * from email_notification where bank_code='" + bankcode + "' and stmdate='" + stdate + "')" +
                                           " union (select * from email_notification_arc where bank_code='" + bankcode + "' and stmdate='" + stdate + "')", ref reply);
           else if (status == "3")
               ds = objProvider.ReturnData("select * from email_notification where bank_code='" + bankcode + "' and stmdate='" + stdate + "'", ref reply);
            if (ds != null) 
           {
               if (ds.Tables.Count > 0) 
               {
                   if (ds.Tables[0].Rows.Count > 0)
                   {
                       objEstList = new EStatementList();

                       for (int i = 0; i < ds.Tables[0].Rows.Count; i++) 
                       {
                           EStatementInfo objEst = new EStatementInfo();
                           objEst.BANK_CODE = ds.Tables[0].Rows[i]["BANK_CODE"].ToString();
                           objEst.STMDATE = ds.Tables[0].Rows[i]["STMDATE"].ToString();
                           objEst.MONTH = ds.Tables[0].Rows[i]["STM_MONTH"].ToString();
                           switch (objEst.MONTH)
                           {
                               case "01":
                                   objEst.MONTH = "January";
                                   break;
                               case "02":
                                   objEst.MONTH = "February";
                                   break;
                               case "03":
                                   objEst.MONTH = "March";
                                   break;
                               case "04":
                                   objEst.MONTH = "April";
                                   break;
                               case "05":
                                   objEst.MONTH = "May";
                                   break;
                               case "06":
                                   objEst.MONTH = "June";
                                   break;
                               case "07":
                                   objEst.MONTH = "July";
                                   break;
                               case "08":
                                   objEst.MONTH = "August";
                                   break;
                               case "09":
                                   objEst.MONTH = "September";
                                   break;
                               case "10":
                                   objEst.MONTH = "October";
                                   break;
                               case "11":
                                   objEst.MONTH = "November";
                                   break;
                               case "12":
                                   objEst.MONTH = "December";
                                   break;
                           }

                           objEst.YEAR = ds.Tables[0].Rows[i]["STM_YEAR"].ToString();
                           objEst.PAN_NUMBER = ds.Tables[0].Rows[i]["PAN_NUMBER"].ToString();
                           objEst.MAILADDRESS = ds.Tables[0].Rows[i]["MAILADDRESS"].ToString();
                           objEst.FILE_LOCATION = ds.Tables[0].Rows[i]["FILE_LOCATION"].ToString();
                           objEst.MAILSUBJECT = ds.Tables[0].Rows[i]["MAILSUBJECT"].ToString();
                           objEst.MAILBODY = ds.Tables[0].Rows[i]["MAILBODY"].ToString();
                           objEst.STATUS = ds.Tables[0].Rows[i]["STATUS"].ToString();
                           if (objEst.STATUS == "1")
                           {
                               objEst.STATUS = "Statement Generated";
                           }
                           else if (objEst.STATUS == "0")
                           {
                               objEst.STATUS = "Mail Sent Successfully";
                           }
                           else if (objEst.STATUS == "2")
                           {
                               objEst.STATUS = "Mail is not Sent";
                           }

                           else if (objEst.STATUS == "8")
                           {
                               objEst.STATUS = "No Mail Address Found";
                           }


                           objEstList.Add(objEst);
                       }
                       return objEstList;
                   }
                   else
                       return null;
               }
               else
                   return null;
           }
           else
               return null;
        }

        public bool AlreadyProcessedEStatements(string bankcode, string stdate, string pan, ref string reply)
        {
            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);

            DataSet ds = new DataSet();
            ds = objProvider.ReturnData("select * from email_notification where bank_code='" + bankcode + "' and pan_number='" + pan + "' and stmdate='" + stdate + "' ", ref reply);
            if (ds != null)
            {
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        return true;
                    }
                    else
                        return false;
                }
                else
                    return false;
            }
            else
                return false;
        }

        public string AddEStatement(EStatementInfo objEst, ref string reply)
        {
            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
            string _reply = string.Empty;
            _reply = objProvider.RunQuery("insert into email_notification values('" + objEst.BANK_CODE + "','" + objEst.PAN_NUMBER + "','" + objEst.STMDATE + "','" + objEst.MONTH + "','" + objEst.YEAR + "','" + objEst.FILE_LOCATION + "','" + objEst.MAILADDRESS + "','" + objEst.MAILSUBJECT + "','" + objEst.MAILBODY + "','" + objEst.STATUS + "')");
            return _reply;
        }

        public string UpdateEStatement(EStatementInfo objEst, ref string reply)
        {
            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
            string _reply = string.Empty;
            _reply = objProvider.RunQuery("update email_notification set  status='" + objEst.STATUS + "' where bank_code='" + objEst.BANK_CODE + "' and pan_number='" + objEst.PAN_NUMBER + "' and stmdate='" + objEst.STMDATE + "'");
            return _reply;
        }

        #endregion
        //
        public string ArchiveEStatement(ref string reply)
        {
            int qStatus = 0;
            string _reply = string.Empty;
            try
            {
                ConStr = new ConnectionStringBuilder(1);
                SPExecute objProvider = new SPExecute(ConStr.ConnectionString_DBConfig);
                
                qStatus = objProvider.ExecuteNonQuery("sp_ArchievePreviousData", null);

            }
            catch (Exception ex)
            {
                reply = "Error: " + ex.Message;
            }
            if (qStatus >= 0)
                reply = "Success";

            return _reply;
        }
    }
}
