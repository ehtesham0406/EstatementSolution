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

        public EStatementList GetAllEStatements(string bankcode, string startdate, string enddate, string status, ref string reply)
        {
            EStatementList objEstList = null;
           ConStr = new ConnectionStringBuilder(1);
           objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);

           DataSet ds = new DataSet();
           //Without Date Format
           if (status == "1")
               ds = objProvider.ReturnData("(select * from email_notification where bank_code='" + bankcode + "' and status='" + status + "' and startdate='" + startdate + "' and enddate='" + enddate +"') " +
           " union (select * from email_notification where bank_code='" + bankcode + "' and status='" + status + "' and startdate='" + startdate + "' and enddate='" + enddate + "')", ref reply);
           else if (status == "2")
               ds = objProvider.ReturnData("(select * from email_notification where bank_code='" + bankcode + "' and startdate='" + startdate + "' and enddate='" + enddate + "')" +
                  " union (select * from email_notification_arc where bank_code='" + bankcode + "' and startdate='" + startdate + "' and enddate='" + enddate + "')", ref reply);
           else if (status == "3")
               ds = objProvider.ReturnData("select * from email_notification where bank_code='" + bankcode + "' and startdate='" + startdate + "' and enddate='" + enddate + "'", ref reply);
          

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
                           objEst.STARTDATE = ds.Tables[0].Rows[i]["STARTDATE"].ToString();
                           objEst.ENDDATE = ds.Tables[0].Rows[i]["ENDDATE"].ToString();
                           objEst.IDCLIENT = ds.Tables[0].Rows[i]["IDCLIENT"].ToString();
                           objEst.PAN = ds.Tables[0].Rows[i]["PAN"].ToString();
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

       

        public string AddEStatement(EStatementInfo objEst, ref string reply) 
        {
            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
            string _reply = string.Empty;

            _reply = objProvider.RunQuery("insert into email_notification values('" + objEst.BANK_CODE + "','" + objEst.IDCLIENT + "','" + objEst.PAN + "','" + objEst.STARTDATE + "','" + objEst.ENDDATE + "','" + objEst.FILE_LOCATION + "','" + objEst.MAILADDRESS + "','" + objEst.MAILSUBJECT + "','" + objEst.MAILBODY + "','" + objEst.STATUS + "')");
            return _reply;
        }

        public string UpdateEStatement(EStatementInfo objEst, ref string reply)
        {
            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
            string _reply = string.Empty;
            _reply = objProvider.RunQuery("update email_notification set  status='" + objEst.STATUS + "' where bank_code='" + objEst.BANK_CODE + "' and PAN='" + objEst.PAN + "' and IDCLIENT='" + objEst.IDCLIENT + "' and STARTDATE='" + objEst.STARTDATE + "' and ENDDATE='" + objEst.ENDDATE + "'");
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
