using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator.App_Code
{
    public class EStatementManager
    {
        private EStatementManager() 
        {
 
        }

        public static EStatementManager Instance() 
        {
            return new EStatementManager(); 
        }

        #region IEStatement Members

        public EStatementList GetAllEStatements(string bankcode, string startdate, string enddate, string status, ref string reply)
        {
            return EStatementDataProvider.Instance().GetAllEStatements(bankcode, startdate,enddate, status, ref reply);
        }
       

        public string AddEStatement(EStatementInfo objESt, ref string reply)
        {
            return EStatementDataProvider.Instance().AddEStatement(objESt, ref reply);
        }

        public string UpdateEStatement(EStatementInfo objESt, ref string reply)
        {
            return EStatementDataProvider.Instance().UpdateEStatement(objESt, ref reply);
        }
        public string ArchiveEStatement(ref string reply)
        {
            return EStatementDataProvider.Instance().ArchiveEStatement(ref reply);
        }
        #endregion
    }
}
