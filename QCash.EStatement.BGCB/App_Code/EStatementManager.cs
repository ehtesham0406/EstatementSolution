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

        public EStatementList GetAllEStatements(string bankcode, string stdate, string status, ref string reply)
        {
            return EStatementDataProvider.Instance().GetAllEStatements(bankcode, stdate, status, ref reply);
        }
        public bool AlreadyProcessedEStatements(string bankcode, string stdate, string pan, ref string reply)
        {
            return EStatementDataProvider.Instance().AlreadyProcessedEStatements(bankcode, stdate, pan, ref reply);
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
