using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator.App_Code
{
    interface IEStatement
    {
        EStatementList GetAllEStatements(string bankcode, string stdate, string enddate, string status, ref string reply);
        string AddEStatement(EStatementInfo objESt, ref string reply);
        string UpdateEStatement(EStatementInfo objESt, ref string reply);
        //
        string ArchiveEStatement(ref string reply);
    }
}
