using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QCash.EStatement.NBL.App_Code
{
    public class StatementRegisterInfo
    {
       public string StartPage { get; set; }
       public string EndPage { get; set; }
       public string TotalPage { get; set; }
       public string ClientID { get; set; }
       public string ClientName{get;set;}
       public string StatementNo { get; set; }
       public string PAN { get; set; }
       public string FileName { get; set; }
       public string SL { get; set; }
       public string StatementDate { get; set; }
       public string RefNo { get; set; }
       public string Address { get; set; }
       public string CONTRACTNO { get; set; }
    }

    public class StatementRegisterList : List<StatementRegisterInfo> { }
}
