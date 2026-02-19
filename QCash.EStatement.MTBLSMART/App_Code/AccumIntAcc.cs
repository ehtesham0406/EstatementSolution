using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator
{
    public class AccumIntAcc
    {
        private string _STATEMENTNO;

        public string STATEMENTNO
        {
            get { return _STATEMENTNO; }
            set { _STATEMENTNO = value; }
        }

        private string _ACCUM_INT_RRELEASE;

        public string ACCUM_INT_RRELEASE
        {
            get { return _ACCUM_INT_RRELEASE; }
            set { _ACCUM_INT_RRELEASE = value; }
        }

        private string _ACCUM_INT_EBALANCE;

        public string ACCUM_INT_EBALANCE
        {
            get { return _ACCUM_INT_EBALANCE; }
            set { _ACCUM_INT_EBALANCE = value; }
        }

        private string _ACCUM_INT_SBALANCE;

        public string ACCUM_INT_SBALANCE
        {
            get { return _ACCUM_INT_SBALANCE; }
            set { _ACCUM_INT_SBALANCE = value; }
        }

        private string _ACCUM_INT_AMOUNT;

        public string ACCUM_INT_AMOUNT
        {
            get { return _ACCUM_INT_AMOUNT; }
            set { _ACCUM_INT_AMOUNT = value; }
        }

        private string _ACCOUNT_NO;

        public string ACCOUNT_NO
        {
            get { return _ACCOUNT_NO; }
            set { _ACCOUNT_NO = value; }
        }
        private string _AutoID;

        public string AutoID
        {
            get { return _AutoID; }
            set { _AutoID = value; }
        }
        
    }
    public class AccumIntAccList : List<AccumIntAcc>
    { }
}
