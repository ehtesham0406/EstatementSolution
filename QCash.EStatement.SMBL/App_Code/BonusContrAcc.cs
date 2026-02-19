using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator
{
    public class BonusContrAcc
    {
        private string _STATEMENTNO;

        public string STATEMENTNO
        {
            get { return _STATEMENTNO; }
            set { _STATEMENTNO = value; }
        }
        private string _SUM_CREDIT;

        public string SUM_CREDIT
        {
            get { return _SUM_CREDIT; }
            set { _SUM_CREDIT = value; }
        }
        private string _SUM_DEBIT;

        public string SUM_DEBIT
        {
            get { return _SUM_DEBIT; }
            set { _SUM_DEBIT = value; }
        }
        private string _EBALANCE;

        public string EBALANCE
        {
            get { return _EBALANCE; }
            set { _EBALANCE = value; }
        }
        private string _ACCOUNT_NO;

        public string ACCOUNT_NO
        {
            get { return _ACCOUNT_NO; }
            set { _ACCOUNT_NO = value; }
        }
        private string _ACURN;

        public string ACURN
        {
            get { return _ACURN; }
            set { _ACURN = value; }
        }
        private string _ACURC;

        public string ACURC
        {
            get { return _ACURC; }
            set { _ACURC = value; }
        }
        private string _SBALANCE;

        public string SBALANCE
        {
            get { return _SBALANCE; }
            set { _SBALANCE = value; }
        }
    }
    public class BonusContrAccList : List<BonusContrAcc>
    { }
}
