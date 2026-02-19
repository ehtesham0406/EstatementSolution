using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator
{
    public class Account
    {
        private string _STATEMENTNO;

        public string STATEMENTNO
        {
            get { return _STATEMENTNO; }
            set { _STATEMENTNO = value; }
        }
        private string _ACCOUNTNO;

        public string ACCOUNTNO
        {
            get { return _ACCOUNTNO; }
            set { _ACCOUNTNO = value; }
        }
        private string _ACURN;

        public string ACURN
        {
            get { return _ACURN; }
            set { _ACURN = value; }
        }
        private string _SBALANCE;

        public string SBALANCE
        {
            get { return _SBALANCE; }
            set { _SBALANCE = value; }
        }
        private string _ACURC;

        public string ACURC
        {
            get { return _ACURC; }
            set { _ACURC = value; }
        }
        private string _EBALANCE;

        public string EBALANCE
        {
            get { return _EBALANCE; }
            set { _EBALANCE = value; }
        }
        private string _AVAIL_CRD_LIMIT;

        public string AVAIL_CRD_LIMIT
        {
            get { return _AVAIL_CRD_LIMIT; }
            set { _AVAIL_CRD_LIMIT = value; }
        }
        private string _AVAIL_CASH_LIMIT;

        public string AVAIL_CASH_LIMIT
        {
            get { return _AVAIL_CASH_LIMIT; }
            set { _AVAIL_CASH_LIMIT = value; }
        }
        private string _SUM_WITHDRAWAL;

        public string SUM_WITHDRAWAL
        {
            get { return _SUM_WITHDRAWAL; }
            set { _SUM_WITHDRAWAL = value; }
        }
        private string _SUM_INTEREST;

        public string SUM_INTEREST
        {
            get { return _SUM_INTEREST; }
            set { _SUM_INTEREST = value; }
        }
        private string _OVLFEE_AMOUNT;

        public string OVLFEE_AMOUNT
        {
            get { return _OVLFEE_AMOUNT; }
            set { _OVLFEE_AMOUNT = value; }
        }
        private string _OVDFEE_AMOUNT;

        public string OVDFEE_AMOUNT
        {
            get { return _OVDFEE_AMOUNT; }
            set { _OVDFEE_AMOUNT = value; }
        }
        private string _SUM_REVERSE;

        public string SUM_REVERSE
        {
            get { return _SUM_REVERSE; }
            set { _SUM_REVERSE = value; }
        }
        private string _SUM_CREDIT;

        public string SUM_CREDIT
        {
            get { return _SUM_CREDIT; }
            set { _SUM_CREDIT = value; }
        }
        private string _SUM_OTHER;

        public string SUM_OTHER
        {
            get { return _SUM_OTHER; }
            set { _SUM_OTHER = value; }
        }
        private string _SUM_PURCHASE;

        public string SUM_PURCHASE
        {
            get { return _SUM_PURCHASE; }
            set { _SUM_PURCHASE = value; }
        }
        private string _MIN_AMOUNT_DUE;

        public string MIN_AMOUNT_DUE
        {
            get { return _MIN_AMOUNT_DUE; }
            set { _MIN_AMOUNT_DUE = value; }
        }
        private string _CASH_LIMIT;

        public string CASH_LIMIT
        {
            get { return _CASH_LIMIT; }
            set { _CASH_LIMIT = value; }
        }
        private string _CRD_LIMIT;

        public string CRD_LIMIT
        {
            get { return _CRD_LIMIT; }
            set { _CRD_LIMIT = value; }
        }
    }
    public class AccountList : List<Account>
    { }
}
