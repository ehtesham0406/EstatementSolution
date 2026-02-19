using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator
{
    public class StatementInfo
    {
        private string _STATEMENTID;

        public string STATEMENTID
        {
            get { return _STATEMENTID; }
            set { _STATEMENTID = value; }
        }
        private string _BANK_CODE;

        public string BANK_CODE
        {
            get { return _BANK_CODE; }
            set { _BANK_CODE = value; }
        }

        private string _CONTRACTNO;

        public string CONTRACTNO
        {
            get { return _CONTRACTNO; }
            set { _CONTRACTNO = value; }
        }

        private string _IDCLIENT;

        public string IDCLIENT
        {
            get { return _IDCLIENT; }
            set { _IDCLIENT = value; }
        }

        private string _PAN;

        public string PAN
        {
            get { return _PAN; }
            set { _PAN = value; }
        }
        private string _TITLE;

        public string TITLE
        {
            get { return _TITLE; }
            set { _TITLE = value; }
        }
        private string _CLIENTNAME;

        public string CLIENTNAME
        {
            get { return _CLIENTNAME; }
            set { _CLIENTNAME = value; }
        }
        private string _STATEMENTNO;

        public string STATEMENTNO
        {
            get { return _STATEMENTNO; }
            set { _STATEMENTNO = value; }
        }
        private string _ADDRESS;

        public string ADDRESS
        {
            get { return _ADDRESS; }
            set { _ADDRESS = value; }
        }
        private string _CITY;

        public string CITY
        {
            get { return _CITY; }
            set { _CITY = value; }
        }
        private string _ZIP;

        public string ZIP
        {
            get { return _ZIP; }
            set { _ZIP = value; }
        }
        private string _COUNTRY;

        public string COUNTRY
        {
            get { return _COUNTRY; }
            set { _COUNTRY = value; }
        }
        private string _EMAIL;

        public string EMAIL
        {
            get { return _EMAIL; }
            set { _EMAIL = value; }
        }
        private string _MOBILE;

        public string MOBILE
        {
            get { return _MOBILE; }
            set { _MOBILE = value; }
        }

        private string _STARTDATE;
        public string STARTDATE
        {
            get { return _STARTDATE; }
            set { _STARTDATE = value; }
        }

        private string _ENDDATE;
        public string ENDDATE
        {
            get { return _ENDDATE; }
            set { _ENDDATE = value; }
        }
        private string _NEXT_STATEMENT_DATE;

        public string NEXT_STATEMENT_DATE
        {
            get { return _NEXT_STATEMENT_DATE; }
            set { _NEXT_STATEMENT_DATE = value; }
        }
        private string _PAYMENT_DATE;

        public string PAYMENT_DATE
        {
            get { return _PAYMENT_DATE; }
            set { _PAYMENT_DATE = value; }
        }
        private string _STATEMENT_DATE;

        public string STATEMENT_DATE
        {
            get { return _STATEMENT_DATE; }
            set { _STATEMENT_DATE = value; }
        }
        private string _STATEMENT_PERIOD;

        public string STATEMENT_PERIOD
        {
            get { return _STATEMENT_PERIOD; }
            set { _STATEMENT_PERIOD = value; }
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
        private string _STATUS;

        public string STATUS
        {
            get { return _STATUS; }
            set { _STATUS = value; }
        }
        private string _STM_MSG;

        public string STM_MSG
        {
            get { return _STM_MSG; }
            set { _STM_MSG = value; }
        }

        private string _JOBTITLE;

        public string JOBTITLE
        {
            get { return _JOBTITLE; }
            set { _JOBTITLE = value; }
        }
    }
    public class StatementInfoList : List<StatementInfo>
    { }
}
