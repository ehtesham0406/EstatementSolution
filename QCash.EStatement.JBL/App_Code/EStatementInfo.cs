using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator.App_Code
{
    public class EStatementInfo
    {
        private string _BANK_CODE;

        public string BANK_CODE
        {
            get { return _BANK_CODE; }
            set { _BANK_CODE = value; }
        }
        private string _PAN_NUMBER;

        public string PAN_NUMBER
        {
            get { return _PAN_NUMBER; }
            set { _PAN_NUMBER = value; }
        }
        private string _STMDATE;

        public string STMDATE
        {
            get { return _STMDATE; }
            set { _STMDATE = value; }
        }
        private string _MONTH;

        public string MONTH
        {
            get { return _MONTH; }
            set { _MONTH = value; }
        }
        private string _YEAR;

        public string YEAR
        {
            get { return _YEAR; }
            set { _YEAR = value; }
        }
        private string _FILE_LOCATION;

        public string FILE_LOCATION
        {
            get { return _FILE_LOCATION; }
            set { _FILE_LOCATION = value; }
        }
        private string _MAILADDRESS;

        public string MAILADDRESS
        {
            get { return _MAILADDRESS; }
            set { _MAILADDRESS = value; }
        }
        private string _MAILSUBJECT;

        public string MAILSUBJECT
        {
            get { return _MAILSUBJECT; }
            set { _MAILSUBJECT = value; }
        }
        private string _MAILBODY;

        public string MAILBODY
        {
            get { return _MAILBODY; }
            set { _MAILBODY = value; }
        }
        private string _STATUS;

        public string STATUS
        {
            get { return _STATUS; }
            set { _STATUS = value; }
        }
        private string _EBALANCE_BDT;

        public string EBALANCE_BDT
        {
            get { return _EBALANCE_BDT; }
            set { _EBALANCE_BDT = value; }
        }
        private string _EBALANCE_USD;

        public string EBALANCE_USD
        {
            get { return _EBALANCE_USD; }
            set { _EBALANCE_USD = value; }
        }

        private string _MIN_AMOUNT_DUE_BDT;

        public string MIN_AMOUNT_DUE_BDT
        {
            get { return _MIN_AMOUNT_DUE_BDT; }
            set { _MIN_AMOUNT_DUE_BDT = value; }
        }
        private string _MIN_AMOUNT_DUE_USD;

        public string MIN_AMOUNT_DUE_USD
        {
            get { return _MIN_AMOUNT_DUE_USD; }
            set { _MIN_AMOUNT_DUE_USD = value; }
        }

        private string _PAYMENT_DATE;

        public string PAYMENT_DATE
        {
            get { return _PAYMENT_DATE; }
            set { _PAYMENT_DATE = value; }
        }

        private string _REWARD_BALANCE;

        public string REWARD_BALANCE
        {
            get { return _REWARD_BALANCE; }
            set { _REWARD_BALANCE = value; }
        }

        private string _IDCLIENT;

        public string IDCLIENT
        {
            get { return _IDCLIENT; }
            set { _IDCLIENT = value; }
        }

    }
    public class EStatementList : List<EStatementInfo> { }
}
