using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator
{
    public class Statement
    {
        private string _BANK_CODE;

        public string BANK_CODE
        {
            get { return _BANK_CODE; }
            set { _BANK_CODE = value; }
        }
        private string _STATEMENT_DATE;

        public string STATEMENT_DATE
        {
            get { return _STATEMENT_DATE; }
            set { _STATEMENT_DATE = value; }
        }
        private string _STATEMENTNO;

        public string STATEMENTNO
        {
            get { return _STATEMENTNO; }
            set { _STATEMENTNO = value; }
        }
        private string _ACCOUNT;

        public string ACCOUNT
        {
            get { return _ACCOUNT; }
            set { _ACCOUNT = value; }
        }

        private string _COMPANY;

        public string COMPANY
        {
            get { return _COMPANY; }
            set { _COMPANY = value; }
        }

        private string _STARTDATE;

        public string STARTDATE
        {
            get { return _STARTDATE; }
            set { _STARTDATE = value; }
        }

        private string _TELEPHONE;

        public string TELEPHONE
        {
            get { return _TELEPHONE; }
            set { _TELEPHONE = value; }
        }

        private string _CLIENTLAT;

        public string CLIENTLAT
        {
            get { return _CLIENTLAT; }
            set { _CLIENTLAT = value; }
        }

        private string _PERSONALCODE;

        public string PERSONALCODE
        {
            get { return _PERSONALCODE; }
            set { _PERSONALCODE = value; }
        }

        private string _MOBILE;

        public string MOBILE
        {
            get { return _MOBILE; }
            set { _MOBILE = value; }
        }

        private string _STREETADDRESS;

        public string STREETADDRESS
        {
            get { return _STREETADDRESS; }
            set { _STREETADDRESS = value; }
        }

        private string _TOTALIN;

        public string TOTALIN
        {
            get { return _TOTALIN; }
            set { _TOTALIN = value; }
        }

        private string _TOTALOUT;

        public string TOTALOUT
        {
            get { return _TOTALOUT; }
            set { _TOTALOUT = value; }
        }

        private string _ADDRESS;

        public string ADDRESS
        {
            get { return _ADDRESS; }
            set { _ADDRESS = value; }
        }

        private string _COUNTRY;

        public string COUNTRY
        {
            get { return _COUNTRY; }
            set { _COUNTRY = value; }
        }

        private string _ACCOUNTTYPENAME;

        public string ACCOUNTTYPENAME
        {
            get { return _ACCOUNTTYPENAME; }
            set { _ACCOUNTTYPENAME = value; }
        }

        private string _OVERDRAFT;

        public string OVERDRAFT
        {
            get { return _OVERDRAFT; }
            set { _OVERDRAFT = value; }
        }

        private string _CURRFULLNAME;

        public string CURRFULLNAME
        {
            get { return _CURRFULLNAME; }
            set { _CURRFULLNAME = value; }
        }

        private string _STATEMENTTYPE;

        public string STATEMENTTYPE
        {
            get { return _STATEMENTTYPE; }
            set { _STATEMENTTYPE = value; }
        }

        private string _SENDTYPE;

        public string SENDTYPE
        {
            get { return _SENDTYPE; }
            set { _SENDTYPE = value; }
        }

        private string _ENDDATE;

        public string ENDDATE
        {
            get { return _ENDDATE; }
            set { _ENDDATE = value; }
        }

        private string _FAX;

        public string FAX
        {
            get { return _FAX; }
            set { _FAX = value; }
        }

        private string _CLIENT;

        public string CLIENT
        {
            get { return _CLIENT; }
            set { _CLIENT = value; }
        }
        private string _IDCLIENT;

        public string IDCLIENT
        {
            get { return _IDCLIENT; }
            set { _IDCLIENT = value; }
        }
        private string _CURRENCY;

        public string CURRENCY
        {
            get { return _CURRENCY; }
            set { _CURRENCY = value; }
        }
        private string _CURRENCYNAME;

        public string CURRENCYNAME
        {
            get { return _CURRENCYNAME; }
            set { _CURRENCYNAME = value; }
        }
        private string _STARTBALANCE;

        public string STARTBALANCE
        {
            get { return _STARTBALANCE; }
            set { _STARTBALANCE = value; }
        }
        private string _AVAILABLE;

        public string AVAILABLE
        {
            get { return _AVAILABLE; }
            set { _AVAILABLE = value; }
        }
        private string _SEX;

        public string SEX
        {
            get { return _SEX; }
            set { _SEX = value; }
        }
        private string _PAGER;

        public string PAGER
        {
            get { return _PAGER; }
            set { _PAGER = value; }
        }
        private string _EMPLOYEENO;

        public string EMPLOYEENO
        {
            get { return _EMPLOYEENO; }
            set { _EMPLOYEENO = value; }
        }
        private string _JOBTITLE;

        public string JOBTITLE
        {
            get { return _JOBTITLE; }
            set { _JOBTITLE = value; }
        }

        private string _EMAIL;

        public string EMAIL
        {
            get { return _EMAIL; }
            set { _EMAIL = value; }
        }
        private string _ENDBALANCE;

        public string ENDBALANCE
        {
            get { return _ENDBALANCE; }
            set { _ENDBALANCE = value; }
        }
        private string _DEBITRESERVE;

        public string DEBITRESERVE
        {
            get { return _DEBITRESERVE; }
            set { _DEBITRESERVE = value; }
        }
        private string _TITLE;

        public string TITLE
        {
            get { return _TITLE; }
            set { _TITLE = value; }
        }
         private string _MAIN_CARD;
        public string MAIN_CARD
        {
            get { return _MAIN_CARD; }
            set { _MAIN_CARD = value; }
        }

        private string _PROMOTIONALTEXT;

        public string PROMOTIONALTEXT
        {
            get { return _PROMOTIONALTEXT; }
            set { _PROMOTIONALTEXT = value; }
        }
   }

    public class StatementList : List<Statement>
    { }
}
