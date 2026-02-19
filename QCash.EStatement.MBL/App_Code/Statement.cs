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
        private string _CARD_LIST;

        public string CARD_LIST
        {
            get { return _CARD_LIST; }
            set { _CARD_LIST = value; }
        }
        private string _CATEGORY;

        public string CATEGORY
        {
            get { return _CATEGORY; }
            set { _CATEGORY = value; }
        }
        private string _CITY;

        public string CITY
        {
            get { return _CITY; }
            set { _CITY = value; }
        }
        private string _CONTRACTNO;

        public string CONTRACTNO
        {
            get { return _CONTRACTNO; }
            set { _CONTRACTNO = value; }
        }
        private string _COUNTRY;

        public string COUNTRY
        {
            get { return _COUNTRY; }
            set { _COUNTRY = value; }
        }
        private string _EDUCATION;
        private string _EMAIL;

        public string EMAIL
        {
            get { return _EMAIL; }
            set { _EMAIL = value; }
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
        private string _FAX;

        public string FAX
        {
            get { return _FAX; }
            set { _FAX = value; }
        }
        private string _JOBTITLE;
        public string JOBTITLE
        {
            get { return _JOBTITLE; }
            set { _JOBTITLE = value; }
        }
        //private string _JOBTITLE;
        private string _CLIENTLAT;
        private string _MAIN_CARD;

        public string MAIN_CARD
        {
            get { return _MAIN_CARD; }
            set { _MAIN_CARD = value; }
        }
        private string _MARITALSTATUS;
        private string _MOBILE;

        public string MOBILE
        {
            get { return _MOBILE; }
            set { _MOBILE = value; }
        }
        private string _NEXT_STATEMENT_DATE;

        public string NEXT_STATEMENT_DATE
        {
            get { return _NEXT_STATEMENT_DATE; }
            set { _NEXT_STATEMENT_DATE = value; }
        }
        private string _OCCUPATION;
        private string _OFFICEPHONE;
        private string _PAYMENT_DATE;

        public string PAYMENT_DATE
        {
            get { return _PAYMENT_DATE; }
            set { _PAYMENT_DATE = value; }
        }
        private string _PAGER;
        private string _PAGERNO;
        private string _PERSONALCODE;
        private string _PREVEMPLOYMENT;
        private string _PROMOTIONALTEXT;
        private string _REGION;

        public string REGION
        {
            get { return _REGION; }
            set { _REGION = value; }
        }
        private string _STATEMENT_DATE;

        public string STATEMENT_DATE
        {
            get { return _STATEMENT_DATE; }
            set { _STATEMENT_DATE = value; }
        }
        private string _STATEMENT_PERIOD;
        private string _SEX;

        public string SEX
        {
            get { return _SEX; }
            set { _SEX = value; }
        }
        private string _STATEMENTTYPE;
        private string _SENDTYPE;
        private string _STREETADDRESS;

        public string STREETADDRESS
        {
            get { return _STREETADDRESS; }
            set { _STREETADDRESS = value; }
        }
        private string _TPN;
        private string _TELEPHONE;

        public string TELEPHONE
        {
            get { return _TELEPHONE; }
            set { _TELEPHONE = value; }
        }
        private string _TITLE;

        public string TITLE
        {
            get { return _TITLE; }
            set { _TITLE = value; }
        }
        private string _ZIP;

        public string ZIP
        {
            get { return _ZIP; }
            set { _ZIP = value; }
        }
    }

    public class StatementList : List<Statement>
    { }
}
