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
    }
    public class EStatementList : List<EStatementInfo> { }
}
