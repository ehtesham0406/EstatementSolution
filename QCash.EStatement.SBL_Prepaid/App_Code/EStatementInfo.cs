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
        private string _PAN;

        public string PAN
        {
            get { return _PAN; }
            set { _PAN = value; }
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
        private string _IDCLIENT;

        public string IDCLIENT
        {
            get { return _IDCLIENT; }
            set { _IDCLIENT = value; }
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
