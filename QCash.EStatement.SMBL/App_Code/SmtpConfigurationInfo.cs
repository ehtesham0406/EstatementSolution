using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator.App_Code
{
    public class SmtpConfigurationInfo
    {
        private string _FIID;

        public string FIID
        {
            get { return _FIID; }
            set { _FIID = value; }
        }
        private string _Smtp_Server;

        public string Smtp_Server
        {
            get { return _Smtp_Server; }
            set { _Smtp_Server = value; }
        }
        private int _Smtp_Port;

        public int Smtp_Port
        {
            get { return _Smtp_Port; }
            set { _Smtp_Port = value; }
        }
        private int _EnableSSL;

        public int EnableSSL
        {
            get { return _EnableSSL; }
            set { _EnableSSL = value; }
        }
        private string _From_Address;

        public string From_Address
        {
            get { return _From_Address; }
            set { _From_Address = value; }
        }
        private string _From_User;

        public string From_User
        {
            get { return _From_User; }
            set { _From_User = value; }
        }
        private string _From_Password;

        public string From_Password
        {
            get { return _From_Password; }
            set { _From_Password = value; }
        }
        private int _Status;

        public int Status
        {
            get { return _Status; }
            set { _Status = value; }
        }
    }
    public class SmtpConfigurationList : List<SmtpConfigurationInfo> { }
}
