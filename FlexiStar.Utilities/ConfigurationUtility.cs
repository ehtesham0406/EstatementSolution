using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace FlexiStar.Utilities
{
    public class ConfigurationUtility
    {
        private ConfigurationUtility()
        {
            
        }

        public static ConfigurationUtility CreateInstance()
        {
           return new ConfigurationUtility();
        }
       
        private string _IPAddress;
        public string IPAddress
        {
            get { return _IPAddress; }
            set { _IPAddress = value; }
        }
        
        private int _Port;
        public int Port
        {
            get { return _Port; }
            set { _Port = value; }
        }
        
        private string _FilePath;
        public string FilePath
        {
            get { return _FilePath; }
            set { _FilePath = value; }
        }
        
        private string _PortFileName;
        public string PortFileName
        {
            get { return _PortFileName; }
            set { _PortFileName = value; }
        }
        
        private string _TraceFileName;
        public string TraceFileName
        {
            get { return _TraceFileName; }
            set { _TraceFileName = value; }
        }
        
        private string _ExceptionFileName;
        public string ExceptionFileName
        {
            get { return _ExceptionFileName; }
            set { _ExceptionFileName = value; }
        }
               
    }
}
