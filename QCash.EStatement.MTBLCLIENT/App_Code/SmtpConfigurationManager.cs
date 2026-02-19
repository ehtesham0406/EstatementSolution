using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator.App_Code
{
    public class SmtpConfigurationManager
    {
        public static SmtpConfigurationManager Instance()
        {
            return new SmtpConfigurationManager();
        }

        #region ISmtpConfiguration Members

        public string SaveSmtpConfiguration(SmtpConfigurationInfo objSmtpConfig)
        {
            return SmtpConfigurationDataProvider.Instance().SaveSmtpConfiguration(objSmtpConfig);
        }

        public SmtpConfigurationList GetSmtpConfiguration(string Fid, int status)
        {
            return SmtpConfigurationDataProvider.Instance().GetSmtpConfiguration(Fid, status);
        }

        public string UpdateSmtpConfiguration(SmtpConfigurationInfo objSmtpConfig)
        {
            return SmtpConfigurationDataProvider.Instance().UpdateSmtpConfiguration(objSmtpConfig);
        }

        #endregion
    }
}
