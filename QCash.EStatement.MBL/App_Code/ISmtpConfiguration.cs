using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator.App_Code
{
    interface ISmtpConfiguration
    {
        string SaveSmtpConfiguration(SmtpConfigurationInfo objSmtpConfig);
        SmtpConfigurationList GetSmtpConfiguration(string Fid, int status);
        string UpdateSmtpConfiguration(SmtpConfigurationInfo objSmtpConfig);
    }
}
