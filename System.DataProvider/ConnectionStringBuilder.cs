using System;
using System.Configuration;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.DataProvider;
using FlexiStar.Utilities.EncryptionEngine;

namespace System.Connection
{
    public class ConnectionStringBuilder
    {
        private const string ConnectionTimeOut = "Connection Timeout = 60000";
        private string databaseName;
        private string DefaultConnectionName;
        private string DefaultConnectionString;
        private DataSet dsDBConfig;
        private eConnectionLibrary eProviderType;
        private string password;
        private string serverName;
        private string userID;

        private string _ConnectionString_DBConfig;
        private string _ConnectionString_AppConfig;
        private int constr = 0;

        public ConnectionStringBuilder(int _constr)
        {
            constr = _constr;
        }
        //
        public string ConnectionString_DBConfig
        {
            get 
            {
                this.dsDBConfig = new DataSet();
                if (!this.MakeDefaultConnectionString())
                {
                    if (this.DefaultConnectionString.Trim().Length <= 0)
                    {
                        throw new Exception("DAL is not configured properly\n.There is no link between Application Server and Database Server. Please configure the Application Server properly or ask your MIS Officer or technical person to do so.");
                    }
                    this.Initailize(this.DefaultConnectionString);
                }
                this._ConnectionString_DBConfig = DefaultConnectionString;
                return _ConnectionString_DBConfig;
            }
        }
        // Methods
        public ConnectionStringBuilder()
		{
			
		}
        //
        public string ConnectionString_AppConfig
        {
            get
            {
                this._ConnectionString_AppConfig = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                return _ConnectionString_AppConfig;
            }
        }
        //
        private bool MakeDefaultConnectionString()
        {
            try
            {
                this.dsDBConfig = DBConfig.dsConnection;
                if (this.dsDBConfig.Tables.Count == 0)
                {
                    return false;
                }
                if (this.dsDBConfig.Tables[0].Rows.Count == 0)
                {
                    return false;
                }
                if (this.dsDBConfig.Tables[0].Rows.Count == 1)
                {
                    if (!this.MakeConnectionString(this.dsDBConfig.Tables[0].Rows[0]))
                    {
                        return false;
                    }
                    return true;
                }
                DataRow[] rowArray = this.dsDBConfig.Tables[0].Select("IsDefault=1");
                if (rowArray.Length == 0)
                {
                    if (!this.MakeConnectionString(this.dsDBConfig.Tables[0].Rows[0]))
                    {
                        return false;
                    }
                    return true;
                }
                if (rowArray.Length == 1)
                {
                    if (!this.MakeConnectionString(rowArray[0]))
                    {
                        return false;
                    }
                    return true;
                }
                if (rowArray.Length > 1)
                {
                    if (!this.MakeConnectionString(rowArray[0]))
                    {
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }
            return false;
        }
        //
        private bool MakeConnectionString(DataRow drDefault)
        {
            try
            {
                Encryption encryption = new Encryption();
                string strConn = "";
                string str2 = "";
                if (drDefault["ConnectionName"] == DBNull.Value)
                {
                    return false;
                }
                if (drDefault["ConnectionName"].ToString().Length == 0)
                {
                    return false;
                }
                this.DefaultConnectionName = drDefault["ConnectionName"].ToString();
                this.userID = drDefault["UserName"].ToString();
                if (drDefault["UserName"] != DBNull.Value)
                {
                    str2 = drDefault["UserName"].ToString();
                }
                if (drDefault["ServerName"] == DBNull.Value)
                {
                    strConn = "Data Source=(local);";
                    this.serverName = "(local)";
                }
                else if (drDefault["ServerName"].ToString().Length == 0)
                {
                    strConn = "Data Source=(local);";
                    this.serverName = "(local)";
                }
                else
                {
                    this.serverName = encryption.DecryptWord(drDefault["ServerName"].ToString(), str2);
                    strConn = "Data Source=" + this.serverName + ";";
                }
                if (drDefault["UserName"].ToString().Trim().Length == 0)
                {
                    strConn = strConn + "user id=;";
                }
                else
                {
                    strConn = strConn + "user id=" + drDefault["UserName"].ToString() + ";";
                }
                if (drDefault["Password"].ToString().Trim().Length == 0)
                {
                    strConn = strConn + "pwd=;";
                    this.password = "";
                }
                else
                {
                    this.password = encryption.DecryptWord(drDefault["Password"].ToString(), str2);
                    strConn = strConn + "pwd=" + this.password + ";";
                }
                if (drDefault["DatabaseName"].ToString().Trim().Length == 0)
                {
                    return false;
                }
                this.databaseName = encryption.DecryptWord(drDefault["DatabaseName"].ToString(), str2);
                strConn = (strConn + "Initial Catalog=" + this.databaseName) + "; Connection Timeout = 60000";
                eConnectionLibrary eType = (eConnectionLibrary)Convert.ToInt16(drDefault["Provider"].ToString());
                this.Initailize(strConn, eType);
                return true;
            }
            catch
            {
                return false;
            }
        }
        //
        private void Initailize(string strConn)
        {
            this.Initailize(strConn, eConnectionLibrary.SQlClient);
        }

        private void Initailize(string strConn, eConnectionLibrary eType)
        {
            this.eProviderType = eType;
            this.DefaultConnectionString = strConn;
        }

        public enum eConnectionLibrary : short
        {
            ODBC = 2,
            Oledb = 0,
            OledbMSAcess2000 = 6,
            OledbMSSQL = 4,
            OledbOracle = 5,
            OracleClient = 3,
            SQlClient = 1
        }
    }
}
