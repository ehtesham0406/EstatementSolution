using System;
using System.Collections.Generic;
using System.Text;
using UUL.UDBH.SQLData;
using UUL.UDBH.ORACLEDATA;
using ITOSystem.DatabaseEngine.ConnectionManager;

namespace ITOSystem.DatabaseEngine.AccessPoint
{
    public enum ExecuteType
    {
        INSERT,
        UPDATE,
        DELETE
    }
    public  class DataHandler
    {
        #region ***** Priate variable & Instance Declearation *****
        private DBSQLManupulation _dbHelper;
        private DBOracleDataManupulation _dbOraHelper;
        private DBSQLManupulation _dbTranHelper;
        ConnectionProvider connectionManager = new ConnectionProvider(); 
        #endregion

        #region ***** Public methods ****
        public DataHandler()
        {
        }
        public DBSQLManupulation DBManupulation
        {
            get
            {
                _dbHelper = new DBSQLManupulation(connectionManager.connectionStirng());
                return _dbHelper;
            }
        }
        public DBOracleDataManupulation DBOracleManupulation
        {
            get
            {
                _dbOraHelper = new DBOracleDataManupulation(connectionManager.connectionOracleStirng());
                return _dbOraHelper;
            }
        }
        public DBSQLManupulation DBTransactionManupulation
        {
            get
            {
                _dbTranHelper = new DBSQLManupulation(connectionManager.connectionOracleStirng());
                return _dbTranHelper;
            }
        }

        #endregion
    }
}
