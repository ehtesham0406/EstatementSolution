using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.CompilerServices;
//
namespace Common
{
    public class SPExecute
    {
        private enum SqlConnectionOwnership
        {
            Internal = 0,
            External = 1,
        }
        private static string ConStr;
        private SqlConnection SqlCon;
        private SqlDataAdapter adapter;
        private static Hashtable paramCache;

        static SPExecute()
        {
            SPExecute.paramCache = new Hashtable();
        }
        public SPExecute(string ConnStr)
        {
            this.adapter = new SqlDataAdapter();
            SPExecute.ConStr = ConnStr;
            this.SqlCon = new SqlConnection(ConStr);

        }

        private object IIf(bool Expression, object TruePart, object FalsePart)
        {
            object obj;
            bool bl = !Expression;
            if (!bl)
            {
                obj = TruePart;
            }
            else
            {
                return FalsePart;
            }
            return obj;
        }
        private SqlParameter[] CloneParameters(SqlParameter[] originalParameters)
        {
            int num1 = originalParameters.Length - 1;
            SqlParameter[] parameterArray1 = new SqlParameter[num1 + 1];
            int num2 = num1;
            for (int num3 = 0; num3 <= num2; num3++)
            {
                parameterArray1[num3] = (SqlParameter)((ICloneable)originalParameters[num3]).Clone();
            }
            return parameterArray1;

        }
        private SqlParameter[] DiscoverSpParameterSet(SqlConnection connection, string spName, bool includeReturnValueParameter, params object[] parameterValues)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            if ((spName == null) || (spName.Length == 0))
            {
                throw new ArgumentNullException("spName");
            }
            SqlCommand command1 = new SqlCommand(spName, connection);
            command1.CommandType = CommandType.StoredProcedure;
            connection.Open();
            SqlCommandBuilder.DeriveParameters(command1);
            connection.Close();
            if (!includeReturnValueParameter)
            {
                command1.Parameters.RemoveAt(0);
            }
            SqlParameter[] parameterArray1 = new SqlParameter[(command1.Parameters.Count - 1) + 1];
            command1.Parameters.CopyTo(parameterArray1, 0);
            foreach (SqlParameter parameter1 in parameterArray1)
            {
                parameter1.Value = DBNull.Value;
            }
            return parameterArray1;

        }
        private void AssignParameterValues(SqlParameter[] commandParameters, object[] parameterValues)
        {
            try
            {
                if ((commandParameters != null) || (parameterValues != null))
                {
                    //if (commandParameters.Length != parameterValues.Length)
                    //{
                    //    throw new ArgumentException("Parameter count does not match Parameter Value count.");
                    //}
                    int num1 = commandParameters.Length - 1;
                    int num2 = num1;
                    for (int num3 = 0; num3 <= num2; num3++)
                    {
                        if (parameterValues[num3] is IDbDataParameter)
                        {
                            IDbDataParameter parameter1 = (IDbDataParameter)parameterValues[num3];
                            if (parameter1.Value == null)
                            {
                                commandParameters[num3].Value = DBNull.Value;
                            }
                            else
                            {
                                commandParameters[num3].Value = RuntimeHelpers.GetObjectValue(parameter1.Value);
                            }
                        }
                        else if (parameterValues[num3] == null)
                        {
                            commandParameters[num3].Value = DBNull.Value;
                        }
                        else
                        {
                            commandParameters[num3].Value = RuntimeHelpers.GetObjectValue(parameterValues[num3]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }

        }
        private void PrepareCommand(SqlCommand command, SqlConnection connection, SqlTransaction transaction, CommandType commandType, string commandText, System.Data.SqlClient.SqlParameter[] commandParameters, ref bool mustCloseConnection)
        {
            if (command == null)
            {
                throw new ArgumentNullException("command");
            }
            if ((commandText == null) || (commandText.Length == 0))
            {
                throw new ArgumentNullException("commandText");
            }
            if (connection.State != ConnectionState.Open)
            {
                connection.Open();
                mustCloseConnection = true;
            }
            else
            {
                mustCloseConnection = false;
            }
            command.Connection = connection;
            command.CommandText = commandText;
            if (transaction != null)
            {
                if (transaction.Connection == null)
                {
                    throw new ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction");
                }
                command.Transaction = transaction;
            }
            command.CommandType = commandType;
            if (commandParameters != null)
            {
                this.AttachParameters(command, commandParameters);
            }

        }
        private void AttachParameters(SqlCommand command, SqlParameter[] commandParameters)
        {
            if (command == null)
            {
                throw new ArgumentNullException("command");
            }
            if (commandParameters != null)
            {
                foreach (SqlParameter parameter1 in commandParameters)
                {
                    if (parameter1 != null)
                    {
                        if (((parameter1.Direction == ParameterDirection.InputOutput) || (parameter1.Direction == ParameterDirection.Input)) && (parameter1.Value == null))
                        {
                            parameter1.Value = DBNull.Value;
                        }
                        command.Parameters.Add(parameter1);
                    }
                }
            }

        }
        private SqlParameter[] GetSpParameterSet(SqlConnection connection, string spName)
        {
            return GetSpParameterSet(connection, spName, false);
        }
        private SqlParameter[] GetSpParameterSet(SqlConnection connection, string spName, bool includeReturnValueParameter)
        {
            SqlParameter[] parameterArray1;
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            try
            {
                this.SqlCon = (SqlConnection)((ICloneable)connection).Clone();
                parameterArray1 = this.GetSpParameterSetInternal(this.SqlCon, spName, includeReturnValueParameter);
            }
            finally
            {
                if (this.SqlCon != null)
                {
                    this.SqlCon.Dispose();
                }
            }
            return parameterArray1;

        }
        private SqlParameter[] GetSpParameterSetInternal(SqlConnection connection, string spName, bool includeReturnValueParameter)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            if ((spName == null) || (spName.Length == 0))
            {
                throw new ArgumentNullException("spName");
            }
            string text1 = connection.ConnectionString + ":" + spName + this.IIf(includeReturnValueParameter, ":include ReturnValue Parameter", "").ToString();
            SqlParameter[] parameterArray1 = (SqlParameter[])SPExecute.paramCache[text1];
            if (parameterArray1 == null)
            {
                SqlParameter[] parameterArray2 = this.DiscoverSpParameterSet(connection, spName, includeReturnValueParameter, new object[0]);
                SPExecute.paramCache[text1] = parameterArray2;
                parameterArray1 = parameterArray2;
            }
            return this.CloneParameters(parameterArray1);

        }
        private SqlParameter[] GetSpParameterSet(string connectionString, string spName, bool includeReturnValueParameter)
        {
            SqlParameter[] parameterArray1;
            if ((connectionString == null) || (connectionString.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            try
            {
                this.SqlCon = new SqlConnection(connectionString);
                parameterArray1 = this.GetSpParameterSetInternal(this.SqlCon, spName, includeReturnValueParameter);
            }
            finally
            {
                if (this.SqlCon != null)
                {
                    this.SqlCon.Dispose();
                }
            }
            return parameterArray1;

        }
        private SqlParameter[] GetSpParameterSet(string connectionString, string spName)
        {
            return GetSpParameterSet(connectionString, spName, false);
        }
        private int ExecuteNonQuery(SqlConnection connection, string spName, params object[] parameterValues)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            if ((spName == null) || (spName.Length == 0))
            {
                throw new ArgumentNullException("spName");
            }
            if ((parameterValues != null) && (parameterValues.Length > 0))
            {
                SqlParameter[] parameterArray1 = this.GetSpParameterSet(connection, spName);
                this.AssignParameterValues(parameterArray1, parameterValues);
                return this.ExecuteNonQuery(connection, CommandType.StoredProcedure, spName, parameterArray1);
            }
            return this.ExecuteNonQuery(connection, CommandType.StoredProcedure, spName, new SqlParameter[0]);
        }


        private int ExecuteNonQuery(SqlConnection connection, CommandType commandType, string commandText, params System.Data.SqlClient.SqlParameter[] commandParameters)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            SqlCommand command1 = new SqlCommand();
            bool flag1 = false;
            this.PrepareCommand(command1, connection, null, commandType, commandText, commandParameters, ref flag1);
            int num1 = command1.ExecuteNonQuery();
            command1.Parameters.Clear();
            if (flag1)
            {
                connection.Close();
            }
            return num1;

        }
        private int ExecuteNonQuery(string connectionString, CommandType commandType, string commandText, params System.Data.SqlClient.SqlParameter[] commandParameters)
        {
            int num1;
            if ((connectionString == null) || (connectionString.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            try
            {
                this.SqlCon = new SqlConnection(connectionString);
                this.SqlCon.Open();
                num1 = this.ExecuteNonQuery(this.SqlCon, commandType, commandText, commandParameters);
            }
            finally
            {
                if (this.SqlCon != null)
                {
                    this.SqlCon.Dispose();
                }
            }
            return num1;

        }
        private int ExecuteNonQuery(string connectionString, CommandType commandType, string commandText)
        {
            return ExecuteNonQuery(connectionString, commandType, commandText, null);
        }
        public int ExecuteNonQuery(string spName, params object[] parameterValues)
        {
            if ((SPExecute.ConStr == null) || (SPExecute.ConStr.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            if ((spName == null) || (spName.Length == 0))
            {
                throw new ArgumentNullException("spName");
            }
            if ((parameterValues != null) && (parameterValues.Length > 0))
            {
                SqlParameter[] parameterArray1 = this.GetSpParameterSet(SPExecute.ConStr, spName);
                this.AssignParameterValues(parameterArray1, parameterValues);
              return  this.ExecuteNonQuery(SPExecute.ConStr, CommandType.StoredProcedure, spName, parameterArray1);
            }
            else
            {
               return this.ExecuteNonQuery(SPExecute.ConStr, CommandType.StoredProcedure, spName);
            }

        }
        private SqlDataReader ExecuteReader(SqlConnection connection, SqlTransaction transaction, CommandType commandType, string commandText, System.Data.SqlClient.SqlParameter[] commandParameters, SPExecute.SqlConnectionOwnership connectionOwnership)
        {
            SqlDataReader dataReader;
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            bool flag1 = false;
            SqlCommand command1 = new SqlCommand();
            try
            {
                SqlDataReader reader2;
                this.PrepareCommand(command1, connection, transaction, commandType, commandText, commandParameters, ref flag1);
                if (connectionOwnership == SPExecute.SqlConnectionOwnership.External)
                {
                    reader2 = command1.ExecuteReader();
                }
                else
                {
                    reader2 = command1.ExecuteReader(CommandBehavior.CloseConnection);
                }
                bool flag2 = true;
                foreach (SqlParameter parameter1 in command1.Parameters)
                {
                    if (parameter1.Direction != ParameterDirection.Input)
                    {
                        flag2 = false;
                    }
                }
                if (flag2)
                {
                    command1.Parameters.Clear();
                }
                dataReader = reader2;
            }
            catch (Exception)
            {
                if (flag1)
                {
                    connection.Close();
                }
                return null;
            }
            return dataReader;

        }
        private SqlDataReader ExecuteReader(string connectionString, CommandType commandType, string commandText)
        {
            return this.ExecuteReader(connectionString, commandType, commandText, null);

        }
        private SqlDataReader ExecuteReader(string connectionString, CommandType commandType, string commandText, params System.Data.SqlClient.SqlParameter[] commandParameters)
        {
            SqlDataReader dataReader;
            if ((connectionString == null) || (connectionString.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            try
            {
                this.SqlCon = new SqlConnection(connectionString);
                this.SqlCon.Open();
                dataReader = this.ExecuteReader(this.SqlCon, null, commandType, commandText, commandParameters, SPExecute.SqlConnectionOwnership.Internal);
            }
            catch (Exception)
            {
                if (this.SqlCon != null)
                {
                    this.SqlCon.Dispose();
                }
                return null;
            }
            return dataReader;

        }
        public SqlDataReader ExecuteReader(string spName, params object[] parameterValues)
        {
            if ((SPExecute.ConStr == null) || (SPExecute.ConStr.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            if ((spName == null) || (spName.Length == 0))
            {
                throw new ArgumentNullException("spName");
            }
            if ((parameterValues != null) && (parameterValues.Length > 0))
            {
                SqlParameter[] parameterArray1 = this.GetSpParameterSet(SPExecute.ConStr, spName);
                this.AssignParameterValues(parameterArray1, parameterValues);
                return this.ExecuteReader(SPExecute.ConStr, CommandType.StoredProcedure, spName, parameterArray1);
            }
            return this.ExecuteReader(SPExecute.ConStr, CommandType.StoredProcedure, spName);

        }
        private DataSet ExecuteDataset(string connectionString, CommandType commandType, string commandText)
        {
            return ExecuteDataset(connectionString, commandType, commandText, null);
        }
        private DataSet ExecuteDataset(SqlConnection connection, CommandType commandType, string commandText, params System.Data.SqlClient.SqlParameter[] commandParameters)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            SqlCommand command1 = new SqlCommand();
            DataSet dataSet = new DataSet();
            bool flag1 = false;
            this.PrepareCommand(command1, connection, null, commandType, commandText, commandParameters, ref flag1);
            try
            {
                this.adapter = new SqlDataAdapter(command1);
                this.adapter.Fill(dataSet);
                command1.Parameters.Clear();
            }
            catch (Exception)
            {
                if (this.adapter != null)
                {
                    this.adapter.Dispose();
                }
            }
            if (flag1)
            {
                connection.Close();
            }
            return dataSet;

        }
        private DataSet ExecuteDataset(string connectionString, CommandType commandType, string commandText, params System.Data.SqlClient.SqlParameter[] commandParameters)
        {
            DataSet dataSet;
            if ((connectionString == null) || (connectionString.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            try
            {
                this.SqlCon = new SqlConnection(connectionString);
                this.SqlCon.Open();
                dataSet = this.ExecuteDataset(this.SqlCon, commandType, commandText, commandParameters);
            }
            finally
            {
                if (this.SqlCon != null)
                {
                    this.SqlCon.Dispose();
                }
            }
            return dataSet;

        }
        public DataSet ExecuteDataset(string spName, params object[] parameterValues)
        {
            if ((SPExecute.ConStr == null) || (SPExecute.ConStr.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            if ((spName == null) || (spName.Length == 0))
            {
                throw new ArgumentNullException("spName");
            }
            if ((parameterValues != null) && (parameterValues.Length > 0))
            {
                SqlParameter[] parameterArray1 = this.GetSpParameterSet(SPExecute.ConStr, spName);
                this.AssignParameterValues(parameterArray1, parameterValues);
                return this.ExecuteDataset(SPExecute.ConStr, CommandType.StoredProcedure, spName, parameterArray1);
            }
            return this.ExecuteDataset(SPExecute.ConStr, CommandType.StoredProcedure, spName);

        }

    }
}