using System;
using System.Collections;
using System.Data;
using System.Data.OracleClient;
using System.Runtime.CompilerServices;

namespace Common
{
    public class OracleSPExecute
    {
        private enum OracleConnectionOwnership
        {
            Internal = 0,
            External = 1,
        }
        private static string ConStr;
        private OracleConnection OracleCon;
        private OracleDataAdapter adapter;
        private static Hashtable paramCache;

        static OracleSPExecute()
        {
            OracleSPExecute.paramCache = new Hashtable();
        }
        public OracleSPExecute(string ConnStr)
        {
            this.adapter = new OracleDataAdapter();
            OracleSPExecute.ConStr = ConnStr;
            this.OracleCon = new OracleConnection(ConStr);

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
        private OracleParameter[] CloneParameters(OracleParameter[] originalParameters)
        {
            int num1 = originalParameters.Length - 1;
            OracleParameter[] parameterArray1 = new OracleParameter[num1 + 1];
            int num2 = num1;
            for (int num3 = 0; num3 <= num2; num3++)
            {
                parameterArray1[num3] = (OracleParameter)((ICloneable)originalParameters[num3]).Clone();
            }
            return parameterArray1;

        }
        private OracleParameter[] DiscoverSpParameterSet(OracleConnection connection, string spName, bool includeReturnValueParameter, ref string reply, params object[] parameterValues)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            if ((spName == null) || (spName.Length == 0))
            {
                throw new ArgumentNullException("spName");
            }
            OracleParameter[] parameterArray1 = null;

            try
            {
                OracleCommand command1 = new OracleCommand(spName, connection);
                command1.CommandType = CommandType.StoredProcedure;
                connection.Open();
                OracleCommandBuilder.DeriveParameters(command1);
                connection.Close();
                if (!includeReturnValueParameter)
                {
                    command1.Parameters.RemoveAt(0);
                }
                parameterArray1 = new OracleParameter[(command1.Parameters.Count - 1) + 1];
                command1.Parameters.CopyTo(parameterArray1, 0);
                foreach (OracleParameter parameter1 in parameterArray1)
                {
                    parameter1.Value = DBNull.Value;
                }
            }
            catch (Exception ex)
            {
                reply = ex.Message;
                return parameterArray1;
            }
            return parameterArray1;

        }
        private void AssignParameterValues(OracleParameter[] commandParameters, object[] parameterValues)
        {
            int num1 = commandParameters.Length - 1;
            int num2 = parameterValues.Length;
            try
            {
                if ((commandParameters != null) || (parameterValues != null))
                {
                    int j = 0;
                    for (int num3 = 0; num3 <= num1; num3++)
                    {

                        if (commandParameters[num3].Direction != ParameterDirection.Output)
                        {
                            for (int n = j; n < j + 1; n++)
                            {
                                commandParameters[num3].Value = RuntimeHelpers.GetObjectValue(parameterValues[n]);

                            }
                            j++;
                        }
                    }
                }
            }
            catch (Exception)
            {

            }

        }
        private void PrepareCommand(OracleCommand command, OracleConnection connection, OracleTransaction transaction, CommandType commandType, string commandText, System.Data.OracleClient.OracleParameter[] commandParameters, ref bool mustCloseConnection)
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
        private void AttachParameters(OracleCommand command, OracleParameter[] commandParameters)
        {
            if (command == null)
            {
                throw new ArgumentNullException("command");
            }
            if (commandParameters != null)
            {
                foreach (OracleParameter parameter1 in commandParameters)
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
        //
        private OracleParameter[] GetSpParameterSet(OracleConnection connection, string spName, ref string reply)
        {
            return GetSpParameterSet(connection, spName, true, ref reply);
        }
        private OracleParameter[] GetSpParameterSet(OracleConnection connection, string spName, bool includeReturnValueParameter, ref string reply)
        {
            OracleParameter[] parameterArray1;
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            try
            {
                this.OracleCon = (OracleConnection)((ICloneable)connection).Clone();
                parameterArray1 = this.GetSpParameterSetInternal(this.OracleCon, spName, includeReturnValueParameter, ref reply);
            }
            finally
            {
                if (this.OracleCon != null)
                {
                    this.OracleCon.Dispose();
                }
            }
            return parameterArray1;

        }
        private OracleParameter[] GetSpParameterSetInternal(OracleConnection connection, string spName, bool includeReturnValueParameter, ref string reply)
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
            OracleParameter[] parameterArray1 = (OracleParameter[])OracleSPExecute.paramCache[text1];
            if (parameterArray1 == null)
            {
                OracleParameter[] parameterArray2 = this.DiscoverSpParameterSet(connection, spName, includeReturnValueParameter, ref reply, new object[0]);
                OracleSPExecute.paramCache[text1] = parameterArray2;
                parameterArray1 = parameterArray2;
            }
            if (parameterArray1 != null)
                return this.CloneParameters(parameterArray1);
            else
                return null;

        }
        private OracleParameter[] GetSpParameterSet(string connectionString, string spName, bool includeReturnValueParameter, ref string reply)
        {
            OracleParameter[] parameterArray1;
            if ((connectionString == null) || (connectionString.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            try
            {
                this.OracleCon = new OracleConnection(connectionString);
                parameterArray1 = this.GetSpParameterSetInternal(this.OracleCon, spName, includeReturnValueParameter, ref reply);
            }
            finally
            {
                if (this.OracleCon != null)
                {
                    this.OracleCon.Dispose();
                }
            }
            return parameterArray1;

        }
        private OracleParameter[] GetSpParameterSet(string connectionString, string spName, ref string reply)
        {
            return GetSpParameterSet(connectionString, spName, true, ref reply);
        }
        //
        private int ExecuteNonQuery(OracleConnection connection, string spName, ref string reply, params object[] parameterValues)
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
                OracleParameter[] parameterArray1 = this.GetSpParameterSet(connection, spName, ref reply);

                if (parameterArray1 != null)
                {
                    this.AssignParameterValues(parameterArray1, parameterValues);
                    return this.ExecuteNonQuery(connection, CommandType.StoredProcedure, spName, parameterArray1);
                }
            }
            return 0;
        }
        private int ExecuteNonQuery(OracleConnection connection, CommandType commandType, string commandText, params System.Data.OracleClient.OracleParameter[] commandParameters)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            OracleCommand command1 = new OracleCommand();
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
        private int ExecuteNonQuery(string connectionString, CommandType commandType, string commandText, params System.Data.OracleClient.OracleParameter[] commandParameters)
        {
            int num1;
            if ((connectionString == null) || (connectionString.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            try
            {
                this.OracleCon = new OracleConnection(connectionString);
                this.OracleCon.Open();
                num1 = this.ExecuteNonQuery(this.OracleCon, commandType, commandText, commandParameters);
            }
            finally
            {
                if (this.OracleCon != null)
                {
                    this.OracleCon.Dispose();
                }
            }
            return num1;

        }
        private int ExecuteNonQuery(string connectionString, CommandType commandType, string commandText)
        {
            return ExecuteNonQuery(connectionString, commandType, commandText, null);
        }
        public void ExecuteNonQuery(string spName, ref string reply, params object[] parameterValues)
        {
            if ((OracleSPExecute.ConStr == null) || (OracleSPExecute.ConStr.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            if ((spName == null) || (spName.Length == 0))
            {
                throw new ArgumentNullException("spName");
            }
            if ((parameterValues != null) && (parameterValues.Length > 0))
            {
                OracleParameter[] parameterArray1 = this.GetSpParameterSet(OracleSPExecute.ConStr, spName, ref reply);
                this.AssignParameterValues(parameterArray1, parameterValues);
                this.ExecuteNonQuery(OracleSPExecute.ConStr, CommandType.StoredProcedure, spName, parameterArray1);
            }
            else
            {
                this.ExecuteNonQuery(OracleSPExecute.ConStr, CommandType.StoredProcedure, spName);
            }

        }
        //
        public OracleParameter[] ExecuteNonQueryParam(string spName, ref string reply, params object[] parameterValues)
        {
            OracleParameter[] param = null;
            if ((OracleSPExecute.ConStr == null) || (OracleSPExecute.ConStr.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            if ((spName == null) || (spName.Length == 0))
            {
                throw new ArgumentNullException("spName");
            }
            if ((parameterValues != null) && (parameterValues.Length > 0))
            {
                OracleParameter[] parameterArray1 = this.GetSpParameterSet(OracleSPExecute.ConStr, spName, ref reply);
                if (parameterArray1 != null)
                {
                    this.AssignParameterValues(parameterArray1, parameterValues);
                    param = this.ExecuteNonQueryParam(OracleSPExecute.ConStr, CommandType.StoredProcedure, spName, ref reply, parameterArray1);
                }
            }
            return param;
        }
        private OracleParameter[] ExecuteNonQueryParam(string constr, CommandType commandType, string commandText, ref string reply, params System.Data.OracleClient.OracleParameter[] commandParameters)
        {
            OracleParameter[] param = null;

            try
            {
                OracleCon = new OracleConnection(constr);
                OracleCon.Open();

                param = this.ExecuteNonQueryParam(OracleCon, commandType, commandText, ref reply, commandParameters);
            }
            catch (Exception ex)
            {
                reply = ex.Message;
            }
            return param;

        }
        private OracleParameter[] ExecuteNonQueryParam(OracleConnection connection, CommandType commandType, string commandText, ref string reply, params System.Data.OracleClient.OracleParameter[] commandParameters)
        {
            OracleParameter[] param = null;
            try
            {
                if (connection == null)
                {
                    throw new ArgumentNullException("connection");
                }
                OracleCommand command1 = new OracleCommand();
                bool flag1 = false;
                this.PrepareCommand(command1, connection, null, commandType, commandText, commandParameters, ref flag1);
                int num1 = command1.ExecuteNonQuery();
                command1.Parameters.Clear();
                if (flag1)
                {
                    connection.Close();
                }
                param = commandParameters;
                return param;
            }
            catch (Exception ex)
            {
                reply = ex.Message;
                return commandParameters;
            }
        }
        //
        private OracleDataReader ExecuteReader(OracleConnection connection, OracleTransaction transaction, CommandType commandType, string commandText, System.Data.OracleClient.OracleParameter[] commandParameters, OracleSPExecute.OracleConnectionOwnership connectionOwnership)
        {
            OracleDataReader dataReader;
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            bool flag1 = false;
            OracleCommand command1 = new OracleCommand();
            try
            {
                OracleDataReader reader2;
                this.PrepareCommand(command1, connection, transaction, commandType, commandText, commandParameters, ref flag1);
                if (connectionOwnership == OracleSPExecute.OracleConnectionOwnership.External)
                {
                    reader2 = command1.ExecuteReader();
                }
                else
                {
                    reader2 = command1.ExecuteReader();//CommandBehavior.CloseConnection);
                }
                bool flag2 = true;
                foreach (OracleParameter parameter1 in command1.Parameters)
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
        private OracleDataReader ExecuteReader(string connectionString, CommandType commandType, string commandText)
        {
            return this.ExecuteReader(connectionString, commandType, commandText, null);

        }
        private OracleDataReader ExecuteReader(string connectionString, CommandType commandType, string commandText, params System.Data.OracleClient.OracleParameter[] commandParameters)
        {
            OracleDataReader dataReader;
            if ((connectionString == null) || (connectionString.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            try
            {
                this.OracleCon = new OracleConnection(connectionString);
                this.OracleCon.Open();
                dataReader = this.ExecuteReader(this.OracleCon, null, commandType, commandText, commandParameters, OracleSPExecute.OracleConnectionOwnership.Internal);
            }
            catch (Exception)
            {
                if (this.OracleCon != null)
                {
                    this.OracleCon.Dispose();
                }
                return null;
            }
            return dataReader;

        }
        public OracleDataReader ExecuteReader(string spName, ref string reply, params object[] parameterValues)
        {
            if ((OracleSPExecute.ConStr == null) || (OracleSPExecute.ConStr.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            if ((spName == null) || (spName.Length == 0))
            {
                throw new ArgumentNullException("spName");
            }
            if ((parameterValues != null) && (parameterValues.Length > 0))
            {
                OracleParameter[] parameterArray1 = this.GetSpParameterSet(OracleSPExecute.ConStr, spName, ref reply);
                this.AssignParameterValues(parameterArray1, parameterValues);
                return this.ExecuteReader(OracleSPExecute.ConStr, CommandType.StoredProcedure, spName, parameterArray1);
            }
            return this.ExecuteReader(OracleSPExecute.ConStr, CommandType.StoredProcedure, spName);

        }
        //**//
        private DataSet ExecuteDataset(string connectionString, CommandType commandType, string commandText, ref string reply)
        {
            OracleParameter[] commandParameters = this.GetSpParameterSet(OracleSPExecute.ConStr, commandText, ref reply);
            return ExecuteDataset(connectionString, commandType, commandText, commandParameters);
        }
        private DataSet ExecuteDataset(OracleConnection connection, CommandType commandType, string commandText, params System.Data.OracleClient.OracleParameter[] commandParameters)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            OracleCommand command1 = new OracleCommand();
            DataSet dataSet = new DataSet();
            bool flag1 = false;
            this.PrepareCommand(command1, connection, null, commandType, commandText, commandParameters, ref flag1);
            try
            {
                command1.ExecuteNonQuery();
                this.adapter = new OracleDataAdapter(command1);
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
        private DataSet ExecuteDataset(string connectionString, CommandType commandType, string commandText, params System.Data.OracleClient.OracleParameter[] commandParameters)
        {
            DataSet dataSet;
            if ((connectionString == null) || (connectionString.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            try
            {
                this.OracleCon = new OracleConnection(connectionString);
                this.OracleCon.Open();
                dataSet = this.ExecuteDataset(this.OracleCon, commandType, commandText, commandParameters);
            }
            finally
            {
                if (this.OracleCon != null)
                {
                    this.OracleCon.Dispose();
                }
            }
            return dataSet;

        }
        public DataSet ExecuteDataset(string spName, ref string reply, params object[] parameterValues)
        {
            if ((OracleSPExecute.ConStr == null) || (OracleSPExecute.ConStr.Length == 0))
            {
                throw new ArgumentNullException("connectionString");
            }
            if ((spName == null) || (spName.Length == 0))
            {
                throw new ArgumentNullException("spName");
            }
            if ((parameterValues != null) && (parameterValues.Length > 0))
            {
                OracleParameter[] parameterArray1 = this.GetSpParameterSet(OracleSPExecute.ConStr, spName, ref reply);
                this.AssignParameterValues(parameterArray1, parameterValues);
                return this.ExecuteDataset(OracleSPExecute.ConStr, CommandType.StoredProcedure, spName, parameterArray1);
            }
            return this.ExecuteDataset(OracleSPExecute.ConStr, CommandType.StoredProcedure, spName);

        }

        public OracleConnection OpenConnection(string constr)
        {
            OracleConnection OracleCon = new OracleConnection(constr);
            OracleCon.Open();
            return OracleCon;
        }
        public void CloseConnection(OracleConnection OraCon)
        {
            if (OraCon != null)
            {
                OraCon.Close();
                OraCon.Dispose();
            }
        }
    }
}