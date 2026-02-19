using System;
using System.Data;
using System.Data.OracleClient;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Drawing;
using System.IO;
//using System.Xaml;
using System.Collections;

using ITOSystem.DatabaseEngine.AccessPoint;

namespace ITOSystem.ApplicationSystem.Core
{
    public class Utility
    {
        #region **** Declearation *****
        private DataHandler dataHandler = new DataHandler();
        #endregion

        #region **** Constuctor *****
        #endregion

        #region **** Private Methods *****
        #endregion

        #region **** Public Methods *****

        public bool ValidComboValue(System.Windows.Forms.ComboBox combo)
        {
            bool isvalid = false;
            if (combo.Items.Count == 0)
            {
                isvalid = true;
            }

            for (int i = 0; i < combo.Items.Count; i++)
            {
                if ((combo.Text.Trim() == combo.GetItemText(combo.Items[i])) || (combo.Text.Trim() == ""))
                {
                    isvalid = true;
                }
            }

            return isvalid;
        }
        public void PopulateComboBox(System.Windows.Forms.ComboBox comboBox, DataTable referenceData)
        {
            try
            {

                comboBox.DataSource = referenceData;
                comboBox.ValueMember = referenceData.Columns[0].ColumnName;
                comboBox.DisplayMember = referenceData.Columns[1].ColumnName;
            }
            catch (Exception errorException)
            {

                throw errorException;
            }

        }
        public void PopulateComboBox(System.Windows.Forms.ComboBox comboBox, string sqlQueryString)
        {
            try
            {
                DataTable dataTable = new DataTable();

                OracleCommand cmd = new OracleCommand();
                OracleParameter pOUTTABLE = new OracleParameter("pOUTTABLE", OracleType.Cursor);
                pOUTTABLE.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(new OracleParameter("pObject_name", OracleType.VarChar)).Value = sqlQueryString.Split('\'')[1];
                cmd.Parameters.Add(pOUTTABLE);

                dataTable = dataHandler.DBOracleManupulation.RecordSet(sqlQueryString.Split('\'')[0].Trim(), cmd);

                comboBox.DataSource = dataTable;
                comboBox.ValueMember = dataTable.Columns[0].ColumnName;
                comboBox.DisplayMember = dataTable.Columns[1].ColumnName;
            }
            catch (Exception errorException)
            {

                throw errorException;
            }

        }

        public void PopulateComboBox(System.Windows.Forms.ComboBox comboBox, string sqlQueryString, string displayMember)
        {
            try
            {
                PopulateComboBox(comboBox, sqlQueryString, "", displayMember);
            }
            catch (Exception errorException)
            {

                throw errorException;
            }

        }
        public void PopulateComboBox(System.Windows.Forms.ComboBox comboBox, string sqlQueryString, string valueMember, string displayMember)
        {
            try
            {
                DataTable dataTable = new DataTable();

                OracleCommand cmd = new OracleCommand();
                OracleParameter pOUTTABLE = new OracleParameter("pOUTTABLE", OracleType.Cursor);
                pOUTTABLE.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(new OracleParameter("pObject_name", OracleType.VarChar)).Value = sqlQueryString.Split('\'')[1];
                cmd.Parameters.Add(pOUTTABLE);

                dataTable = dataHandler.DBOracleManupulation.RecordSet(sqlQueryString.Split('\'')[0].Trim(), cmd);
                comboBox.DataSource = dataTable;
                if (valueMember != "")
                {
                    comboBox.ValueMember = valueMember;
                }
                comboBox.DisplayMember = displayMember;
            }
            catch (Exception errorException)
            {

                throw errorException;
            }

        }
        public void PopulateComboBox(System.Windows.Forms.ComboBox comboBox, string sqlQueryString, int displayMember)
        {
            try
            {
                PopulateComboBox(comboBox, sqlQueryString, 0, displayMember);
            }
            catch (Exception errorException)
            {

                throw errorException;
            }

        }
        public void PopulateComboBox(System.Windows.Forms.ComboBox comboBox, string sqlQueryString, int valueMember, int displayMember)
        {
            try
            {
                DataTable dataTable = new DataTable();
                dataTable = dataHandler.DBOracleManupulation.RecordSet(sqlQueryString);
                comboBox.DataSource = dataTable;
                comboBox.ValueMember = dataTable.Columns[valueMember].ToString();
                comboBox.DisplayMember = dataTable.Columns[displayMember].ToString();
            }
            catch (Exception errorException)
            {

                throw errorException;
            }

        }
        public void PopulateCombo(System.Windows.Controls.ComboBox comboBox, string sqlQueryString)
        {
            try
            {
                DataTable dataTable = new DataTable();

                OracleCommand cmd = new OracleCommand();
                OracleParameter pOUTTABLE = new OracleParameter("pOUTTABLE", OracleType.Cursor);
                pOUTTABLE.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(new OracleParameter("pObject_name", OracleType.VarChar)).Value = sqlQueryString.Split('\'')[1];
                cmd.Parameters.Add(pOUTTABLE);

                dataTable = dataHandler.DBOracleManupulation.RecordSet(sqlQueryString.Split('\'')[0].Trim(), cmd);
                comboBox.DataContext = ((IListSource)dataTable).GetList();
                comboBox.SelectedValuePath = dataTable.Columns[0].ColumnName.ToString();
                comboBox.DisplayMemberPath = dataTable.Columns[1].ColumnName.ToString();
            }
            catch (Exception)
            {
            }
        }
        public DataTable GetData(string sqlQueryString)
        {
            try
            {
                DataTable dataTable = new DataTable();
                dataTable = dataHandler.DBOracleManupulation.RecordSet(sqlQueryString);
                return dataTable;
            }
            catch (Exception errorException)
            {

                throw errorException;
            }

        }
        #endregion


    }
}
