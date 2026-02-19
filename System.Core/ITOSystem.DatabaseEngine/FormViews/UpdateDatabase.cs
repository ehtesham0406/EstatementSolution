using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Data.SqlClient;

using ITOSystem.DatabaseEngine.AccessPoint;


namespace ITOSystem.DatabaseEngine.FormViews
{
    public partial class UpdateDatabase : Form
    {
        #region **** Members & Instances Declaration ****
        private DataHandler dataHanadler = new DataHandler();
        #endregion

        #region  ***** Constructor *****
        public UpdateDatabase()
        {
            InitializeComponent();
        }
        #endregion

        #region ***** Private Methods *****
        //private bool ExecuteCommand()
        //{
        //    try
        //    {
        //        string[] totalString = Regex.Split(rtfScript.Text.Trim(), "GO");
        //        SqlConnection connectionCommand = new SqlConnection("Data Source=sqlserver;Initial Catalog=BRACStoreInventory;User ID=mijan;pwd=123");
        //        SqlCommand myCommand = new SqlCommand();
        //        myCommand.CommandType = CommandType.StoredProcedure;
        //        myCommand.CommandText = "sp_ExecuteCommand";
        //        myCommand.Connection = connectionCommand;
        //        connectionCommand.Open();
        //        SqlParameter param1 = new SqlParameter("@CommandString", totalString[0].ToString());
        //        myCommand.Parameters.Add(param1);
        //        myCommand.ExecuteNonQuery();
        //        connectionCommand.Close();
        //        return true;
        //    }
        //    catch (Exception errorException)
        //    {
        //        MessageBox.Show(errorException.Message);
        //        connectionCommand.Close();
        //        return false;
        //    }
           
        //}
        #endregion

        #region ***** Events *****
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                    rtfScript.LoadFile(openFileDialog.FileName, RichTextBoxStreamType.PlainText);
      

            }
            catch (Exception errorException)
            {
                MessageBox.Show(errorException.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        } 

        private void btnExecute_Click(object sender, EventArgs e)
        {
            try
            {
                string quaryString = "";
                string[] totalString = Regex.Split(rtfScript.Text.Trim(), "GO");
                for (int totalQuary = 0; totalQuary < totalString.Length; totalQuary++)
                {
                    quaryString = "";
                    quaryString = totalString[totalQuary].ToString();
                    quaryString.Replace(@"\n\", " ");
                    quaryString.Replace(@"\n\t", " ");
                    quaryString.Replace(@"\n", " ");

                    dataHanadler.DBManupulation.Execute( quaryString.ToString().Trim() ,UUL.UDBH.QuaryType.Complex);
                }
                
                MessageBox.Show("Database Update Successfuly,effect " + totalString.Length +" Statement in Database", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception errorException)
            {
                MessageBox.Show(errorException.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}