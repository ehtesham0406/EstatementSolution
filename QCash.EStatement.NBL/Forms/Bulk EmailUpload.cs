using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Configuration;
using System.Data.SqlClient;

namespace QCash.EStatement.NBL.Forms
{
    public partial class EmailUpload : Form
    {
        private string _fiid = string.Empty;
        private string _Email = string.Empty;
        public EmailUpload(string fiid)
        {
            InitializeComponent();
            _fiid = fiid;
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            //declare variables - edit these based on your particular situation
            string ssqltable = "EMAIL_ADDRESS";
            _Email = ConfigurationManager.AppSettings["EmailPath"].ToString();  // excel file name
            // make sure your sheet name is correct, here sheet name is sheet1,
            // so you can change your sheet name if have    different
            string myexceldataquery = "select * from [email$]";  //  excel sheet name





            try
            {
                //create our connection strings
                // string sexcelconnectionstring = @"provider=Microsoft.ACE.OLEDB.12.0;data source=" + "D:\\sheet1.xlsx" +
                // ";extended properties=" + "\"excel 8.0;hdr=yes;\"";

                string sexcelconnectionstring = @"provider=Microsoft.ACE.OLEDB.12.0;data source=" + _Email +
               ";extended properties=" + "\"excel 8.0;hdr=yes;\"";
                string ssqlconnectionstring = "Data Source=RABBY-LAPTOP;Initial Catalog=EStatementMasked_NBL;Integrated Security=True";
                //execute a query to erase any previous data from our destination table
                //  string sclearsql = "delete from " + ssqltable;
                //SqlConnection sqlconn = new SqlConnection(ssqlconnectionstring);
                // SqlCommand sqlcmd = new SqlCommand(sclearsql, sqlconn);
                // sqlconn.Open();
                // sqlcmd.ExecuteNonQuery();
                //sqlconn.Close();
                //series of commands to bulk copy data from the excel file into our sql table
                OleDbConnection oledbconn = new OleDbConnection(sexcelconnectionstring);
                OleDbCommand oledbcmd = new OleDbCommand(myexceldataquery, oledbconn);
                oledbconn.Open();
                OleDbDataReader dr = oledbcmd.ExecuteReader();
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = ssqltable;

                // while (dr.Read())
                //{
                bulkcopy.WriteToServer(dr);
                // }
                dr.Close();
                oledbconn.Close();
                MessageBox.Show("File imported into database Successfully.");
                btnUpload.Enabled = false;
               
            }
            catch (Exception ex)
            {
                //handle exception
            }

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        
    }
}
