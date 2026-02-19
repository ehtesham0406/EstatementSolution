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
using System.Common;
using System.Connection;
using System.Net.Mail;


namespace QCash.EStatement.NBL.Forms
{
    public partial class AddEmail : Form
    {
        private string _fiid = string.Empty;
        private string _EmailID = string.Empty;
        private ConnectionStringBuilder ConStr = null;
        private SqlDbProvider objProvider = null;
        string reply = string.Empty;
        string sql = string.Empty;
        public AddEmail(string fiid)
        {
            InitializeComponent();
            _fiid = fiid;
        }

        private bool IsValid(string emailaddress)
        {
            if (!string.IsNullOrEmpty(emailaddress))
            {
                try
                {
                    MailAddress m = new MailAddress(emailaddress);
                    return true;
                }
                catch (FormatException ex)
                {

                    MessageBox.Show(ex.Message);
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private void btnEmailAdd_Click(object sender, EventArgs e)
        {
           
            

            try
            {
                ConStr = new ConnectionStringBuilder(1);
                objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);

                if (txtemailid.Text.ToString().Trim()== "")
                {
                    MessageBox.Show("Please Add Email Address.");
                    txtemailid.Focus();
                    return;
                }

                if (IsValid(txtemailid.Text.ToString().Trim()))
                {
                    sql = "INSERT INTO EMAIL_ADDRESS(MAILADDRESS) VALUES('" + txtemailid.Text.ToString().Trim() + "')";

                    reply = objProvider.RunQuery(sql);

                    MessageBox.Show("Email ID Added Successfully.");
                    txtemailid.Clear();
                    txtemailid.Focus();
                }
                else 
                {
                    MessageBox.Show("Invalid Email ID.");
                    txtemailid.Focus();
                }
               
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
