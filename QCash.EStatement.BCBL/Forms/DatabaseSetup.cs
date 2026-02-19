using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using FlexiStar.Utilities.EncryptionEngine;

namespace QCash.EStatement
{
    public partial class DatabaseSetup : Form
    {
        private string configPath = string.Empty;
        private enum eProvider : short { Oledb, SQlClient, ODBC, OracleClient, OledbMSSQL, OledbOracle, OledbMSAcess2000 };
        private const string XMLFileName = "DBConfig";
        private DataSet dsDB;

        public DatabaseSetup()
        {
            InitializeComponent();

            this.Load += new EventHandler(DatabaseSetup_Load);
            this.grdDBSetup.InitializeLayout +=new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(grdDB_InitializeLayout);
            this.btnSave.Click +=new EventHandler(btnSave_Click);
            this.btnCancel.Click +=new EventHandler(btnCancel_Click);
            this.btnClose.Click += new EventHandler(btnClose_Click);
        }

        void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        void DatabaseSetup_Load(object sender, EventArgs e)
        {
            try
            {
                this.loadInitialData();
                this.fillCombo();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information");
            }
        }
        //
        private void fillCombo()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("DBTypeID", typeof(short));
            dt.Columns.Add("DBType", typeof(string));
            dt.Rows.Add(new object[] { eProvider.SQlClient, eProvider.SQlClient.ToString() });
            this.drpProvider.DataSource = dt;
            this.drpProvider.ValueMember = "DBTypeID";
            this.drpProvider.DisplayMember = "DBType";
            this.drpConnection.DataSource = new string[] { "ES" };
        }
        //
        private void getXMLFileData()
        {
            try
            {
                this.configPath = AppDomain.CurrentDomain.BaseDirectory;
                this.dsDB = new DataSet("dsDB");
                this.dsDB.ReadXmlSchema(configPath + XMLFileName + ".xsd");
                this.dsDB.ReadXml(configPath + XMLFileName + ".xml");
                if (this.dsDB.Tables.Count > 0)
                {
                    this.dsDB.Tables[0].TableName = "DB";
                    this.dsDB.Tables["DB"].Columns["IsDefault"].DefaultValue = false;
                    this.dsDB.Tables["DB"].Columns["Provider"].DefaultValue = eProvider.SQlClient;
                    if (!this.dsDB.Tables["DB"].Columns.Contains("Test"))
                    {
                        this.dsDB.Tables["DB"].Columns.Add(new DataColumn("Test", typeof(string)));
                    }
                    foreach (DataRow drDB in this.dsDB.Tables["DB"].Rows)
                    {
                        this.decryptRow(drDB);
                    }
                    this.dsDB.AcceptChanges();
                    grdDBSetup.DataSource = this.dsDB;
                }
                else
                {
                    throw new Exception(this.configPath + " does not contain any DBConfig.XML file.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        //
        private void decryptRow(DataRow drDB)
        {
            Encryption objEncrypt = new Encryption();
            string strUserID = "";
            if (drDB["UserName"].ToString().Length > 0)
            {
                strUserID = drDB["UserName"].ToString();
            }
            if (drDB["ServerName"].ToString().Trim().Length == 0)
            {
                drDB["ServerName"] = "(local)";
            }
            else
            {
                drDB["ServerName"] = objEncrypt.DecryptWord(drDB["ServerName"].ToString(), strUserID);
            }
            if (drDB["Password"].ToString().Length > 0)
            {
                drDB["Password"] = objEncrypt.DecryptWord(drDB["Password"].ToString(), strUserID);
            }
            if (drDB["DatabaseName"].ToString().Length > 0)
            {
                drDB["DatabaseName"] = objEncrypt.DecryptWord(drDB["DatabaseName"].ToString(), strUserID);
            }
            drDB["Test"] = "Test";
        }
        //
        private void encryptRow(DataRow drDB)
        {
            Encryption objEncrypt = new Encryption();
            string strUserID = "";
            if (drDB["UserName"].ToString().Length > 0)
            {
                strUserID = drDB["UserName"].ToString();
            }
            if (drDB["ServerName"].ToString().Trim().Length > 0)
            {
                drDB["ServerName"] = objEncrypt.EncryptWord(drDB["ServerName"].ToString(), strUserID);
            }
            drDB["DatabaseName"] = objEncrypt.EncryptWord(drDB["DatabaseName"].ToString(), strUserID);
            if (drDB["Password"].ToString().Length > 0)
            {
                drDB["Password"] = objEncrypt.EncryptWord(drDB["Password"].ToString(), strUserID);
            }
        }
        //
        private bool checkPree()
        {
            foreach (DataRow drDB in this.dsDB.Tables["DB"].Rows)
            {
                if (drDB.RowState == DataRowState.Deleted) continue;
                if (drDB["DatabaseName"].ToString().Trim().Length == 0)
                {
                    MessageBox.Show("DataBase name can't be empty.", "Information");
                    return false;
                }
                if (drDB["ConnectionName"].ToString().Trim().Length == 0)
                {
                    MessageBox.Show("Connection Name name can't be empty.", "Information");
                    return false;
                }
                if (drDB["Provider"].ToString().Trim().Length == 0)
                {
                    MessageBox.Show("Provider Name name can't be empty.", "Information");
                    return false;
                }
                this.encryptRow(drDB);
            }
            if (this.dsDB.Tables["DB"].Columns.Contains("Test"))
            {
                this.dsDB.Tables["DB"].Columns.Remove("Test");
            }
            return true;
        }
        //
        private void btnSave_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (!this.checkPree())
                    return;
                this.dsDB.WriteXmlSchema(configPath + XMLFileName + ".xsd");
                this.dsDB.WriteXml(configPath + XMLFileName + ".xml");
                this.loadInitialData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information");
            }
        }
        //
        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            this.loadInitialData();
        }
        //
        private void grdDB_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
            e.Layout.Bands[0].Columns["ConnectionName"].ValueList = this.drpConnection;
            e.Layout.Bands[0].Columns["ConnectionName"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            e.Layout.Bands[0].Columns["Provider"].ValueList = this.drpProvider;
            e.Layout.Bands[0].Columns["Provider"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
            e.Layout.Bands[0].Columns["Provider"].CellActivation = Activation.NoEdit;
            //this.Initializer.makeCellAsButton(e.Layout.Bands[0].Columns["Test"]);
            e.Layout.Bands[0].Columns["Test"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
            e.Layout.Bands[0].Columns["Test"].ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.OnMouseEnter;
            e.Layout.Bands[0].Columns["Test"].CellButtonAppearance.Cursor = Cursors.Hand;
            //
            Infragistics.Win.UltraWinEditors.UltraTextEditor txtPassword = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
            txtPassword.PasswordChar = '*';
            txtPassword.Text = string.Empty;
            e.Layout.Bands[0].Columns["Password"].EditorControl = txtPassword;
            e.Layout.Bands[0].Columns["ConnectionName"].Header.Caption = "Connection";
            e.Layout.Bands[0].Columns["ServerName"].Header.Caption = "SQL Server";
            e.Layout.Bands[0].Columns["DatabaseName"].Header.Caption = "Database";
            e.Layout.Bands[0].Columns["IsDefault"].Header.Caption = "Default";
            e.Layout.Bands[0].Columns["IsDefault"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
            e.Layout.Bands[0].Columns["Test"].Header.Caption = "";
            e.Layout.Bands[0].Columns["Test"].Width = 50;
            e.Layout.Bands[0].Columns["Test"].Hidden = true;
            e.Layout.Override.AllowAddNew = AllowAddNew.TemplateOnBottom;
        }
		//
        private void loadInitialData()
        {
            this.getXMLFileData();
        }
    }
}
