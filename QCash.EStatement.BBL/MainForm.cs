using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Infragistics.Win.UltraWinToolbars;
using Infragistics.Win.UltraWinTabControl;
using Infragistics.Win.UltraWinTabs;
using QCash.EStatement;
using System.Configuration;
using System.IO;
using System.Collections;
using System.Common;
using System.Connection;
using QCash.EStatement.BBL.App_Code;


namespace StatementGenerator
{
    public partial class MainForm : Form
    {
        //private int numForms;
        //private int newFormNum;
        private string FIID = string.Empty;
        private string CardMailerPath = string.Empty;
        private SqlDbProvider objProvider = null;
        private ConnectionStringBuilder ConStr = null;

        public MainForm()
        {
            InitializeComponent();

            //this.mainToolbarsManager.ToolClick += new ToolClickEventHandler(mainToolbarsManager_ToolClick);
            this.Load += new EventHandler(MainForm_Load);
            this.tsmExit.Click += new EventHandler(tsmExit_Click);
            this.tsmProcess.Click += new EventHandler(tsmProcess_Click);
            this.tsmArchieve.Click += new EventHandler(tsmArchieve_Click);
            //
            this.tsmSMTP.Click += new EventHandler(tsmSMTP_Click);
            this.tsmDatabase.Click += new EventHandler(tsmDatabase_Click); 
            //
            this.tsmSentStatus.Click += new EventHandler(tsmSentStatus_Click);
          
           
        }
                

        void tsmSentStatus_Click(object sender, EventArgs e)
        {
            AddReportForm(FIID);
        }
       

        void tsmDatabase_Click(object sender, EventArgs e)
        {
            DatabaseSetupForm(FIID);
        }

        void tsmSMTP_Click(object sender, EventArgs e)
        {
            AddConfigurationForm(FIID);
        }

        void tsmArchieve_Click(object sender, EventArgs e)
        {
            DataMaintainForm(FIID);
        }

        void tsmProcess_Click(object sender, EventArgs e)
        {
            AddEStatementForm(FIID);
        }

        void tsmExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
      
        void MainForm_Load(object sender, EventArgs e)
        {
            FIID = ConfigurationManager.AppSettings["FIID"].ToString();
        }

 
        private DatabaseSetup DatabaseSetupForm(string _fiid)
        {
            DatabaseSetup newForm = new DatabaseSetup();
            Form[] _forms = this.MdiChildren;

            bool flag = IfExistForm(_forms, newForm);

            if (!flag)
            { // Add new form to MDI parent
                newForm.MdiParent = this;
                newForm.Show();
            }
            return newForm;
        }
        //
        private DataMaintain DataMaintainForm(string _fiid)
        {
            DataMaintain newForm = new DataMaintain(_fiid);
            Form[] _forms = this.MdiChildren;

            bool flag = IfExistForm(_forms, newForm);

            if (!flag)
            { // Add new form to MDI parent
                newForm.MdiParent = this;
                newForm.Show();
            }
            return newForm;
        }
        //
        private SMTPConfiguration AddConfigurationForm(string _fiid)
        {
            SMTPConfiguration newForm = new SMTPConfiguration();
            Form[] _forms = this.MdiChildren;

            bool flag = IfExistForm(_forms, newForm);

            if (!flag)
            { // Add new form to MDI parent
                newForm.MdiParent = this;
                newForm.Show();
            }
            return newForm;
        }

        private StatementGenerator AddEStatementForm(string _fiid)
        {
            StatementGenerator newForm = new StatementGenerator(_fiid);
            Form[] _forms = this.MdiChildren;

            bool flag = IfExistForm(_forms, newForm);

            if (!flag)
            { // Add new form to MDI parent
                newForm.MdiParent = this;
                newForm.Show();
            }
            return newForm;
        }

        private ReportViewer AddReportForm(string _fiid)
        {
            ReportViewer newForm = new ReportViewer(_fiid);
            Form[] _forms = this.MdiChildren;

            bool flag = IfExistForm(_forms, newForm);

            if (!flag)
            { // Add new form to MDI parent
                newForm.MdiParent = this;
                newForm.Show();
            }
            return newForm;
        }
       
        private bool IfExistForm(Form [] objForms, Form _form)
        {
            bool flag = false;
            for (int i = 0; i < objForms.Length; i++)
            {
                if (objForms[i].Text == _form.Text)
                {
                    flag = true;
                    
                    break;

                }
                else
                    flag = false;
            }
            return flag;
        }

        private void tsmSentStatus_Click_1(object sender, EventArgs e)
        {

        }

        public void readData(string filepath)
        {
            TextObj _TextObj = new TextObj();
            var fileNames = Directory.GetFiles(@"C:\Users\rabby\Desktop\AnalysisFile");
           
            const string lineToFind = "()";
            string[] insertData = new string[12];

            foreach (var fileName in fileNames)
            {
                
                int insertCount = 0;
                using (var reader = new StreamReader(fileName))
                {
                    string lineRead;
                    string linecheck = "";

                    while ((lineRead = reader.ReadLine()) != null)
                    {
                        #region If
                        if (lineRead != "")
                        {
                            linecheck = lineRead.Trim();
                            if (linecheck != lineToFind)
                            {

                                #region Switch
                                switch (insertCount)
                                {
                                    case 0:
                                        {
                                            _TextObj.IDClient = linecheck;
                                            insertCount++;
                                            break;
                                        }
                                    case 1:
                                        {
                                            _TextObj.PAN = linecheck;
                                            insertCount++;
                                            break;
                                        }
                                    case 2:
                                        {
                                            _TextObj.SDate = linecheck;
                                            insertCount++;
                                            break;
                                        }
                                    case 3:
                                        {
                                            _TextObj.Branch = linecheck;
                                            insertCount++;
                                            break;
                                        }
                                    case 4:
                                        {
                                            _TextObj.AmountLimit = linecheck;
                                            insertCount++;
                                            break;
                                        }
                                    case 5:
                                        {
                                            _TextObj.Client = linecheck;
                                            insertCount++;
                                            break;
                                        }
                                    case 6:
                                        {
                                            _TextObj.CardType = linecheck;
                                            insertCount++;
                                            break;
                                        }
                                    case 7:
                                        {
                                            _TextObj.Code = linecheck;
                                            insertCount++;
                                            break;
                                        }
                                    case 8:
                                        {
                                            _TextObj.Address1 = linecheck;
                                            insertCount++;
                                            break;
                                        }
                                    case 9:
                                        {
                                            _TextObj.Address2 = linecheck;
                                            insertCount++;
                                            break;
                                        }
                                    case 10:
                                        {
                                            _TextObj.Country = linecheck;
                                            insertCount++;
                                            break;
                                        }
                                    case 11:
                                        {
                                            _TextObj.Mobile = linecheck;
                                            insertCount++;
                                            break;
                                        }


                                }

                                #endregion

                            }
                            else
                            {
                                //insertDatabase
                                string sql = "Insert into CardMailerInfo(Information) " +
                          "values('" + _TextObj.IDClient + "')";
                                string reply = objProvider.RunQuery(sql);
                                _TextObj = new TextObj();
                                insertCount = 0;
                               // line++;
                               // line++;


                            }
                           // line++;
                        }
                        #endregion
                    }



                }


            }


        }

    }
}
