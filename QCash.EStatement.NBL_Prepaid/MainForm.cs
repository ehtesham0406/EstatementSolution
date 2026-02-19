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

namespace StatementGenerator
{
    public partial class MainForm : Form
    {
        private int numForms;
        private int newFormNum;
        private string FIID = string.Empty;

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

        //void mainToolbarsManager_ToolClick(object sender, ToolClickEventArgs e)
        //{
        //    switch (e.Tool.Key)
        //    {

        //        case "Create":
        //            {
        //                AddEStatementForm(FIID);
        //                break;
        //            }
        //        case "Configuration":
        //            {
        //                AddConfigurationForm(FIID);
        //                break;
        //            }
        //        case "DataMaintain":
        //            {
        //                DataMaintainForm(FIID);
        //                break;
        //            }
        //        case "DBSetup":
        //            {
        //                DatabaseSetupForm(FIID);
        //                break;
        //            }
        //        case "SentStatus":
        //            {
        //                AddReportForm(FIID);
        //                break;
        //            }
        //        case "Exit":
        //            {
        //                Application.Exit();
        //                break;
        //            }
        //    }

        //}
        ////
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

        private EStatementGenerator AddEStatementForm(string _fiid)
        {
            EStatementGenerator newForm = new EStatementGenerator(_fiid);
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

    }
}
