using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using StatementGenerator.App_Code;
using FlexiStar.Utilities;
using System.Connection;
using System.Common;
using System.Threading;
using System.Configuration;
using Infragistics.Win.UltraWinGrid;

namespace QCash.EStatement
{
    public partial class DataMaintain : Form
    {
        private ConnectionStringBuilder ConStr = null;
        private SqlDbProvider objProvider = null;

        //
        delegate void SetTextCallback(string text);
        private SetTextCallback _addText = null;
        //
        private string Bank_Code = string.Empty;
        private string _LogPath = string.Empty;
        private string StmDate = string.Empty;

        private string _XLSourcePath = string.Empty;

        private System.Drawing.Printing.PrintDocument c_pdSetup = null;

        Thread tdSend = null;
        bool stopSend = false;

        private string _fiid = string.Empty;

        public DataMaintain(string fiid)
        {
            InitializeComponent();

            _addText = new SetTextCallback(Output);

            this.btnSubmit.Click += new EventHandler(btnSubmit_Click);
            this.Load += new EventHandler(DataMaintain_Load);
            this.btnArchive.Click += new EventHandler(btnArchive_Click);
            this.btnExport.Click += new EventHandler(btnExport_Click);
            this.btnDelete.Click += new EventHandler(btnDelete_Click);
            this.btnClose.Click += new EventHandler(btnClose_Click);

            this.grdEmailData.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(grdEmailData_InitializeLayout);
            this.grdEmailData.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(grdEmailData_CellChange);
            this.grdEmailData.DoubleClickHeader += new Infragistics.Win.UltraWinGrid.DoubleClickHeaderEventHandler(grdEmailData_DoubleClickHeader);

            this.rbAll.Click += new EventHandler(rbAll_Click);
            this.rbDatewise.Click += new EventHandler(rbDatewise_Click);

            _fiid = fiid;
        }

        void rbDatewise_Click(object sender, EventArgs e)
        {
            rbAll.Checked = false;
            rbDatewise.Checked = true;
            dtpStmDate.Enabled = true;
        }

        void rbAll_Click(object sender, EventArgs e)
        {
            rbAll.Checked = true;
            rbDatewise.Checked = false;
            dtpStmDate.Enabled = false;
        }

        void DataMaintain_Load(object sender, EventArgs e)
        {
            _XLSourcePath = ConfigurationManager.AppSettings[3].ToString();
            c_pdSetup = new System.Drawing.Printing.PrintDocument();
            //
            this.grdEmailData.Text = "EStatement Information for " + _fiid + " Cardholders";
        }

        private void Output(string text)
        {
            try
            {
                if (text != "")
                {
                    if (text.Contains('\0'))
                    {
                        text.Replace("\0", "\r\n");
                    }
                    txtAnalyzer.AppendText(text);
                    txtAnalyzer.AppendText("\r\n");
                    txtAnalyzer.ScrollBars = ScrollBars.Both;
                    txtAnalyzer.WordWrap = false;
                }
                else
                    txtAnalyzer.Text = text;
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
            }
        }

        void grdEmailData_DoubleClickHeader(object sender, Infragistics.Win.UltraWinGrid.DoubleClickHeaderEventArgs e)
        {
            if (e.Header.Column.Key == "Checks")
            {
                foreach (UltraGridRow aRow in grdEmailData.Rows)
                {
                    if (aRow.Cells["Checks"].Value == null)
                        aRow.Cells["Checks"].Value = true;
                    else if (aRow.Cells["Checks"].Value.ToString() == "True")
                        aRow.Cells["Checks"].Value = false;
                    else if (aRow.Cells["Checks"].Value.ToString() == "False")
                        aRow.Cells["Checks"].Value = true;
                }
            }
        }

        void grdEmailData_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
        {
            if (StringComparer.Ordinal.Equals(e.Cell.Column.Key, @"Checks"))
            {
                if (e.Cell.Value == null)
                    e.Cell.Value = true;
                else if (e.Cell.Value.ToString() == "True")
                    e.Cell.Value = false;
                else if (e.Cell.Value.ToString() == "False")
                    e.Cell.Value = true;
            }
            else return;
        }

        void grdEmailData_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
            e.Layout.Bands[0].ColHeaderLines = 2;

            //
            UltraGridLayout layout = e.Layout;
            UltraGridOverride ov = layout.Override;
            ov.FilterUIType = FilterUIType.FilterRow;
            ov.FilterEvaluationTrigger = FilterEvaluationTrigger.OnCellValueChange;
            //
            UltraGridColumn ugc = e.Layout.Bands[0].Columns.Add(@"Checks", "Select\nAll");
            ugc.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
            ugc.CellActivation = Activation.AllowEdit;
            ugc.Header.VisiblePosition = 0;
            ugc.Width = 50;
            //
            e.Layout.Bands[0].Columns["BANK_CODE"].Header.Caption = "Bank Code";
            e.Layout.Bands[0].Columns["BANK_CODE"].Width = 70;
            e.Layout.Bands[0].Columns["BANK_CODE"].CellActivation = Activation.ActivateOnly;
            //
            e.Layout.Bands[0].Columns["STMDATE"].Header.Caption = "Statement \nDate";
            e.Layout.Bands[0].Columns["STMDATE"].Width = 70;
            e.Layout.Bands[0].Columns["STMDATE"].CellActivation = Activation.ActivateOnly;
            //
            e.Layout.Bands[0].Columns["MONTH"].Header.Caption = "Statement \nMonth";
            e.Layout.Bands[0].Columns["MONTH"].Width = 70;
            e.Layout.Bands[0].Columns["MONTH"].CellActivation = Activation.ActivateOnly;
            //
            e.Layout.Bands[0].Columns["YEAR"].Header.Caption = "Year";
            e.Layout.Bands[0].Columns["YEAR"].Width = 50;
            e.Layout.Bands[0].Columns["YEAR"].CellActivation = Activation.ActivateOnly;
            e.Layout.Bands[0].Columns["YEAR"].Hidden = true;
            //
            e.Layout.Bands[0].Columns["PAN_NUMBER"].Header.Caption = "PAN \nNumber";
            e.Layout.Bands[0].Columns["PAN_NUMBER"].Width = 120;
            e.Layout.Bands[0].Columns["PAN_NUMBER"].CellActivation = Activation.ActivateOnly;
            //
            e.Layout.Bands[0].Columns["MAILADDRESS"].Header.Caption = "Mail\nAddress";
            e.Layout.Bands[0].Columns["MAILADDRESS"].Width = 130;
            e.Layout.Bands[0].Columns["MAILADDRESS"].CellActivation = Activation.AllowEdit;
            //
            e.Layout.Bands[0].Columns["FILE_LOCATION"].Header.Caption = "File \nLocation";
            e.Layout.Bands[0].Columns["FILE_LOCATION"].Width = 100;
            e.Layout.Bands[0].Columns["FILE_LOCATION"].CellActivation = Activation.ActivateOnly;
            //
            e.Layout.Bands[0].Columns["MAILSUBJECT"].Header.Caption = "Mail \nSubject";
            e.Layout.Bands[0].Columns["MAILSUBJECT"].Width = 100;
            e.Layout.Bands[0].Columns["MAILSUBJECT"].CellActivation = Activation.ActivateOnly;
            //
            e.Layout.Bands[0].Columns["MAILBODY"].Header.Caption = "Mail \nBody";
            e.Layout.Bands[0].Columns["MAILBODY"].Width = 100;
            e.Layout.Bands[0].Columns["MAILBODY"].CellActivation = Activation.ActivateOnly;
            //
            e.Layout.Bands[0].Columns["STATUS"].Header.Caption = "Status";
            e.Layout.Bands[0].Columns["STATUS"].Width = 50;
            e.Layout.Bands[0].Columns["STATUS"].CellActivation = Activation.AllowEdit;
        }

        void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        void btnExport_Click(object sender, EventArgs e)
        {
            
        }

        void btnDelete_Click(object sender, EventArgs e)
        {
            
        }

        void btnArchive_Click(object sender, EventArgs e)
        {
            string reply = string.Empty;
            EStatementManager.Instance().ArchiveEStatement(ref reply);

            if (reply.Contains("Error"))
            {
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", reply);
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh24:mm:ss") + " : " + reply });
            }
        }

        void btnSubmit_Click(object sender, EventArgs e)
        {
            string reply = string.Empty;
            try
            {
                if (StmDate == "")
                    StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy");
                else StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy");

                MsgLogWriter objLW = new MsgLogWriter();

                EStatementList objESList = EStatementManager.Instance().GetAllEStatements(_fiid, StmDate, "3", ref reply);
                if (objESList != null)
                {
                    if (objESList.Count > 0)
                    {
                        grdEmailData.DataSource = objESList;
                    }
                }
            }
            catch (Exception ex)
            {
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh24:mm:ss") + " : " + ex.Message });
            }
        }
    }
}
