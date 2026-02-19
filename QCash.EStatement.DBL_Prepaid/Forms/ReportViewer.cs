using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FlexiStar.Utilities;
using StatementGenerator.App_Code;
using System.Net.Mail;
using System.Threading;
using System.Connection;
using System.Common;
using Infragistics.Win.UltraWinGrid;
using Infragistics.Win;
using Infragistics.Shared;
using QCash.EStatement;
using System.Configuration;
using FlexiStar.Utilities.EncryptionEngine;
using System.IO;
namespace StatementGenerator
{
    public partial class ReportViewer : Form
    {
        private ConnectionStringBuilder ConStr = null;
        private SqlDbProvider objProvider = null;
        public static string  FiD = string.Empty;
        EStatementGenerator eg = new EStatementGenerator(FiD);

        //
        delegate void SetTextCallback(string text);
        private SetTextCallback _addText = null;
        //
        private string Bank_Code = string.Empty;
        private string _LogPath = string.Empty;
        private string StmDate = string.Empty;

        private string _XLSourcePath = string.Empty;
        private string _AdditionalAttachment = string.Empty;
        private string _Mail = string.Empty;

        private System.Drawing.Printing.PrintDocument c_pdSetup = null;

        Thread tdSendMail = null;
        bool stopSend = false;
               
        private string _fiid = string.Empty;

        public ReportViewer(string fiid)
        {
            InitializeComponent();

            _addText = new SetTextCallback(Output);
            //
            this.btnSearch.Click += new EventHandler(btnSearch_Click);
            this.btnSendMail.Click += new EventHandler(btnSendMail_Click);

            this.btnClose.Click += new EventHandler(btnClose_Click);
            this.Load += new EventHandler(ReportViewer_Load);
            this.btnExport.Click += new EventHandler(btnExport_Click);
            this.btnPrint.Click += new EventHandler(btnPrint_Click);

            this.grdEmailData.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(grdEmailData_InitializeLayout);
            this.grdEmailData.CellChange += new CellEventHandler(grdEmailData_CellChange);
            this.grdEmailData.DoubleClickHeader += new DoubleClickHeaderEventHandler(grdEmailData_DoubleClickHeader);

            _fiid = fiid;
        }

        void btnPrint_Click(object sender, EventArgs e)
        {
            ////display grid in Print Preview
            //grdEmailData.PrintPreview(UltraGrid1.DisplayLayout, c_pdSetup);
            //print grid to default printer
            grdEmailData.Print(grdEmailData.DisplayLayout, c_pdSetup);
        }

        void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                MsgLogWriter objLW = new MsgLogWriter();
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + "Exporting EStatement Data..." });
                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Exporting EStatement Data...");
                
                string file_path = _XLSourcePath + "\\" + System.DateTime.Now.ToString("dd.MM.yyyy") + "_GridData.xls";
                this.excelExporter.Export(this.grdEmailData, file_path);

                //mailProgress.PerformStep();
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + "Export Complete, File Location " + file_path });
                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Export Complete, File Location " + file_path);
                
                
            }
            catch (Exception ex)
            {
               txtAnalyzer.Invoke(_addText, new object[] { ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
            }
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

        void grdEmailData_DoubleClickHeader(object sender, DoubleClickHeaderEventArgs e)
        {
            if (e.Header.Column.Key == "Checks")
            {
                bool isSet = false, setValue = false;

                foreach (UltraGridRow aRow in grdEmailData.Rows)
                {
                    if (aRow.IsFilteredOut)
                    {
                        aRow.Cells["Checks"].Value = false;
                    }
                    if (!aRow.IsFilteredOut)
                    {
                        if (!isSet)
                        {
                            isSet = true;
                            try
                            {
                                setValue = aRow.Cells["Checks"].Value.ToString() == "False";
                            }
                            catch
                            {
                                setValue = true;
                            }
                        }
                        aRow.Cells["Checks"].Value = setValue;
                    }
                }
            }
        }

        void grdEmailData_CellChange(object sender, CellEventArgs e)
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

        void ReportViewer_Load(object sender, EventArgs e)
        {
            _XLSourcePath = ConfigurationManager.AppSettings[3].ToString();
            _AdditionalAttachment = ConfigurationManager.AppSettings[8].ToString();
            c_pdSetup = new System.Drawing.Printing.PrintDocument();
            _LogPath = ConfigurationManager.AppSettings[5].ToString();
            //
            this.grdEmailData.Text = "EStatement Information for " + _fiid + " Cardholders";
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

        void btnSendMail_Click(object sender, EventArgs e)
        {

            tdSendMail = new Thread(new ThreadStart(SendEmail));
            tdSendMail.IsBackground = true;
            tdSendMail.Start();


        }

        void btnSearch_Click(object sender, EventArgs e)
        {
            string reply = string.Empty;
            try
            {
                //if (StmDate == "")
                //{
                //    string p = dtpStmDate.Value.ToString();
                //    //StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy");
                //    StmDate = getNumberFormat(p);
                //}
                //else StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy"); objOp.OpDate = getNumberFormat(objOp.OpDate);
               // else StmDate = getNumberFormat(dtpStmDate.Value.ToString());

              //  string curdate = eg.getNumberFormat1(dtpStmDate.Value.ToString());

               // string Month = curdate.Split('-')[1].ToString();
              //  string Year = curdate.Split('-')[2].ToString();
              //  string StmDate = "01" + '-' + Month + '-' + Year;
                string StmDate = getNumberFormat(dtpStmDate.Value.ToString());

                MsgLogWriter objLW = new MsgLogWriter();

                EStatementList objESList = EStatementManager.Instance().GetAllEStatements(_fiid, StmDate, "2", ref reply);
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

        void SendEmail()
        {
            string reply = string.Empty;
            try
            {
                EStatementList objESList = new EStatementList();


                foreach (UltraGridRow aRow in grdEmailData.Rows)
                {
                    EStatementInfo objES = null;
                    if (aRow.Cells["Checks"].Value != null)
                    {
                        if (aRow.Cells["Checks"].Value.ToString() == "True")
                        {
                            objES = (EStatementInfo)aRow.ListObject;
                            objESList.Add(objES);
                        }
                    }
                }
                if (StmDate == "")
                    //StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy");
                StmDate = getNumberFormat(dtpStmDate.Value.ToString());

                MsgLogWriter objLW = new MsgLogWriter();

                if (objESList != null)
                {
                    if (objESList.Count > 0)
                    {
                        SmtpConfigurationManager objSmtpMan = new SmtpConfigurationManager();
                        SmtpConfigurationList objSmtpList = new SmtpConfigurationList();

                        Encryption objEnc = new Encryption();

                        objSmtpList = objSmtpMan.GetSmtpConfiguration(_fiid, 1);

                        if (objSmtpList != null)
                        {
                            if (objSmtpList.Count > 0)
                            {
                                int count = 0;
                                for (int i = 0; i < objESList.Count; i++)
                                {
                                    string email = objESList[i].MAILADDRESS;
                                    //if (objESList[i].MAILADDRESS != "")
                                    if (!string.IsNullOrEmpty(email) && IsValid(email))
                                    {
                                        try
                                        {
                                            MailMessage mail = new MailMessage();
                                            mail.From = new MailAddress(objSmtpList[0].From_Address);
                                            mail.Subject = objESList[i].MAILSUBJECT;
                                            mail.Body = objESList[i].MAILBODY;
                                            mail.To.Add(objESList[i].MAILADDRESS.Trim());
                                            System.Net.Mail.Attachment attachment;
                                            attachment = new System.Net.Mail.Attachment(objESList[i].FILE_LOCATION);
                                            mail.Attachments.Add(attachment);

                                            //=-=-=-=-=-=-=-=-=-=-=-=-=--=--=-=-=-=-=-=
                                            _Mail = ConfigurationManager.AppSettings["Mail"].ToString();
                                            StreamReader reader = new StreamReader(_Mail + @"\\Template.html");
                                            string readFile = reader.ReadToEnd();
                                            string myString = "";
                                            myString = readFile;
                                            mail.Body = myString;

                                            mail.IsBodyHtml = true;
                                            //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=

                                            //For Additional Attatchment
                                            string[] filePaths = Directory.GetFiles(_AdditionalAttachment);
                                            if (filePaths.Length != 0)
                                            {
                                                for (int x = 0; x < filePaths.Length; x++)
                                                {
                                                    attachment = new System.Net.Mail.Attachment(filePaths[x]);
                                                    mail.Attachments.Add(attachment);
                                                }
                                            }

                                            SmtpClient SmtpServer = new SmtpClient(objSmtpList[0].Smtp_Server);
                                            SmtpServer.Port = objSmtpList[0].Smtp_Port;
                                            SmtpServer.Credentials = new System.Net.NetworkCredential(objSmtpList[0].From_User, objEnc.DecryptWord(objSmtpList[0].From_Password));
                                            SmtpServer.EnableSsl = Convert.ToBoolean(objSmtpList[0].EnableSSL);

                                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + "Sending EStatement to " + mail.To.ToString() }); ;
                                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Sending EStatement " + mail.To.ToString());

                                            SmtpServer.Send(mail);

                                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + "mail Send to " + mail.To.ToString() }); ;
                                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : mail Send to " + mail.To.ToString());


                                            objESList[i].STATUS = "0";  // Statement Generated and Mail Sent Successfully
                                            EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                            count++;
                                        }
                                        catch (Exception ex)
                                        {
                                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + ex.Message });
                                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + ex.Message);

                                            objESList[i].STATUS = "8"; // Estatement Generated and mail sent but no acknowledged received from mail server.
                                            EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                        }
                                    }
                                    else
                                    {
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + "Invalid or No Mail Address Found to send the Estatement " + objESList[i].FILE_LOCATION + " " + " PAN : " + objESList[i].PAN_NUMBER + " and Email : " + email }); ;
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Invalid or No Mail Address Found to send the Estatement " + objESList[i].FILE_LOCATION + " " + " PAN : " + objESList[i].PAN_NUMBER + " and Email : " + email);

                                        objESList[i].STATUS = "2";  // No Mail Address Found
                                        EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                    }
                                }
                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " e-statements have been mailed out of " + objESList.Count + "." });
                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " e-statements have been mailed out of " + objESList.Count + ".");
                            
                            }
                        }
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

        private void txtAnalyzer_TextChanged(object sender, EventArgs e)
        {

        }
        public string getNumberFormat(string vDate)
        {
            string[] omitSpace = vDate.Split(' ');
            string[] date = omitSpace[0].Split('/');
            DateTime dt = new DateTime(Int32.Parse(date[2]), Int32.Parse(date[0]), Int32.Parse(date[1]));
            string formatedDate = string.Format("{0:dd-MMM-yyyy}", dt);
            return formatedDate;
        }
        private bool IsValid(string emailAddress)
        {
            if (string.IsNullOrEmpty(emailAddress) || emailAddress.Trim().Length == 0)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: The email address is either null or empty" });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", "Error: The email address is either null or empty");
                return false;
            }

            // Check if the email address is 'na@na.na'
            if (emailAddress.Trim().ToUpper() == "NA@NA.NA")
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: The email address is 'na@na.na' and is not valid" });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", "Error: The email address is 'na@na.na' and is not valid");
                return false;
            }

            // Regular expression to strictly validate email addresses
            //string emailRegex = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
            // if (!Regex.IsMatch(emailAddress, emailRegex))
            //{
            //txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: Email is null or empty" });
            // MsgLogWriter objLW = new MsgLogWriter();
            //objLW.logTrace(_LogPath, "EStatement.log", "Email contains invalid characters or is in incorrect format.");
            //return false;
            //}

            try
            {
                var mailAddress = new MailAddress(emailAddress);
                return true;
            }
            catch (FormatException ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return false;
            }
        }

        
    }
}
