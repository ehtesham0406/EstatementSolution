using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Connection;
using System.Common;
using CrystalDecisions.Shared;
using System.Configuration;
using System.Net.Mail;
using StatementGenerator.App_Code;
using System.IO;
using Infragistics.Win.UltraWinTabControl;
using System.Threading;
using FlexiStar.Utilities;
using CrystalDecisions.CrystalReports.Engine;
using FlexiStar.Utilities.EncryptionEngine;
using QCash.EStatement.MMBL_Prepaid.Reports;
using PdfSharp.Pdf.Security;
using PdfSharp.Pdf.IO;
using System.Globalization;

// SERACH FOR  CHANGE_HERE FOR MODIFY

namespace StatementGenerator
{
    public partial class EStatementGenerator : Form
    {
        private ConnectionStringBuilder ConStr = null;
        private SqlDbProvider objProvider = null;

        //
        delegate void SetTextCallback(string text);
        private SetTextCallback _addText = null;
        //
        private string Bank_Code = string.Empty;
        private string _LogPath = string.Empty;
        private string _XMLProcessedPath = string.Empty;
        private string _XMLSourcePath = string.Empty;
        private string _EStatementProcessedPath = string.Empty;
        private string _AdditionalAttachment = string.Empty;
        private string StartDate = string.Empty;
        private string EndDate = string.Empty;
        private string stmMessage = string.Empty;

        int pdfCount = 0; //pdf counter

        Thread tdGenerate = null;
        Thread tdSendMail = null;
        string prePan = string.Empty;
        string preDoc = string.Empty;

        private string _fiid = string.Empty;

        public EStatementGenerator(string fiid)
        {
            InitializeComponent();

            _addText = new SetTextCallback(Output);

            this.Load += new EventHandler(ReportViewer_Load);
            this.btnLoad.Click += new EventHandler(btnLoad_Click);
            this.btnGenerate.Click += new EventHandler(btnGenerate_Click);
            this.btnSendMail.Click += new EventHandler(btnSendMail_Click);
            this.btnClose.Click += new EventHandler(btnClose_Click);
            //
            btnGenerate.Visible = false;

            _fiid = fiid;
        }

        void btnClose_Click(object sender, EventArgs e)
        {
            if (tdGenerate != null)
            {
                if (tdGenerate.ThreadState == ThreadState.Background)
                {
                    tdGenerate.Abort();
                    Thread.Sleep(1000);
                    this.Close();
                }
                else
                {
                    tdGenerate = null;
                    this.Close();
                }
            }
            else if (tdSendMail != null)
            {
                if (tdSendMail.ThreadState == ThreadState.Background)
                {
                    tdSendMail.Abort();
                    Thread.Sleep(1000);
                    this.Close();
                }
                else
                {
                    tdSendMail = null;
                    this.Close();
                }
            }
            else
                this.Close();
        }

        void btnLoad_Click(object sender, EventArgs e)
        {
            stmMessage = txtStmMsg.Text;
            btnLoad.Enabled = false;
            tdGenerate = new Thread(new ThreadStart(GenerateEStatement));
            tdGenerate.IsBackground = true;
            tdGenerate.Start();

            
        }
        private void GenerateEStatement()
        {
            if (txtEmailSubject.Text != "")
            {
                if (txtEmailBody.Text != "")
                {
                    string reply = string.Empty;
                    EStatementManager.Instance().ArchiveEStatement(ref reply);

                    if (reply.Contains("Error"))
                    {
                        MsgLogWriter objLW = new MsgLogWriter();
                        objLW.logTrace(_LogPath, "EStatement.log", reply);
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + reply });
                    }
                    else if (reply == "Success")
                    {
                        MsgLogWriter objLW = new MsgLogWriter();
                        objLW.logTrace(_LogPath, "EStatement.log", "Successfully archived previous data !!!");
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + "Successfully archived previous data !!!" });

                        ProcessData();
                       // ProcessFolderFiles( "SIBL", "success");
                    }
                }
                else
                {
                    MessageBox.Show("Set Email Body", "Warning !!!");
                }
            }
            else
            {
                MessageBox.Show("Set Email Subject", "Warning !!!");
            }
        }
        private void ProcessData()
        {
            string _bankCode = string.Empty;
            string _bankName = string.Empty;

            string _reply = string.Empty;

            DirectoryInfo di = new DirectoryInfo(_XMLSourcePath);
            DirectoryInfo[] dia = di.GetDirectories();

            for (int fcount = 0; fcount < dia.Length; fcount++)
            {
                if (dia[fcount].FullName.Contains("MMBL"))
                {
                    _bankName = "MMBL";
                    _bankCode = "7058";
                    _XMLSourcePath = dia[fcount].FullName;
                    
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }


                else
                {
                    MsgLogWriter objLW = new MsgLogWriter();
                    objLW.logTrace(_LogPath, "EStatement.log", "Not an CSV data !!!");
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + "Not an CSV data !!!" });
                }

                Bank_Code = _bankName;

            }


        }
        void btnSendMail_Click(object sender, EventArgs e)
        {
            btnSendMail.Enabled = false;

            tdSendMail = new Thread(new ThreadStart(SendMail));
            tdSendMail.IsBackground = true;
            tdSendMail.Start();
        }
        private void SendMail()
        {
            string reply = string.Empty;
            try
            {
                if (StartDate == "" && EndDate == "")
                {

                    StartDate = dtpStartDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                    EndDate = dtpEndDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                }
                
                else 
                    {
                        StartDate = dtpStartDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                        EndDate = dtpEndDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                    }


                MsgLogWriter objLW = new MsgLogWriter();

                EStatementList objESList = EStatementManager.Instance().GetAllEStatements(_fiid, StartDate,EndDate, "1", ref reply);
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
                                    if (objESList[i].MAILADDRESS != "")
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

                                            //attachment = new System.Net.Mail.Attachment(_AdditionalAttachment);
                                            //mail.Attachments.Add(attachment);
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

                                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + "Sending EStatement to " + mail.To.ToString() }); ;
                                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Sending EStatement " + mail.To.ToString());

                                            SmtpServer.Send(mail);

                                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + "mail Send to " + mail.To.ToString() }); ;
                                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : mail Send to " + mail.To.ToString());


                                            objESList[i].STATUS = "0";   // Mail Sent Successfully
                                            EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                            count++;
                                        }
                                        catch (Exception ex)
                                        {
                                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message);

                                            objESList[i].STATUS = "2"; // Mail is not Sent
                                            EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                        }
                                    }
                                    else
                                    {
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + "No Mail Address Found to send the Estatement " + objESList[i].FILE_LOCATION }); ;
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : No Mail Address Found to send the Estatement " + objESList[i].FILE_LOCATION);

                                        objESList[i].STATUS = "8";   //  No Mail Address Found
                                        EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                    }
                                }
                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Estatement has mailed out of " + objESList.Count + "." });
                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Estatement has mailed" + objESList.Count + ".");
                            }
                        }
                    }
                }
                else
                {
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : There is no Estatement has generate on selected  date." });
                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : There is no Estatement has generate on selected  date.");

                }
            }
            
            catch (Exception ex)
            {
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
            }
        }
        //
        void btnGenerate_Click(object sender, EventArgs e)
        {
            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
            string reply = string.Empty;
            MsgLogWriter objLW = new MsgLogWriter();

            DataTable dtCardbdt = new DataTable();
            dtCardbdt = objProvider.ReturnData("select * from Qry_Card_Account where Curr=50 order by Statementno  ASC", ref reply).Tables[0];

            if (dtCardbdt.Rows.Count > 0)
            {
                txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Estatement BDT." });
                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Estatement BDT.");

                //Process pdf for BDT
                ProcessStatementBDT(dtCardbdt);
            }

            DataTable dtCardusd = new DataTable();
            dtCardusd = objProvider.ReturnData("select * from Qry_Card_Account where Curr=840 order by Statementno  ASC", ref reply).Tables[0];
            if (dtCardusd != null)
            {
                if (dtCardusd.Rows.Count > 0)
                {
                    txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Estatement USD." });
                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Estatement USD.");
                    //Process pdf for USD
                    ProcessStatementUSD(dtCardusd);
                }
            }
        }

        //Process pdf for BDT
        private void ProcessStatementBDT(DataTable dtCards)
        {
            DataSet ds = new DataSet();
            DataTable stmdt = new DataTable();

            string reply = string.Empty;
            string filePath = string.Empty;
            string fileName = string.Empty;

            int count = 0;

            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
            ds = objProvider.ReturnData("select * from statement_BDT order by Statementno  ASC", ref reply);

            MsgLogWriter objLW = new MsgLogWriter();

            if (ds != null)
            {
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        DataTable dtAllRows = ds.Tables[0];

                        FileInfo objFile = new FileInfo(_EStatementProcessedPath);

                        if (!Directory.Exists(_EStatementProcessedPath))
                            Directory.CreateDirectory(_EStatementProcessedPath);

                        filePath = _EStatementProcessedPath + "\\EStatement of " + System.DateTime.Now.ToString("ddMMyyyy");

                        if (!Directory.Exists(filePath))
                            Directory.CreateDirectory(filePath);

                        DataRow dr;

                        txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement.");

                        for (int j = 0; j < dtCards.Rows.Count; j++)//dtCards.Rows.Count
                        {
                            //if (dtCards.Rows[j]["EMAIL"].ToString().Trim() != "")
                            //{
                                //if (IsValid(dtCards.Rows[j]["EMAIL"].ToString().Trim()))
                                //{
                                    try
                                    {
                                        pdfCount = pdfCount + 1;
                                        stmdt = new DataTable();
                                        stmdt = objProvider.ReturnData("select * from statement_BDT where IDCLIENT='" + dtCards.Rows[j]["IDCLIENT"].ToString() + "' and ACCOUNTNO= '" + dtCards.Rows[j]["ACCOUNTNO"].ToString() + "' ORDER BY [AutoID]", ref reply).Tables[0];
                                        //stmdt = objProvider.ReturnData("select * from statement_DUAL where CONTRACTNO = '" + dtCards.Rows[j]["CONTRACTNO"].ToString() + "' ", ref reply).Tables[0];


                                        if (stmdt.Rows.Count > 0)
                                        {
                                            EStatement objst = new EStatement();
                                            objst.SetDataSource(stmdt);

                                         
                                            string acc_no = dtCards.Rows[j]["ACCOUNTNO"].ToString();
                                            fileName = _fiid + "_" + dtCards.Rows[j]["STATEMENTNO"].ToString() + "_" + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["ACCOUNTNO"].ToString().Substring(acc_no.Length - 5, 5) + "_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["StartDate"].ToString().Replace('/', '-') + "_050_" + pdfCount + ".PDF";

                                           

                                            System.IO.Stream st = objst.ExportToStream(ExportFormatType.PortableDocFormat);

                                            PdfSharp.Pdf.PdfDocument document = PdfReader.Open(st);

                                            PdfSecuritySettings securitySettings = document.SecuritySettings;

                                            // Setting one of the passwords automatically sets the security level to 
                                            // PdfDocumentSecurityLevel.Encrypted128Bit.
                                            string card_no = dtCards.Rows[j]["PAN"].ToString();
                                            securitySettings.UserPassword = dtCards.Rows[j]["PAN"].ToString().Substring(card_no.Length - 4, 4);
                                            securitySettings.OwnerPassword = "owner";



                                            // Don´t use 40 bit encryption unless needed for compatibility reasons
                                            //securitySettings.DocumentSecurityLevel = PdfDocumentSecurityLevel.Encrypted40Bit;

                                            // Restrict some rights.            
                                            securitySettings.PermitAccessibilityExtractContent = false;
                                            securitySettings.PermitAnnotations = false;
                                            securitySettings.PermitAssembleDocument = false;
                                            securitySettings.PermitExtractContent = false;
                                            securitySettings.PermitFormsFill = true;
                                            securitySettings.PermitFullQualityPrint = false;
                                            securitySettings.PermitModifyDocument = true;
                                            securitySettings.PermitPrint = true;

                                            // Save the document...
                                            document.Save(filePath + "\\" + fileName);
                                            
                                            
                                            EStatementInfo objEst = new EStatementInfo();
                                            objEst.BANK_CODE = stmdt.Rows[0]["bank_code"].ToString();
                                            objEst.STARTDATE = stmdt.Rows[0]["STARTDATE"].ToString();
                                            objEst.ENDDATE = stmdt.Rows[0]["ENDDATE"].ToString();
                                            objEst.IDCLIENT = stmdt.Rows[0]["IDCLIENT"].ToString();
                                            objEst.PAN = dtCards.Rows[j]["pan"].ToString();

                                            if (stmdt.Rows.Count > 0)
                                                objEst.MAILADDRESS = stmdt.Rows[0]["EMAIL"].ToString();
                                            else
                                                objEst.MAILADDRESS = null;

                                            objEst.FILE_LOCATION = filePath + "\\" + fileName;
                                            objEst.MAILSUBJECT = txtEmailSubject.Text.Replace("'", "''");
                                            objEst.MAILBODY = txtEmailBody.Text.Replace("'", "''");
                                            objEst.STATUS = "1";

                                            reply = EStatementManager.Instance().AddEStatement(objEst, ref reply);

                                            if (reply == "Success")
                                            {
                                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Estatement has been created for Card# " + objEst.PAN.Substring(0, 6) + "******" + objEst.PAN.Substring(12, 4) });
                                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Estatement has been created for Card# " + objEst.PAN.Substring(0, 6) + "******" + objEst.PAN.Substring(12, 4));
                                                count++;
                                            }
                                            else
                                            {
                                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Message " + reply });
                                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + "Message " + reply);
                                            }
                                            if (count % 10 == 0)
                                            {
                                                objst.Dispose();
                                                GC.Collect();
                                                Thread.Sleep(1000);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + ex.Message);
                                    }
                                //}
                               // else
                               // {
                                   // txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                                   // objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                                //}
                            //}
                            //else
                           // {
                               // txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                               // objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                            //}
                        }
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Estatement has processed out of " + dtCards.Rows.Count + "." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString()  + " Estatement has processed out of " + dtCards.Rows.Count + ".");
                    }
                }
            }
        }
        
        //Process pdf for USD
        private void ProcessStatementUSD(DataTable dtCards)
        {
            DataSet ds = new DataSet();
            DataTable stmdt = new DataTable();

            string reply = string.Empty;
            string filePath = string.Empty;
            string fileName = string.Empty;

            int count = 0;

            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
            ds = objProvider.ReturnData("select * from statement_USD order by Statementno  ASC", ref reply);

            MsgLogWriter objLW = new MsgLogWriter();

            if (ds != null)
            {
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        DataTable dtAllRows = ds.Tables[0];

                        FileInfo objFile = new FileInfo(_EStatementProcessedPath);

                        if (!Directory.Exists(_EStatementProcessedPath))
                            Directory.CreateDirectory(_EStatementProcessedPath);

                        filePath = _EStatementProcessedPath + "\\EStatement of " + System.DateTime.Now.ToString("ddMMyyyy");

                        if (!Directory.Exists(filePath))
                            Directory.CreateDirectory(filePath);

                        DataRow dr;

                        txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement.");

                        for (int j = 0; j < dtCards.Rows.Count; j++)//dtCards.Rows.Count
                        {
                           // if (dtCards.Rows[j]["EMAIL"].ToString().Trim() != "")
                           // {
                               // if (IsValid(dtCards.Rows[j]["EMAIL"].ToString().Trim()))

                                //{
                                    try
                                    {
                                        pdfCount = pdfCount + 1;  
                                        stmdt = new DataTable();
                                        
                                        stmdt = objProvider.ReturnData("select * from statement_USD where IDCLIENT='" + dtCards.Rows[j]["IDCLIENT"].ToString() + "' and ACCOUNTNO= '" + dtCards.Rows[j]["ACCOUNTNO"].ToString() + "' ORDER BY [AutoID]", ref reply).Tables[0];
                                        if (stmdt.Rows.Count > 0)
                                        {
                                            EStatement objst = new EStatement();
                                            objst.SetDataSource(stmdt);

                                            
                                            string acc_no = dtCards.Rows[j]["ACCOUNTNO"].ToString();
                                            fileName = _fiid + "_" + dtCards.Rows[j]["STATEMENTNO"].ToString() + "_" + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["ACCOUNTNO"].ToString().Substring(acc_no.Length - 5, 5) + "_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["StartDate"].ToString().Replace('/', '-') + "_840_" + pdfCount + ".PDF";


                                            System.IO.Stream st = objst.ExportToStream(ExportFormatType.PortableDocFormat);

                                            PdfSharp.Pdf.PdfDocument document = PdfReader.Open(st);

                                            PdfSecuritySettings securitySettings = document.SecuritySettings;

                                            // Setting one of the passwords automatically sets the security level to 
                                            // PdfDocumentSecurityLevel.Encrypted128Bit.
                                            string card_no = dtCards.Rows[j]["PAN"].ToString();
                                            securitySettings.UserPassword = dtCards.Rows[j]["PAN"].ToString().Substring(card_no.Length - 4, 4);
                                            securitySettings.OwnerPassword = "owner";



                                            // Don´t use 40 bit encryption unless needed for compatibility reasons
                                            //securitySettings.DocumentSecurityLevel = PdfDocumentSecurityLevel.Encrypted40Bit;

                                            // Restrict some rights.            
                                            securitySettings.PermitAccessibilityExtractContent = false;
                                            securitySettings.PermitAnnotations = false;
                                            securitySettings.PermitAssembleDocument = false;
                                            securitySettings.PermitExtractContent = false;
                                            securitySettings.PermitFormsFill = true;
                                            securitySettings.PermitFullQualityPrint = false;
                                            securitySettings.PermitModifyDocument = true;
                                            securitySettings.PermitPrint = true;

                                            // Save the document...
                                            document.Save(filePath + "\\" + fileName);


                                            EStatementInfo objEst = new EStatementInfo();
                                            objEst.BANK_CODE = stmdt.Rows[0]["bank_code"].ToString();
                                            objEst.STARTDATE = stmdt.Rows[0]["STARTDATE"].ToString();
                                            objEst.ENDDATE = stmdt.Rows[0]["ENDDATE"].ToString();
                                            objEst.IDCLIENT = stmdt.Rows[0]["IDCLIENT"].ToString();
                                            objEst.PAN = dtCards.Rows[j]["pan"].ToString();




                                            if (stmdt.Rows.Count > 0)
                                                objEst.MAILADDRESS = stmdt.Rows[0]["EMAIL"].ToString();
                                            else
                                                objEst.MAILADDRESS = null;

                                            objEst.FILE_LOCATION = filePath + "\\" + fileName;
                                            objEst.MAILSUBJECT = txtEmailSubject.Text.Replace("'", "''");
                                            objEst.MAILBODY = txtEmailBody.Text.Replace("'", "''");
                                            objEst.STATUS = "1";

                                            reply = EStatementManager.Instance().AddEStatement(objEst, ref reply);

                                            if (reply == "Success")
                                            {
                                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Estatement has been created for Card# " + objEst.PAN.Substring(0, 6) + "******" + objEst.PAN.Substring(12, 4) });
                                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Estatement has been created for Card# " + objEst.PAN.Substring(0, 6) + "******" + objEst.PAN.Substring(12, 4));
                                                count++;
                                            }
                                            else
                                            {
                                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Message " + reply });
                                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + "Message " + reply);
                                            }
                                            if (count % 10 == 0)
                                            {
                                                objst.Dispose();
                                                GC.Collect();
                                                Thread.Sleep(1000);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + ex.Message);
                                    }
                                //}
                                //else
                                //{
                                   // txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                                   // objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                                //}
                            //}
                            //else
                            //{
                                //txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                                //objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                            //}
                        }
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Estatement has processed out of " + dtCards.Rows.Count + "." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Estatement has processed out of " + dtCards.Rows.Count + ".");
                    }
                }
            }
        }
        //
        void ReportViewer_Load(object sender, EventArgs e)
        {
            mailProgress.Visible = false;

             _XMLProcessedPath = ConfigurationManager.AppSettings[2].ToString();
            _XMLSourcePath = ConfigurationManager.AppSettings[3].ToString();
            _EStatementProcessedPath = ConfigurationManager.AppSettings[4].ToString();
            _LogPath = ConfigurationManager.AppSettings[5].ToString();
            _AdditionalAttachment = ConfigurationManager.AppSettings[8].ToString();
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

        private void ProcessFolderFiles(string _SourcePath, string BankCode, string BankName, ref string _reply)
        {
            #region Files of a Directory
            string reply = string.Empty;

            try
            {


                MsgLogWriter objLW = new MsgLogWriter();


                DirectoryInfo dir = new DirectoryInfo(_SourcePath);
                FileInfo[] fi = dir.GetFiles();

               

                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + fi.Length.ToString() + " files found to process.." });
                objLW.logTrace(_LogPath, "EStatement.log", " : Total " + fi.Length.ToString() + " files found to process..");

                for (int f = 0; f < fi.Length; f++)
                {
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + fi[f].Name + " on process.." });
                    objLW.logTrace(_LogPath, "EStatement.log", " : " + fi[f].Name + " on process..");

                   // DataSet dsXML = getDataFromXML(fi[f].FullName);
                    string fullPath = fi[f].FullName;

                    
                   
                    #region Operation On Data
                    if (fullPath != null)
                    {
                        
                            ConStr = new ConnectionStringBuilder(1);
                            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);

                            string sql = string.Empty;

                            try
                            {
                                objProvider.RunQuery("insert into CSV_DATA_ARC select * from CSV_DATA");
                                objProvider.RunQuery("Truncate table  CSV_DATA");
                               // objProvider.RunQuery("BULK INSERT CSV_DATA FROM '" + fullPath + "' WITH(ROWTERMINATOR='\\n',FIELDTERMINATOR='|')");
                                reply = objProvider.RunQuery("BULK INSERT CSV_DATA FROM '" + fullPath + "' WITH(ROWTERMINATOR='\\n',FIELDTERMINATOR='|')");

                                if (reply == "Success")
                                {
                                    objProvider.RunQuery("Delete from  STATEMENT");
                                    objProvider.RunQuery("insert  into  STATEMENT (BANK_CODE,STATEMENTNO,CLIENT ,IDCLIENT,STREETADDRESS,EMAIL,MOBILE,REGION,ZIP,COUNTRY,STARTDATE,ENDDATE,MAIN_CARD,JOBTITLE,COMPANYNAME)select distinct BANK_CODE,STATEMENTNO,CLIENT ,IDCLIENT,ADDRESS AS STREETADDRESS,EMAIL,MOBILE,CITY AS REGION,ZIP,COUNTRY,STARTDATE,ENDDATE,PAN,JOBTITLE,COMPANYNAME from csv_data");

                                    objProvider.RunQuery("UPDATE  STATEMENT SET StatementMessage='" + stmMessage + "' WHERE BANK_CODE='"+_fiid+"'");  
                                    objProvider.RunQuery("Delete from  ACCOUNT");
                                    objProvider.RunQuery(" insert  into  ACCOUNT (STATEMENTNO,ACCOUNTNO,ACURC,SBALANCE,EBALANCE) select distinct STATEMENTNO,ACCOUNTNO,AC_CURR,SBALANCE,EBALANCE from csv_data");


                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " :  Card record has been found to process.." });
                                    objLW.logTrace(_LogPath, "EStatement.log", " :  Card record has been found to process..");
                                }

                                else
                                {

                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " Error: " + reply });
                                    objLW = new MsgLogWriter();
                                    objLW.logTrace(_LogPath, "EStatement.log", "Error. " + reply);
                                }


                                

                            }
                            catch (IOException ex)
                            {
                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " Error: " + ex.Message });
                                objLW = new MsgLogWriter();
                                objLW.logTrace(_LogPath, "EStatement.log", "Error. " + ex.Message);
                            }

                        
                           
                    }
                    #endregion

                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + fi[f].Name + " process complete.." });
                    objLW.logTrace(_LogPath, "EStatement.log", " : " + fi[f].Name + " process complete..");
                    txtAnalyzer.Invoke(_addText, new object[] { "\n" });

                    btnGenerate_Click(null, null);
                }
                if (Directory.Exists(_SourcePath))
                {
                    try
                    {
                        Directory.Move(dir.FullName, _XMLProcessedPath + "\\" + dir.Name);

                    }
                    catch (IOException ex)
                    {
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Source Directory moving Error. Error: " + ex.Message });
                        objLW = new MsgLogWriter();
                        objLW.logTrace(_LogPath, "EStatement.log", "Source Directory moving Error. " + ex.Message);
                    }
                }
                // return true;
            }
            catch (Exception ex)
            {
                _reply = ex.StackTrace;
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
            }

            #endregion
        }


        

        private DataSet getDataFromXML(string _filename)
        {
            try
            {
                System.Data.DataSet ds = new System.Data.DataSet();
                ds.ReadXml(_filename);
                return ds;
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return null;
            }

        }

       
        private bool IsValid(string emailaddress)
        {
            try
            {
                MailAddress m = new MailAddress(emailaddress);
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
