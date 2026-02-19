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
using QCash.EStatement.LBF.Reports;
using CrystalDecisions.CrystalReports.Engine;
using FlexiStar.Utilities.EncryptionEngine;
using QCash.EStatement.LBF;
using Infragistics.Documents.PDF;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Security;
using Common;
using System.Text.RegularExpressions;
using System.Net;
using System.Globalization;
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
        private string StmDate = string.Empty;
        private string stmMessage = string.Empty;

        string vPAN = string.Empty;
        string prePan = string.Empty;
        string preDoc = string.Empty;
        
        int pdfCount = 0;

        Thread tdGenerate = null;
        Thread tdSendMail = null;

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
                        objLW.logTrace(_LogPath, "EStatement.log", "Successfully archive previous data !!!");
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + "Successfully archive previous data !!!" });

                        ProcessData();
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
                if (StmDate == "")
                    StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                else StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);

                MsgLogWriter objLW = new MsgLogWriter();

                EStatementList objESList = EStatementManager.Instance().GetAllEStatements(_fiid, StmDate, "1", ref reply);

                if (objESList != null && objESList.Count > 0)
                {
                    SmtpConfigurationManager objSmtpMan = new SmtpConfigurationManager();
                    SmtpConfigurationList objSmtpList = objSmtpMan.GetSmtpConfiguration(_fiid, 1);
                    Encryption objEnc = new Encryption();
                    if (objSmtpList != null && objSmtpList.Count > 0)
                    {
                        int count = 0;

                        SmtpClient SmtpServer = new SmtpClient(objSmtpList[0].Smtp_Server);

                        try
                        {
                            SmtpServer.Port = objSmtpList[0].Smtp_Port;
                            SmtpServer.Credentials = new NetworkCredential(objSmtpList[0].From_User, objEnc.DecryptWord(objSmtpList[0].From_Password));
                            SmtpServer.EnableSsl = Convert.ToBoolean(objSmtpList[0].EnableSSL);

                            foreach (var objES in objESList)
                            {
                                if (!string.IsNullOrEmpty(objES.MAILADDRESS))
                                {
                                    try
                                    {

                                        MailMessage mail = new MailMessage();
                                        mail.From = new MailAddress(objSmtpList[0].From_Address);
                                        mail.Subject = objES.MAILSUBJECT;
                                        mail.Body = objES.MAILBODY;
                                        mail.To.Add(objES.MAILADDRESS.Trim());
                                        System.Net.Mail.Attachment attachment;
                                        attachment = new System.Net.Mail.Attachment(objES.FILE_LOCATION);
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


                                        // Capture the start time
                                        //DateTime startTime = DateTime.Now;

                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : " + "Sending EStatement to " + mail.To.ToString() }); ;
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : Sending EStatement " + mail.To.ToString());


                                        SmtpServer.Send(mail);

                                        // Capture the end time
                                        // DateTime endTime = DateTime.Now;

                                        // Calculate the elapsed time
                                        //TimeSpan elapsedTime = endTime - startTime;

                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : " + "mail Send to " + mail.To.ToString() }); ;
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : mail Send to " + mail.To.ToString());

                                        objES.STATUS = "0";  // Statement Generated and Mail Sent Successfully
                                        EStatementManager.Instance().UpdateEStatement(objES, ref reply);
                                        count++;
                                    }
                                    catch (Exception ex)
                                    {
                                        txtAnalyzer.Invoke(_addText, DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : Error: " + ex.Message);
                                        objLW.logTrace(_LogPath, "EStatement.log", DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : Error: " + ex.Message);

                                        objES.STATUS = "2"; // Mail is not Sent
                                        EStatementManager.Instance().UpdateEStatement(objES, ref reply);
                                    }
                                }
                                else
                                {
                                    txtAnalyzer.Invoke(_addText, DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : No Mail Address Found to send the Estatement " + objES.FILE_LOCATION);
                                    objLW.logTrace(_LogPath, "EStatement.log", DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : No Mail Address Found to send the Estatement " + objES.FILE_LOCATION);

                                    objES.STATUS = "8";   // No Mail Address Found
                                    EStatementManager.Instance().UpdateEStatement(objES, ref reply);
                                }
                            }

                            txtAnalyzer.Invoke(_addText, DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : Total " + count + " Estatement(s) have been mailed out of " + objESList.Count + ".");
                            objLW.logTrace(_LogPath, "EStatement.log", DateTime.Now.ToString("MMMM dd, yyyy h:mm: ss tt") + " : Total " + count + " Estatement(s) have been mailed out of " + objESList.Count + ".");
                        }
                        catch (Exception ex)
                        {
                            // MsgLogWriter objLW = new MsgLogWriter();
                            objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                            txtAnalyzer.Invoke(_addText, DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : Error: " + ex.Message);
                        }
                    }
                }
                else
                {
                    txtAnalyzer.Invoke(_addText, DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : There are no EStatements generated on that statement date.");
                    objLW.logTrace(_LogPath, "EStatement.log", DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : There are no EStatements generated on that statement date.");
                }
            }
            catch (Exception ex)
            {
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                txtAnalyzer.Invoke(_addText, DateTime.Now.ToString("MMMM dd, yyyy h:mm:ss tt") + " : Error: " + ex.Message);
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
            dtCardbdt = objProvider.ReturnData("select * from Qry_Card_Account where Curr='BDT'  ORDER BY CAST(STATEMENTNO AS INT) ASC", ref reply).Tables[0];// where Curr='BDT'

            if (dtCardbdt.Rows.Count > 0)
            {
                txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Estatement." });//Processing Estatement BDT
                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Estatement.");//Processing Estatement BDT.

                //Process pdf for BDT
                ProcessStatementBDT(dtCardbdt);
            }

            /*DataTable dtCardusd = new DataTable();
            dtCardusd = objProvider.ReturnData("select * from Qry_Card_Account where Curr='USD'", ref reply).Tables[0];
            if (dtCardusd != null)
            {
                if (dtCardusd.Rows.Count > 0)
                {
                    txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Estatement USD." });
                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Estatement USD.");
                    //Process pdf for USD
                    ProcessStatementUSD(dtCardusd);
                }
            }*/
        }
        private void ProcessStatementDUAL(DataTable dtCards)
        {
            
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
            ds = objProvider.ReturnData(" SELECT * FROM STATEMENT_BDT ORDER BY  CAST( STATEMENTNO AS INT) ASC", ref reply);

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
                            if (dtCards.Rows[j]["EMAIL"].ToString().Trim() != "")
                            {
                                if (IsValid(dtCards.Rows[j]["EMAIL"].ToString().Trim()))
                                {
                                    try
                                    {
                                        pdfCount = pdfCount + 1;
                                        stmdt = new DataTable();
                                        stmdt = objProvider.ReturnData("select * from statement_BDT where CONTRACTNO='" + dtCards.Rows[j]["CONTRACTNO"].ToString()+"' order by SL,autoid", ref reply).Tables[0];
                                        //if ((dtCards.Rows[j]["pan"].ToString()) != vPAN)
                                        //{
                                        //    vPAN = dtCards.Rows[j]["pan"].ToString();
                                            if (stmdt.Rows.Count > 0)
                                            {
                                                // For VISA GOLD and VISA PLATINUM
                                                /*EStatement objst = new EStatement();
                                                EStatementPlatinum objstPlatinum = new EStatementPlatinum();

                                                if (dtCards.Rows[j]["EMAIL"].ToString().Trim() == "rtte")
                                                {
                                                    objst.SetDataSource(stmdt);
                                                }
                                                else
                                                {
                                                    objstPlatinum.SetDataSource(stmdt);
                                                }

                                                fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) + ".pdf";

                                                if (dtCards.Rows[j]["EMAIL"].ToString().Trim() == "rtte")
                                                {
                                                    objst.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);
                                                }
                                                else
                                                    objstPlatinum.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);*/

                                                
                                                EStatement objst = new EStatement();
                                                objst.SetDataSource(stmdt);
                                                //fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) + ".pdf";
                                                //fileName = "VISA_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 4) + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + '_' + pdfCount + ".pdf";
                                               // string Bin = dtCards.Rows[j]["pan"].ToString().Substring(0, 6);
                                                //fileName = _fiid + "_" + Bin + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + "_" + pdfCount + ".pdf";

                                                string acc_no = dtCards.Rows[j]["ACCOUNTNO"].ToString();
                                                fileName = _fiid + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["CONTRACTNO"].ToString().Substring(0, 4) + dtCards.Rows[j]["ACCOUNTNO"].ToString().Substring(acc_no.Length - 5, 5) + "_" + dtCards.Rows[j]["idclient"].ToString() + "_" + "050" + ".PDF";
 
                                                
                                                
                                                
                                                //////////////Modified File name/////////////////////
                                                //string Bin = dtCards.Rows[j]["pan"].ToString().Substring(0, 6);

                                                //if (Bin == "416248")
                                                //{
                                                //fileName = _fiid + "_" + "Visa Signature" + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + "_" + pdfCount + ".pdf";
                                                //}
                                                //else
                                                //{
                                                //    fileName = _fiid + "_" + "Visa Gold" + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + "_" + pdfCount + ".pdf";

                                                //}
                                                System.IO.Stream st = objst.ExportToStream(ExportFormatType.PortableDocFormat);

                                                PdfSharp.Pdf.PdfDocument document = PdfReader.Open(st);

                                                PdfSecuritySettings securitySettings = document.SecuritySettings;

                                                // Setting one of the passwords automatically sets the security level to 
                                                // PdfDocumentSecurityLevel.Encrypted128Bit.
                                                string card_no = dtCards.Rows[j]["pan"].ToString();
                                                securitySettings.UserPassword = dtCards.Rows[j]["pan"].ToString().Substring(card_no.Length - 4, 4);
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
                                                securitySettings.PermitPrint = true ;

                                                // Save the document...
                                                document.Save(filePath + "\\" + fileName);

                                                // objst.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);


                                                EStatementInfo objEst = new EStatementInfo();
                                                objEst.BANK_CODE = stmdt.Rows[0]["bank_code"].ToString();
                                                objEst.STMDATE = stmdt.Rows[0]["STATEMENT_DATE"].ToString();
                                                StmDate = stmdt.Rows[0]["STATEMENT_DATE"].ToString();

                                                string[] drdate = stmdt.Rows[0]["STATEMENT_DATE"].ToString().Split('/');

                                                if (drdate.Length == 3)
                                                {
                                                    objEst.MONTH = drdate[1].ToString();
                                                    objEst.YEAR = drdate[2].ToString();
                                                }
                                                else
                                                {
                                                    objEst.MONTH = null;
                                                    objEst.YEAR = null;
                                                }
                                                objEst.PAN_NUMBER = dtCards.Rows[j]["pan"].ToString();

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
                                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4) });
                                                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4));
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
                                    //}
                                    catch (Exception ex)
                                    {
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + ex.Message);
                                    }
                                }
                                else
                                {
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                                }
                            }
                            else
                            {
                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                            }
                        }
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Estatement has processed out of " + dtCards.Rows.Count + "." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Estatement has processed" + dtCards.Rows.Count + ".");
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
            ds = objProvider.ReturnData("select * from statement_USD", ref reply);

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
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + "Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement.");

                        for (int j = 0; j < dtCards.Rows.Count; j++)//dtCards.Rows.Count
                        {
                            if (dtCards.Rows[j]["EMAIL"].ToString().Trim() != "")
                            {
                                if (IsValid(dtCards.Rows[j]["EMAIL"].ToString().Trim()))
                                {
                                    try
                                    {
                                        stmdt = new DataTable();
                                        stmdt = objProvider.ReturnData("select * from statement_USD where pan='" + dtCards.Rows[j]["pan"].ToString() + "'", ref reply).Tables[0];
                                        stmdt = objProvider.ReturnData("select * from statement_USD where CONTRACTNO='" + dtCards.Rows[j]["CONTRACTNO"].ToString() + "'", ref reply).Tables[0];
                                        if (stmdt.Rows.Count > 0)
                                        {
                                            EStatement objst = new EStatement();
                                            //EStatementPlatinum objstPlatinum = new EStatementPlatinum();

                                            if (dtCards.Rows[j]["EMAIL"].ToString().Trim() == "rtte")
                                            {
                                                objst.SetDataSource(stmdt);
                                            }
                                            else
                                            {
                                                //objstPlatinum.SetDataSource(stmdt);
                                            }

                                            //fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) + "_USD.pdf";

                                            string acc_no = dtCards.Rows[j]["ACCOUNTNO"].ToString();
                                            fileName = _fiid + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["CONTRACTNO"].ToString().Substring(0, 4) + dtCards.Rows[j]["ACCOUNTNO"].ToString().Substring(acc_no.Length - 5, 5) + "_" + dtCards.Rows[j]["idclient"].ToString() + "_" + "840" + ".PDF";
 
                                                
                                           // fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) + ".pdf";
                                            objst.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);

                                            EStatementInfo objEst = new EStatementInfo();
                                            objEst.BANK_CODE = stmdt.Rows[0]["bank_code"].ToString();
                                            objEst.STMDATE = stmdt.Rows[0]["STATEMENT_DATE"].ToString();
                                            StmDate = stmdt.Rows[0]["STATEMENT_DATE"].ToString();

                                            string[] drdate = stmdt.Rows[0]["STATEMENT_DATE"].ToString().Split('/');

                                            if (drdate.Length == 3)
                                            {
                                                objEst.MONTH = drdate[1].ToString();
                                                objEst.YEAR = drdate[2].ToString();
                                            }
                                            else
                                            {
                                                objEst.MONTH = null;
                                                objEst.YEAR = null;
                                            }
                                            objEst.PAN_NUMBER = dtCards.Rows[j]["pan"].ToString();

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
                                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4) });
                                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4));
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
                                }
                                else
                                {
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                                }
                            }
                            else
                            {
                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                            }
                        }
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Estatement has processed out of " + dtCards.Rows.Count + "." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Estatement has processed" + dtCards.Rows.Count + ".");
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

        private void ProcessData()
        {
            string _bankCode = string.Empty;
            string _bankName = string.Empty;

            string _reply = string.Empty;

            #region Folder Searching in File name path

            DirectoryInfo di = new DirectoryInfo(_XMLSourcePath);
            DirectoryInfo[] dia = di.GetDirectories();


            for (int fcount = 0; fcount < dia.Length; fcount++)
            {
                
               
                 if (dia[fcount].FullName.Contains("LBF"))
                {
                    _bankName = "LBF";
                    _bankCode = "9935";
                    _XMLSourcePath = dia[fcount].FullName;
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }
                
                else
                {
                    MsgLogWriter objLW = new MsgLogWriter();
                    objLW.logTrace(_LogPath, "EStatement.log", "Not an XML data !!!");
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + "Not an XML data !!!" });
                }

                Bank_Code = _bankName;

            }
            #endregion

        }

        //private string CalcuateBalance()
        //{
        //    int qStatus = 0;
        //    string _reply = string.Empty;
        //    try
        //    {
        //        ConStr = new ConnectionStringBuilder(1);
        //        SPExecute objProvider = new SPExecute(ConStr.ConnectionString_DBConfig);

        //        qStatus = objProvider.ExecuteNonQuery("UPDATE_LIMIT", null);

        //    }
        //    catch (Exception ex)
        //    {
                
        //    }

        //    return _reply;
        //}

        private void ProcessFolderFiles(string _SourcePath, string BankCode, string BankName, ref string _reply)
        {
            #region Files of a Directory
            string reply = string.Empty;

            try
            {
                MsgLogWriter objLW = new MsgLogWriter();
                DirectoryInfo dir = new DirectoryInfo(_SourcePath);
                FileInfo[] fi = dir.GetFiles("*.xml"); // Only XML files

                int totalXmlFiles = fi.Length;

                txtAnalyzer.Invoke(_addText, new object[] 
        { 
            DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + totalXmlFiles.ToString() + " XML files found to process.." 
        });
                objLW.logTrace(_LogPath, "EStatement.log", " : Total " + totalXmlFiles.ToString() + " XML files found to process..");

                // Now process each XML
                foreach (FileInfo file in fi)
                {
                    string xmlFileName = file.Name;

                    txtAnalyzer.Invoke(_addText, new object[]
            {
                DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + xmlFileName + " on process.."
            });
                    objLW.logTrace(_LogPath, "EStatement.log", " : " + xmlFileName + " on process..");

                    using (FileStream fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read))
                    {
                        DataSet dsXML = new DataSet();
                        dsXML.ReadXml(fs);

                        if (StmDate == "")
                            StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                        else
                            StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);

                        // Validate Statement Date
                        DataTable dtStatement = dsXML.Tables["Statement"];
                        if (dtStatement != null && dtStatement.Rows.Count > 0)
                        {
                            string xmlStatementDate = dtStatement.Rows[0]["STATEMENT_DATE"].ToString();

                            if (xmlStatementDate != StmDate)
                            {
                                txtAnalyzer.Invoke(_addText, new object[]
                        {
                            DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + 
                            " : XML Statement Date (" + xmlStatementDate + 
                            ") does not match user input (" + StmDate + "). Stopping process."
                        });

                                objLW.logTrace(_LogPath, "EStatement.log",
                                    "XML Statement Date (" + xmlStatementDate +
                                    ") does not match user input (" + StmDate + "). Stopping process."
                                );
                                return;
                            }
                        }

                        #region Operation On Data
                        if (dsXML != null && dsXML.Tables.Count > 0)
                        {
                            ConStr = new ConnectionStringBuilder(1);
                            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);

                            // Clear previous data
                            objProvider.RunQuery("Delete from AccumIntAcc");
                            objProvider.RunQuery("Delete from BonusContrAcc");
                            objProvider.RunQuery("insert into statement_info_arc select * from statement_info");
                            objProvider.RunQuery("insert into statement_details_arc select * from statement_details");
                            objProvider.RunQuery("Truncate table statement_details");
                            objProvider.RunQuery("Delete from statement_info");

                            for (int i = 0; i < dsXML.Tables.Count; i++)
                            {
                                if (dsXML.Tables[i].TableName == "Statement")
                                {
                                    GetCardHolderPersonalInfo(dsXML.Tables[i], BankName, ref reply);
                                    txtAnalyzer.Invoke(_addText, new object[]
                            {
                                DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") +  " : CardHolder Personal Info Saved. " + reply});
                             }
                                else if (dsXML.Tables[i].TableName == "Operation")
                                {
                                    reply = GetCardHolderTransactionInfo(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[]
                            {
                                DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + 
                                " : CardHolder Transaction Info Saved. " + reply
                            });
                                }
                                else if (dsXML.Tables[i].TableName == "Account")
                                {
                                    reply = GetCardHolderAccountInfo(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[]
                            {
                                DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + 
                                " : CardHolder Account Info Saved. " + reply
                            });
                                }
                                else if (dsXML.Tables[i].TableName == "Card")
                                {
                                    reply = GetCardHolderCardInfo(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[]
                            {
                                DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + 
                                " : CardHolder Card Info Saved. " + reply
                            });
                                }
                                else if (dsXML.Tables[i].TableName == "BonusContrAcc")
                                {
                                    reply = GetBonusContrAccInfo(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[]
                            {
                                DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + 
                                " : Reward Info Saved. " + reply
                            });
                                }
                                else if (dsXML.Tables[i].TableName == "AccumIntAcc")
                                {
                                    reply = GetAccumIntAcc(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[]
                            {
                                DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + 
                                " : Suspense Info Saved. " + reply
                            });
                                }
                            }

                            if (reply == "Success")
                            {
                                for (int i = 0; i < dsXML.Tables.Count; i++)
                                {
                                    if (dsXML.Tables[i].TableName == "Operation")
                                        GenerateStatementInfo(dsXML, BankName, ref reply);
                                }
                            }

                      /*      txtAnalyzer.Invoke(_addText, new object[]
                   {
                        DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + 
                        " : Total " + dsXML.Tables["Card"].Rows.Count.ToString() + 
                        " Card record found.."
                    });*/
                        }
                        #endregion

                        btnGenerate_Click(null, null);

                        txtAnalyzer.Invoke(_addText, new object[]
                {
                    DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + 
                    " : " + xmlFileName + " process complete..\n"
                });
                        objLW.logTrace(_LogPath, "EStatement.log", " : " + xmlFileName + " process complete..");
                    }
                }

                // Move processed folder
                if (Directory.Exists(_SourcePath))
                {
                    try
                    {
                        Directory.Move(dir.FullName, _XMLProcessedPath + "\\" + dir.Name);
                    }
                    catch (IOException ex)
                    {
                        txtAnalyzer.Invoke(_addText, new object[]
                {
                    DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + 
                    " : Source Directory moving Error. " + ex.Message
                });
                    }
                }
            }
            catch (Exception ex)
            {
                _reply = ex.StackTrace;
                txtAnalyzer.Invoke(_addText, new object[]
        {
            DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + ex.Message
        });
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

        private StatementList GetCardHolderPersonalInfo(DataTable dtStatement, string BankCode, ref string errMsg)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            Statement objSt = null;
            StatementList objStList = new StatementList();
            int @vSL = 0;
            try
            {
               

                //Clear Previous Data
                objProvider.RunQuery("Delete from " + dtStatement.TableName);
                //Clear Previous Data
                objProvider.RunQuery("Delete from CARD");

                for (int k = 0; k < dtStatement.Rows.Count; k++)
                {
                    objSt = new Statement();
                    objSt.BANK_CODE = BankCode;

                    for (int j = 0; j < dtStatement.Columns.Count; j++)
                    {
                        #region setting properties values

                        if (dtStatement.Columns[j].ColumnName == "StatementNo")
                        {
                            objSt.STATEMENTNO = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Address")
                        {
                            objSt.ADDRESS = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "CARD_LIST")
                        {
                            objSt.CARD_LIST = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "City")
                        {
                            objSt.CITY = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Region")
                        {
                            objSt.REGION = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Country")
                        {
                            objSt.COUNTRY = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Email")
                        {
                            objSt.EMAIL = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "StartDate")
                        {
                            objSt.STARTDATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "EndDate")
                        {
                            objSt.ENDDATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Client")
                        {
                            objSt.CLIENT = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "ContractNo")
                        {
                            objSt.CONTRACTNO = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "IdClient")
                        {
                            objSt.IDCLIENT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Fax")
                        {
                            objSt.FAX = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "MAIN_CARD")
                        {
                            objSt.MAIN_CARD = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Mobile")
                        {
                            objSt.MOBILE = dtStatement.Rows[k][j].ToString().Replace("'", "").Replace("(", "").Replace(")", "").Replace("8800", "880");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "NEXT_STATEMENT_DATE")
                        {
                            objSt.NEXT_STATEMENT_DATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "PAYMENT_DATE")
                        {
                            objSt.PAYMENT_DATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "STATEMENT_DATE")
                        {
                            objSt.STATEMENT_DATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "StreetAddress")
                        {
                            objSt.STREETADDRESS = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Telephone")
                        {
                            objSt.TELEPHONE = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Title")
                        {
                            objSt.TITLE = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "ZIP")
                        {
                            objSt.ZIP = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "PromotionalText")
                        {
                          

                            objSt.PromotionalText = dtStatement.Rows[k][j].ToString().Replace("'", "''");

                            string[] lines = objSt.PromotionalText.Split('|');

                            // Card is always at index 4
                            if (lines.Length > 4 && !string.IsNullOrEmpty(lines[4]))
                            {
                                if (string.IsNullOrEmpty(objSt.MAIN_CARD) || objSt.MAIN_CARD.Length < 16)
                                {
                                    
                                    @vSL = @vSL + 1;
                                   
                                    objSt.MAIN_CARD = lines[4];
                                    objSt.MAIN_CARD = objSt.MAIN_CARD.ToString().Substring(0, 6) + "******" + objSt.MAIN_CARD.ToString().Substring(objSt.MAIN_CARD.Length - 4, 4);

                                    sql = "Insert into Card(STATEMENTNO,PAN,MBR,CLIENTNAME,SLNO)" +
                       " Values('" + objSt.STATEMENTNO + "','" + objSt.MAIN_CARD + "','" + '0' + "','" + objSt.CLIENT + "','" + @vSL + "')";

                                    reply = objProvider.RunQuery(sql);
                                   
                                }
                            }



                        }
                        #endregion
                    }
                    objStList.Add(objSt);

                    sql = "Insert into Statement(BANK_CODE,STATEMENTNO,ADDRESS,CARD_LIST,CITY,COUNTRY,EMAIL," +
                          "STARTDATE,ENDDATE,CLIENT,CONTRACTNO,IDCLIENT,FAX,MAIN_CARD,MOBILE," +
                          "NEXT_STATEMENT_DATE,PAYMENT_DATE,REGION,STATEMENT_DATE,SEX,STREETADDRESS,TELEPHONE,TITLE,ZIP,PromotionalText) " +
                          "values('" + objSt.BANK_CODE + "','" + objSt.STATEMENTNO + "','" + objSt.ADDRESS + "','" + objSt.CARD_LIST + "','" + objSt.CITY + "','" + objSt.COUNTRY + "','" + objSt.EMAIL + "'," +
                          "'" + objSt.STARTDATE + "','" + objSt.ENDDATE + "','" + objSt.CLIENT + "','" + objSt.CONTRACTNO + "','" + objSt.IDCLIENT + "','" + objSt.FAX + "','" + objSt.MAIN_CARD + "','" + objSt.MOBILE + "'," +
                          "'" + objSt.NEXT_STATEMENT_DATE + "','" + objSt.PAYMENT_DATE + "','" + objSt.REGION + "','" + objSt.STATEMENT_DATE + "','" + objSt.SEX + "','" + objSt.STREETADDRESS + "'," +
                          "'" + objSt.TELEPHONE + "','" + objSt.TITLE + "','" + objSt.ZIP + "','" + objSt.PromotionalText + "')";

                    reply = objProvider.RunQuery(sql);
                    //if (!reply.Contains("Success"))
                        errMsg=reply;
                }
                return objStList;
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                errMsg = "Error: " + ex.StackTrace;
                return objStList;
            }

        }
        //
        private string GetCardHolderTransactionInfo(DataTable dtOperation)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            Operation objOp = null;
            //OperationList objOpList = new OperationList();

            try
            {
                //Clear Previous Data
                objProvider.RunQuery("Delete from " + dtOperation.TableName);

                for (int k = 0; k < dtOperation.Rows.Count; k++)
                {
                    objOp = new Operation();
                    //objSt.BANK_CODE = BankCode;

                    for (int j = 0; j < dtOperation.Columns.Count; j++)
                    {
                        #region setting properties values

                        if (dtOperation.Columns[j].ColumnName == "StatementNo")
                        {
                            objOp.STATEMENTNO = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "O")
                        {
                            objOp.OpID = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "OD")
                        {
                            objOp.OpDate = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "TD")
                        {
                            objOp.TD = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "A")
                        {
                            if (dtOperation.Rows[k][j].ToString() == "" || dtOperation.Rows[k][j].ToString() == null)
                                objOp.Amount = "0.00";
                            else
                                objOp.Amount = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "ACURC")
                        {
                            objOp.ACURCode = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "ACURN")
                        {
                            objOp.ACURName = dtOperation.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "D")
                        {
                            objOp.D = dtOperation.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "DE")
                        {
                            objOp.DE = dtOperation.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "CF")
                        {
                            objOp.CF = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "DOCNO")
                        {
                            objOp.DOCNO = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "NO")
                        {
                            objOp.NO = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "ACCOUNT")
                        {
                            objOp.ACCOUNT = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "ACC")
                        {
                            objOp.ACC = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "FR")
                        {
                            objOp.FR = dtOperation.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "APPROVAL")
                        {
                            objOp.APPROVAL = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "MN")
                        {
                            objOp.MN = dtOperation.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "S")
                        {
                            objOp.S = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "TERMN")
                        {
                            objOp.TERMN = dtOperation.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "TL")
                        {
                            objOp.TL = dtOperation.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "P")
                        {
                            objOp.P = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "OCC")
                        {
                            objOp.OCCode = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "OC")
                        {
                            objOp.OCName = dtOperation.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "AMOUNTSIGN")
                        {
                            objOp.AMOUNTSIGN = dtOperation.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "SERIALNO")
                        {
                            objOp.SERIALNO = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "OA")
                        {
                            if (dtOperation.Rows[k][j].ToString() == "" || dtOperation.Rows[k][j].ToString() == null)
                                objOp.OA = "0.00";
                            else
                                objOp.OA = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        #endregion
                    }
                    //objOpList.Add(objOp);

                    sql = "Insert into Operation(STATEMENTNO,O,OD,TD,A,ACURC,ACURN,D,DE,P,OA,OCC,OC,TL,TERMN,CF,S,MN,DOCNO,NO,ACCOUNT,ACC,FR,APPROVAL,AMOUNTSIGN) " +
                    "Values('" + objOp.STATEMENTNO + "','" + objOp.OpID + "','" + objOp.OpDate + "','" + objOp.TD + "','" + objOp.Amount + "'," +
                    "'" + objOp.ACURCode + "','" + objOp.ACURName + "','" + objOp.D + "','" + objOp.DE + "','" + objOp.P + "','" + objOp.OA + "'," +
                    "'" + objOp.OCCode + "','" + objOp.OCName + "','" + objOp.TL + "','" + objOp.TERMN + "','" + objOp.CF + "','" + objOp.S + "'," +
                    "'" + objOp.MN + "','" + objOp.DOCNO + "','" + objOp.NO + "','" + objOp.ACCOUNT + "','" + objOp.ACC + "','" + objOp.FR + "','" + objOp.APPROVAL + "','" + objOp.AMOUNTSIGN + "') ";

                    reply = objProvider.RunQuery(sql);
                    if (!reply.Contains("Success"))
                        return reply;
                }
                return reply;
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return "Error: " + ex.StackTrace;
            }
        }
        ////////New Table
        private string GetAccumIntAcc(DataTable dtGetAccumIntAcc)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            AccumIntAcc objOp = null;
            AccumIntAccList objOpList = new AccumIntAccList();

            try
            {
                //Clear Previous Data
                objProvider.RunQuery("Delete from " + dtGetAccumIntAcc.TableName);

                for (int k = 0; k < dtGetAccumIntAcc.Rows.Count; k++)
                {
                    objOp = new AccumIntAcc();

                    for (int j = 0; j < dtGetAccumIntAcc.Columns.Count; j++)
                    {
                        #region setting properties values

                        if (dtGetAccumIntAcc.Columns[j].ColumnName == "StatementNo")
                        {
                            objOp.STATEMENTNO = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtGetAccumIntAcc.Columns[j].ColumnName == "ACCUM_INT_RRELEASE")
                        {
                            objOp.ACCUM_INT_RRELEASE = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtGetAccumIntAcc.Columns[j].ColumnName == "ACCUM_INT_EBALANCE")
                        {
                            objOp.ACCUM_INT_EBALANCE = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtGetAccumIntAcc.Columns[j].ColumnName == "ACCUM_INT_SBALANCE")
                        {
                            objOp.ACCUM_INT_SBALANCE = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtGetAccumIntAcc.Columns[j].ColumnName == "ACCUM_INT_AMOUNT")
                        {
                            objOp.ACCUM_INT_AMOUNT = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtGetAccumIntAcc.Columns[j].ColumnName == "ACCOUNT_NO")
                        {
                            objOp.ACCOUNT_NO = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtGetAccumIntAcc.Columns[j].ColumnName == "AutoID")
                        {
                            objOp.AutoID = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        #endregion
                    }
                    //objOpList.Add(objOp);

                    sql = "Insert into AccumIntAcc(STATEMENTNO,ACCUM_INT_RRELEASE,ACCUM_INT_EBALANCE,ACCUM_INT_SBALANCE,ACCUM_INT_AMOUNT,ACCOUNT_NO) " +
                    "Values('" + objOp.STATEMENTNO + "','" + objOp.ACCUM_INT_RRELEASE + "','" + objOp.ACCUM_INT_EBALANCE + "','" + objOp.ACCUM_INT_SBALANCE + "','" + objOp.ACCUM_INT_AMOUNT + "'," +
                    "'" + objOp.ACCOUNT_NO + "," + objOp.AutoID + "') ";

                    reply = objProvider.RunQuery(sql);
                    if (!reply.Contains("Success"))
                        return reply;
                }
                return reply;
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return "Error: " + ex.StackTrace;
            }
        }
        ////New Table end
        //
        private string GetBonusContrAccInfo(DataTable dtBonusContrAcc)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            BonusContrAcc objOp = null;
            //OperationList objOpList = new OperationList();

            try
            {
                //Clear Previous Data
                objProvider.RunQuery("Delete from " + dtBonusContrAcc.TableName);

                for (int k = 0; k < dtBonusContrAcc.Rows.Count; k++)
                {
                    objOp = new BonusContrAcc();
                    //objSt.BANK_CODE = BankCode;

                    for (int j = 0; j < dtBonusContrAcc.Columns.Count; j++)
                    {
                        #region setting properties values

                        if (dtBonusContrAcc.Columns[j].ColumnName == "StatementNo")
                        {
                            objOp.STATEMENTNO = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtBonusContrAcc.Columns[j].ColumnName == "SUM_CREDIT")
                        {
                            objOp.SUM_CREDIT = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtBonusContrAcc.Columns[j].ColumnName == "SUM_DEBIT")
                        {
                            objOp.SUM_DEBIT = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtBonusContrAcc.Columns[j].ColumnName == "EBALANCE")
                        {
                            objOp.EBALANCE = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        /*else if (dtBonusContrAcc.Columns[j].ColumnName == "A")
                        {
                            if (dtBonusContrAcc.Rows[k][j].ToString() == "" || dtBonusContrAcc.Rows[k][j].ToString() == null)
                                objOp.Amount = "0.00";
                            else
                                objOp.Amount = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                        }*/
                        else if (dtBonusContrAcc.Columns[j].ColumnName == "ACCOUNT_NO")
                        {
                            objOp.ACCOUNT_NO = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtBonusContrAcc.Columns[j].ColumnName == "ACURN")
                        {
                            objOp.ACURN = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtBonusContrAcc.Columns[j].ColumnName == "ACURC")
                        {
                            objOp.ACURC = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtBonusContrAcc.Columns[j].ColumnName == "SBALANCE")
                        {
                            objOp.SBALANCE = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                        }
                        #endregion
                    }
                    //objOpList.Add(objOp);

                    sql = "Insert into BONUSCONTRACC(STATEMENTNO,SUM_CREDIT,SUM_DEBIT,EBALANCE,ACCOUNT_NO,ACURN,ACURC,SBALANCE) " +
                    "Values('" + objOp.STATEMENTNO + "','" + objOp.SUM_CREDIT + "','" + objOp.SUM_DEBIT + "','" + objOp.EBALANCE + "','" + objOp.ACCOUNT_NO + "'," +
                    "'" + objOp.ACURN + "','" + objOp.ACURC + "','" + objOp.SBALANCE + "') ";

                    reply = objProvider.RunQuery(sql);
                    if (!reply.Contains("Success"))
                        return reply;
                }
                return reply;
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return "Error: " + ex.StackTrace;
            }
        }

        //
        private string GetCardHolderAccountInfo(DataTable dtAccount)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            Account objAc = null;
            AccountList objAcList = new AccountList();

            try
            {
                //Clear Previous Data
                objProvider.RunQuery("Delete from " + dtAccount.TableName);
                objAc = new Account();

                for (int k = 0; k < dtAccount.Rows.Count; k++)
                {
                    objAc = new Account();

                    for (int j = 0; j < dtAccount.Columns.Count; j++)
                    {
                        #region setting properties values

                        if (dtAccount.Columns[j].ColumnName == "StatementNo")
                        {
                            objAc.STATEMENTNO = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "ACCOUNTNO")
                        {
                            objAc.ACCOUNTNO = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "ACURN")
                        {
                            objAc.ACURN = dtAccount.Rows[k][j].ToString().Replace("'", "''"); 
                        }
                        else if (dtAccount.Columns[j].ColumnName == "SBALANCE")
                        {
                            objAc.SBALANCE = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "ACURC")
                        {
                            objAc.ACURC = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "EBALANCE")
                        {
                            objAc.EBALANCE = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "AVAIL_CRD_LIMIT")
                        {
                            objAc.AVAIL_CRD_LIMIT = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "AVAIL_CASH_LIMIT")
                        {
                            objAc.AVAIL_CASH_LIMIT = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "SUM_WITHDRAWAL")
                        {
                            objAc.SUM_WITHDRAWAL = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "SUM_INTEREST")
                        {
                            objAc.SUM_INTEREST = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "OVLFEE_AMOUNT")
                        {
                            objAc.OVLFEE_AMOUNT = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "OVDFEE_AMOUNT")
                        {
                            objAc.OVDFEE_AMOUNT = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "SUM_REVERSE")
                        {
                            objAc.SUM_REVERSE = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "SUM_CREDIT")
                        {
                            objAc.SUM_CREDIT = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "SUM_OTHER")
                        {
                            objAc.SUM_OTHER = dtAccount.Rows[k][j].ToString();
                        }
                        if (dtAccount.Columns[j].ColumnName == "SUM_PURCHASE")
                        {
                            objAc.SUM_PURCHASE = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "MIN_AMOUNT_DUE")
                        {
                            objAc.MIN_AMOUNT_DUE = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "CASH_LIMIT")
                        {
                            objAc.CASH_LIMIT = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "CRD_LIMIT")
                        {
                            objAc.CRD_LIMIT = dtAccount.Rows[k][j].ToString();
                        }
                        #endregion
                    }
                    objAcList.Add(objAc);

                    sql = "Insert into Account(STATEMENTNO,ACCOUNTNO,ACURN,SBALANCE,ACURC,EBALANCE,AVAIL_CRD_LIMIT,AVAIL_CASH_LIMIT," +
                        "SUM_WITHDRAWAL,SUM_INTEREST,OVLFEE_AMOUNT,OVDFEE_AMOUNT,SUM_REVERSE,SUM_CREDIT,SUM_OTHER,SUM_PURCHASE,MIN_AMOUNT_DUE,CASH_LIMIT,CRD_LIMIT)" +
                        " Values('" + objAc.STATEMENTNO + "','" + objAc.ACCOUNTNO + "','" + objAc.ACURN + "','" + objAc.SBALANCE + "','" + objAc.ACURC + "'," +
                        "'" + objAc.EBALANCE + "','" + objAc.AVAIL_CRD_LIMIT + "','" + objAc.AVAIL_CASH_LIMIT + "','" + objAc.SUM_WITHDRAWAL + "'," +
                        "'" + objAc.SUM_INTEREST + "','" + objAc.OVLFEE_AMOUNT + "','" + objAc.OVDFEE_AMOUNT + "','" + objAc.SUM_REVERSE + "'," +
                        "'" + objAc.SUM_CREDIT + "','" + objAc.SUM_OTHER + "','" + objAc.SUM_PURCHASE + "','" + objAc.MIN_AMOUNT_DUE + "','" + objAc.CASH_LIMIT + "','" + objAc.CRD_LIMIT + "')";

                                        
                    reply = objProvider.RunQuery(sql);
                    if (!reply.Contains("Success"))
                        return reply;
                }
                return reply;
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return "Error: " + ex.StackTrace;
            }
        }
        //
        private string GetCardHolderCardInfo(DataTable dtCard)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            Card objCard = null;
            int @vSL = 0;

            DataSet dsSL = new DataSet();

            dsSL = objProvider.ReturnData(
                "SELECT ISNULL(MAX(SLNO), 0) AS MAXSL FROM Card",
                ref reply
            );

            if (dsSL != null && dsSL.Tables.Count > 0 && dsSL.Tables[0].Rows.Count > 0)
            {
                int.TryParse(dsSL.Tables[0].Rows[0]["MAXSL"].ToString(), out vSL);
            }

            CardList objCardList = new CardList();

            try
            {
                //Clear Previous Data
              //  objProvider.RunQuery("Delete from " + dtCard.TableName);

                for (int k = 0; k < dtCard.Rows.Count; k++)
                {
                    objCard = new Card();
                    vSL++; 

                    for (int j = 0; j < dtCard.Columns.Count; j++)
                    {
                        #region setting properties values

                        if (dtCard.Columns[j].ColumnName == "StatementNo")
                        {
                            objCard.STATEMENTNO = dtCard.Rows[k][j].ToString();
                        }
                        else if (dtCard.Columns[j].ColumnName == "PAN")
                        {
                            objCard.PAN = dtCard.Rows[k][j].ToString();
                        }
                        else if (dtCard.Columns[j].ColumnName == "MBR")
                        {
                            objCard.MBR = dtCard.Rows[k][j].ToString();
                        }
                        else if (dtCard.Columns[j].ColumnName == "CLIENTNAME")
                        {
                           //objCard.CLIENTNAME = dtCard.Rows[k][j].ToString().Replace("'", "");
                           objCard.CLIENTNAME = Regex.Replace(dtCard.Rows[k][j].ToString().Replace("'", "''"), @"\s+", " "); //to ignore extra blank spaces 
                        }

                        #endregion
                    }
                    objCardList.Add(objCard);

                    sql = "Insert into Card(STATEMENTNO,PAN,MBR,CLIENTNAME,SLNO)" +
                        " Values('" + objCard.STATEMENTNO + "','" + objCard.PAN + "','" + objCard.MBR + "','" + objCard.CLIENTNAME + "','" + @vSL + "')";

                    reply = objProvider.RunQuery(sql);
                    if (!reply.Contains("Success"))
                        return reply;
                }
                return reply;
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return "Error: " + ex.StackTrace;
            }
        }

        //
        private void GenerateStatementInfo(DataSet dsStatement,string BankName, ref string errMsg)
        {
            string reply = string.Empty;
            errMsg = string.Empty;

            try
            {
                DataTable dtOperation = dsStatement.Tables["Operation"];
                DataSet dsBDT = objProvider.ReturnData("select * from Qry_Card_Account where Curr='BDT'", ref reply);

                if (dsBDT != null)
                {
                    if (dsBDT.Tables.Count > 0)
                    {
                        if (dsBDT.Tables[0].Rows.Count > 0)
                        {
                            DataTable dtStatementBDT = dsBDT.Tables[0];
                            ProcessBDTCurrency(dtStatementBDT, dtOperation, BankName, ref errMsg);
                        }
                    }
                }

                reply = string.Empty;
                errMsg = string.Empty;
                DataSet dsUSD = objProvider.ReturnData("select * from Qry_Card_Account where Curr='USD'", ref reply);

                if (dsUSD != null)
                {
                    if (dsUSD.Tables.Count > 0)
                    {
                        if (dsUSD.Tables[0].Rows.Count > 0)
                        {
                            DataTable dtStatementUSD = dsUSD.Tables[0];
                            ProcessUSDCurrency(dtStatementUSD, dtOperation, BankName, ref errMsg);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                errMsg = ex.StackTrace;
            }

        }

        private void ProcessBDTCurrency(DataTable dtStatement, DataTable dtOperation, string BankName, ref string errMsg)
        {
            #region BDT
            string reply = string.Empty;
            string sql = string.Empty;
            StatementInfo objSt = null;
            //StatementInfoList objStList = new StatementInfoList();

            for (int k = 0; k < dtStatement.Rows.Count; k++)
            {
                //if(k==167)
                //    sql = string.Empty;
                try
                {
                    objSt = new StatementInfo();

                    objSt.BANK_CODE = BankName;

                    for (int j = 0; j < dtStatement.Columns.Count; j++)
                    {
                        #region setting properties values

                        if (dtStatement.Columns[j].ColumnName.ToUpper() == "STATEMENTNO")
                        {
                            objSt.STATEMENTNO = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CONTRACTNO")
                        {
                            objSt.CONTRACTNO = dtStatement.Rows[k][j].ToString();
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "IDCLIENT")
                        {
                            objSt.IDCLIENT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ADDRESS")
                        {
                            objSt.ADDRESS = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "PROMOTIONALTEXT")
                        {
                            objSt.PROMOTIONALTEXT = dtStatement.Rows[k][j].ToString().Replace("'", "''"); ;
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "PAN")
                        {
                            if (dtStatement.Rows[k][j].ToString().Length >= 16)
                               objSt.PAN = dtStatement.Rows[k][j].ToString().Substring(0, 16);

                               
                            else
                            {
                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Card Not fount for the Contract " + objSt.CONTRACTNO });
                                MsgLogWriter objLW = new MsgLogWriter();
                                objLW.logTrace(_LogPath, "EStatement.log", "Card Not fount for the Contract " + objSt.CONTRACTNO);
                                continue;
                            }
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "REGION")
                        {
                            objSt.CITY = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ZIP")
                        {
                            objSt.ZIP = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "COUNTRY")
                        {
                            objSt.COUNTRY = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "EMAIL")
                        {
                            objSt.EMAIL = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "MOBILE")
                        {
                            objSt.MOBILE = dtStatement.Rows[k][j].ToString().Replace("(", "").Replace(")", "").Replace("8800", "880");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "TITLE")
                        {
                            objSt.TITLE = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "JOBTITLE")
                        {
                            objSt.JOBTITLE = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CLIENT")
                        {
                            objSt.CLIENTNAME = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ACCOUNTNO")
                        {
                            objSt.ACCOUNTNO = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CURR")
                        {
                            objSt.ACURN = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "PBAL")
                        {
                            objSt.SBALANCE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "TOTINTEREST")
                        {
                            objSt.SUM_INTEREST = dtStatement.Rows[k][j].ToString();
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "STARTDATE")
                        {
                            objSt.STARTDATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ENDDATE")
                        {
                            objSt.ENDDATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "NEXT_STATEMENT_DATE")
                        {
                            objSt.NEXT_STATEMENT_DATE = dtStatement.Rows[k][j].ToString();
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "PAYDATE")
                        {
                            objSt.PAYMENT_DATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "STDATE")
                        {
                            objSt.STATEMENT_DATE = dtStatement.Rows[k][j].ToString();
                            objSt.STATEMENTID = dtStatement.Rows[k][j].ToString().Replace("/", ""); ;
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ACURC")
                        {
                            objSt.ACURC = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "OVLFEE_AMOUNT")
                        {
                            objSt.OVLFEE_AMOUNT = dtStatement.Rows[k][j].ToString().Replace("-", "");
                        }

                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ODAMOUNT")
                        {
                            objSt.OVDFEE_AMOUNT = dtStatement.Rows[k][j].ToString().Replace("-", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "MINPAY")
                        {
                            objSt.MIN_AMOUNT_DUE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "TOTLIMIT")
                        {
                            objSt.CRD_LIMIT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "TOTPURCHASE")
                        {
                            objSt.SUM_PURCHASE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "SUM_REVERSE")
                        {
                            objSt.SUM_REVERSE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "SUM_CREDIT")
                        {
                            objSt.SUM_CREDIT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "SUM_OTHER")
                        {
                            objSt.SUM_OTHER = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CASHADV")
                        {
                            objSt.SUM_WITHDRAWAL = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "AVLIMIT")
                        {
                            objSt.AVAIL_CRD_LIMIT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "AVCASHLIMIT")
                        {
                            objSt.AVAIL_CASH_LIMIT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "LASTBAL")
                        {
                            objSt.EBALANCE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CASH_LIMIT")
                        {
                            objSt.CASH_LIMIT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        #endregion
                    }

                    objSt.STM_MSG = txtStmMsg.Text.ToString().Replace("'","''");
                    objSt.STATUS = "1";

                    sql = "Insert into STATEMENT_INFO(STATEMENTID,BANK_CODE,CONTRACTNO,IDCLIENT,PAN,TITLE,CLIENTNAME,JOBTITLE,STATEMENTNO,ADDRESS,CITY,ZIP,COUNTRY," +
                        "EMAIL,MOBILE,STARTDATE,ENDDATE,NEXT_STATEMENT_DATE,PAYMENT_DATE,STATEMENT_DATE,ACCOUNTNO,ACURN,SBALANCE,ACURC,EBALANCE,AVAIL_CRD_LIMIT," +
                        "AVAIL_CASH_LIMIT,SUM_WITHDRAWAL,SUM_INTEREST,OVLFEE_AMOUNT,OVDFEE_AMOUNT,SUM_REVERSE,SUM_CREDIT,SUM_OTHER,SUM_PURCHASE," +
                        "MIN_AMOUNT_DUE,CASH_LIMIT,CRD_LIMIT,STM_MSG,STATUS,PromotionalText) VALUES('" + objSt.STATEMENTID + "'," +
                        "'" + objSt.BANK_CODE + "','" + objSt.CONTRACTNO + "','" + objSt.IDCLIENT + "','" + objSt.PAN + "','" + objSt.TITLE + "','" + objSt.CLIENTNAME + "','" + objSt.JOBTITLE + "','" + objSt.STATEMENTNO + "'," +
                        "'" + objSt.ADDRESS + "','" + objSt.CITY + "','" + objSt.ZIP + "','" + objSt.COUNTRY + "','" + objSt.EMAIL + "','" + objSt.MOBILE + "','" + objSt.STARTDATE + "','" + objSt.ENDDATE + "'," +
                        "'" + objSt.NEXT_STATEMENT_DATE + "','" + objSt.PAYMENT_DATE + "','" + objSt.STATEMENT_DATE + "','" + objSt.ACCOUNTNO + "','" + objSt.ACURN + "'," +
                        "'" + objSt.SBALANCE + "','" + objSt.ACURC + "','" + objSt.EBALANCE + "','" + objSt.AVAIL_CRD_LIMIT + "','" + objSt.AVAIL_CASH_LIMIT + "'," +
                        "'" + objSt.SUM_WITHDRAWAL + "','" + objSt.SUM_INTEREST + "','" + objSt.OVLFEE_AMOUNT + "','" + objSt.OVDFEE_AMOUNT + "','" + objSt.SUM_REVERSE + "'," +
                        "'" + objSt.SUM_CREDIT + "','" + objSt.SUM_OTHER + "','" + objSt.SUM_PURCHASE + "','" + objSt.MIN_AMOUNT_DUE + "','" + objSt.CASH_LIMIT + "'," +
                        "'" + objSt.CRD_LIMIT + "','" + objSt.STM_MSG + "','" + objSt.STATUS  + "','" + objSt.PROMOTIONALTEXT +"')";

                    reply = objProvider.RunQuery(sql);
                    //DataTable dtOperation = dsStatement.Tables["Operation"];
                    string trn_Date = string.Empty;
                    if (dtOperation != null && dtOperation.Columns.Contains("ACCOUNT"))
                    {
                        #region  ACCOUNT CHECK

                        if (dtOperation.Rows.Count > 0)
                        {
                            #region  dtOperation Row Check

                            DataRow[] dr = dtOperation.Select("STATEMENTNO='" + objSt.STATEMENTNO + "' AND ACCOUNT='" + objSt.ACCOUNTNO + "'");
                            if (dr.Length > 0)
                            {

                                #region  STATEMENTNO,ACCOUNT  >> FOUND
                                //  double feesnCharges = 0.00;
                                

                                for (int l = 0; l < dr.Length; l++)
                                {
                                    #region setting properties values
                                    List<string> INTlist = new List<string>() {"MARK UP PROFIT (R)", "INTEREST ON FEES & CHARGES","INTEREST ON INTEREST","INTEREST ON ATM TRANSACTION", "INTEREST ON POS TRANSACTION", "INTEREST ON CARD CHEQUE","CHARGE INTEREST FOR 0", "CHARGE INTEREST FOR 1", "CHARGE INTEREST FOR 2", "CHARGE INTEREST FOR 3", "CHARGE INTEREST FOR 4", "CHARGE INTEREST FOR 5", "CHARGE INTEREST FOR 6", "CHARGE INTEREST FOR 7", "CHARGE INTEREST FOR 8", "CHARGE INTEREST FOR 9", "CHARGE INTEREST FOR 10", "CHARGE INTEREST FOR 11", "CHARGE INTEREST FOR 0 OPERATIONS GROUP", "CHARGE INTEREST FOR 1 OPERATIONS GROUP", "CHARGE INTEREST FOR 2 OPERATIONS GROUP", "CHARGE INTEREST FOR 3 OPERATIONS GROUP", "CHARGE INTEREST FOR 4 OPERATIONS GROUP", "CHARGE INTEREST FOR 5 OPERATIONS GROUP", "CHARGE INTEREST FOR 6 OPERATIONS GROUP", "CHARGE INTEREST FOR 7 OPERATIONS GROUP", "INTEREST ON FUND TRANSFER", "INTEREST ON BALANCE TRANSFER", "INTEREST ON EMI", "INTEREST ON FT", "INTEREST ON BT", "INTEREST ON BANK POS TRANSACTION",
                                    "INTEREST ON BPOS TRANSACTION","CHARGE INTEREST FOR INTEREST OPERATIONS", "CHARGE INTEREST FOR POS OPERATIONS", "CHARGE INTEREST FOR ATM OPERATIONS", "LATE PAYMENT CHARGE FOR GROUP 1", "LATE PAYMENT CHARGE FOR GROUP 2", "LATE PAYMENT CHARGE FOR GROUP 3", "CHARGE OF A DEBT FOR CREDIT OVERDRAFTING" ,"INTEREST ON SERVICE FEE","INTEREST ON PREVIOUS BALANCE","EMI INTER"};
                                      
                                    if (INTlist.Contains(dr[l]["D"].ToString().ToUpper()) == false)
                                    {
                                        StatementDetails objSTD = new StatementDetails();
                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                        objSTD.PAN = objSt.PAN;

                                        if (dr[l].Table.Columns.Contains("ACCOUNT"))
                                            objSTD.ACCOUNTNO = dr[l]["ACCOUNT"].ToString();

                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;

                                        if (dr[l].Table.Columns.Contains("TD"))
                                            objSTD.TRNDATE = dr[l]["TD"].ToString();

                                        if (dr[l].Table.Columns.Contains("OD"))
                                            objSTD.POSTDATE = dr[l]["OD"].ToString();

                                        if (dr[l].Table.Columns.Contains("ACURN"))
                                            objSTD.ACURN = dr[l]["ACURN"].ToString();

                                        if (dr[l].Table.Columns.Contains("FR"))
                                            objSTD.FR = dr[l]["FR"].ToString().Replace("'", "''");

                                        if (dr[l].Table.Columns.Contains("DE"))
                                            objSTD.DE = dr[l]["DE"].ToString().Replace("'", "''");

                                        if (dr[l].Table.Columns.Contains("SERIALNO"))
                                            objSTD.SERIALNO = dr[l]["SERIALNO"].ToString();

                                        if (dr[l].Table.Columns.Contains("P"))   //Add new column from Operation 06.02.2017
                                        {
                                            if (dr[l]["P"].ToString() == "" || dr[l]["P"].ToString() == null) // NULL P TAG
                                            {
                                                if (prePan != objSt.PAN && preDoc == dr[l]["DOCNO"].ToString())  // PARENT P TAG CHECK
                                                {
                                                    objSTD.P = prePan;
                                                    prePan = objSTD.P;
                                                }
                                                else
                                                {
                                                    objSTD.P = objSt.PAN;
                                                    prePan = objSt.PAN;
                                                }


                                            }

                                            else
                                            {
                                                objSTD.P = dr[l]["P"].ToString();
                                                prePan = dr[l]["P"].ToString();
                                            }
                                        }

                                        if (dr[l].Table.Columns.Contains("DOCNO"))   //Add new column from Operation 06.02.2017
                                        {
                                            objSTD.DOCNO = dr[l]["DOCNO"].ToString();
                                            preDoc = dr[l]["DOCNO"].ToString();
                                        }

                                        if (dr[l].Table.Columns.Contains("NO"))   //Add new column from Operation 06.02.2017
                                        {
                                            objSTD.NO = dr[l]["NO"].ToString();
                                        }

                                        if (dr[l].Table.Columns.Contains("OCC"))
                                        {
                                            DataTable dtOcc = new DataTable();
                                            dtOcc = objProvider.ReturnData("select * from CURRENCYCODE", ref reply).Tables[0];// where Curr='BDT'
                                            DataRow[] drr = dtOcc.Select();
                                            string sp = string.Empty;
                                            string Sc = string.Empty;
                                            for (int x = 0; x <= 183; x++)
                                            {
                                                sp = dr[l]["OCC"].ToString();
                                                Sc = drr[x]["OCC"].ToString();
                                                if (dr[l]["OCC"].ToString() == drr[x]["OCC"].ToString())
                                                    objSTD.OC = drr[x]["Name"].ToString();
                                            }
                                        }
                                        else
                                            objSTD.OC = "";// dr[l]["OC"].ToString();



                                        if (dr[l].Table.Columns.Contains("AMOUNTSIGN"))
                                            objSTD.AMOUNTSIGN = dr[l]["AMOUNTSIGN"].ToString();

                                        if (dr[l].Table.Columns.Contains("ACURN"))
                                        {
                                            if (dr[l]["A"].ToString() == "" || dr[l]["A"].ToString() == null)
                                                objSTD.AMOUNT = "0.00";
                                            else
                                                objSTD.AMOUNT = dr[l]["A"].ToString();
                                        }
                                        else objSTD.AMOUNT = "0.00";

                                        if (dr[l].Table.Columns.Contains("OCC"))
                                        {
                                            if (dr[l]["OA"].ToString() == "" || dr[l]["OA"].ToString() == null)
                                                objSTD.ORGAMOUNT = "0.00";
                                            else
                                                objSTD.ORGAMOUNT = dr[l]["OA"].ToString();
                                        }
                                        else objSTD.ORGAMOUNT = "0.00";

                                        //Remmove Terminal Name when Fee and VAT Impose
                                        //Sum Charges amount with Fees & Charges. 

                                        #region  #region Monthly EMI ,TRANSFER TO EMI,EMI CANCELLED,EMI

                                        if ((dr[l]["D"].ToString().ToUpper().Contains("MONTHLY EMI")) || (dr[l]["D"].ToString().ToUpper().Contains("TRANSFER TO EMI")) || (dr[l]["D"].ToString().ToUpper().Contains("EMI CANCELLED")))
                                        {
                                            if (dr[l].Table.Columns.Contains("FR"))
                                            {
                                                if (dr[l]["FR"].ToString() == "" || dr[l]["FR"].ToString() == null)
                                                    if (dr[l].Table.Columns.Contains("TL"))
                                                    {
                                                        objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                                    }
                                                    else
                                                    {
                                                        objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                                    }
                                                else
                                                {
                                                    string data = dr[l]["FR"].ToString().Replace("'", "''");
                                                    bool contains = data.IndexOf("[VALUE NOT DEFINED]", StringComparison.OrdinalIgnoreCase) >= 0;
                                                    if (contains == true)
                                                    {
                                                        string[] list = data.Split(':');
                                                        objSTD.TRNDESC = list[0];
                                                    }
                                                    else
                                                    {
                                                        objSTD.TRNDESC = data.Replace("\n", "").Replace("\r", "");
                                                    }

                                                }
                                            }
                                            else
                                               // objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                            objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''"); // modify

                                        }

                                        #endregion

                                        #region CHEQUE TRANSACTION
                                        else if ((dr[l]["D"].ToString().ToUpper().Contains("CHEQUE TRANSACTION")) || (dr[l]["D"].ToString().ToUpper().Contains("CARD CHEQUE TRANSACTION")))
                                        {
                                            if (dr[l].Table.Columns.Contains("SERIALNO"))
                                            {
                                                if (dr[l]["SERIALNO"].ToString() == "" || dr[l]["SERIALNO"].ToString() == null)
                                                {
                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + " [CHQ NO:" + "]";
                                                }
                                                else
                                                {
                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " [CHQ NO:" + dr[l]["SERIALNO"].ToString().Replace("'", "") + "]";
                                                }
                                            }
                                            else
                                            {
                                                objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + " [CHQ NO:" + "]";
                                            }
                                        }

                                        #endregion

                                        #region Rest of Txn
                                        else
                                        {
                                            if (dr[l].Table.Columns.Contains("TL"))
                                            {
                                                if (dr[l]["FR"].ToString().ToUpper().Contains("A 10") || dr[l]["FR"].ToString().ToUpper().Contains("A 64") || dr[l]["FR"].ToString().ToUpper().Contains("P 14") || dr[l]["FR"].ToString().ToUpper().Contains("P 32") || dr[l]["FR"].ToString().ToUpper().Contains("P 33") || dr[l]["FR"].ToString().ToUpper().Contains("F 29") || dr[l]["FR"].ToString().ToUpper().Contains("P 13"))
                                                {
                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                                    // objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''") + " " + "(" + objSTD.OC + " " + dr[l]["OA"].ToString() + ")";

                                                }
                                                else
                                                {
                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                                }

                                             /*   if (dr[l]["D"].ToString().ToUpper().Contains("PURCHASE"))
                                                {
                                                    if (dr[l]["D"].ToString().Trim().Length > 8)
                                                    {

                                                        objSTD.TRNDESC = (dr[l]["D"].ToString().ToUpper().Replace("PURCHASE", "")).Trim() + " " + dr[l]["TL"].ToString().Replace("'", "''");

                                                    }
                                                    else
                                                    {

                                                        objSTD.TRNDESC = (dr[l]["D"].ToString().ToUpper().Replace("PURCHASE", "")).Trim() + dr[l]["TL"].ToString().Replace("'", "''");
                                                    }
                                                } */

                                                if (dr[l]["D"].ToString().ToUpper().Contains("PURCHASE"))
                                                {

                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");

                                                    
                                                } 


                                            }

                                            else
                                            {
                                                objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                            }

                                        }

                                        #endregion

                                        #region PAYMENT CASH DEPOSIT
                                        
                                        if ((objSTD.TRNDESC.ToUpper().Contains("PAYMENT RECEIVED (THANK YOU)")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT RECEIVED, THANK YOU.")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT RECEIVED, THANK YOU")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT RECEIVED [AUTO DEBIT]")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT RECEIVED [CASH]")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT CASH DEPOSIT")) || (objSTD.TRNDESC.ToUpper().Contains("CREDIT CASH DEPOSIT")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT THROUGH BRANCHES  (CASH)")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT BY CHEQUE (MAIL)")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT THROUGH AUTO DEBIT")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT THROUGH CHEQUE")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT THROUGH FC")) || (objSTD.TRNDESC.ToUpper().Contains("VISA PAYMENT")) || (objSTD.TRNDESC.ToUpper().Contains("MC PAYMENT")))
                                        {
                                            objSTD.TRNDESC = "PAYMENT RECEIVED (THANK YOU)";
                                            //objSTD.TRNDATE = dr[l]["OD"].ToString();
                                        }

                                        #endregion

                                        #region APPROVAL
                                        if (dr[l].Table.Columns.Contains("APPROVAL"))
                                        {
                                            objSTD.APPROVAL = dr[l]["APPROVAL"].ToString().Replace("'", "''");

                                            if (dr[l]["APPROVAL"].ToString() != "" && objSTD.TRNDATE == "")
                                            {
                                                objSTD.TRNDATE = dr[l]["OD"].ToString();
                                            }
                                        }
                                        #endregion

                                        #region CASH ADVANCE

                                        try
                                        {
                                            if ((dr[l]["D"].ToString().ToUpper().Trim() == ("CASH ADVANCE")))
                                            {

                                                objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                            }
                                        }

                                        catch (Exception ex)
                                        {
                                            objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                        }

                                        #endregion

                                        #region INTEREST CHARGES TRANSACTION

                                        if ((dr[l]["D"].ToString().ToUpper().Trim() == ("INTEREST CHARGES")))
                                        {

                                            objSTD.TRNDESC = "INTEREST CHARGE";
                                        }


                                        #endregion

                                        //objSTD.AMOUNTSIGN = dr[l]["AMOUNTSIGN"].ToString();
                                        if (dr[l].Table.Columns.Contains("TD"))
                                            objSTD.TRNDATE = dr[l]["TD"].ToString();

                                        if (!dr[l].Table.Columns.Contains("P"))   //Add new column from Operation 06.02.2017
                                        {
                                            objSTD.P = objSt.PAN;
                                        }

                                        sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,OC,ORGAMOUNT,AMOUNTSIGN,APPROVAL,FR,SERIALNO,DE,P,DOCNO,NO)" +
                                            " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                            "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.OC + "','" + objSTD.ORGAMOUNT + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.APPROVAL + "','" + objSTD.FR + "','" + objSTD.SERIALNO + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";

                                        reply = objProvider.RunQuery(sql);
                                        if (!reply.Contains("Success"))
                                            errMsg = reply;
                                    }

                                    #endregion
                                }

                               
                                if (objSt.SUM_INTEREST != "0.00")
                                {
                                    #region SUM_INTEREST

                                    StatementDetails objSTD = new StatementDetails();
                                    objSTD.STATEMENTID = objSt.STATEMENTID;
                                    objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                    objSTD.IDCLIENT = objSt.IDCLIENT;
                                    objSTD.PAN = objSt.PAN;
                                    objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                    objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                    objSTD.ACURN = objSt.ACURN;
                                    objSTD.TRNDESC = "INTEREST CHARGES";
                                    //objSTD.TRNDESC = "Profit Charges";
                                    objSTD.AMOUNT = "-" + objSt.SUM_INTEREST;//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                    objSTD.TRNDATE = trn_Date;
                                    objSTD.POSTDATE = trn_Date;
                                    DataTable dtCardbdt = new DataTable();
                                    dtCardbdt = objProvider.ReturnData("SELECT *  FROM  STATEMENT_DETAILS where STATEMENTNO='" + objSt.STATEMENTNO + "' AND P <>'" + objSt.PAN + "' AND ACURN = '" + objSt.ACURN + "'", ref reply).Tables[0];// where Curr='BDT'

                                    if (dtCardbdt.Rows.Count <= 0)
                                    {
                                        objSTD.P = objSt.PAN;
                                    }
                                    else
                                    {
                                        objSTD.P = "000000******0000";
                                    }

                                    sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,FR,SERIALNO,DE,P,DOCNO,NO)" +
                                            " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                            "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.FR + "','" + objSTD.SERIALNO + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";

                                    reply = objProvider.RunQuery(sql);
                                    if (!reply.Contains("Success"))
                                        errMsg = reply;



                                #endregion
                                }
                                else
                                {
                                    DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW", ref reply);

                                    if (dsAcI != null)
                                    {
                                        if (dsAcI.Tables.Count > 0)
                                        {
                                            if (dsAcI.Tables[0].Rows.Count > 0)
                                            {
                                                DataTable dtAcI = dsAcI.Tables[0]; ;
                                                for (int x = 0; x < dtAcI.Rows.Count; x++)
                                                {
                                                    StatementDetails objSTD = new StatementDetails();

                                                    objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                    if (objSTD.CONTRACTNO == dtAcI.Rows[x][1].ToString())
                                                    {
                                                        if (dtAcI.Rows[x][0].ToString() != "0.00")
                                                        {
                                                            objSTD.STATEMENTID = objSt.STATEMENTID;
                                                            objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                            objSTD.IDCLIENT = objSt.IDCLIENT;
                                                            objSTD.PAN = objSt.PAN;
                                                            objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                            objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                            objSTD.ACURN = objSt.ACURN;
                                                            objSTD.TRNDESC = "INTEREST CHARGES";
                                                            objSTD.AMOUNT = "-" + dtAcI.Rows[x][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                            objSTD.TRNDATE = trn_Date;
                                                            objSTD.POSTDATE = trn_Date;
                                                            DataTable dtCardbdt = new DataTable();
                                                            dtCardbdt = objProvider.ReturnData("SELECT *  FROM  STATEMENT_DETAILS where STATEMENTNO='" + objSt.STATEMENTNO + "' AND P <>'" + objSt.PAN + "' AND ACURN = '" + objSt.ACURN + "'", ref reply).Tables[0];// where Curr='BDT'

                                                            if (dtCardbdt.Rows.Count <= 0)
                                                            {
                                                                objSTD.P = objSt.PAN;
                                                            }
                                                            else
                                                            {
                                                                objSTD.P = "000000******0000";
                                                            }

                                                            sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,FR,SERIALNO,DE,P,DOCNO,NO)" +
                                                                    " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                                    "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.FR + "','" + objSTD.SERIALNO + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";


                                                            reply = objProvider.RunQuery(sql);
                                                            if (!reply.Contains("Success"))
                                                                errMsg = reply;

                                                            decimal tempIntAmtI = 0;
                                                            decimal tempIntAmt = 0;
                                                            decimal tempTotalIntAmt = 0;
                                                            string st = string.Empty;

                                                            DataTable dt = new DataTable();
                                                            dt = objProvider.ReturnData("select AMOUNT from STATEMENT_DETAILS WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES' ", ref reply).Tables[0];
                                                            //tempIntAmtI = Convert.ToInt32(dt.Rows[0][0])*(-1);
                                                            tempIntAmtI = Convert.ToDecimal(dt.Rows[0][0]) * (-1);
                                                            st = dtAcI.Rows[x][0].ToString();
                                                            tempIntAmt = Convert.ToDecimal(st);
                                                            tempTotalIntAmt = tempIntAmtI + tempIntAmt;

                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion

                            }
                            else
                            {

                                //string trn_Date = string.Empty;


                                //New View add
                                // DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW", ref reply);
                                DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW", ref reply);
                                if (dsAcI != null)
                                {
                                    if (dsAcI.Tables.Count > 0)
                                    {
                                        if (dsAcI.Tables[0].Rows.Count > 0)
                                        {
                                            DataTable dtAcI = dsAcI.Tables[0]; ;
                                            for (int x = 0; x < dtAcI.Rows.Count; x++)
                                            {
                                                StatementDetails objSTD = new StatementDetails();

                                                objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                if (objSTD.CONTRACTNO == dtAcI.Rows[x][1].ToString())
                                                {
                                                    if (dtAcI.Rows[x][0].ToString() != "0.00")
                                                    {
                                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                                        objSTD.PAN = objSt.PAN;
                                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                        objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                        objSTD.ACURN = objSt.ACURN;
                                                        objSTD.TRNDESC = "INTEREST CHARGES";
                                                        objSTD.AMOUNT = "-" + dtAcI.Rows[x][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                        objSTD.TRNDATE = trn_Date;
                                                        objSTD.POSTDATE = trn_Date;
                                                        DataTable dtCardbdt = new DataTable();
                                                        dtCardbdt = objProvider.ReturnData("SELECT *  FROM  STATEMENT_DETAILS where STATEMENTNO='" + objSt.STATEMENTNO + "' AND P <>'" + objSt.PAN + "' AND ACURN = '" + objSt.ACURN + "'", ref reply).Tables[0];// where Curr='BDT'

                                                        if (dtCardbdt.Rows.Count <= 0)
                                                        {
                                                            objSTD.P = objSt.PAN; 
                                                        }
                                                        else
                                                        {
                                                            objSTD.P = "000000******0000";
                                                        }

                                                        sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,FR,SERIALNO,DE,P,DOCNO,NO)" +
                                                                " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                                "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.FR + "','" + objSTD.SERIALNO + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";


                                                        reply = objProvider.RunQuery(sql);
                                                        if (!reply.Contains("Success"))
                                                            errMsg = reply;

                                                        decimal tempIntAmtI = 0;
                                                        decimal tempIntAmt = 0;
                                                        decimal tempTotalIntAmt = 0;
                                                        string st = string.Empty;

                                                        DataTable dt = new DataTable();
                                                        dt = objProvider.ReturnData("select AMOUNT from STATEMENT_DETAILS WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES' ", ref reply).Tables[0];
                                                        //tempIntAmtI = Convert.ToInt32(dt.Rows[0][0])*(-1);
                                                        tempIntAmtI = Convert.ToDecimal(dt.Rows[0][0]) * (-1);
                                                        st = dtAcI.Rows[x][0].ToString();
                                                        tempIntAmt = Convert.ToDecimal(st);
                                                        tempTotalIntAmt = tempIntAmtI + tempIntAmt;

                                                    }
                                                }
                                            }
                                        }
                                    }
                                }



                            }

                            #endregion
                        }

                            
                    }
                    else
                    {
                        if (dtOperation.Rows.Count > 0)
                        {

                            DataRow[] dr = dtOperation.Select("STATEMENTNO='" + objSt.STATEMENTNO + "'");
                            if (dr.Length > 0)
                            {

                               // string trn_Date = string.Empty;
                                //New View add
                                // DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW", ref reply);
                                DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW", ref reply);

                                if (dsAcI != null)
                                {
                                    if (dsAcI.Tables.Count > 0)
                                    {
                                        if (dsAcI.Tables[0].Rows.Count > 0)
                                        {
                                            DataTable dtAcI = dsAcI.Tables[0]; ;
                                            for (int x = 0; x < dtAcI.Rows.Count; x++)
                                            {
                                                StatementDetails objSTD = new StatementDetails();

                                                objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                if (objSTD.CONTRACTNO == dtAcI.Rows[x][1].ToString())
                                                {
                                                    if (dtAcI.Rows[x][0].ToString() != "0.00")
                                                    {
                                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                                        objSTD.PAN = objSt.PAN;
                                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                        objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                        objSTD.ACURN = objSt.ACURN;
                                                        objSTD.TRNDESC = "INTEREST CHARGES";
                                                        objSTD.AMOUNT = "-" + dtAcI.Rows[x][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                        objSTD.TRNDATE = trn_Date;
                                                        objSTD.POSTDATE = trn_Date;
                                                        DataTable dtCardbdt = new DataTable();
                                                        dtCardbdt = objProvider.ReturnData("SELECT *  FROM  STATEMENT_DETAILS where STATEMENTNO='" + objSt.STATEMENTNO + "' AND P <>'" + objSt.PAN + "' AND ACURN = '" + objSt.ACURN + "'", ref reply).Tables[0];// where Curr='BDT'

                                                        if (dtCardbdt.Rows.Count <= 0)
                                                        {
                                                            objSTD.P = objSt.PAN;
                                                        }
                                                        else
                                                        {
                                                            objSTD.P = "000000******0000";
                                                        }

                                                        sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,FR,SERIALNO,DE,P,DOCNO,NO)" +
                                                                " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                                "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.FR + "','" + objSTD.SERIALNO + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";


                                                        reply = objProvider.RunQuery(sql);
                                                        if (!reply.Contains("Success"))
                                                            errMsg = reply;

                                                        decimal tempIntAmtI = 0;
                                                        decimal tempIntAmt = 0;
                                                        decimal tempTotalIntAmt = 0;
                                                        string st = string.Empty;

                                                        DataTable dt = new DataTable();
                                                        dt = objProvider.ReturnData("select AMOUNT from STATEMENT_DETAILS WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES' ", ref reply).Tables[0];
                                                        //tempIntAmtI = Convert.ToInt32(dt.Rows[0][0])*(-1);
                                                        tempIntAmtI = Convert.ToDecimal(dt.Rows[0][0]) * (-1);
                                                        st = dtAcI.Rows[x][0].ToString();
                                                        tempIntAmt = Convert.ToDecimal(st);
                                                        tempTotalIntAmt = tempIntAmtI + tempIntAmt;

                                                    }
                                                }
                                            }
                                        }
                                    }
                                }


                            }
                        }



                    }

                        #endregion

                    }
                catch (Exception ex)
                {
                    errMsg = "Error: " + ex.Message;
                }
            }
        }
            #endregion BDT

        private void ProcessUSDCurrency(DataTable dtStatement, DataTable dtOperation, string BankName, ref string errMsg)
        {
            #region USD
            string reply = string.Empty;
            string sql = string.Empty;
            StatementInfo objSt = null;
            //StatementInfoList objStList = new StatementInfoList();

            for (int k = 0; k < dtStatement.Rows.Count; k++)
            {
                objSt = new StatementInfo();

                objSt.BANK_CODE = BankName;

                try
                {

                    for (int j = 0; j < dtStatement.Columns.Count; j++)
                    {
                        #region setting properties values

                        if (dtStatement.Columns[j].ColumnName.ToUpper() == "STATEMENTNO")
                        {
                            objSt.STATEMENTNO = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ADDRESS")
                        {
                            objSt.ADDRESS = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CONTRACTNO")
                        {
                            objSt.CONTRACTNO = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "IDCLIENT")
                        {
                            objSt.IDCLIENT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "PAN")
                        {
                            objSt.PAN = dtStatement.Rows[k][j].ToString().Replace("'", "").Substring(0, 16);
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "REGION")
                        {
                            objSt.CITY = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ZIP")
                        {
                            objSt.ZIP = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "COUNTRY")
                        {
                            objSt.COUNTRY = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "EMAIL")
                        {
                            objSt.EMAIL = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "MOBILE")
                        {
                            objSt.MOBILE = dtStatement.Rows[k][j].ToString().Replace("'", "").Replace("(", "").Replace(")", "").Replace("8800", "880");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "TITLE")
                        {
                            objSt.TITLE = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "JOBTITLE")
                        {
                            objSt.JOBTITLE = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CLIENT")
                        {
                            objSt.CLIENTNAME = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ACCOUNTNO")
                        {
                            objSt.ACCOUNTNO = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CURR")
                        {
                            objSt.ACURN = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "PBAL")
                        {
                            objSt.SBALANCE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "TOTINTEREST")
                        {
                            objSt.SUM_INTEREST = dtStatement.Rows[k][j].ToString();
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "STARTDATE")
                        {
                            objSt.STARTDATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ENDDATE")
                        {
                            objSt.ENDDATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "NEXT_STATEMENT_DATE")
                        {
                            objSt.NEXT_STATEMENT_DATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "PAYDATE")
                        {
                            objSt.PAYMENT_DATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "STDATE")
                        {
                            objSt.STATEMENT_DATE = dtStatement.Rows[k][j].ToString();
                            objSt.STATEMENTID = dtStatement.Rows[k][j].ToString().Replace("/", ""); ;
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ACURC")
                        {
                            objSt.ACURC = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "OVLFEE_AMOUNT")
                        {
                            objSt.OVLFEE_AMOUNT = dtStatement.Rows[k][j].ToString().Replace("-", "");
                        }

                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ODAMOUNT")
                        {
                            objSt.OVDFEE_AMOUNT = dtStatement.Rows[k][j].ToString().Replace("-", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "MINPAY")
                        {
                            objSt.MIN_AMOUNT_DUE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "TOTLIMIT")
                        {
                            objSt.CRD_LIMIT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "TOTPURCHASE")
                        {
                            objSt.SUM_PURCHASE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "SUM_REVERSE")
                        {
                            objSt.SUM_REVERSE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "SUM_CREDIT")
                        {
                            objSt.SUM_CREDIT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "SUM_OTHER")
                        {
                            objSt.SUM_OTHER = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CASHADV")
                        {
                            objSt.SUM_WITHDRAWAL = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "AVLIMIT")
                        {
                            objSt.AVAIL_CRD_LIMIT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "AVCASHLIMIT")
                        {
                            objSt.AVAIL_CASH_LIMIT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "LASTBAL")
                        {
                            objSt.EBALANCE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CASH_LIMIT")
                        {
                            objSt.CASH_LIMIT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        #endregion
                    }

                    objSt.STM_MSG = txtStmMsg.Text.ToString().Replace("'","''");
                    objSt.STATUS = "1";

                    sql = "Insert into STATEMENT_INFO(STATEMENTID,BANK_CODE,CONTRACTNO,IDCLIENT,PAN,TITLE,CLIENTNAME,JOBTITLE,STATEMENTNO,ADDRESS,CITY,ZIP,COUNTRY," +
                        "EMAIL,MOBILE,STARTDATE,ENDDATE,NEXT_STATEMENT_DATE,PAYMENT_DATE,STATEMENT_DATE,ACCOUNTNO,ACURN,SBALANCE,ACURC,EBALANCE,AVAIL_CRD_LIMIT," +
                        "AVAIL_CASH_LIMIT,SUM_WITHDRAWAL,SUM_INTEREST,OVLFEE_AMOUNT,OVDFEE_AMOUNT,SUM_REVERSE,SUM_CREDIT,SUM_OTHER,SUM_PURCHASE," +
                        "MIN_AMOUNT_DUE,CASH_LIMIT,CRD_LIMIT,STM_MSG,STATUS) VALUES('" + objSt.STATEMENTID + "'," +
                        "'" + objSt.BANK_CODE + "','" + objSt.CONTRACTNO + "','" + objSt.IDCLIENT + "','" + objSt.PAN + "','" + objSt.TITLE + "','" + objSt.CLIENTNAME + "','" + objSt.JOBTITLE + "','" + objSt.STATEMENTNO + "'," +
                        "'" + objSt.ADDRESS + "','" + objSt.CITY + "','" + objSt.ZIP + "','" + objSt.COUNTRY + "','" + objSt.EMAIL + "','" + objSt.MOBILE + "','" + objSt.STARTDATE + "','" + objSt.ENDDATE + "'," +
                        "'" + objSt.NEXT_STATEMENT_DATE + "','" + objSt.PAYMENT_DATE + "','" + objSt.STATEMENT_DATE + "','" + objSt.ACCOUNTNO + "','" + objSt.ACURN + "'," +
                        "'" + objSt.SBALANCE + "','" + objSt.ACURC + "','" + objSt.EBALANCE + "','" + objSt.AVAIL_CRD_LIMIT + "','" + objSt.AVAIL_CASH_LIMIT + "'," +
                        "'" + objSt.SUM_WITHDRAWAL + "','" + objSt.SUM_INTEREST + "','" + objSt.OVLFEE_AMOUNT + "','" + objSt.OVDFEE_AMOUNT + "','" + objSt.SUM_REVERSE + "'," +
                        "'" + objSt.SUM_CREDIT + "','" + objSt.SUM_OTHER + "','" + objSt.SUM_PURCHASE + "','" + objSt.MIN_AMOUNT_DUE + "','" + objSt.CASH_LIMIT + "'," +
                        "'" + objSt.CRD_LIMIT + "','" + objSt.STM_MSG + "','" + objSt.STATUS + "')";

                    reply = objProvider.RunQuery(sql);
                    //DataTable dtOperation = dsStatement.Tables["Operation"];
                    string trn_Date = string.Empty;
                    if (dtOperation != null && dtOperation.Columns.Contains("ACCOUNT"))
                    {
                        #region  ACCOUNT CHECK

                        if (dtOperation.Rows.Count > 0)
                        {
                            #region  dtOperation Row Check

                            DataRow[] dr = dtOperation.Select("STATEMENTNO='" + objSt.STATEMENTNO + "' AND ACCOUNT='" + objSt.ACCOUNTNO + "'");
                            if (dr.Length > 0)
                            {

                                #region  STATEMENTNO,ACCOUNT  >> FOUND
                                //  double feesnCharges = 0.00;


                                for (int l = 0; l < dr.Length; l++)
                                {
                                    #region setting properties values
                                    List<string> INTlist = new List<string>() { "INTEREST ON FEES & CHARGES","INTEREST ON INTEREST","INTEREST ON ATM TRANSACTION", "INTEREST ON POS TRANSACTION", "INTEREST ON CARD CHEQUE","CHARGE INTEREST FOR 0", "CHARGE INTEREST FOR 1", "CHARGE INTEREST FOR 2", "CHARGE INTEREST FOR 3", "CHARGE INTEREST FOR 4", "CHARGE INTEREST FOR 5", "CHARGE INTEREST FOR 6", "CHARGE INTEREST FOR 7", "CHARGE INTEREST FOR 8", "CHARGE INTEREST FOR 9", "CHARGE INTEREST FOR 10", "CHARGE INTEREST FOR 11", "CHARGE INTEREST FOR 0 OPERATIONS GROUP", "CHARGE INTEREST FOR 1 OPERATIONS GROUP", "CHARGE INTEREST FOR 2 OPERATIONS GROUP", "CHARGE INTEREST FOR 3 OPERATIONS GROUP", "CHARGE INTEREST FOR 4 OPERATIONS GROUP", "CHARGE INTEREST FOR 5 OPERATIONS GROUP", "CHARGE INTEREST FOR 6 OPERATIONS GROUP", "CHARGE INTEREST FOR 7 OPERATIONS GROUP", "INTEREST ON FUND TRANSFER", "INTEREST ON BALANCE TRANSFER", "INTEREST ON EMI", "INTEREST ON FT", "INTEREST ON BT", "INTEREST ON BANK POS TRANSACTION",
                                    "INTEREST ON BPOS TRANSACTION","CHARGE INTEREST FOR INTEREST OPERATIONS", "CHARGE INTEREST FOR POS OPERATIONS", "CHARGE INTEREST FOR ATM OPERATIONS", "LATE PAYMENT CHARGE FOR GROUP 1", "LATE PAYMENT CHARGE FOR GROUP 2", "LATE PAYMENT CHARGE FOR GROUP 3", "CHARGE OF A DEBT FOR CREDIT OVERDRAFTING" ,"INTEREST ON SERVICE FEE","INTEREST ON PREVIOUS BALANCE","EMI INTER"};
                                     
                                    if (INTlist.Contains(dr[l]["D"].ToString().ToUpper()) == false)
                                    {
                                        StatementDetails objSTD = new StatementDetails();
                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                        objSTD.PAN = objSt.PAN;

                                        if (dr[l].Table.Columns.Contains("ACCOUNT"))
                                            objSTD.ACCOUNTNO = dr[l]["ACCOUNT"].ToString();

                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;

                                        if (dr[l].Table.Columns.Contains("TD"))
                                            objSTD.TRNDATE = dr[l]["TD"].ToString();

                                        if (dr[l].Table.Columns.Contains("OD"))
                                            objSTD.POSTDATE = dr[l]["OD"].ToString();

                                        if (dr[l].Table.Columns.Contains("ACURN"))
                                            objSTD.ACURN = dr[l]["ACURN"].ToString();

                                        if (dr[l].Table.Columns.Contains("FR"))
                                            objSTD.FR = dr[l]["FR"].ToString().Replace("'", "''");

                                        if (dr[l].Table.Columns.Contains("DE"))
                                            objSTD.DE = dr[l]["DE"].ToString().Replace("'", "''");

                                        if (dr[l].Table.Columns.Contains("SERIALNO"))
                                            objSTD.SERIALNO = dr[l]["SERIALNO"].ToString();

                                        if (dr[l].Table.Columns.Contains("P"))   //Add new column from Operation 06.02.2017
                                        {
                                            if (dr[l]["P"].ToString() == "" || dr[l]["P"].ToString() == null)
                                            {
                                                objSTD.P = objSt.PAN;
                                            }

                                            else
                                            {

                                                objSTD.P = dr[l]["P"].ToString();
                                            }
                                        }

                                        if (dr[l].Table.Columns.Contains("DOCNO"))   //Add new column from Operation 06.02.2017
                                        {
                                            objSTD.DOCNO = dr[l]["DOCNO"].ToString();
                                        }

                                        if (dr[l].Table.Columns.Contains("NO"))   //Add new column from Operation 06.02.2017
                                        {
                                            objSTD.NO = dr[l]["NO"].ToString();
                                        }

                                        if (dr[l].Table.Columns.Contains("OCC"))
                                        {
                                            DataTable dtOcc = new DataTable();
                                            dtOcc = objProvider.ReturnData("select * from CURRENCYCODE", ref reply).Tables[0];// where Curr='BDT'
                                            DataRow[] drr = dtOcc.Select();
                                            string sp = string.Empty;
                                            string Sc = string.Empty;
                                            for (int x = 0; x <= 183; x++)
                                            {
                                                sp = dr[l]["OCC"].ToString();
                                                Sc = drr[x]["OCC"].ToString();
                                                if (dr[l]["OCC"].ToString() == drr[x]["OCC"].ToString())
                                                    objSTD.OC = drr[x]["Name"].ToString();
                                            }
                                        }
                                        else
                                            objSTD.OC = "";// dr[l]["OC"].ToString();



                                        if (dr[l].Table.Columns.Contains("AMOUNTSIGN"))
                                            objSTD.AMOUNTSIGN = dr[l]["AMOUNTSIGN"].ToString();

                                        if (dr[l].Table.Columns.Contains("ACURN"))
                                        {
                                            if (dr[l]["A"].ToString() == "" || dr[l]["A"].ToString() == null)
                                                objSTD.AMOUNT = "0.00";
                                            else
                                                objSTD.AMOUNT = dr[l]["A"].ToString();
                                        }
                                        else objSTD.AMOUNT = "0.00";

                                        if (dr[l].Table.Columns.Contains("OCC"))
                                        {
                                            if (dr[l]["OA"].ToString() == "" || dr[l]["OA"].ToString() == null)
                                                objSTD.ORGAMOUNT = "0.00";
                                            else
                                                objSTD.ORGAMOUNT = dr[l]["OA"].ToString();
                                        }
                                        else objSTD.ORGAMOUNT = "0.00";

                                        //Remmove Terminal Name when Fee and VAT Impose
                                        //Sum Charges amount with Fees & Charges. 

                                        #region  #region Monthly EMI ,TRANSFER TO EMI,EMI CANCELLED,EMI

                                        if ((dr[l]["D"].ToString().ToUpper().Contains("MONTHLY EMI")) || (dr[l]["D"].ToString().ToUpper().Contains("TRANSFER TO EMI")) || (dr[l]["D"].ToString().ToUpper().Contains("EMI CANCELLED")))
                                        {
                                            if (dr[l].Table.Columns.Contains("FR"))
                                            {
                                                if (dr[l]["FR"].ToString() == "" || dr[l]["FR"].ToString() == null)
                                                    if (dr[l].Table.Columns.Contains("TL"))
                                                    {
                                                        objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                                    }
                                                    else
                                                    {
                                                        objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                                    }
                                                else
                                                {
                                                    string data = dr[l]["FR"].ToString().Replace("'", "''");
                                                    bool contains = data.IndexOf("[VALUE NOT DEFINED]", StringComparison.OrdinalIgnoreCase) >= 0;
                                                    if (contains == true)
                                                    {
                                                        string[] list = data.Split(':');
                                                        objSTD.TRNDESC = list[0];
                                                    }
                                                    else
                                                    {
                                                        objSTD.TRNDESC = data.Replace("\n", "").Replace("\r", "");
                                                    }

                                                }
                                            }
                                            else
                                                // objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                                objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''"); // modify

                                        }

                                        #endregion

                                        #region CHEQUE TRANSACTION
                                        else if ((dr[l]["D"].ToString().ToUpper().Contains("CHEQUE TRANSACTION")) || (dr[l]["D"].ToString().ToUpper().Contains("CARD CHEQUE TRANSACTION")))
                                        {
                                            if (dr[l].Table.Columns.Contains("SERIALNO"))
                                            {
                                                if (dr[l]["SERIALNO"].ToString() == "" || dr[l]["SERIALNO"].ToString() == null)
                                                {
                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + " [CHQ NO:" + "]";
                                                }
                                                else
                                                {
                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " [CHQ NO:" + dr[l]["SERIALNO"].ToString().Replace("'", "") + "]";
                                                }
                                            }
                                            else
                                            {
                                                objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + " [CHQ NO:" + "]";
                                            }
                                        }

                                        #endregion

                                        #region Rest of Txn
                                        else
                                        {
                                            if (dr[l].Table.Columns.Contains("TL"))
                                            {
                                                if (dr[l]["FR"].ToString().ToUpper().Contains("A 10") || dr[l]["FR"].ToString().ToUpper().Contains("A 64") || dr[l]["FR"].ToString().ToUpper().Contains("P 14")  || dr[l]["FR"].ToString().ToUpper().Contains("P 32") || dr[l]["FR"].ToString().ToUpper().Contains("P 33") || dr[l]["FR"].ToString().ToUpper().Contains("F 29") || dr[l]["FR"].ToString().ToUpper().Contains("P 13"))
                                                {
                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                                    // objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''") + " " + "(" + objSTD.OC + " " + dr[l]["OA"].ToString() + ")";

                                                }
                                                else
                                                {
                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                                }
                                              /*  if (dr[l]["D"].ToString().ToUpper().Contains("PURCHASE"))
                                                {
                                                    if (dr[l]["D"].ToString().Trim().Length > 8)
                                                    {

                                                        objSTD.TRNDESC = (dr[l]["D"].ToString().ToUpper().Replace("PURCHASE", "")).Trim() + " " + dr[l]["TL"].ToString().Replace("'", "''");

                                                    }
                                                    else
                                                    {

                                                        objSTD.TRNDESC = (dr[l]["D"].ToString().ToUpper().Replace("PURCHASE", "")).Trim() + dr[l]["TL"].ToString().Replace("'", "''");
                                                    }
                                                }*/
                                                if (dr[l]["D"].ToString().ToUpper().Contains("PURCHASE"))
                                                {

                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");


                                                } 


                                            }

                                            else
                                            {
                                                objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                            }

                                        }

                                        #endregion

                                        #region PAYMENT CASH DEPOSIT

                                        if ((objSTD.TRNDESC.ToUpper().Contains("PAYMENT RECEIVED (THANK YOU)")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT RECEIVED, THANK YOU.")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT RECEIVED, THANK YOU")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT RECEIVED [AUTO DEBIT]")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT RECEIVED [CASH]")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT CASH DEPOSIT")) || (objSTD.TRNDESC.ToUpper().Contains("CREDIT CASH DEPOSIT")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT THROUGH BRANCHES  (CASH)")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT BY CHEQUE (MAIL)")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT THROUGH AUTO DEBIT")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT THROUGH CHEQUE")) || (objSTD.TRNDESC.ToUpper().Contains("PAYMENT THROUGH FC")) || (objSTD.TRNDESC.ToUpper().Contains("VISA PAYMENT")) || (objSTD.TRNDESC.ToUpper().Contains("MC PAYMENT")))
                                        {
                                            objSTD.TRNDESC = "PAYMENT RECEIVED (THANK YOU)";
                                            //objSTD.TRNDATE = dr[l]["OD"].ToString();
                                        }

                                        #endregion

                                        #region APPROVAL
                                        if (dr[l].Table.Columns.Contains("APPROVAL"))
                                        {
                                            objSTD.APPROVAL = dr[l]["APPROVAL"].ToString().Replace("'", "''");

                                            if (dr[l]["APPROVAL"].ToString() != "" && objSTD.TRNDATE == "")
                                            {
                                                objSTD.TRNDATE = dr[l]["OD"].ToString();
                                            }
                                        }
                                        #endregion
                                        #region CASH ADVANCE

                                        try
                                        {
                                            if ((dr[l]["D"].ToString().ToUpper().Trim() == ("CASH ADVANCE")))
                                            {

                                                objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                            }
                                        }

                                        catch (Exception ex)
                                        {
                                            objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                        }

                                        #endregion

                                        //objSTD.AMOUNTSIGN = dr[l]["AMOUNTSIGN"].ToString();
                                        if (dr[l].Table.Columns.Contains("TD"))
                                            objSTD.TRNDATE = dr[l]["TD"].ToString();

                                        if (!dr[l].Table.Columns.Contains("P"))   //Add new column from Operation 06.02.2017
                                        {
                                            objSTD.P = objSt.PAN;
                                        }

                                        sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,OC,ORGAMOUNT,AMOUNTSIGN,APPROVAL,FR,SERIALNO,DE,P,DOCNO,NO)" +
                                            " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                            "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.OC + "','" + objSTD.ORGAMOUNT + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.APPROVAL + "','" + objSTD.FR + "','" + objSTD.SERIALNO + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";

                                        reply = objProvider.RunQuery(sql);
                                        if (!reply.Contains("Success"))
                                            errMsg = reply;
                                    }

                                    #endregion
                                }


                                if (objSt.SUM_INTEREST != "0.00")
                                {
                                    #region SUM_INTEREST

                                    StatementDetails objSTD = new StatementDetails();
                                    objSTD.STATEMENTID = objSt.STATEMENTID;
                                    objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                    objSTD.IDCLIENT = objSt.IDCLIENT;
                                    objSTD.PAN = objSt.PAN;
                                    objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                    objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                    objSTD.ACURN = objSt.ACURN;
                                    objSTD.TRNDESC = "INTEREST CHARGES";
                                    //objSTD.TRNDESC = "Profit Charges";
                                    objSTD.AMOUNT = "-" + objSt.SUM_INTEREST;//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                    objSTD.TRNDATE = trn_Date;
                                    objSTD.POSTDATE = trn_Date;
                                    objSTD.P = objSt.PAN;

                                    sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,FR,SERIALNO,DE,P,DOCNO,NO)" +
                                            " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                            "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.FR + "','" + objSTD.SERIALNO + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";

                                    reply = objProvider.RunQuery(sql);
                                    if (!reply.Contains("Success"))
                                        errMsg = reply;



                                    #endregion
                                }
                                else
                                {
                                    DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW", ref reply);

                                    if (dsAcI != null)
                                    {
                                        if (dsAcI.Tables.Count > 0)
                                        {
                                            if (dsAcI.Tables[0].Rows.Count > 0)
                                            {
                                                DataTable dtAcI = dsAcI.Tables[0]; ;
                                                for (int x = 0; x < dtAcI.Rows.Count; x++)
                                                {
                                                    StatementDetails objSTD = new StatementDetails();

                                                    objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                    if (objSTD.CONTRACTNO == dtAcI.Rows[x][1].ToString())
                                                    {
                                                        if (dtAcI.Rows[x][0].ToString() != "0.00")
                                                        {
                                                            objSTD.STATEMENTID = objSt.STATEMENTID;
                                                            objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                            objSTD.IDCLIENT = objSt.IDCLIENT;
                                                            objSTD.PAN = objSt.PAN;
                                                            objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                            objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                            objSTD.ACURN = objSt.ACURN;
                                                            objSTD.TRNDESC = "INTEREST CHARGES";
                                                            objSTD.AMOUNT = "-" + dtAcI.Rows[x][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                            objSTD.TRNDATE = trn_Date;
                                                            objSTD.POSTDATE = trn_Date;
                                                            objSTD.P = objSt.PAN;

                                                            sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,FR,SERIALNO,DE,P,DOCNO,NO)" +
                                                                    " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                                    "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.FR + "','" + objSTD.SERIALNO + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";


                                                            reply = objProvider.RunQuery(sql);
                                                            if (!reply.Contains("Success"))
                                                                errMsg = reply;

                                                            decimal tempIntAmtI = 0;
                                                            decimal tempIntAmt = 0;
                                                            decimal tempTotalIntAmt = 0;
                                                            string st = string.Empty;

                                                            DataTable dt = new DataTable();
                                                            dt = objProvider.ReturnData("select AMOUNT from STATEMENT_DETAILS WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES' ", ref reply).Tables[0];
                                                            //tempIntAmtI = Convert.ToInt32(dt.Rows[0][0])*(-1);
                                                            tempIntAmtI = Convert.ToDecimal(dt.Rows[0][0]) * (-1);
                                                            st = dtAcI.Rows[x][0].ToString();
                                                            tempIntAmt = Convert.ToDecimal(st);
                                                            tempTotalIntAmt = tempIntAmtI + tempIntAmt;

                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion

                            }
                            else
                            {

                                //string trn_Date = string.Empty;


                                //New View add
                                // DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW", ref reply);
                                DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW", ref reply);
                                if (dsAcI != null)
                                {
                                    if (dsAcI.Tables.Count > 0)
                                    {
                                        if (dsAcI.Tables[0].Rows.Count > 0)
                                        {
                                            DataTable dtAcI = dsAcI.Tables[0]; ;
                                            for (int x = 0; x < dtAcI.Rows.Count; x++)
                                            {
                                                StatementDetails objSTD = new StatementDetails();

                                                objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                if (objSTD.CONTRACTNO == dtAcI.Rows[x][1].ToString())
                                                {
                                                    if (dtAcI.Rows[x][0].ToString() != "0.00")
                                                    {
                                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                                        objSTD.PAN = objSt.PAN;
                                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                        objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                        objSTD.ACURN = objSt.ACURN;
                                                        objSTD.TRNDESC = "INTEREST CHARGES";
                                                        objSTD.AMOUNT = "-" + dtAcI.Rows[x][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                        objSTD.TRNDATE = trn_Date;
                                                        objSTD.POSTDATE = trn_Date;
                                                        objSTD.P = objSt.PAN;

                                                        sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,FR,SERIALNO,DE,P,DOCNO,NO)" +
                                                                " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                                "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.FR + "','" + objSTD.SERIALNO + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";


                                                        reply = objProvider.RunQuery(sql);
                                                        if (!reply.Contains("Success"))
                                                            errMsg = reply;

                                                        decimal tempIntAmtI = 0;
                                                        decimal tempIntAmt = 0;
                                                        decimal tempTotalIntAmt = 0;
                                                        string st = string.Empty;

                                                        DataTable dt = new DataTable();
                                                        dt = objProvider.ReturnData("select AMOUNT from STATEMENT_DETAILS WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES' ", ref reply).Tables[0];
                                                        //tempIntAmtI = Convert.ToInt32(dt.Rows[0][0])*(-1);
                                                        tempIntAmtI = Convert.ToDecimal(dt.Rows[0][0]) * (-1);
                                                        st = dtAcI.Rows[x][0].ToString();
                                                        tempIntAmt = Convert.ToDecimal(st);
                                                        tempTotalIntAmt = tempIntAmtI + tempIntAmt;

                                                    }
                                                }
                                            }
                                        }
                                    }
                                }



                            }

                            #endregion
                        }


                    }
                    else
                    {
                        if (dtOperation.Rows.Count > 0)
                        {

                            DataRow[] dr = dtOperation.Select("STATEMENTNO='" + objSt.STATEMENTNO + "'");
                            if (dr.Length > 0)
                            {

                                // string trn_Date = string.Empty;
                                //New View add
                                // DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW", ref reply);
                                DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW", ref reply);

                                if (dsAcI != null)
                                {
                                    if (dsAcI.Tables.Count > 0)
                                    {
                                        if (dsAcI.Tables[0].Rows.Count > 0)
                                        {
                                            DataTable dtAcI = dsAcI.Tables[0]; ;
                                            for (int x = 0; x < dtAcI.Rows.Count; x++)
                                            {
                                                StatementDetails objSTD = new StatementDetails();

                                                objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                if (objSTD.CONTRACTNO == dtAcI.Rows[x][1].ToString())
                                                {
                                                    if (dtAcI.Rows[x][0].ToString() != "0.00")
                                                    {
                                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                                        objSTD.PAN = objSt.PAN;
                                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                        objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                        objSTD.ACURN = objSt.ACURN;
                                                        objSTD.TRNDESC = "INTEREST CHARGES";
                                                        objSTD.AMOUNT = "-" + dtAcI.Rows[x][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                        objSTD.TRNDATE = trn_Date;
                                                        objSTD.POSTDATE = trn_Date;
                                                        objSTD.P = objSt.PAN;

                                                        sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,FR,SERIALNO,DE,P,DOCNO,NO)" +
                                                                " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                                "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.FR + "','" + objSTD.SERIALNO + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";


                                                        reply = objProvider.RunQuery(sql);
                                                        if (!reply.Contains("Success"))
                                                            errMsg = reply;

                                                        decimal tempIntAmtI = 0;
                                                        decimal tempIntAmt = 0;
                                                        decimal tempTotalIntAmt = 0;
                                                        string st = string.Empty;

                                                        DataTable dt = new DataTable();
                                                        dt = objProvider.ReturnData("select AMOUNT from STATEMENT_DETAILS WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES' ", ref reply).Tables[0];
                                                        //tempIntAmtI = Convert.ToInt32(dt.Rows[0][0])*(-1);
                                                        tempIntAmtI = Convert.ToDecimal(dt.Rows[0][0]) * (-1);
                                                        st = dtAcI.Rows[x][0].ToString();
                                                        tempIntAmt = Convert.ToDecimal(st);
                                                        tempTotalIntAmt = tempIntAmtI + tempIntAmt;

                                                    }
                                                }
                                            }
                                        }
                                    }
                                }


                            }
                        }



                    }

                        #endregion

                }
                catch (Exception ex)
                {
                    errMsg = "Error: " + ex.Message;
                }
            }
        }
            #endregion USD

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
