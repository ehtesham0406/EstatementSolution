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
using QCash.EStatement.DBLPrepaid.Reports;
using Infragistics.Documents.PDF;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Security;
using Common;
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
        private string _Mail = string.Empty;
        private string StmDate = string.Empty;
        private string stmMessage = string.Empty;

        string vPAN = string.Empty;

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
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh24:mm:ss") + " : " + reply });
                    }
                    else if (reply == "Success")
                    {
                        MsgLogWriter objLW = new MsgLogWriter();
                        objLW.logTrace(_LogPath, "EStatement.log", "Successfully archive previous data !!!");
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh24:mm:ss") + " : " + "Successfully archive previous data !!!" });

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
                string StmDate = getNumberFormat1(dtpStmDate.Value.ToString());

               // string Month = curdate.Split('-')[1].ToString();
              //  string Year = curdate.Split('-')[2].ToString();
              //  string StmDate = "01" + '-' + Month + '-' + Year;

                MsgLogWriter objLW = new MsgLogWriter();

                EStatementList objESList = EStatementManager.Instance().GetAllEStatements(_fiid, StmDate, "1", ref reply);
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
                                   // if (objESList[i].MAILADDRESS != "")
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

                                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + "Sending EStatement to " + mail.To.ToString() }); ;
                                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Sending EStatement " + mail.To.ToString());

                                            SmtpServer.Send(mail);

                                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + "mail Send to " + mail.To.ToString() }); ;
                                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : mail Send to " + mail.To.ToString());


                                            objESList[i].STATUS = "0"; // Estatement Generated and mail sent.
                                            EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                            count++;
                                        }
                                        catch (Exception ex)
                                        {
                                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message);

                                            objESList[i].STATUS = "2";  // Estatement Generated and mail sent but no acknowledged received from mail server.
                                            EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                        }
                                    }
                                    else
                                    {
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + "Invalid or No Mail Address Found to send the Estatement " + objESList[i].FILE_LOCATION + " " + " PAN : " + objESList[i].PAN_NUMBER + " and Email : " + email }); ;
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Invalid or No Mail Address Found to send the Estatement " + objESList[i].FILE_LOCATION + " " + " PAN : " + objESList[i].PAN_NUMBER + " and Email : " + email);

                                        objESList[i].STATUS = "8";  // no mail address
                                        EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                    }
                                }
                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " e-statements have been mailed out of " + objESList.Count + "." });
                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " e-statements have been mailed out of " + objESList.Count + ".");
                            }
                        }
                    }
                }
                else
                {
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : There is no Estatement has generate on that statement date." });
                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : There is no Estatement has generate on that statement date.");

                }
            }

            catch (Exception ex)
            {
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh24:mm:ss") + " : Error: " + ex.Message });
            }
        }

        void btnGenerate_Click(object sender, EventArgs e)
        {
            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
            string reply = string.Empty;
            MsgLogWriter objLW = new MsgLogWriter();

            DataTable dtCardbdt = new DataTable();
            dtCardbdt = objProvider.ReturnData("select * from STATEMENT", ref reply).Tables[0];// where Curr='BDT'

            if (dtCardbdt.Rows.Count > 0)
            {
                txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Processing Estatement." });//Processing Estatement BDT
                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Processing Estatement.");//Processing Estatement BDT.

                //Process pdf for BDT
                ProcessStatementBDT(dtCardbdt);
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
            ds = objProvider.ReturnData("select * from statement_DUAL", ref reply);
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

                        txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement.");

                        for (int j = 0; j < dtCards.Rows.Count; j++)//dtCards.Rows.Count
                        {
                            //if (dtCards.Rows[j]["EMAIL"].ToString().Trim() != "")
                            //{
                               // if (IsValid(dtCards.Rows[j]["EMAIL"].ToString().Trim()))
                               // {
                                    try
                                    {
                                        pdfCount = pdfCount + 1;
                                        stmdt = new DataTable();
                                        stmdt = objProvider.ReturnData("select * from Statement_DUAL where main_card='" + dtCards.Rows[j]["main_card"].ToString() + "' ORDER BY [AutoID]", ref reply).Tables[0];
                                        if ((dtCards.Rows[j]["main_card"].ToString()) != vPAN)
                                        {

                                            vPAN = dtCards.Rows[j]["main_card"].ToString();
                                            if (stmdt.Rows.Count > 0)
                                            {

                                                EStatement objst = new EStatement();
                                                objst.SetDataSource(stmdt);
                                               // fileName = dtCards.Rows[j]["idclient"].ToString() + "_" + DateTime.Now.ToShortDateString().Replace('/', '-') + "_" + dtCards.Rows[j]["main_card"].ToString().Substring(0, 6) + '_' + pdfCount + ".pdf";
                                                fileName = dtCards.Rows[j]["idclient"].ToString() + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + "_" + dtCards.Rows[j]["main_card"].ToString().Substring(0, 6) + '_' + pdfCount + ".pdf";  
                                                System.IO.Stream st = objst.ExportToStream(ExportFormatType.PortableDocFormat);

                                                PdfSharp.Pdf.PdfDocument document = PdfReader.Open(st);

                                                PdfSecuritySettings securitySettings = document.SecuritySettings;

                                                string card_no = dtCards.Rows[j]["main_card"].ToString();
                                                securitySettings.UserPassword = dtCards.Rows[j]["main_card"].ToString().Substring(card_no.Length - 4, 4);
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
                                                objEst.STMDATE = stmdt.Rows[0]["STATEMENT_DATE"].ToString();
                                                objEst.IDCLIENT = stmdt.Rows[0]["IDCLIENT"].ToString();
                                                StmDate = stmdt.Rows[0]["STATEMENT_DATE"].ToString();

                                                string[] drdate = stmdt.Rows[0]["STATEMENT_DATE"].ToString().Split('/', '-');
                                                
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
                                                objEst.PAN_NUMBER = dtCards.Rows[j]["main_card"].ToString();

                                                if (stmdt.Rows.Count > 0)
                                                    objEst.MAILADDRESS = stmdt.Rows[0]["EMAIL"].ToString();
                                                else
                                                    objEst.MAILADDRESS = null;

                                                objEst.FILE_LOCATION = filePath + "\\" + fileName;
                                                objEst.MAILSUBJECT = txtEmailSubject.Text.Replace("'", "''");
                                                objEst.STATUS = "1";

                                                reply = EStatementManager.Instance().AddEStatement(objEst, ref reply);

                                                if (reply == "Success")
                                                {
                                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4) });
                                                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4));
                                                    count++;
                                                }
                                                else
                                                {
                                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Message " + reply });
                                                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + "Message " + reply);
                                                }
                                                if (count % 10 == 0)
                                                {
                                                    objst.Dispose();
                                                    GC.Collect();
                                                    Thread.Sleep(1000);
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + ex.Message);
                                    }
                               // }
                                //else
                                //{
                                   // txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["main_card"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["main_card"].ToString().Substring(12, 4) });
                                   // objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["main_card"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["main_card"].ToString().Substring(12, 4));

                                //}
                             
                           // } 
                        }
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " Estatement has processed out of " + dtCards.Rows.Count + "." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " Estatement has processed out of" + dtCards.Rows.Count + ".");
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
            ds = objProvider.ReturnData("select * from statement_DUAL", ref reply);

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

                        txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + "Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement.");

                        for (int j = 0; j < dtCards.Rows.Count; j++)//dtCards.Rows.Count
                        {
                           
                               
                                    try
                                    {
                                        stmdt = new DataTable();
                                        stmdt = objProvider.ReturnData("select * from statement_DUAL where IDCLIENT='" + dtCards.Rows[j]["IDCLIENT"].ToString() + "'", ref reply).Tables[0];
                                        if (stmdt.Rows.Count > 0)
                                        {
                                            EStatement objst = new EStatement();
                                            EStatement objstPlatinum = new EStatement();

                                            if (dtCards.Rows[j]["EMAIL"].ToString().Trim() == "rtte")
                                            {
                                                objst.SetDataSource(stmdt);
                                            }
                                            else
                                            {
                                                objstPlatinum.SetDataSource(stmdt);
                                            }

                                            fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["main_card"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["main_card"].ToString().Substring(12, 4) + ".pdf";
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
                                            objEst.PAN_NUMBER = dtCards.Rows[j]["main_card"].ToString();

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
                                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4) });
                                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4));
                                                count++;
                                            }
                                            else
                                            {
                                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Message " + reply });
                                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + "Message " + reply);
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
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + ex.Message);
                                    }
                               
                           
                        }
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " Estatement has processed out of " + dtCards.Rows.Count + "." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " Estatement has processed" + dtCards.Rows.Count + ".");
                    }
                }
            }
        }
        
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
                if (dia[fcount].FullName.Contains("DBL"))
                {
                    _bankName = "DBL";
                    _bankCode = "3";
                    _XMLSourcePath = dia[fcount].FullName;
                    //
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }
                else
                {
                    MsgLogWriter objLW = new MsgLogWriter();
                    objLW.logTrace(_LogPath, "EStatement.log", "Not an XML data !!!");
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh24:mm:ss") + " : " + "Not an XML data !!!" });
                }

                Bank_Code = _bankName;

            }
            #endregion

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

                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + fi.Length.ToString() + " files found to process.." });
                objLW.logTrace(_LogPath, "EStatement.log", " : Total " + fi.Length.ToString() + " files found to process..");
                
                for (int f = 0; f < fi.Length; f++)
                {
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + fi[f].Name + " on process.." });
                    objLW.logTrace(_LogPath, "EStatement.log", " : " + fi[f].Name + " on process..");

                    DataSet dsXML = getDataFromXML(fi[f].FullName);
                    EStatementManager.Instance().ArchiveEStatementCardData(ref reply);

                    #region Operation On Data
                    if (dsXML != null)
                    {
                       
                        if (dsXML.Tables.Count > 0)
                        {
                            ConStr = new ConnectionStringBuilder(1);
                            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);

                            string sql = string.Empty;

                            for (int i = 0; i < dsXML.Tables.Count; i++)
                            {
                                if (dsXML.Tables[i].TableName == "Statement")
                                {
                                    GetCardHolderPersonalInfo(dsXML.Tables[i], BankName, ref reply);
                                    GetCardHolderCardInfo(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : CardHolder Personal Info data Saved from XML. " + reply });
                                    objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Personal Info data Saved from XML. " + reply);
                                }
                                else if (dsXML.Tables[i].TableName == "Operation")
                                {
                                    reply = GetCardHolderTransactionInfo(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : CardHolder Transaction Info data Saved from XML. " + reply });
                                    objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Transaction Info data Saved from XML. " + reply);
                                }

                                //else if (dsXML.Tables[i].TableName == "Card")
                                //{
                                //    reply = GetCardHolderCardInfo(dsXML.Tables[i]);
                                //    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : CardHolder Card Info data Saved from XML. " + reply });
                                //    objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Card Info data Saved from XML. " + reply);
                                //}


                            }                           
                        }
                        DataSet dsCard = objProvider.ReturnData("select * from card", ref reply);
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + dsCard.Tables[0].Rows.Count.ToString() + " Card record has been found to process.." });
                        objLW.logTrace(_LogPath, "EStatement.log", " : Total " + dsCard.Tables[0].Rows.Count.ToString() + " Card record has been found to process..");
                        




                   }
                    #endregion

                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + fi[f].Name + " process complete.." });
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
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Source Directory moving Error. Error: " + ex.Message });
                        objLW = new MsgLogWriter();
                        objLW.logTrace(_LogPath, "EStatement.log", "Source Directory moving Error. " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                _reply = ex.StackTrace;
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + ex.Message });
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
          //  string STATEMENT_DATE = DateTime.Now.ToShortDateString().Replace('/', '-');

            string StmDate = getNumberFormat1(dtpStmDate.Value.ToString());

           // string Month = curdate.Split('-')[1].ToString();
           // string Year = curdate.Split('-')[2].ToString();
           // string StmDate = "01" +'-'+ Month +'-'+Year;


            try
            {
                //Clear Previous Data
                objProvider.RunQuery("Delete from " + dtStatement.TableName);


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
                        if (dtStatement.Columns[j].ColumnName == "Account")
                        {
                            objSt.ACCOUNT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Company")
                        {
                            objSt.COMPANY = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "StartDate")
                        {
                            objSt.STARTDATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                            objSt.STARTDATE = getNumberFormat(objSt.STARTDATE);
                        }

                        if (dtStatement.Columns[j].ColumnName == "Telephone")
                        {
                            objSt.TELEPHONE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "ClientLat")
                        {
                            objSt.CLIENTLAT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "PersonalCode")
                        {
                            objSt.PERSONALCODE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Mobile")
                        {
                            objSt.MOBILE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "StreetAddress")
                        {
                            objSt.STREETADDRESS = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "TotalIn")
                        {
                            objSt.TOTALIN = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "TotalOut")
                        {
                            objSt.TOTALOUT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Address")
                        {
                            objSt.ADDRESS = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Country")
                        {
                            objSt.COUNTRY = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "AccountTypeName")
                        {
                            objSt.ACCOUNTTYPENAME = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Overdraft")
                        {
                            objSt.OVERDRAFT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "CURRFULLNAME")
                        {
                            objSt.CURRFULLNAME = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "StatementType")
                        {
                            objSt.STATEMENTTYPE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "SendType")
                        {
                            objSt.SENDTYPE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "EndDate")
                        {
                            objSt.ENDDATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                            objSt.ENDDATE = getNumberFormat(objSt.ENDDATE);
                        }
                        if (dtStatement.Columns[j].ColumnName == "Fax")
                        {
                            objSt.FAX = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Client")
                        {
                            objSt.CLIENT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "IdClient")
                        {
                            objSt.IDCLIENT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Currency")
                        {
                            objSt.CURRENCY = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "CurrencyName")
                        {
                            objSt.CURRENCYNAME = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "StartBalance")
                        {
                            objSt.STARTBALANCE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Available")
                        {
                            objSt.AVAILABLE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Sex")
                        {
                            objSt.SEX = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Pager")
                        {
                            objSt.PAGER = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "EmployeeNo")
                        {
                            objSt.EMPLOYEENO = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "JobTitle")
                        {
                            objSt.JOBTITLE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Email")
                        {
                            objSt.EMAIL = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "EndBalance")
                        {
                            objSt.ENDBALANCE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "DebitReserve")
                        {
                            objSt.DEBITRESERVE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "Title")
                        {
                            objSt.TITLE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        if (dtStatement.Columns[j].ColumnName == "PromotionalText")
                        {

                         //   string mAIN_cARD = dtStatement.Rows[k][j].ToString().Substring(0, dtStatement.Rows[k][j].ToString().Length-1).Replace("'", "");

                           // objSt.MAIN_CARD = dtStatement.Rows[k][j].ToString().Substring(0, dtStatement.Rows[k][j].ToString().Length - 1).Replace("'", "");
                           
                            
                            //objSt.MAIN_CARD = mAIN_cARD.Substring(0, 6) + "******" + mAIN_cARD.Substring(mAIN_cARD.Length - 4, 4);
                            objSt.PROMOTIONALTEXT = dtStatement.Rows[k]["PromotionalText"].ToString().Replace("'", "");
                            string value = objSt.PROMOTIONALTEXT;
                            string[] lines = value.Split(new char[] { '|' });

                            if (lines.Length <= 1)
                            {
                              objSt.MAIN_CARD = dtStatement.Rows[k][j].ToString().Substring(0, dtStatement.Rows[k][j].ToString().Length - 1).Replace("'", "");
                           
                            }

                            else
                            {

                                    if (!string.IsNullOrEmpty(lines[0]))
                                    {
                                        objSt.MAIN_CARD = lines[0];
                                    }
                                    else
                                    {
                                        objSt.MAIN_CARD = null;
                                    }


                                    if (!string.IsNullOrEmpty(lines[1]))
                                    {
                                        string cn = lines[1];
                                        //objSt.COMPANY = cn.Substring(0, cn.ToString().Length - 1).Replace("'", "");
                                        objSt.COMPANY = cn.Replace("'", "");
                                    }
                                    else
                                    {
                                        objSt.COMPANY = null;
                                    }
                                }

                           


                        }




                    }
                        #endregion

                    sql = "Insert into STATEMENT(BANK_CODE,STATEMENTNO,STATEMENT_DATE,ACCOUNT,COMPANY,STARTDATE,TELEPHONE,CLIENTLAT,PERSONALCODE,MOBILE,STREETADDRESS,TOTALIN,TOTALOUT,ADDRESS,COUNTRY," +
                            "ACCOUNTTYPENAME,OVERDRAFT,CURRFULLNAME,STATEMENTTYPE,SENDTYPE,ENDDATE,FAX,CLIENT,IDCLIENT,CURRENCY,CURRENCYNAME,STARTBALANCE," +
                                "AVAILABLE,SEX,PAGER,EMPLOYEENO,JOBTITLE,EMAIL,ENDBALANCE,DEBITRESERVE,TITLE,MAIN_CARD)" +
                          "values('" + objSt.BANK_CODE + "','" + objSt.STATEMENTNO + "','" + StmDate + "','" + objSt.ACCOUNT + "','" + objSt.COMPANY + "','" + objSt.STARTDATE + "','" + objSt.TELEPHONE + "','" + objSt.CLIENTLAT + "'," +
                          "'" + objSt.PERSONALCODE + "','" + objSt.MOBILE + "','" + objSt.STREETADDRESS + "','" + objSt.TOTALIN + "','" + objSt.TOTALOUT + "','" + objSt.ADDRESS + "','" + objSt.COUNTRY + "','" + objSt.ACCOUNTTYPENAME + "'," +
                          "'" + objSt.OVERDRAFT + "','" + objSt.CURRFULLNAME + "','" + objSt.STATEMENTTYPE + "','" + objSt.SENDTYPE + "','" + objSt.ENDDATE + "','" + objSt.FAX + "'," +
                          "'" + objSt.CLIENT + "','" + objSt.IDCLIENT + "','" + objSt.CURRENCY + "','" + objSt.CURRENCYNAME + "','" + objSt.STARTBALANCE + "','" + objSt.AVAILABLE + "','" + objSt.SEX + "','" + objSt.PAGER + "','" + objSt.EMPLOYEENO + "','" + objSt.JOBTITLE + "','" + objSt.EMAIL + "','" + objSt.ENDBALANCE + "','" + objSt.DEBITRESERVE + "','" + objSt.TITLE + "','" + objSt.MAIN_CARD+ "')";

                    reply = objProvider.RunQuery(sql);
                    if (!reply.Contains("Success"))
                    errMsg = reply;
                }
                return objStList;
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                errMsg = "Error: " + ex.StackTrace;
                return objStList;
            }

        }

        private string GetCardHolderCardInfo(DataTable dtCard)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            Card objCard = null;
            CardList objCardList = new CardList();
            string tableName = "Card";

            try
            {
                //Clear Previous Data
                objProvider.RunQuery("Delete from " + tableName);

                for (int k = 0; k < dtCard.Rows.Count; k++)
                {
                    objCard = new Card();

                    for (int j = 0; j < dtCard.Columns.Count; j++)
                    {
                        #region setting properties values

                        switch (dtCard.Columns[j].ColumnName)
                        {
                            case "StatementNo":
                                objCard.STATEMENTNO = dtCard.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "PromotionalText":
                                //objCard.PAN = dtCard.Rows[k][j].ToString();
                                objCard.PAN = dtCard.Rows[k][j].ToString().Substring(0, dtCard.Rows[k][j].ToString().Length - 1).Replace("'", "");
                           
                                break;
                            //case "MBR":
                                //objCard.MBR = dtCard.Rows[k][j].ToString();
                                //break;
                            case "Client":
                                objCard.CLIENTNAME = dtCard.Rows[k][j].ToString().Replace("'", "");
                                break;

                        }

                        #endregion
                    }
                    objCardList.Add(objCard);

                    sql = "Insert into Card(STATEMENTNO,PAN,MBR,CLIENTNAME)" +
                        " Values('" + objCard.STATEMENTNO + "','" + objCard.PAN + "','" + objCard.MBR + "','" + objCard.CLIENTNAME + "')";

                    reply = objProvider.RunQuery(sql);
                    if (!reply.Contains("Success"))
                        return reply;
                }
                return reply;
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return "Error: " + ex.StackTrace;
            }
        }
        
        private string GetCardHolderTransactionInfo(DataTable dtOperation)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            Operation objOp = null;

            try
            {
                //Clear Previous Data
                objProvider.RunQuery("Delete from " + dtOperation.TableName);

                for (int k = 0; k < dtOperation.Rows.Count; k++)
                {
                    objOp = new Operation();

                    for (int j = 0; j < dtOperation.Columns.Count; j++)
                    {
                        #region setting properties values

                        if (dtOperation.Columns.Contains("StatementNo"))
                        {
                            objOp.STATEMENTNO = dtOperation.Rows[k]["StatementNo"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("O"))
                        {
                            objOp.OpID = dtOperation.Rows[k]["O"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("OD"))                  
                        {
                            objOp.OpDate = dtOperation.Rows[k]["OD"].ToString().Replace("'", "");
                            if (dtOperation.Rows[k]["OD"].ToString() == "" || dtOperation.Rows[k]["OD"].ToString() == null)
                            {
                                objOp.OpDate = "";
                            }
                            else
                            {
                                objOp.OpDate = getNumberFormat(objOp.OpDate);                               
                            }
                        }
                        if (dtOperation.Columns.Contains("TD"))
                        {
                            objOp.TD = dtOperation.Rows[k]["TD"].ToString().Replace("'", "");
                            if (dtOperation.Rows[k]["TD"].ToString() == "" || dtOperation.Rows[k]["TD"].ToString() == null)
                            {
                                objOp.TD = "";
                            }
                            else
                            {
                                objOp.TD = getNumberFormat(objOp.TD);                                
                            }
                        }
                        if (dtOperation.Columns.Contains("A"))
                        {
                            if (dtOperation.Rows[k]["A"].ToString() == "" || dtOperation.Rows[k]["A"].ToString() == null)
                                objOp.Amount = "0.00";
                            else
                                objOp.Amount = dtOperation.Rows[k]["A"].ToString().Replace("'", "");
                        }

                        if (dtOperation.Columns.Contains("ACURC"))
                        {
                            objOp.ACURCode = dtOperation.Rows[k]["ACURC"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("ACURN"))
                        {
                            objOp.ACURName = dtOperation.Rows[k]["ACURN"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("D"))
                        {
                            objOp.D = dtOperation.Rows[k]["D"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("DE"))
                        {
                            objOp.DE = dtOperation.Rows[k]["DE"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("CF"))
                        {
                            objOp.CF = dtOperation.Rows[k]["CF"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("DOCNO"))
                        {
                            objOp.DOCNO = dtOperation.Rows[k]["DOCNO"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("NO"))
                        {
                            objOp.NO = dtOperation.Rows[k]["NO"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("ACCOUNT"))
                        {
                            objOp.ACCOUNT = dtOperation.Rows[k]["ACCOUNT"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("ACC"))
                        {
                            objOp.ACC = dtOperation.Rows[k]["ACC"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("FR"))
                        {
                            objOp.FR = dtOperation.Rows[k]["FR"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("APPROVAL"))
                        {
                            objOp.APPROVAL = dtOperation.Rows[k]["APPROVAL"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("MN"))
                        {
                            objOp.MN = dtOperation.Rows[k]["MN"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("S"))
                        {
                            objOp.S = dtOperation.Rows[k]["S"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("TERMN"))
                        {
                            objOp.TERMN = dtOperation.Rows[k]["TERMN"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("TL"))
                        {
                            objOp.TL = dtOperation.Rows[k]["TL"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("P"))
                        {
                            objOp.P = dtOperation.Rows[k]["P"].ToString().Replace("'", "");
                        }
                       // if (dtOperation.Columns.Contains("SERIALNO"))
                        ////{
                        //    objOp.SERIALNO = dtOperation.Rows[k]["SERIALNO"].ToString().Replace("'", "");
                       // }

                        if (dtOperation.Columns.Contains("OCC"))
                        {
                            objOp.OCCode = dtOperation.Rows[k]["OCC"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("OC"))
                        {
                            objOp.OCName = dtOperation.Rows[k]["OC"].ToString().Replace("'", "");
                        }
                        if (dtOperation.Columns.Contains("AMOUNTSIGN"))
                        {
                            objOp.AMOUNTSIGN = dtOperation.Rows[k]["AMOUNTSIGN"].ToString().Replace("'", "");
                        }
                        //else if (dtOperation.Columns[j].ColumnName == "OA")
                        //{
                        //    if (dtOperation.Rows[k][j].ToString() == "" || dtOperation.Rows[k][j].ToString() == null)
                        //        objOp.OA = "0.00";
                        //    else
                        //        objOp.OA = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        //}

                     
                        if (dtOperation.Columns.Contains("OA"))
                        {
                            if (dtOperation.Rows[k]["OA"].ToString() == "" || dtOperation.Rows[k]["OA"].ToString() == null)
                                objOp.OA = "0.00";
                            else
                                objOp.OA = dtOperation.Rows[k]["OA"].ToString().Replace("'", "");
                        }

                        else objOp.OA = "0.00";
                        




                        #endregion
                    }
                    //objOpList.Add(objOp);

                    sql = "Insert into Operation(STATEMENTNO,O,OD,TD,A,ACURC,ACURN,D,DE,PAN,OA,OCC,OC,TL,TERMN,CF,S,MN,DOCNO,NO,ACCOUNT,ACC,FR,APPROVAL,AMOUNTSIGN) " +
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
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return "Error: " + ex.StackTrace;
            }
        }

        private void GenerateStatementInfo(DataSet dsStatement, string BankName, ref string errMsg)
        {
            string reply = string.Empty;
            errMsg = string.Empty;

            try
            {
                DataTable dtOperation = dsStatement.Tables["Operation"];
                DataSet dsBDT = objProvider.ReturnData("select * from statement_DUAL where ACURN='BDT'", ref reply);

                if (dsBDT != null)
                {
                    if (dsBDT.Tables.Count > 0)
                    {
                        if (dsBDT.Tables[0].Rows.Count > 0)
                        {
                            DataTable dtStatementBDT = dsBDT.Tables[0];
                        }
                    }
                }

                reply = string.Empty;
                errMsg = string.Empty;
                DataSet dsUSD = objProvider.ReturnData("select * from statement_DUAL where ACURN='USD'", ref reply);

                if (dsBDT != null)
                {
                    if (dsBDT.Tables.Count > 0)
                    {
                        if (dsBDT.Tables[0].Rows.Count > 0)
                        {
                            DataTable dtStatementBDT = dsBDT.Tables[0];
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                errMsg = ex.StackTrace;
            }

        }        

        private bool IsValid_OLD(string emailaddress)
        {
            try
            {
                MailAddress m = new MailAddress(emailaddress);
                return true;
            }
            catch (FormatException ex)
            {
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return false;
            }
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
        private string getNumberFormat(string vDate)
        {
            string[] omitSpace = vDate.Split(' ');
            string[] date = omitSpace[0].Split('/');
            DateTime dt = new DateTime(Int32.Parse(date[2]), Int32.Parse(date[1]), Int32.Parse(date[0]));
            string formatedDate = string.Format("{0:dd-MMM-yyyy}", dt);
            return formatedDate;
        }
        public string getNumberFormat1(string vDate)
        {
            string[] omitSpace = vDate.Split(' ');
            string[] date = omitSpace[0].Split('/');
            DateTime dt = new DateTime(Int32.Parse(date[2]), Int32.Parse(date[0]), Int32.Parse(date[1]));
            string formatedDate = string.Format("{0:dd-MMM-yyyy}", dt);
            return formatedDate;
        }

    }
}
