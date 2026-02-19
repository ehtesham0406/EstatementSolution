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
using QCash.EStatement.DBL_MC.Reports;
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
                string p = dtpStmDate.Value.ToString();
                if (StmDate == "")
                {
                    // StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy");
                    StmDate = getNumberFormat1(dtpStmDate.Value.ToString());
                }
                else //StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy");
                {
                    StmDate = getNumberFormat1(dtpStmDate.Value.ToString());
                }

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


                                            objESList[i].STATUS = "0";
                                            EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                            count++;
                                        }
                                        catch (Exception ex)
                                        {
                                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message);

                                            objESList[i].STATUS = "8";
                                            EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                        }
                                    }
                                    else
                                    {
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + "No Mail Address Found to send the Estatement " + objESList[i].FILE_LOCATION }); ;
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : No Mail Address Found to send the Estatement " + objESList[i].FILE_LOCATION);

                                        objESList[i].STATUS = "8";
                                        EStatementManager.Instance().UpdateEStatement(objESList[i], ref reply);
                                    }
                                }
                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " Estatement has mailed out of " + objESList.Count + "." });
                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " Estatement has mailed" + objESList.Count + ".");
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
        //
        void btnGenerate_Click(object sender, EventArgs e)
        {
            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
            string reply = string.Empty;
            MsgLogWriter objLW = new MsgLogWriter();

            DataTable dtCardbdt = new DataTable();
            dtCardbdt = objProvider.ReturnData("select * from Qry_Card_Account where Curr='BDT'", ref reply).Tables[0];// where Curr='BDT'

            if (dtCardbdt.Rows.Count > 0)
            {
                txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Processing Estatement." });//Processing Estatement BDT
                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Processing Estatement.");//Processing Estatement BDT.

                //Process pdf for BDT
                ProcessStatementBDT(dtCardbdt);
            }

            /*DataTable dtCardusd = new DataTable();
            dtCardusd = objProvider.ReturnData("select * from Qry_Card_Account where Curr='USD'", ref reply).Tables[0];
            if (dtCardusd != null)
            {
                if (dtCardusd.Rows.Count > 0)
                {
                    txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Processing Estatement USD." });
                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Processing Estatement USD.");
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
            ds = objProvider.ReturnData("select * from Statement_DUAL", ref reply);
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
                            if (dtCards.Rows[j]["EMAIL"].ToString().Trim() != "")
                            {
                               // if (dtCards.Rows[j]["PAN"].ToString().Trim() != "-")
                               //{
                                if (IsValid(dtCards.Rows[j]["EMAIL"].ToString().Trim()))
                                {
                                    try
                                    {
                                        pdfCount = pdfCount + 1;
                                        stmdt = new DataTable();
                                        stmdt = objProvider.ReturnData("select * from Statement_DUAL where CONTRACTNO='" + dtCards.Rows[j]["CONTRACTNO"].ToString() + "'", ref reply).Tables[0];
                                       // if ((dtCards.Rows[j]["pan"].ToString()) != vPAN)
                                       // {

                                            vPAN = dtCards.Rows[j]["pan"].ToString();


                                            #region
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
                                                //string Bin = dtCards.Rows[j]["pan"].ToString().Substring(0, 6);
                                                fileName = dtCards.Rows[j]["idclient"].ToString() + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + '_' + pdfCount + ".pdf";
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
                                                securitySettings.PermitPrint = true;

                                                // Save the document...
                                                document.Save(filePath + "\\" + fileName);

                                                // objst.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);


                                                EStatementInfo objEst = new EStatementInfo();
                                                objEst.BANK_CODE = stmdt.Rows[0]["bank_code"].ToString();
                                                objEst.STMDATE = stmdt.Rows[0]["STATEMENT_DATE"].ToString();
                                                //objEst.STMDATE = getNumberFormat(stmdt.Rows[0]["STATEMENT_DATE"].ToString());
                                                StmDate = stmdt.Rows[0]["STATEMENT_DATE"].ToString();

                                                string[] drdate = stmdt.Rows[0]["STATEMENT_DATE"].ToString().Split('/','-');
                                                
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

                                            #endregion


                                       // }
                                    }
                                    catch (Exception ex)
                                    {
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + ex.Message);
                                    }
                                }
                                else
                                {
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                                }
                             //} //card
                            } //email
                            else
                            {
                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                            }
                            //else
                            //{
                            //    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + ": WARNING! No Email Address present. \n : Estatement Generatede but email will not be sent for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                            //    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + ": WARNING! No Email Address present. \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));
                            //    try
                            //    {
                            //        pdfCount = pdfCount + 1;
                            //        stmdt = new DataTable();
                            //        stmdt = objProvider.ReturnData("select * from Statement_DUAL where CONTRACTNO='" + dtCards.Rows[j]["CONTRACTNO"].ToString() + "'", ref reply).Tables[0];
                            //        if ((dtCards.Rows[j]["pan"].ToString()) != vPAN)
                            //        {

                            //            vPAN = dtCards.Rows[j]["pan"].ToString();
                            //            if (stmdt.Rows.Count > 0)
                            //            {
                            //                // For VISA GOLD and VISA PLATINUM
                            //                /*EStatement objst = new EStatement();
                            //                EStatementPlatinum objstPlatinum = new EStatementPlatinum();

                            //                if (dtCards.Rows[j]["EMAIL"].ToString().Trim() == "rtte")
                            //                {
                            //                    objst.SetDataSource(stmdt);
                            //                }
                            //                else
                            //                {
                            //                    objstPlatinum.SetDataSource(stmdt);
                            //                }

                            //                fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) + ".pdf";

                            //                if (dtCards.Rows[j]["EMAIL"].ToString().Trim() == "rtte")
                            //                {
                            //                    objst.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);
                            //                }
                            //                else
                            //                    objstPlatinum.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);*/


                            //                EStatement objst = new EStatement();
                            //                objst.SetDataSource(stmdt);
                            //                //fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) + ".pdf";
                            //                //fileName = "VISA_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 4) + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + '_' + pdfCount + ".pdf";
                            //                //string Bin = dtCards.Rows[j]["pan"].ToString().Substring(0, 6);
                            //                fileName = _fiid + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 4) + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + '_' + pdfCount + ".pdf";
                            //                System.IO.Stream st = objst.ExportToStream(ExportFormatType.PortableDocFormat);

                            //                PdfSharp.Pdf.PdfDocument document = PdfReader.Open(st);

                            //                PdfSecuritySettings securitySettings = document.SecuritySettings;

                            //                // Setting one of the passwords automatically sets the security level to 
                            //                // PdfDocumentSecurityLevel.Encrypted128Bit.
                            //                string card_no = dtCards.Rows[j]["pan"].ToString();
                            //                securitySettings.UserPassword = dtCards.Rows[j]["pan"].ToString().Substring(card_no.Length - 4, 4);
                            //                securitySettings.OwnerPassword = "owner";

                            //                // Don´t use 40 bit encryption unless needed for compatibility reasons
                            //                //securitySettings.DocumentSecurityLevel = PdfDocumentSecurityLevel.Encrypted40Bit;

                            //                // Restrict some rights.            
                            //                securitySettings.PermitAccessibilityExtractContent = false;
                            //                securitySettings.PermitAnnotations = false;
                            //                securitySettings.PermitAssembleDocument = false;
                            //                securitySettings.PermitExtractContent = false;
                            //                securitySettings.PermitFormsFill = true;
                            //                securitySettings.PermitFullQualityPrint = false;
                            //                securitySettings.PermitModifyDocument = true;
                            //                securitySettings.PermitPrint = true;

                            //                // Save the document...
                            //                document.Save(filePath + "\\" + fileName);

                            //                // objst.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);


                            //                EStatementInfo objEst = new EStatementInfo();
                            //                objEst.BANK_CODE = stmdt.Rows[0]["bank_code"].ToString();
                            //                objEst.STMDATE = stmdt.Rows[0]["STATEMENT_DATE"].ToString();
                            //                StmDate = stmdt.Rows[0]["STATEMENT_DATE"].ToString();

                            //                string[] drdate = stmdt.Rows[0]["STATEMENT_DATE"].ToString().Split('/');

                            //                if (drdate.Length == 3)
                            //                {
                            //                    objEst.MONTH = drdate[1].ToString();
                            //                    objEst.YEAR = drdate[2].ToString();
                            //                }
                            //                else
                            //                {
                            //                    objEst.MONTH = null;
                            //                    objEst.YEAR = null;
                            //                }
                            //                objEst.PAN_NUMBER = dtCards.Rows[j]["pan"].ToString();

                            //                if (stmdt.Rows.Count > 0)
                            //                    objEst.MAILADDRESS = stmdt.Rows[0]["EMAIL"].ToString();
                            //                else
                            //                    objEst.MAILADDRESS = null;

                            //                objEst.FILE_LOCATION = filePath + "\\" + fileName;
                            //                objEst.MAILSUBJECT = txtEmailSubject.Text.Replace("'", "''");
                            //                objEst.MAILBODY = txtEmailBody.Text.Replace("'", "''");
                            //                objEst.STATUS = "1";

                            //                reply = EStatementManager.Instance().AddEStatement(objEst, ref reply);

                            //                if (reply == "Success")
                            //                {
                            //                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4) });
                            //                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4));
                            //                    count++;
                            //                }
                            //                else
                            //                {
                            //                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Message " + reply });
                            //                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + "Message " + reply);
                            //                }
                            //                if (count % 10 == 0)
                            //                {
                            //                    objst.Dispose();
                            //                    GC.Collect();
                            //                    Thread.Sleep(1000);
                            //                }
                            //            }
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                            //        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + ex.Message);
                            //    }
                            //}
                        }
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " Estatement has processed out of " + dtCards.Rows.Count + "." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " Estatement has processed" + dtCards.Rows.Count + ".");
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

                        txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + "Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement.");

                        for (int j = 0; j < dtCards.Rows.Count; j++)//dtCards.Rows.Count
                        {
                            if (dtCards.Rows[j]["EMAIL"].ToString().Trim() != "")
                            {
                                if (IsValid(dtCards.Rows[j]["EMAIL"].ToString().Trim()))
                                {
                                    try
                                    {
                                        stmdt = new DataTable();
                                        stmdt = objProvider.ReturnData("select * from statement_USD where IDCLIENT='" + dtCards.Rows[j]["IDCLIENT"].ToString() + "'", ref reply).Tables[0];
                                        if (stmdt.Rows.Count > 0)
                                        {
                                            EStatement objst = new EStatement();
                                            EStatementPlatinum objstPlatinum = new EStatementPlatinum();

                                            if (dtCards.Rows[j]["EMAIL"].ToString().Trim() == "rtte")
                                            {
                                                objst.SetDataSource(stmdt);
                                            }
                                            else
                                            {
                                                objstPlatinum.SetDataSource(stmdt);
                                            }

                                            //fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) + "_USD.pdf";
                                            fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) + ".pdf";
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
                                else
                                {
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Invalid Email Address present " + dtCards.Rows[j]["EMAIL"].ToString().Trim() + " \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                                }
                            }
                            else
                            {
                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : No Email Address present !!!\n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));

                            }
                            //else
                            //{
                            //    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + ": WARNING! No Email Address present. \n : Estatement Generatede but email will not be sent for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4) });
                            //    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + ": WARNING! No Email Address present. \n : Estatement has not been created for Card# " + dtCards.Rows[j]["PAN"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["PAN"].ToString().Substring(12, 4));
                            //    try
                            //    {
                            //        pdfCount = pdfCount + 1;
                            //        stmdt = new DataTable();
                            //        stmdt = objProvider.ReturnData("select * from Statement_DUAL where CONTRACTNO='" + dtCards.Rows[j]["CONTRACTNO"].ToString() + "'", ref reply).Tables[0];
                            //        if ((dtCards.Rows[j]["pan"].ToString()) != vPAN)
                            //        {

                            //            vPAN = dtCards.Rows[j]["pan"].ToString();
                            //            if (stmdt.Rows.Count > 0)
                            //            {
                            //                // For VISA GOLD and VISA PLATINUM
                            //                /*EStatement objst = new EStatement();
                            //                EStatementPlatinum objstPlatinum = new EStatementPlatinum();

                            //                if (dtCards.Rows[j]["EMAIL"].ToString().Trim() == "rtte")
                            //                {
                            //                    objst.SetDataSource(stmdt);
                            //                }
                            //                else
                            //                {
                            //                    objstPlatinum.SetDataSource(stmdt);
                            //                }

                            //                fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) + ".pdf";

                            //                if (dtCards.Rows[j]["EMAIL"].ToString().Trim() == "rtte")
                            //                {
                            //                    objst.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);
                            //                }
                            //                else
                            //                    objstPlatinum.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);*/


                            //                EStatement objst = new EStatement();
                            //                objst.SetDataSource(stmdt);
                            //                //fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) + ".pdf";
                            //                //fileName = "VISA_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 4) + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + '_' + pdfCount + ".pdf";
                            //                //string Bin = dtCards.Rows[j]["pan"].ToString().Substring(0, 6);
                            //                fileName = _fiid + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 4) + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + '_' + pdfCount + ".pdf";
                            //                System.IO.Stream st = objst.ExportToStream(ExportFormatType.PortableDocFormat);

                            //                PdfSharp.Pdf.PdfDocument document = PdfReader.Open(st);

                            //                PdfSecuritySettings securitySettings = document.SecuritySettings;

                            //                // Setting one of the passwords automatically sets the security level to 
                            //                // PdfDocumentSecurityLevel.Encrypted128Bit.
                            //                string card_no = dtCards.Rows[j]["pan"].ToString();
                            //                securitySettings.UserPassword = dtCards.Rows[j]["pan"].ToString().Substring(card_no.Length - 4, 4);
                            //                securitySettings.OwnerPassword = "owner";

                            //                // Don´t use 40 bit encryption unless needed for compatibility reasons
                            //                //securitySettings.DocumentSecurityLevel = PdfDocumentSecurityLevel.Encrypted40Bit;

                            //                // Restrict some rights.            
                            //                securitySettings.PermitAccessibilityExtractContent = false;
                            //                securitySettings.PermitAnnotations = false;
                            //                securitySettings.PermitAssembleDocument = false;
                            //                securitySettings.PermitExtractContent = false;
                            //                securitySettings.PermitFormsFill = true;
                            //                securitySettings.PermitFullQualityPrint = false;
                            //                securitySettings.PermitModifyDocument = true;
                            //                securitySettings.PermitPrint = true;

                            //                // Save the document...
                            //                document.Save(filePath + "\\" + fileName);

                            //                // objst.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);


                            //                EStatementInfo objEst = new EStatementInfo();
                            //                objEst.BANK_CODE = stmdt.Rows[0]["bank_code"].ToString();
                            //                objEst.STMDATE = stmdt.Rows[0]["STATEMENT_DATE"].ToString();
                            //                StmDate = stmdt.Rows[0]["STATEMENT_DATE"].ToString();

                            //                string[] drdate = stmdt.Rows[0]["STATEMENT_DATE"].ToString().Split('/');

                            //                if (drdate.Length == 3)
                            //                {
                            //                    objEst.MONTH = drdate[1].ToString();
                            //                    objEst.YEAR = drdate[2].ToString();
                            //                }
                            //                else
                            //                {
                            //                    objEst.MONTH = null;
                            //                    objEst.YEAR = null;
                            //                }
                            //                objEst.PAN_NUMBER = dtCards.Rows[j]["pan"].ToString();

                            //                if (stmdt.Rows.Count > 0)
                            //                    objEst.MAILADDRESS = stmdt.Rows[0]["EMAIL"].ToString();
                            //                else
                            //                    objEst.MAILADDRESS = null;

                            //                objEst.FILE_LOCATION = filePath + "\\" + fileName;
                            //                objEst.MAILSUBJECT = txtEmailSubject.Text.Replace("'", "''");
                            //                objEst.MAILBODY = txtEmailBody.Text.Replace("'", "''");
                            //                objEst.STATUS = "1";

                            //                reply = EStatementManager.Instance().AddEStatement(objEst, ref reply);

                            //                if (reply == "Success")
                            //                {
                            //                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4) });
                            //                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4));
                            //                    count++;
                            //                }
                            //                else
                            //                {
                            //                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Message " + reply });
                            //                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + "Message " + reply);
                            //                }
                            //                if (count % 10 == 0)
                            //                {
                            //                    objst.Dispose();
                            //                    GC.Collect();
                            //                    Thread.Sleep(1000);
                            //                }
                            //            }
                            //        }
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                            //        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + ex.Message);
                            //    }
                            //}
                        }
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " Estatement has processed out of " + dtCards.Rows.Count + "." });
                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + count.ToString() + " Estatement has processed" + dtCards.Rows.Count + ".");
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
                if (dia[fcount].FullName.Contains("DBL"))
                {
                    _bankName = "DBL";
                    _bankCode = "3";
                    _XMLSourcePath = dia[fcount].FullName;
                    //
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }
                //else if (dia[fcount].FullName.Contains("TBL"))
                //{
                //    _bankName = "TBL";
                //    _bankCode = "5";
                //    _XMLSourcePath = dia[fcount].FullName;
                //    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                //}
                else if (dia[fcount].FullName.Contains("IFIC"))
                {
                    _bankName = "IFIC";
                    _bankCode = "8";
                    _XMLSourcePath = dia[fcount].FullName;
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }
                else if (dia[fcount].FullName.Contains("JBL"))
                {
                    _bankName = "JBL";
                    _bankCode = "9";
                    _XMLSourcePath = dia[fcount].FullName;
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }
                else if (dia[fcount].FullName.Contains("SBL"))
                {
                    _bankName = "SBL";
                    _bankCode = "10";
                    _XMLSourcePath = dia[fcount].FullName;
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }
                else if (dia[fcount].FullName.Contains("MDBL"))
                {
                    _bankName = "MDBL";
                    _bankCode = "17";
                    _XMLSourcePath = dia[fcount].FullName;
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }
                else if (dia[fcount].FullName.Contains("LBF"))
                {
                    _bankName = "LBF";
                    _bankCode = "12";
                    _XMLSourcePath = dia[fcount].FullName;
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }
                else if (dia[fcount].FullName.Contains("SEBL"))
                {
                    _bankName = "SEBL";
                    _bankCode = "15";
                    _XMLSourcePath = dia[fcount].FullName;
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }
                else if (dia[fcount].FullName.Contains("BAL"))
                {
                    _bankName = "BAL";
                    _bankCode = "4";
                    _XMLSourcePath = dia[fcount].FullName;
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }
                else if (dia[fcount].FullName.Contains("NRBB"))
                {
                    _bankName = "NRBB";
                    _bankCode = "16";
                    _XMLSourcePath = dia[fcount].FullName;
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }

                else if (dia[fcount].FullName.Contains("SIBL"))
                {
                    _bankName = "SIBL";
                    _bankCode = "19";
                    _XMLSourcePath = dia[fcount].FullName;
                    ProcessFolderFiles(_XMLSourcePath, _bankCode, _bankName, ref _reply);
                }
                else if (dia[fcount].FullName.Contains("MBL"))
                {
                    _bankName = "MBL";
                    _bankCode = "21";
                    _XMLSourcePath = dia[fcount].FullName;
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
                FileInfo[] fi = dir.GetFiles();

                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + fi.Length.ToString() + " files found to process.." });
                objLW.logTrace(_LogPath, "EStatement.log", " : Total " + fi.Length.ToString() + " files found to process..");
                //if (fi.Length == 1)
                //{
                for (int f = 0; f < fi.Length; f++)
                {
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + fi[f].Name + " on process.." });
                    objLW.logTrace(_LogPath, "EStatement.log", " : " + fi[f].Name + " on process..");

                    DataSet dsXML = getDataFromXML(fi[f].FullName);

                    #region Operation On Data
                    if (dsXML != null)
                    {
                        if (dsXML.Tables.Count > 0)
                        {
                            ConStr = new ConnectionStringBuilder(1);
                            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);

                            string sql = string.Empty;

                            //Clear Previous AccumIntAcc Data
                            objProvider.RunQuery("Delete from AccumIntAcc");
                            //Clear Previous BonusContrAcc Data
                            objProvider.RunQuery("Delete from  BonusContrAcc");

                            for (int i = 0; i < dsXML.Tables.Count; i++)
                            {
                                if (dsXML.Tables[i].TableName == "Statement")
                                {
                                    GetCardHolderPersonalInfo(dsXML.Tables[i], BankName, ref reply);
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : CardHolder Personal Info data Saved from XML. " + reply });
                                    objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Personal Info data Saved from XML. " + reply);
                                }
                                else if (dsXML.Tables[i].TableName == "Operation")
                                {
                                    reply = GetCardHolderTransactionInfo(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : CardHolder Transaction Info data Saved from XML. " + reply });
                                    objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Transaction Info data Saved from XML. " + reply);
                                }
                                else if (dsXML.Tables[i].TableName == "Account")
                                {
                                    reply = GetCardHolderAccountInfo(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : CardHolder Account Info data Saved from XML. " + reply });
                                    objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Account Info data Saved from XML. " + reply);
                                }
                                else if (dsXML.Tables[i].TableName == "Card")
                                {
                                    reply = GetCardHolderCardInfo(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : CardHolder Card Info data Saved from XML. " + reply });
                                    objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Card Info data Saved from XML. " + reply);
                                }
                                else if (dsXML.Tables[i].TableName == "BonusContrAcc")
                                {
                                    reply = GetBonusContrAccInfo(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : CardHolder Reward Info data Saved from XML. " + reply });
                                    objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Card Info data Saved from XML. " + reply);
                                }
                                else if (dsXML.Tables[i].TableName == "AccumIntAcc")
                                {
                                    reply = GetAccumIntAcc(dsXML.Tables[i]);
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : CardHolder Suspense Account Info data Saved from XML. " + reply });
                                    objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Card Info data Saved from XML. " + reply);
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
                        }
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Total " + dsXML.Tables["Card"].Rows.Count.ToString() + " Card record has been found to process.." });
                        objLW.logTrace(_LogPath, "EStatement.log", " : Total " + dsXML.Tables["Card"].Rows.Count.ToString() + " Card record has been found to process..");
                    }
                    #endregion

                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : " + fi[f].Name + " process complete.." });
                    objLW.logTrace(_LogPath, "EStatement.log", " : " + fi[f].Name + " process complete..");
                    txtAnalyzer.Invoke(_addText, new object[] { "\n" });

                    //CalcuateBalance();

                    btnGenerate_Click(null, null);
                }
                //}
                //else 
                //{
                //    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " More than one XML found to process.." });
                //}
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
                // return true;
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

            try
            {
                //Shorcut method need to implement next
                objProvider.RunQuery("Delete from dbo.AccumIntAcc");
                objProvider.RunQuery("Delete from dbo.BonusContrAcc");

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
                        else if (dtStatement.Columns[j].ColumnName == "Address")
                        {
                            objSt.ADDRESS = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "CARD_LIST")
                        {
                            objSt.CARD_LIST = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "City")
                        {
                            objSt.CITY = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Region")
                        {
                            objSt.REGION = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Country")
                        {
                            objSt.COUNTRY = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Email")
                        {
                            objSt.EMAIL = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "StartDate")
                        {
                            objSt.STARTDATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                            objSt.STARTDATE = getNumberFormat(dtStatement.Rows[k][j].ToString().Replace("'", ""));
                        }
                        else if (dtStatement.Columns[j].ColumnName == "EndDate")
                        {
                            objSt.ENDDATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                            objSt.ENDDATE = getNumberFormat(dtStatement.Rows[k][j].ToString().Replace("'", ""));
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Client")
                        {
                            objSt.CLIENT = dtStatement.Rows[k][j].ToString().Replace("'", "");
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
                            objSt.FAX = dtStatement.Rows[k][j].ToString().Replace("'", "");
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
                            objSt.NEXT_STATEMENT_DATE = getNumberFormat(dtStatement.Rows[k][j].ToString().Replace("'", ""));
                        }
                        else if (dtStatement.Columns[j].ColumnName == "PAYMENT_DATE")
                        {
                            objSt.PAYMENT_DATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                            objSt.PAYMENT_DATE = getNumberFormat(dtStatement.Rows[k][j].ToString().Replace("'", ""));
                        }
                        else if (dtStatement.Columns[j].ColumnName == "STATEMENT_DATE")
                        {
                            objSt.STATEMENT_DATE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                            objSt.STATEMENT_DATE = getNumberFormat(dtStatement.Rows[k][j].ToString().Replace("'", ""));
                        }
                        if (dtStatement.Columns[j].ColumnName == "StreetAddress")
                        {
                            objSt.STREETADDRESS = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Telephone")
                        {
                            objSt.TELEPHONE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Title")
                        {
                            objSt.TITLE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "ZIP")
                        {
                            objSt.ZIP = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "JobTitle")
                        {
                            objSt.JOBTITLE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "PromotionalText")
                        {
                            objSt.PTEXT = dtStatement.Rows[k][j].ToString().Replace("<COMPANYNAME>", "").Replace("</COMPANYNAME>", ""); 
                            int pindex = objSt.PTEXT.IndexOf("<");
                            if (pindex > 0)
                            {
                                objSt.PPROMOTIONALTEXT = objSt.PTEXT.Substring(0, pindex);
                            }

                        }
                        #endregion
                    }
                    objStList.Add(objSt);

                    sql = "Insert into Statement(BANK_CODE,STATEMENTNO,ADDRESS,CARD_LIST,CITY,COUNTRY,EMAIL," +
                          "STARTDATE,ENDDATE,CLIENT,CONTRACTNO,IDCLIENT,FAX,MAIN_CARD,MOBILE," +
                          "NEXT_STATEMENT_DATE,PAYMENT_DATE,REGION,STATEMENT_DATE,SEX,STREETADDRESS,TELEPHONE,TITLE,ZIP,JOBTITLE,PPROMOTIONALTEXT) " +
                          "values('" + objSt.BANK_CODE + "','" + objSt.STATEMENTNO + "','" + objSt.ADDRESS + "','" + objSt.CARD_LIST + "','" + objSt.CITY + "','" + objSt.COUNTRY + "','" + objSt.EMAIL + "'," +
                          "'" + objSt.STARTDATE + "','" + objSt.ENDDATE + "','" + objSt.CLIENT + "','" + objSt.CONTRACTNO + "','" + objSt.IDCLIENT + "','" + objSt.FAX + "','" + objSt.MAIN_CARD + "','" + objSt.MOBILE + "'," +
                          "'" + objSt.NEXT_STATEMENT_DATE + "','" + objSt.PAYMENT_DATE + "','" + objSt.REGION + "','" + objSt.STATEMENT_DATE + "','" + objSt.SEX + "','" + objSt.STREETADDRESS + "'," +
                          "'" + objSt.TELEPHONE + "','" + objSt.TITLE + "','" + objSt.ZIP + "','" + objSt.JOBTITLE + "','" + objSt.PPROMOTIONALTEXT + "')";

                    reply = objProvider.RunQuery(sql);
                    //if (!reply.Contains("Success"))
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
                            if (dtOperation.Rows[k][j].ToString() == "" || dtOperation.Rows[k][j].ToString() == null)
                            {
                                objOp.OpDate = "";
                            }
                            else
                            {
                                objOp.OpDate = getNumberFormat(objOp.OpDate);                               
                            }
                        }
                        else if (dtOperation.Columns[j].ColumnName == "TD")
                        {
                            objOp.TD = dtOperation.Rows[k][j].ToString().Replace("'", "");
                            if (dtOperation.Rows[k][j].ToString() == "" || dtOperation.Rows[k][j].ToString() == null)
                            {
                                objOp.TD = "";
                            }
                            else
                            {
                                objOp.TD = getNumberFormat(objOp.TD);                                
                            }
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
                            objOp.ACURName = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "D")
                        {
                            objOp.D = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "DE")
                        {
                            objOp.DE = dtOperation.Rows[k][j].ToString().Replace("'", "");
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
                            objOp.FR = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "APPROVAL")
                        {
                            objOp.APPROVAL = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "MN")
                        {
                            objOp.MN = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "S")
                        {
                            objOp.S = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "TERMN")
                        {
                            objOp.TERMN = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "TL")
                        {
                            objOp.TL = dtOperation.Rows[k][j].ToString().Replace("'", "");
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
                            objOp.OCName = dtOperation.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtOperation.Columns[j].ColumnName == "AMOUNTSIGN")
                        {
                            objOp.AMOUNTSIGN = dtOperation.Rows[k][j].ToString().Replace("'", "");
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
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return "Error: " + ex.StackTrace;
            }
        }

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
                            objOp.ACURN = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
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
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
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
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
                MsgLogWriter objLW = new MsgLogWriter();
                objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                return "Error: " + ex.StackTrace;
            }
        }
        ////New Table end
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
                            objAc.ACURN = dtAccount.Rows[k][j].ToString();
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
                        //else if (dtAccount.Columns[j].ColumnName == "AVAIL_CRD_LIMIT")
                        //{
                        //    objAc.AVAIL_CRD_LIMIT = dtAccount.Rows[k][j].ToString();
                        //}
                        else if (dtAccount.Columns[j].ColumnName == "AVAIL_CRD_LIMIT")
                        {
                            objAc.AVAIL_CRD_LIMIT = dtAccount.Rows[k][j].ToString();
                            if (objAc.AVAIL_CRD_LIMIT == "0.00" && (k > 0))
                            {
                                objAc.AVAIL_CRD_LIMIT = dtAccount.Rows[k - 1][j].ToString();
                                objAc.INDICATOR = "BDT";
                            }
                        }
                        //else if (dtAccount.Columns[j].ColumnName == "AVAIL_CASH_LIMIT")
                        //{
                        //    objAc.AVAIL_CASH_LIMIT = dtAccount.Rows[k][j].ToString();
                        //}
                        else if (dtAccount.Columns[j].ColumnName == "AVAIL_CASH_LIMIT")
                        {
                            objAc.AVAIL_CASH_LIMIT = dtAccount.Rows[k][j].ToString();
                            if (objAc.AVAIL_CASH_LIMIT == "0.00" && (k > 0))
                            {
                                objAc.AVAIL_CASH_LIMIT = dtAccount.Rows[k - 1][j].ToString();
                                objAc.INDICATOR = "BDT";
                            }
                        }
                        else if (dtAccount.Columns[j].ColumnName == "INSTALL_UNPAID_AMOUNT")
                        {
                            objAc.INSTALL_UNPAID_AMOUNT = dtAccount.Rows[k][j].ToString();
                        }
                        else if (dtAccount.Columns[j].ColumnName == "INSTALL_MONTH_PAYM")
                        {
                            objAc.INSTALL_MONTH_PAYM = dtAccount.Rows[k][j].ToString();
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
                            if (objAc.CRD_LIMIT == "0.00" && (k > 0))
                            {
                                objAc.CRD_LIMIT = dtAccount.Rows[k-1][j].ToString();
                                objAc.INDICATOR = "BDT";
                            }
                        }
                        #endregion
                    }
                    objAcList.Add(objAc);

                    sql = "Insert into Account(STATEMENTNO,ACCOUNTNO,ACURN,SBALANCE,ACURC,EBALANCE,AVAIL_CRD_LIMIT,AVAIL_CASH_LIMIT,SUM_WITHDRAWAL,SUM_INTEREST,OVLFEE_AMOUNT,OVDFEE_AMOUNT,SUM_REVERSE,SUM_CREDIT,SUM_OTHER,SUM_PURCHASE,MIN_AMOUNT_DUE,CASH_LIMIT,CRD_LIMIT,INSTALL_UNPAID_AMOUNT,INSTALL_MONTH_PAYM,INDICATOR)" +
                        " Values('" + objAc.STATEMENTNO + "','" + objAc.ACCOUNTNO + "','" + objAc.ACURN + "','" + objAc.SBALANCE + "','" + objAc.ACURC + "'," +
                        "'" + objAc.EBALANCE + "','" + objAc.AVAIL_CRD_LIMIT + "','" + objAc.AVAIL_CASH_LIMIT + "','" + objAc.SUM_WITHDRAWAL + "'," +
                        "'" + objAc.SUM_INTEREST + "','" + objAc.OVLFEE_AMOUNT + "','" + objAc.OVDFEE_AMOUNT + "','" + objAc.SUM_REVERSE + "'," +
                        "'" + objAc.SUM_CREDIT + "','" + objAc.SUM_OTHER + "','" + objAc.SUM_PURCHASE + "','" + objAc.MIN_AMOUNT_DUE + "','" + objAc.CASH_LIMIT + "','" + objAc.CRD_LIMIT + "','" + objAc.INSTALL_UNPAID_AMOUNT + "','" + objAc.INSTALL_MONTH_PAYM + "','" + objAc.INDICATOR + "')";


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
        //
        private string GetCardHolderCardInfo(DataTable dtCard)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            Card objCard = null;
            CardList objCardList = new CardList();

            try
            {
                //Clear Previous Data
                objProvider.RunQuery("Delete from " + dtCard.TableName);

                for (int k = 0; k < dtCard.Rows.Count; k++)
                {
                    objCard = new Card();

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
                            objCard.CLIENTNAME = dtCard.Rows[k][j].ToString().Replace("'", "");
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

        //
        private void GenerateStatementInfo(DataSet dsStatement, string BankName, ref string errMsg)
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
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
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

            for (int k = 0; k < dtStatement.Rows.Count; k++)
            {
                try
                {
                    objSt = new StatementInfo();

                    objSt.BANK_CODE = BankName;

                    for (int j = 0; j < dtStatement.Columns.Count; j++)
                    {
                        #region setting properties values

                        //if (dtStatement.Columns[j].ColumnName.Contains("INSTALL_UNPAID_AMOUNT"))
                        //{
                        //    objSt.INSTALL_UNPAID_AMOUNT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        //}
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
                            objSt.ADDRESS = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "PAN")
                        {
                            if (dtStatement.Rows[k][j].ToString().Length >= 16)
                                objSt.PAN = dtStatement.Rows[k][j].ToString().Substring(0, 16);
                            else
                            {
                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Card Not found for the Contract " + objSt.CONTRACTNO });
                                MsgLogWriter objLW = new MsgLogWriter();
                                objLW.logTrace(_LogPath, "EStatement.log", "Card Not fount for the Contract " + objSt.CONTRACTNO);
                                continue;
                            }
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "REGION")
                        {
                            objSt.CITY = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CITY")
                        {
                            objSt.CITY = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }

                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ZIP")
                        {
                            objSt.ZIP = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "COUNTRY")
                        {
                            objSt.COUNTRY = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "EMAIL")
                        {
                            objSt.EMAIL = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "MOBILE")
                        {
                            objSt.MOBILE = dtStatement.Rows[k][j].ToString().Replace("(", "").Replace(")", "").Replace("8800", "880");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "TITLE")
                        {
                            objSt.TITLE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CLIENT")
                        {
                            objSt.CLIENTNAME = dtStatement.Rows[k][j].ToString().Replace("'", "");
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

                        else if (dtStatement.Columns[j].ColumnName.Contains ("INSTALL_UNPAID_AMOUNT"))
                        {
                            //objSt.INSTALL_UNPAID_AMOUNT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                            if (dtStatement.Rows[k][j].ToString()== "" || dtStatement.Rows[k][j].ToString()== null)
                                objSt.INSTALL_UNPAID_AMOUNT = "0.00";
                            else
                            {
                                objSt.INSTALL_UNPAID_AMOUNT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                            }
                        }

                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "INSTALL_MONTH_PAYM")
                        {
                            objSt.INSTALL_MONTH_PAYM = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }

                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "INDICATOR")
                        {
                            objSt.INDICATOR = dtStatement.Rows[k][j].ToString();
                        }

                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "JOBTITLE")
                        {
                            objSt.JOBTITLE = dtStatement.Rows[k][j].ToString().Replace("'", ""); ;
                        }

                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "PPROMOTIONALTEXT")
                        {
                            objSt.PPROMOTIONALTEXT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        #endregion
                    }

                    objSt.STM_MSG = txtStmMsg.Text.Replace("'", "''");
                    objSt.STATUS = "1";

                    sql = "Insert into STATEMENT_INFO(STATEMENTID,BANK_CODE,CONTRACTNO,IDCLIENT,PAN,TITLE,CLIENTNAME,STATEMENTNO,ADDRESS,CITY,ZIP,COUNTRY," +
                        "EMAIL,MOBILE,STARTDATE,ENDDATE,NEXT_STATEMENT_DATE,PAYMENT_DATE,STATEMENT_DATE,ACCOUNTNO,ACURN,SBALANCE,ACURC,EBALANCE,AVAIL_CRD_LIMIT," +
                        "AVAIL_CASH_LIMIT,SUM_WITHDRAWAL,SUM_INTEREST,OVLFEE_AMOUNT,OVDFEE_AMOUNT,SUM_REVERSE,SUM_CREDIT,SUM_OTHER,SUM_PURCHASE," +
                        "MIN_AMOUNT_DUE,CASH_LIMIT,CRD_LIMIT,STM_MSG,STATUS,INSTALL_UNPAID_AMOUNT,INSTALL_MONTH_PAYM,INDICATOR,JOBTITLE,PPROMOTIONALTEXT) VALUES('" + objSt.STATEMENTID + "'," +
                        "'" + objSt.BANK_CODE + "','" + objSt.CONTRACTNO + "','" + objSt.IDCLIENT + "','" + objSt.PAN + "','" + objSt.TITLE + "','" + objSt.CLIENTNAME + "','" + objSt.STATEMENTNO + "'," +
                        "'" + objSt.ADDRESS + "','" + objSt.CITY + "','" + objSt.ZIP + "','" + objSt.COUNTRY + "','" + objSt.EMAIL + "','" + objSt.MOBILE + "','" + objSt.STARTDATE + "','" + objSt.ENDDATE + "'," +
                        "'" + objSt.NEXT_STATEMENT_DATE + "','" + objSt.PAYMENT_DATE + "','" + objSt.STATEMENT_DATE + "','" + objSt.ACCOUNTNO + "','" + objSt.ACURN + "'," +
                        "'" + objSt.SBALANCE + "','" + objSt.ACURC + "','" + objSt.EBALANCE + "','" + objSt.AVAIL_CRD_LIMIT + "','" + objSt.AVAIL_CASH_LIMIT + "'," +
                        "'" + objSt.SUM_WITHDRAWAL + "','" + objSt.SUM_INTEREST + "','" + objSt.OVLFEE_AMOUNT + "','" + objSt.OVDFEE_AMOUNT + "','" + objSt.SUM_REVERSE + "'," +
                        "'" + objSt.SUM_CREDIT + "','" + objSt.SUM_OTHER + "','" + objSt.SUM_PURCHASE + "','" + objSt.MIN_AMOUNT_DUE + "','" + objSt.CASH_LIMIT + "'," +
                        "'" + objSt.CRD_LIMIT + "','" + objSt.STM_MSG + "','" + objSt.STATUS + "','" + objSt.INSTALL_UNPAID_AMOUNT + "','" + 0 + "','" + objSt.INDICATOR + "','" + objSt.JOBTITLE + "','" + objSt.PPROMOTIONALTEXT + "')";

                    reply = objProvider.RunQuery(sql);
                    if (dtOperation != null && dtOperation.Columns.Contains("ACCOUNT"))
                    {
                        if (dtOperation.Rows.Count > 0)
                        {

                            DataRow[] dr = dtOperation.Select("STATEMENTNO='" + objSt.STATEMENTNO + "' AND ACCOUNT='" + objSt.ACCOUNTNO + "'");
                            if (dr.Length > 0)
                            {
                                double feesnCharges = 0.00;
                                string trn_Date = string.Empty;

                                for (int l = 0; l < dr.Length; l++)
                                {

                                    #region setting properties values
                                    if (!dr[l]["D"].ToString().Contains("INTEREST ON"))
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
                                        {
                                            //objSTD.TRNDATE = dr[l]["TD"].ToString();
                                            // objOp.OpDate = getNumberFormat(objOp.OpDate); 
                                            if (dr[l]["TD"].ToString() == "" || dr[l]["TD"].ToString() == null)
                                                objSTD.TRNDATE = "";
                                            else
                                                objSTD.TRNDATE = getNumberFormat(dr[l]["TD"].ToString());
                                        }

                                        if (dr[l].Table.Columns.Contains("OD"))
                                        {
                                            //objSTD.POSTDATE = dr[l]["OD"].ToString();
                                            // objSTD.POSTDATE = getNumberFormat(dr[l]["OD"].ToString());
                                            if (dr[l]["OD"].ToString() == "" || dr[l]["OD"].ToString() == null)
                                                objSTD.POSTDATE = "";
                                            else
                                                objSTD.POSTDATE = getNumberFormat(dr[l]["OD"].ToString());
                                        }

                                        if (dr[l].Table.Columns.Contains("ACURN"))
                                            objSTD.ACURN = dr[l]["ACURN"].ToString();

                                        if (dr[l].Table.Columns.Contains("OC"))
                                            objSTD.OC = dr[l]["OC"].ToString();

                                        if (dr[l].Table.Columns.Contains("P"))
                                            objSTD.P = dr[l]["P"].ToString();

                                        if (dr[l].Table.Columns.Contains("DOCNO"))
                                            objSTD.DOCNO = dr[l]["DOCNO"].ToString();

                                        if (dr[l].Table.Columns.Contains("DE"))
                                            objSTD.DE = dr[l]["DE"].ToString();


                                        if (dr[l].Table.Columns.Contains("NO"))
                                            objSTD.NO = dr[l]["NO"].ToString();

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
                                        if (dr[l].Table.Columns.Contains("OC"))
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
                                            objSTD.OC = "";//dr[l]["OC"].ToString();

                                        if (dr[l].Table.Columns.Contains("OC"))
                                        {
                                            if (dr[l]["OA"].ToString() == "" || dr[l]["OA"].ToString() == null)
                                                objSTD.ORGAMOUNT = "0.00";
                                            else
                                                objSTD.ORGAMOUNT = dr[l]["OA"].ToString();
                                        }
                                        else objSTD.ORGAMOUNT = "0.00";

                                        //Remmove Terminal Name when Fee and VAT Impose
                                        //Sum Charges amount with Fees & Charges. 
                                        if ((!dr[l]["D"].ToString().ToUpper().Contains("FEE")) || (dr[l]["D"].ToString() != "Charge interest for Installment") || (dr[l]["D"].ToString() != "Credit Shield Premium") || (dr[l]["D"].ToString() != "Monthly Installment"))
                                        {
                                            if (dr[l].Table.Columns.Contains("TL") && dr[l]["DE"].ToString().ToUpper() != "FEE")
                                                objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                            else
                                                objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                        }
                                        else
                                        {
                                            objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                            feesnCharges = feesnCharges + Convert.ToDouble(dr[l]["A"].ToString().Replace("-", ""));
                                            if (dr[l].Table.Columns.Contains("OD"))
                                                objSTD.TRNDATE = dr[l]["OD"].ToString();
                                        }

                                        //if (objSTD.TRNDESC.Contains("Credit cash deposit"))
                                        //{
                                        //    objSTD.TRNDESC = "PAYMENT RECEIVED (THANK YOU)";
                                        //    objSTD.TRNDATE = dr[l]["OD"].ToString();
                                        //}
                                        var Entrylist=new List<String>(){"Credit cash deposit","[MANUAL_TXN[POS]]","[MANUAL_TXN[POS-R]]","CREDIT ADJUSTMENT","DEBIT ADJUSTMENT","Credit acct","CASH BACK[BDT]","ATM INTEREST (REVERSE)","CARD FEE(REVERSE)","CASH ADVANCE FEE(REVERSE)","LATE PAYMENT FEE(REVERSE)","CASH DEPOSIT (REVERSE)","SALES_SLIP_RET_FEE","STATEMENT REPRINT FEE","VAT","REVERSAL-POS PURCHASE","OVER LIMIT FEE(REVERSE)","PIN FEE(REVERSE)","POS INTEREST(REVERSE)","CARD REPLACEMENT FEE(REVERSE)","ATM INTEREST","BAL_TRNS","FUND TRANSFER","INTEREST CHARGE","POS INTEREST"};
                                        if (Entrylist.Contains(objSTD.TRNDESC.Trim(), StringComparer.OrdinalIgnoreCase))
                                        {
                                            if (dr[l]["FR"].ToString() == "" || dr[l]["FR"].ToString() == null)
                                                objSTD.TRNDESC = dr[l]["TRNDESC"].ToString().Replace("'", "''");
                                            else
                                                objSTD.TRNDESC = dr[l]["FR"].ToString().Replace("'", "''");
                                            
                                        }
                                        
                                        if (dr[l].Table.Columns.Contains("APPROVAL"))
                                        {
                                            objSTD.APPROVAL = dr[l]["APPROVAL"].ToString().Replace("'", "");

                                            if (dr[l]["APPROVAL"].ToString() != "" && objSTD.TRNDATE == "")
                                            {
                                                objSTD.TRNDATE = dr[l]["OD"].ToString();
                                            }
                                        }
                                        if (dr[l].Table.Columns.Contains("FR"))
                                            objSTD.FR = dr[l]["FR"].ToString().Replace("'", "''");;
                                        //objSTD.AMOUNTSIGN = dr[l]["AMOUNTSIGN"].ToString();

                                        if (objSTD.FR == "P 10")
                                        {
                                            objSTD.TRNDESC = "Purchase" + " " + dr[l]["TL"].ToString().Replace("'", "''");

                                        }
                                        else if (objSTD.FR == "A 10")
                                        {
                                            objSTD.TRNDESC = "Cash Withdrawal" + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                        }

                                        sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,OC,ORGAMOUNT,AMOUNTSIGN,APPROVAL,FR,P,DOCNO,NO,DE)" +
                                            " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                            "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.OC + "','" + objSTD.ORGAMOUNT + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.APPROVAL + "','" + objSTD.FR + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "','" + objSTD.DE + "')";

                                        reply = objProvider.RunQuery(sql);
                                        if (!reply.Contains("Success"))
                                            errMsg = reply;
                                    }
                                    else if (dr[l]["D"].ToString().Contains("Charge interest"))
                                    {
                                        trn_Date = dr[l]["OD"].ToString();
                                    }

                                    if (dr[l]["D"].ToString() == "Charge interest for Installment")
                                    {
                                        StatementDetails objSTD = new StatementDetails();
                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                        objSTD.PAN = objSt.PAN;
                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                        objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                        objSTD.ACURN = objSt.ACURN;
                                        objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                        //objSTD.AMOUNT = "-" + objSt.SUM_INTEREST;//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                        objSTD.AMOUNT = dr[l]["A"].ToString();
                                        // objSTD.TRNDATE = trn_Date;
                                        //objSTD.POSTDATE = trn_Date;
                                        objSTD.TRNDATE = dr[l]["TD"].ToString();
                                        objSTD.POSTDATE = dr[l]["OD"].ToString();

                                        sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN)" +
                                                " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "')";

                                        reply = objProvider.RunQuery(sql);
                                        if (!reply.Contains("Success"))
                                            errMsg = reply;
                                    }


                                    #endregion
                                }

                                if (objSt.SUM_INTEREST != "0.00")
                                {
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

                                    sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN)" +
                                            " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                            "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "')";

                                    reply = objProvider.RunQuery(sql);
                                    if (!reply.Contains("Success"))
                                        errMsg = reply;
                                }

                                //New View added

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
                                                objSTD.ACURN = objSt.ACURN;
                                                objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                //if (objSTD.CONTRACTNO == dtAcI.Rows[0][1].ToString())
                                                if (objSTD.CONTRACTNO == dtAcI.Rows[x]["CONTRACTNO"].ToString() && objSTD.ACURN == dtAcI.Rows[x]["ACURN"].ToString())// Rocky 27-03-2017
                                                {
                                                    //if (dtAcI.Rows[0][0].ToString() != "0.00")
                                                    if (dtAcI.Rows[x]["ACCUM_INT_AMOUNT"].ToString() != "0.00")// Rocky 27-03-2017
                                                    {
                                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                                        objSTD.PAN = objSt.PAN;
                                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                        objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                        objSTD.ACURN = objSt.ACURN;
                                                        //objSTD.TRNDESC = "Interest Charges";
                                                        objSTD.TRNDESC = "INTEREST CHARGES";
                                                        //objSTD.AMOUNT = dtAcI.Rows[0][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                        objSTD.AMOUNT = dtAcI.Rows[x]["ACCUM_INT_AMOUNT"].ToString();// Rocky 27-03-2017
                                                        objSTD.TRNDATE = trn_Date;
                                                        objSTD.POSTDATE = trn_Date;

                                                        //sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN)" +
                                                        //        " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                        //        "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "')";

                                                        decimal tempIntAmtI = 0;
                                                        decimal tempIntAmt = 0;
                                                        decimal tempTotalIntAmt = 0;
                                                        string st = string.Empty;

                                                        DataTable dt = new DataTable();
                                                        dt = objProvider.ReturnData("select AMOUNT from STATEMENT_DETAILS WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES'  and ACURN = '" + objSTD.ACURN + "' ", ref reply).Tables[0];
                                                        //tempIntAmtI = Convert.ToInt32(dt.Rows[0][0])*(-1);
                                                        if (dt.Rows.Count > 0)
                                                        {
                                                            tempIntAmtI = Convert.ToDecimal(dt.Rows[0][0]) * (-1);
                                                            st = dtAcI.Rows[0][0].ToString();
                                                            tempIntAmt = Convert.ToDecimal(st);
                                                            tempTotalIntAmt = tempIntAmtI + tempIntAmt;

                                                            sql = "Update STATEMENT_DETAILS SET STATEMENTID= '" + objSTD.STATEMENTID + "',CONTRACTNO= '" + objSTD.CONTRACTNO + "',IDCLIENT= '" + objSTD.IDCLIENT + "',PAN= '" + objSTD.PAN + "', " +
                                                        " ACCOUNTNO= '" + objSTD.ACCOUNTNO + "',STATEMENTNO= '" + objSTD.STATEMENTNO + "',TRNDATE= '" + objSTD.TRNDATE + "',POSTDATE= '" + objSTD.POSTDATE + "',TRNDESC= '" + objSTD.TRNDESC + "', " +
                                                        " ACURN= '" + objSTD.ACURN + "',AMOUNT= '" + tempTotalIntAmt + "',APPROVAL= '" + objSTD.APPROVAL + "',AMOUNTSIGN= '" + objSTD.AMOUNTSIGN + "',DE= '" + objSTD.DE + "',P= '" + objSTD.P + "',DOCNO= '" + objSTD.DOCNO + "',FR= '" + objSTD.FR + "',NO= '" + objSTD.NO + "' " +
                                                        " WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND ACURN= '" + objSTD.ACURN + "' AND TRNDESC= 'INTEREST CHARGES' ";


                                                            reply = objProvider.RunQuery(sql);
                                                            if (!reply.Contains("Success"))
                                                                errMsg = reply;
                                                        }
                                                        else
                                                        {
                                                            sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,DE,P,DOCNO,NO)" +
                                                                  " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                                  "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";

                                                            reply = objProvider.RunQuery(sql);
                                                            if (!reply.Contains("Success"))
                                                                errMsg = reply;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }                                
                                }

                                //New View added end
                                
                                if (feesnCharges != 0.00)
                                {
                                    double sumofInterest = Convert.ToDouble(objSt.SUM_INTEREST) + feesnCharges;
                                    sql = "update STATEMENT_INFO SET SUM_INTEREST=" + sumofInterest +
                                    " Where STATEMENTID='" + objSt.STATEMENTID + "' AND PAN='" + objSt.PAN + "' AND ACCOUNTNO='" + objSt.ACCOUNTNO + "' AND STATEMENTNO='" + objSt.STATEMENTNO + "'";

                                    reply = objProvider.RunQuery(sql);
                                    if (!reply.Contains("Success"))
                                        errMsg = reply;
                                }

                            }
                            else
                            {

                                //New View added

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
                                                objSTD.ACURN = objSt.ACURN;
                                                objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                //if (objSTD.CONTRACTNO == dtAcI.Rows[0][1].ToString())
                                                if (objSTD.CONTRACTNO == dtAcI.Rows[x]["CONTRACTNO"].ToString() && objSTD.ACURN == dtAcI.Rows[x]["ACURN"].ToString())// Rocky 27-03-2017
                                                {
                                                    //if (dtAcI.Rows[0][0].ToString() != "0.00")
                                                    if (dtAcI.Rows[x]["ACCUM_INT_AMOUNT"].ToString() != "0.00")// Rocky 27-03-2017
                                                    {
                                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                                        objSTD.PAN = objSt.PAN;
                                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                        objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                        objSTD.ACURN = objSt.ACURN;
                                                        //objSTD.TRNDESC = "Interest Charges";
                                                        objSTD.TRNDESC = "INTEREST CHARGES";
                                                        //objSTD.AMOUNT = dtAcI.Rows[0][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                        objSTD.AMOUNT = dtAcI.Rows[x]["ACCUM_INT_AMOUNT"].ToString();// Rocky 27-03-2017
                                                        objSTD.TRNDATE = objSTD.TRNDATE;
                                                        objSTD.POSTDATE = objSTD.POSTDATE;

                                                        //sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN)" +
                                                        //        " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                        //        "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "')";

                                                        decimal tempIntAmtI = 0;
                                                        decimal tempIntAmt = 0;
                                                        decimal tempTotalIntAmt = 0;
                                                        string st = string.Empty;

                                                        DataTable dt = new DataTable();
                                                        dt = objProvider.ReturnData("select AMOUNT from STATEMENT_DETAILS WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES'  and ACURN = '" + objSTD.ACURN + "' ", ref reply).Tables[0];
                                                        //tempIntAmtI = Convert.ToInt32(dt.Rows[0][0])*(-1);
                                                        if (dt.Rows.Count > 0)
                                                        {
                                                            tempIntAmtI = Convert.ToDecimal(dt.Rows[0][0]) * (-1);
                                                            st = dtAcI.Rows[0][0].ToString();
                                                            tempIntAmt = Convert.ToDecimal(st);
                                                            tempTotalIntAmt = tempIntAmtI + tempIntAmt;

                                                            sql = "Update STATEMENT_DETAILS SET STATEMENTID= '" + objSTD.STATEMENTID + "',CONTRACTNO= '" + objSTD.CONTRACTNO + "',IDCLIENT= '" + objSTD.IDCLIENT + "',PAN= '" + objSTD.PAN + "', " +
                                                                " ACCOUNTNO= '" + objSTD.ACCOUNTNO + "',STATEMENTNO= '" + objSTD.STATEMENTNO + "',TRNDATE= '" + objSTD.TRNDATE + "',POSTDATE= '" + objSTD.POSTDATE + "',TRNDESC= '" + objSTD.TRNDESC + "', " +
                                                                " ACURN= '" + objSTD.ACURN + "',AMOUNT= '" + tempTotalIntAmt + "',APPROVAL= '" + objSTD.APPROVAL + "',AMOUNTSIGN= '" + objSTD.AMOUNTSIGN + "',DE= '" + objSTD.DE + "',P= '" + objSTD.P + "',DOCNO= '" + objSTD.DOCNO + "',NO= '" + objSTD.NO + "' " +
                                                                " WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES' ";


                                                            reply = objProvider.RunQuery(sql);
                                                            if (!reply.Contains("Success"))
                                                                errMsg = reply;
                                                        }
                                                        else
                                                        {
                                                            sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,DE,P,DOCNO,NO)" +
                                                                          " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                                          "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";

                                                            reply = objProvider.RunQuery(sql);
                                                            if (!reply.Contains("Success"))
                                                                errMsg = reply;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                        }

                    }

                    else
                    {


                        //New View add
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
                                        if (objSTD.CONTRACTNO == dtAcI.Rows[0][1].ToString())
                                        {
                                            if (dtAcI.Rows[0][0].ToString() != "0.00")
                                            {
                                                objSTD.STATEMENTID = objSt.STATEMENTID;
                                                objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                objSTD.IDCLIENT = objSt.IDCLIENT;
                                                objSTD.PAN = objSt.PAN;
                                                objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                objSTD.ACURN = objSt.ACURN;
                                                objSTD.TRNDESC = "Interest Charges";
                                                objSTD.AMOUNT = dtAcI.Rows[0][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                objSTD.TRNDATE = objSTD.TRNDATE;
                                                objSTD.POSTDATE = objSTD.POSTDATE;

                                                sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN)" +
                                                        " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                        "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "')";

                                                reply = objProvider.RunQuery(sql);
                                                if (!reply.Contains("Success"))
                                                    errMsg = reply;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //New View add

                }
                catch (Exception ex)
                {
                    errMsg = "Error: " + ex.Message;
                }
            }
            #endregion BDT
        }

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
                            objSt.ADDRESS = dtStatement.Rows[k][j].ToString().Replace("'", "");
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
                            objSt.CITY = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "ZIP")
                        {
                            objSt.ZIP = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "COUNTRY")
                        {
                            objSt.COUNTRY = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "EMAIL")
                        {
                            objSt.EMAIL = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "MOBILE")
                        {
                            objSt.MOBILE = dtStatement.Rows[k][j].ToString().Replace("'", "").Replace("(", "").Replace(")", "").Replace("8800", "880");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "TITLE")
                        {
                            objSt.TITLE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "CLIENT")
                        {
                            objSt.CLIENTNAME = dtStatement.Rows[k][j].ToString().Replace("'", "");
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
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "INSTALL_UNPAID_AMOUNT")
                        {
                            objSt.INSTALL_UNPAID_AMOUNT = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }

                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "INSTALL_MONTH_PAYM")
                        {
                            objSt.INSTALL_MONTH_PAYM = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        //else if (dtStatement.Columns[j].ColumnName.ToUpper() == "INDICATOR")
                        //{
                        //    objSt.INDICATOR = "BDT";
                        //}
                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "INDICATOR")
                        {
                            objSt.INDICATOR = dtStatement.Rows[k][j].ToString();
                        }

                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "JOBTITLE")
                        {
                            objSt.JOBTITLE = dtStatement.Rows[k][j].ToString().Replace("'", ""); ;
                        }

                        else if (dtStatement.Columns[j].ColumnName.ToUpper() == "PPROMOTIONALTEXT")
                        {
                            objSt.PPROMOTIONALTEXT = dtStatement.Rows[k][j].ToString().Replace("'", ""); ;
                        }
                        #endregion
                    }

                    objSt.STM_MSG = txtStmMsg.Text.ToString().Replace("'","''");
                    objSt.STATUS = "1";

                    sql = "Insert into STATEMENT_INFO(STATEMENTID,BANK_CODE,CONTRACTNO,IDCLIENT,PAN,TITLE,CLIENTNAME,STATEMENTNO,ADDRESS,CITY,ZIP,COUNTRY," +
                        "EMAIL,MOBILE,STARTDATE,ENDDATE,NEXT_STATEMENT_DATE,PAYMENT_DATE,STATEMENT_DATE,ACCOUNTNO,ACURN,SBALANCE,ACURC,EBALANCE,AVAIL_CRD_LIMIT," +
                        "AVAIL_CASH_LIMIT,SUM_WITHDRAWAL,SUM_INTEREST,OVLFEE_AMOUNT,OVDFEE_AMOUNT,SUM_REVERSE,SUM_CREDIT,SUM_OTHER,SUM_PURCHASE," +
                        "MIN_AMOUNT_DUE,CASH_LIMIT,CRD_LIMIT,STM_MSG,STATUS,INSTALL_UNPAID_AMOUNT,INSTALL_MONTH_PAYM,INDICATOR,JOBTITLE,PPROMOTIONALTEXT) VALUES('" + objSt.STATEMENTID + "'," +
                        "'" + objSt.BANK_CODE + "','" + objSt.CONTRACTNO + "','" + objSt.IDCLIENT + "','" + objSt.PAN + "','" + objSt.TITLE + "','" + objSt.CLIENTNAME + "','" + objSt.STATEMENTNO + "'," +
                        "'" + objSt.ADDRESS + "','" + objSt.CITY + "','" + objSt.ZIP + "','" + objSt.COUNTRY + "','" + objSt.EMAIL + "','" + objSt.MOBILE + "','" + objSt.STARTDATE + "','" + objSt.ENDDATE + "'," +
                        "'" + objSt.NEXT_STATEMENT_DATE + "','" + objSt.PAYMENT_DATE + "','" + objSt.STATEMENT_DATE + "','" + objSt.ACCOUNTNO + "','" + objSt.ACURN + "'," +
                        "'" + objSt.SBALANCE + "','" + objSt.ACURC + "','" + objSt.EBALANCE + "','" + objSt.AVAIL_CRD_LIMIT + "','" + objSt.AVAIL_CASH_LIMIT + "'," +
                        "'" + objSt.SUM_WITHDRAWAL + "','" + objSt.SUM_INTEREST + "','" + objSt.OVLFEE_AMOUNT + "','" + objSt.OVDFEE_AMOUNT + "','" + objSt.SUM_REVERSE + "'," +
                        "'" + objSt.SUM_CREDIT + "','" + objSt.SUM_OTHER + "','" + objSt.SUM_PURCHASE + "','" + objSt.MIN_AMOUNT_DUE + "','" + objSt.CASH_LIMIT + "'," +
                        "'" + objSt.CRD_LIMIT + "','" + objSt.STM_MSG + "','" + objSt.STATUS + "','" + 0 + "','" + 0 + "','" + objSt.INDICATOR + "','" + objSt.JOBTITLE + "','" + objSt.PPROMOTIONALTEXT + "')";

                    reply = objProvider.RunQuery(sql);
                    //DataTable dtOperation = dsStatement.Tables["Operation"];
                    if (dtOperation != null && dtOperation.Columns.Contains("ACCOUNT"))
                    {
                        if (dtOperation.Rows.Count > 0)
                        {

                            DataRow[] dr = dtOperation.Select("STATEMENTNO='" + objSt.STATEMENTNO + "' AND ACCOUNT='" + objSt.ACCOUNTNO + "'");
                            if (dr.Length > 0)
                            {
                                double feesnCharges = 0.00;
                                string trn_Date = string.Empty;

                                for (int l = 0; l < dr.Length; l++)
                                {
                                    #region setting properties values
                                    if (!dr[l]["D"].ToString().Contains("Charge interest"))
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
                                        {
                                            //objSTD.TRNDATE = dr[l]["TD"].ToString();
                                            // objOp.OpDate = getNumberFormat(objOp.OpDate); 
                                            if (dr[l]["TD"].ToString() == "" || dr[l]["TD"].ToString() == null)
                                                objSTD.TRNDATE = "";
                                            else
                                                objSTD.TRNDATE = getNumberFormat(dr[l]["TD"].ToString());
                                        }

                                        if (dr[l].Table.Columns.Contains("OD"))                                            
                                        {
                                            //objSTD.POSTDATE = dr[l]["OD"].ToString();
                                            // objSTD.POSTDATE = getNumberFormat(dr[l]["OD"].ToString());
                                            if (dr[l]["OD"].ToString() == "" || dr[l]["OD"].ToString() == null)
                                                objSTD.POSTDATE = "";
                                            else
                                                objSTD.POSTDATE = getNumberFormat(dr[l]["OD"].ToString());
                                        }


                                        if (dr[l].Table.Columns.Contains("ACURN"))
                                            objSTD.ACURN = dr[l]["ACURN"].ToString();

                                        if (dr[l].Table.Columns.Contains("OC"))
                                            objSTD.OC = dr[l]["OC"].ToString();

                                        if (dr[l].Table.Columns.Contains("P"))
                                            objSTD.P = dr[l]["P"].ToString();

                                        if (dr[l].Table.Columns.Contains("DOCNO"))
                                            objSTD.DOCNO = dr[l]["DOCNO"].ToString();

                                        if (dr[l].Table.Columns.Contains("DE"))
                                            objSTD.DE = dr[l]["DE"].ToString();


                                        if (dr[l].Table.Columns.Contains("NO"))
                                            objSTD.NO = dr[l]["NO"].ToString();

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
                                        if (dr[l].Table.Columns.Contains("OC"))
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
                                            objSTD.OC = ""; //dr[l]["OC"].ToString();
                                        if (dr[l].Table.Columns.Contains("OC"))
                                        {
                                            if (dr[l]["OA"].ToString() == "" || dr[l]["OA"].ToString() == null)
                                                objSTD.ORGAMOUNT = "0.00";
                                            else
                                                objSTD.ORGAMOUNT = dr[l]["OA"].ToString();
                                        }
                                        else objSTD.ORGAMOUNT = "0.00";

                                        //Remmove Terminal Name when Fee and VAT Impose
                                        //Sum Charges amount with Fees & Charges. 
                                        if ((!dr[l]["D"].ToString().ToUpper().Contains("FEE")) || (dr[l]["D"].ToString() != "Charge interest for Installment") || (dr[l]["D"].ToString() != "Credit Shield Premium") || (dr[l]["D"].ToString() != "Monthly Installment"))
                                        {
                                            if (dr[l].Table.Columns.Contains("TL") && dr[l]["DE"].ToString().ToUpper() != "FEE")
                                                objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                            else
                                                objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                        }
                                        else
                                        {
                                            objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                            feesnCharges = feesnCharges + Convert.ToDouble(dr[l]["A"].ToString().Replace("-", ""));
                                            if (dr[l].Table.Columns.Contains("OD"))
                                                objSTD.TRNDATE = dr[l]["OD"].ToString();
                                        }

                                        //if (objSTD.TRNDESC.Contains("Credit cash deposit"))
                                        //{
                                        //    objSTD.TRNDESC = "PAYMENT RECEIVED (THANK YOU)";
                                        //    objSTD.TRNDATE = dr[l]["OD"].ToString();
                                        //}
                                        
                                        //else if ((objSTD.TRNDESC.Contains("[MANUAL_TXN[POS]]")) || (objSTD.TRNDESC.Contains("[MANUAL_TXN[POS-R]]")) || (objSTD.TRNDESC.Contains("CREDIT ADJUSTMENT")) || (objSTD.TRNDESC.Contains("DEBIT ADJUSTMENT")))
                                        //{
                                        //    objSTD.TRNDESC = dr[l]["FR"].ToString();
                                        //}
                                        var Entrylist = new List<String>() { "Credit cash deposit", "[MANUAL_TXN[POS]]", "[MANUAL_TXN[POS-R]]", "CREDIT ADJUSTMENT", "DEBIT ADJUSTMENT", "Credit acct", "CASH BACK[BDT]", "ATM INTEREST (REVERSE)", "CARD FEE(REVERSE)", "CASH ADVANCE FEE(REVERSE)", "LATE PAYMENT FEE(REVERSE)", "CASH DEPOSIT (REVERSE)", "SALES_SLIP_RET_FEE", "STATEMENT REPRINT FEE", "VAT", "REVERSAL-POS PURCHASE", "OVER LIMIT FEE(REVERSE)", "PIN FEE(REVERSE)", "POS INTEREST(REVERSE)", "CARD REPLACEMENT FEE(REVERSE)", "ATM INTEREST", "BAL_TRNS", "FUND TRANSFER", "INTEREST CHARGE", "POS INTEREST" };
                                        if (Entrylist.Contains(objSTD.TRNDESC.Trim(), StringComparer.OrdinalIgnoreCase))
                                        {
                                            if (dr[l]["FR"].ToString() == "" || dr[l]["FR"].ToString() == null)
                                                objSTD.TRNDESC = dr[l]["TRNDESC"].ToString().Replace("'", "''");
                                            else
                                                objSTD.TRNDESC = dr[l]["FR"].ToString().Replace("'", "''");
                                        }
                                        if (dr[l].Table.Columns.Contains("APPROVAL"))
                                        {
                                            objSTD.APPROVAL = dr[l]["APPROVAL"].ToString().Replace("'", "");

                                            if (dr[l]["APPROVAL"].ToString() != "" && objSTD.TRNDATE == "")
                                            {
                                                objSTD.TRNDATE = dr[l]["OD"].ToString();
                                            }
                                        }
                                        if (dr[l].Table.Columns.Contains("FR"))
                                            objSTD.FR = dr[l]["FR"].ToString().Replace("'", "''");
                                        //objSTD.AMOUNTSIGN = dr[l]["AMOUNTSIGN"].ToString();

                                        if (objSTD.FR == "P 10")
                                        {
                                            objSTD.TRNDESC = "Purchase" + " " + dr[l]["TL"].ToString().Replace("'", "''");

                                        }
                                        else if (objSTD.FR == "A 10")
                                        {
                                            objSTD.TRNDESC = "Cash Withdrawal" + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                        }

                                        sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,OC,ORGAMOUNT,AMOUNTSIGN,APPROVAL,FR,P,DOCNO,NO,DE)" +
                                            " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                            "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.OC + "','" + objSTD.ORGAMOUNT + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.APPROVAL + "','" + objSTD.FR + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "','" + objSTD.DE + "')";

                                        reply = objProvider.RunQuery(sql);
                                        if (!reply.Contains("Success"))
                                            errMsg = reply;
                                    }
                                    else if (dr[l]["D"].ToString().Contains("Charge interest"))
                                    {
                                        trn_Date = dr[l]["OD"].ToString();
                                    }

                                    if (dr[l]["D"].ToString() == "Charge interest for Installment")
                                    {
                                        StatementDetails objSTD = new StatementDetails();
                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                        objSTD.PAN = objSt.PAN;
                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                        objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                        objSTD.ACURN = objSt.ACURN;
                                        objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                        //objSTD.AMOUNT = "-" + objSt.SUM_INTEREST;//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                        objSTD.AMOUNT = dr[l]["A"].ToString();
                                        // objSTD.TRNDATE = trn_Date;
                                        //objSTD.POSTDATE = trn_Date;
                                        objSTD.TRNDATE = dr[l]["TD"].ToString();
                                        objSTD.POSTDATE = dr[l]["OD"].ToString();

                                        sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN)" +
                                                " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "')";

                                        reply = objProvider.RunQuery(sql);
                                        if (!reply.Contains("Success"))
                                            errMsg = reply;
                                    }


                                    #endregion
                                }

                                if (objSt.SUM_INTEREST != "0.00")
                                {
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

                                    sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN)" +
                                            " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                            "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "')";

                                    reply = objProvider.RunQuery(sql);
                                    if (!reply.Contains("Success"))
                                        errMsg = reply;
                                }

                                //New View added

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
                                                objSTD.ACURN = objSt.ACURN;
                                                objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                //if (objSTD.CONTRACTNO == dtAcI.Rows[0][1].ToString())
                                                if (objSTD.CONTRACTNO == dtAcI.Rows[x]["CONTRACTNO"].ToString() && objSTD.ACURN == dtAcI.Rows[x]["ACURN"].ToString())// Rocky 27-03-2017
                                                {
                                                    //if (dtAcI.Rows[0][0].ToString() != "0.00")
                                                    if (dtAcI.Rows[x]["ACCUM_INT_AMOUNT"].ToString() != "0.00")// Rocky 27-03-2017
                                                    {
                                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                                        objSTD.PAN = objSt.PAN;
                                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                        objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                        objSTD.ACURN = objSt.ACURN;
                                                        //objSTD.TRNDESC = "Interest Charges";
                                                        objSTD.TRNDESC = "INTEREST CHARGES";
                                                        //objSTD.AMOUNT = dtAcI.Rows[0][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                        objSTD.AMOUNT = dtAcI.Rows[x]["ACCUM_INT_AMOUNT"].ToString();// Rocky 27-03-2017
                                                        objSTD.TRNDATE = trn_Date;
                                                        objSTD.POSTDATE = trn_Date;

                                                        //sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN)" +
                                                        //        " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                        //        "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "')";

                                                        decimal tempIntAmtI = 0;
                                                        decimal tempIntAmt = 0;
                                                        decimal tempTotalIntAmt = 0;
                                                        string st = string.Empty;

                                                        DataTable dt = new DataTable();
                                                        dt = objProvider.ReturnData("select AMOUNT from STATEMENT_DETAILS WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES'  and ACURN = '" + objSTD.ACURN + "' ", ref reply).Tables[0];
                                                        //tempIntAmtI = Convert.ToInt32(dt.Rows[0][0])*(-1);
                                                        if (dt.Rows.Count > 0)
                                                        {
                                                            tempIntAmtI = Convert.ToDecimal(dt.Rows[0][0]) * (-1);
                                                            st = dtAcI.Rows[0][0].ToString();
                                                            tempIntAmt = Convert.ToDecimal(st);
                                                            tempTotalIntAmt = tempIntAmtI + tempIntAmt;

                                                            sql = "Update STATEMENT_DETAILS SET STATEMENTID= '" + objSTD.STATEMENTID + "',CONTRACTNO= '" + objSTD.CONTRACTNO + "',IDCLIENT= '" + objSTD.IDCLIENT + "',PAN= '" + objSTD.PAN + "', " +
                                                        " ACCOUNTNO= '" + objSTD.ACCOUNTNO + "',STATEMENTNO= '" + objSTD.STATEMENTNO + "',TRNDATE= '" + objSTD.TRNDATE + "',POSTDATE= '" + objSTD.POSTDATE + "',TRNDESC= '" + objSTD.TRNDESC + "', " +
                                                        " ACURN= '" + objSTD.ACURN + "',AMOUNT= '" + tempTotalIntAmt + "',APPROVAL= '" + objSTD.APPROVAL + "',AMOUNTSIGN= '" + objSTD.AMOUNTSIGN + "',DE= '" + objSTD.DE + "',P= '" + objSTD.P + "',DOCNO= '" + objSTD.DOCNO + "',FR= '" + objSTD.FR + "',NO= '" + objSTD.NO + "' " +
                                                        " WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND ACURN= '" + objSTD.ACURN + "' AND TRNDESC= 'INTEREST CHARGES' ";


                                                            reply = objProvider.RunQuery(sql);
                                                            if (!reply.Contains("Success"))
                                                                errMsg = reply;
                                                        }
                                                        else
                                                        {
                                                            sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,DE,P,DOCNO,NO)" +
                                                                  " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                                  "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";

                                                            reply = objProvider.RunQuery(sql);
                                                            if (!reply.Contains("Success"))
                                                                errMsg = reply;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                //New View added end

                                if (feesnCharges != 0.00)
                                {
                                    double sumofInterest = Convert.ToDouble(objSt.SUM_INTEREST) + feesnCharges;
                                    sql = "update STATEMENT_INFO SET SUM_INTEREST=" + sumofInterest +
                                    " Where STATEMENTID='" + objSt.STATEMENTID + "' AND PAN='" + objSt.PAN + "' AND ACCOUNTNO='" + objSt.ACCOUNTNO + "' AND STATEMENTNO='" + objSt.STATEMENTNO + "'";

                                    reply = objProvider.RunQuery(sql);
                                    if (!reply.Contains("Success"))
                                        errMsg = reply;
                                }

                            }
                            else
                            {

                                //New View added

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
                                                objSTD.ACURN = objSt.ACURN;
                                                objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                //if (objSTD.CONTRACTNO == dtAcI.Rows[0][1].ToString())
                                                if (objSTD.CONTRACTNO == dtAcI.Rows[x]["CONTRACTNO"].ToString() && objSTD.ACURN == dtAcI.Rows[x]["ACURN"].ToString())// Rocky 27-03-2017
                                                {
                                                    //if (dtAcI.Rows[0][0].ToString() != "0.00")
                                                    if (dtAcI.Rows[x]["ACCUM_INT_AMOUNT"].ToString() != "0.00")// Rocky 27-03-2017
                                                    {
                                                        objSTD.STATEMENTID = objSt.STATEMENTID;
                                                        objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                        objSTD.IDCLIENT = objSt.IDCLIENT;
                                                        objSTD.PAN = objSt.PAN;
                                                        objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                        objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                        objSTD.ACURN = objSt.ACURN;
                                                        //objSTD.TRNDESC = "Interest Charges";
                                                        objSTD.TRNDESC = "INTEREST CHARGES";
                                                        //objSTD.AMOUNT = dtAcI.Rows[0][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                        objSTD.AMOUNT = dtAcI.Rows[x]["ACCUM_INT_AMOUNT"].ToString();// Rocky 27-03-2017
                                                        objSTD.TRNDATE = objSTD.TRNDATE;
                                                        objSTD.POSTDATE = objSTD.POSTDATE;

                                                        //sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN)" +
                                                        //        " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                        //        "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "')";

                                                        decimal tempIntAmtI = 0;
                                                        decimal tempIntAmt = 0;
                                                        decimal tempTotalIntAmt = 0;
                                                        string st = string.Empty;

                                                        DataTable dt = new DataTable();
                                                        dt = objProvider.ReturnData("select AMOUNT from STATEMENT_DETAILS WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES'  and ACURN = '" + objSTD.ACURN + "' ", ref reply).Tables[0];
                                                        //tempIntAmtI = Convert.ToInt32(dt.Rows[0][0])*(-1);
                                                        if (dt.Rows.Count > 0)
                                                        {
                                                            tempIntAmtI = Convert.ToDecimal(dt.Rows[0][0]) * (-1);
                                                            st = dtAcI.Rows[0][0].ToString();
                                                            tempIntAmt = Convert.ToDecimal(st);
                                                            tempTotalIntAmt = tempIntAmtI + tempIntAmt;

                                                            sql = "Update STATEMENT_DETAILS SET STATEMENTID= '" + objSTD.STATEMENTID + "',CONTRACTNO= '" + objSTD.CONTRACTNO + "',IDCLIENT= '" + objSTD.IDCLIENT + "',PAN= '" + objSTD.PAN + "', " +
                                                                " ACCOUNTNO= '" + objSTD.ACCOUNTNO + "',STATEMENTNO= '" + objSTD.STATEMENTNO + "',TRNDATE= '" + objSTD.TRNDATE + "',POSTDATE= '" + objSTD.POSTDATE + "',TRNDESC= '" + objSTD.TRNDESC + "', " +
                                                                " ACURN= '" + objSTD.ACURN + "',AMOUNT= '" + tempTotalIntAmt + "',APPROVAL= '" + objSTD.APPROVAL + "',AMOUNTSIGN= '" + objSTD.AMOUNTSIGN + "',DE= '" + objSTD.DE + "',P= '" + objSTD.P + "',DOCNO= '" + objSTD.DOCNO + "',NO= '" + objSTD.NO + "' " +
                                                                " WHERE STATEMENTNO= '" + objSTD.STATEMENTNO + "' AND CONTRACTNO= '" + objSTD.CONTRACTNO + "' AND TRNDESC= 'INTEREST CHARGES' ";


                                                            reply = objProvider.RunQuery(sql);
                                                            if (!reply.Contains("Success"))
                                                                errMsg = reply;
                                                        }
                                                        else
                                                        {
                                                            sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN,DE,P,DOCNO,NO)" +
                                                                          " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                                          "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "','" + objSTD.DE + "','" + objSTD.P + "','" + objSTD.DOCNO + "','" + objSTD.NO + "')";

                                                            reply = objProvider.RunQuery(sql);
                                                            if (!reply.Contains("Success"))
                                                                errMsg = reply;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                        }

                    }

                    else
                    {


                        //New View add
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
                                        if (objSTD.CONTRACTNO == dtAcI.Rows[0][1].ToString())
                                        {
                                            if (dtAcI.Rows[0][0].ToString() != "0.00")
                                            {
                                                objSTD.STATEMENTID = objSt.STATEMENTID;
                                                objSTD.CONTRACTNO = objSt.CONTRACTNO;
                                                objSTD.IDCLIENT = objSt.IDCLIENT;
                                                objSTD.PAN = objSt.PAN;
                                                objSTD.STATEMENTNO = objSt.STATEMENTNO;
                                                objSTD.ACCOUNTNO = objSt.ACCOUNTNO;
                                                objSTD.ACURN = objSt.ACURN;
                                                objSTD.TRNDESC = "Interest Charges";
                                                objSTD.AMOUNT = dtAcI.Rows[0][0].ToString();//.PadLeft(objSt.SUM_INTEREST.Length+1,'-');
                                                objSTD.TRNDATE = objSTD.TRNDATE;
                                                objSTD.POSTDATE = objSTD.POSTDATE;

                                                sql = "Insert into STATEMENT_DETAILS(STATEMENTID,CONTRACTNO,IDCLIENT,PAN,ACCOUNTNO,STATEMENTNO,TRNDATE,POSTDATE,TRNDESC,ACURN,AMOUNT,APPROVAL,AMOUNTSIGN)" +
                                                        " VALUES('" + objSTD.STATEMENTID + "','" + objSTD.CONTRACTNO + "','" + objSTD.IDCLIENT + "','" + objSTD.PAN + "','" + objSTD.ACCOUNTNO + "','" + objSTD.STATEMENTNO + "','" + objSTD.TRNDATE + "'," +
                                                        "'" + objSTD.POSTDATE + "','" + objSTD.TRNDESC + "','" + objSTD.ACURN + "','" + objSTD.AMOUNT + "','" + objSTD.APPROVAL + "','" + objSTD.AMOUNTSIGN + "')";

                                                reply = objProvider.RunQuery(sql);
                                                if (!reply.Contains("Success"))
                                                    errMsg = reply;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //New View add

                }
                catch (Exception ex)
                {
                    errMsg = "Error: " + ex.Message;
                }
            }
            #endregion USD
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
                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("dd.MM.yyyy hh:mm:ss") + " : Error: " + ex.Message });
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
        private string getNumberFormat1(string vDate)
        {
            string[] omitSpace = vDate.Split(' ');
            string[] date = omitSpace[0].Split('/');
            DateTime dt = new DateTime(Int32.Parse(date[2]), Int32.Parse(date[0]), Int32.Parse(date[1]));
            string formatedDate = string.Format("{0:dd-MMM-yyyy}", dt);
            return formatedDate;
        }

    }
}
