using System;
using System.Collections.Generic;
using System.Common;
using System.Configuration;
using System.Connection;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Threading;
using System.Windows.Forms;
//using CrystalDecisions.Shared;
using FlexiStar.Utilities;
using FlexiStar.Utilities.EncryptionEngine;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using QCash.EStatement.NBL.App_Code;
using QCash.EStatement.NBL.Reports;
using StatementGenerator.App_Code;
using ThoughtWorks.QRCode.Codec;
using CrystalDecisions.Shared;

namespace StatementGenerator
{
    public partial class StatementGenerator : Form
    {
        #region Declaration
        private ConnectionStringBuilder ConStr = null;
        private SqlDbProvider objProvider = null;
        //
        delegate void SetTextCallback(string text);
        private SetTextCallback _addText = null;
        //
        private string Bank_Code = string.Empty;
        string prePan = string.Empty;
        string preDoc = string.Empty;

        private string _QRPath = string.Empty;
        private int _QRSize = 70;
        private int _QRPositionX = 300;
        private int _QRPositionY = 30;
        private int _QRPage = 30;


        private string _XMLProcessedPath = string.Empty;
        private string _XMLSourcePath = string.Empty;
        private string _LogPath = string.Empty;
        private string _EmailResultPath = string.Empty;
        private string _AdditionalAttachment = string.Empty;
        private string _Mail = string.Empty;


        private string _StatementResultPath = string.Empty;
        private string _StatementSourcePath = string.Empty;
        private string _statementProcessPath = string.Empty;
        private string _StatementLogPath = string.Empty;
        private string _CombineStatementPath = string.Empty;

        private string StmDate = string.Empty;
        private string stmMessage = string.Empty;
        private string _xmlName = string.Empty;


        Thread tdGenerate = null;
        Thread tdSendMail = null;

        private string _fiid = string.Empty;
        int pdfCount = 0;

        #endregion

        #region Constructer
        public StatementGenerator(string fiid)
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

        #endregion


        private void AddQRonPDF(string pdfFile, String QR, int page)
        {
            PdfSharp.Pdf.PdfDocument inputPDFDocument = PdfReader.Open(pdfFile, PdfDocumentOpenMode.Modify);
            if (page > 0)
            {
                PdfPage pp = inputPDFDocument.Pages[page - 1];
                using (XGraphics gfx = XGraphics.FromPdfPage(pp))
                {
                    XImage image = XImage.FromFile(_QRPath + QR + ".jpg");
                    gfx.DrawImage(image, _QRPositionX, _QRPositionY, _QRSize, _QRSize);
                    image.Dispose();

                }
            }
            else if (page < 0)
            {
                PdfPage pp = inputPDFDocument.Pages[inputPDFDocument.PageCount - 1];
                using (XGraphics gfx = XGraphics.FromPdfPage(pp))
                {
                    XImage image = XImage.FromFile(_QRPath + QR + ".jpg");
                    gfx.DrawImage(image, _QRPositionX, _QRPositionY, _QRSize, _QRSize);
                    image.Dispose();

                }
            }
            else
            {
                for (int i = 0; i < inputPDFDocument.PageCount; i++)
                {
                    PdfPage pp = inputPDFDocument.Pages[i];
                    using (XGraphics gfx = XGraphics.FromPdfPage(pp))
                    {
                        XImage image = XImage.FromFile(_QRPath + QR + ".jpg");
                        gfx.DrawImage(image, _QRPositionX, _QRPositionY, _QRSize, _QRSize);
                        image.Dispose();

                    }
                }
            }
            System.IO.File.Delete(pdfFile);
            inputPDFDocument.Save(pdfFile);
        }

        private void MergeMultiplePDFIntoSinglePDF(string path, string[] pdfFiles, DataTable dt, string xmlName)
        {

            PdfSharp.Pdf.PdfDocument document = new PdfSharp.Pdf.PdfDocument();

            int lc = 0;
            foreach (string pdfFile in pdfFiles)
            {
                AddQRonPDF(pdfFile, dt.Rows[lc]["CONTRACTNO"].ToString(), _QRPage);
                lc++;
                PdfSharp.Pdf.PdfDocument inputPDFDocument = PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import);
                document.Version = inputPDFDocument.Version;

                // foreach (PdfPage page in inputPDFDocument.Pages)
                for (int i = 0; i < inputPDFDocument.Pages.Count; i++)
                {
                    document.AddPage(inputPDFDocument.Pages[i]);
                }
                // When document is add in pdf document remove file from folder  
                System.IO.File.Delete(pdfFile);
                //  System.IO.File.Delete(_CombineStatementPath);

            }
            // Set font for paging  
            XFont font = new XFont("Verdana", 7);
            XBrush brush = XBrushes.Black;
            // Create variable that store page count  
            string noPages = document.Pages.Count.ToString();
            string x = string.Empty;
            // Set for loop of document page count and set page number using DrawString function of PdfSharp  

            for (int i = 0; i < document.Pages.Count; ++i)
            {
                PdfPage page = document.Pages[i];
                // Make a layout rectangle.  
                //XRect layoutRectangle = new XRect(240 /*X*/ , page.Height - font.Height - 10 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                XRect layoutRectangle = new XRect(20 /*X*/ , 90 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                {
                    gfx.DrawString("P- " + (i + 1).ToString() + " " + x, font, brush, layoutRectangle, XStringFormats.Center);

                }
            }

            document.Options.CompressContentStreams = true;
            document.Options.NoCompression = false;
            // In the final stage, all documents are merged and save in your output file path.  
            document.Save(path + xmlName);
            // document.Save(@"D:\XML_For_Statement\Statement" + "\\" + xmlName);

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
            bool eStatementChecked = rbtneStatement.Checked;
            bool StatementChecked = rbtnStatement.Checked;
            //Rocky 04th Feb 2019

            if (eStatementChecked)
            {
                if (txtEmailSubject.Text.Length > 100)
                {
                    MessageBox.Show("Email subject should be within 100 character...");
                }
                //else if (txtEmailBody.Text.Length > 1000)
                //{
                //    MessageBox.Show("Email body should be within 1000 character...");
                //}
                //else if (txtStmMsg.Text.Length > 500)
                //{
                //    MessageBox.Show("Message should be within 500 character...");
                //}
                else
                {
                    //stmMessage = txtStmMsg.Text;
                    btnLoad.Enabled = false;
                    tdGenerate = new Thread(new ThreadStart(GenerateEStatement));
                    tdGenerate.IsBackground = true;
                    tdGenerate.Start();
                }
            }
            else if (StatementChecked)
            {
                //stmMessage = txtStmMsg.Text;
                btnLoad.Enabled = false;
                btnSendMail.Enabled = false;
                tdGenerate = new Thread(new ThreadStart(GenerateEStatement));
                tdGenerate.IsBackground = true;
                tdGenerate.Start();
            }


        }

        private void GenerateEStatement()
        {
            bool eStatementChecked = rbtneStatement.Checked;
            bool StatementChecked = rbtnStatement.Checked;

            if (eStatementChecked)
            {

                if (txtEmailSubject.Text != "")
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
                    MessageBox.Show("Set Email Subject", "Warning !!!");
                }
            }
            else if (StatementChecked)
            {

                string reply = string.Empty;
                EStatementManager.Instance().ArchiveStatement(ref reply);

                if (reply.Contains("Error"))
                {
                    MsgLogWriter objLW = new MsgLogWriter();
                    objLW.logTrace(_StatementLogPath, "Statement.log", reply);
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + reply });
                }
                else if (reply == "Success")
                {
                    MsgLogWriter objLW = new MsgLogWriter();
                    objLW.logTrace(_StatementLogPath, "Statement.log", "Successfully archive previous data !!!");
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + "Successfully archive previous data !!!" });

                    ProcessData();
                }


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
                    StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy");
                else StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy");

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
                                            mail.Subject = objESList[i].MAILSUBJECT + " " + objESList[i].PAN_NUMBER.ToString().Substring(0, 6) + "*********" + objESList[i].PAN_NUMBER.ToString().Substring(15, 1);
                                            mail.Body = objESList[i].MAILBODY;
                                            mail.To.Add(objESList[i].MAILADDRESS.Trim());
                                            System.Net.Mail.Attachment attachment;
                                            attachment = new System.Net.Mail.Attachment(objESList[i].FILE_LOCATION);
                                            mail.Attachments.Add(attachment);
                                            //mail.Attachments.Add(attachment);
                                            //=-=-=-=-=-=-=-=-=-=-=-=-=--=--=-=-=-=-=-=
                                            _Mail = ConfigurationManager.AppSettings["Mail"].ToString();
                                            StreamReader reader = new StreamReader(_Mail + @"\\Template.html");
                                            string readFile = reader.ReadToEnd();
                                            string myString = "";

                                            string mon = DateTime.Now.ToString("MMMMM yyyy");
                                            myString = readFile.Replace("##month##", mon);
                                            mail.Body = myString;

                                            mail.IsBodyHtml = true;
                                            //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=

                                            /*   **** imange in email body code ****
                                             var contentID = "Image";
                                            var inlineLogo = new Attachment(@"D:\Image2.jpg");  --change_here
                                            inlineLogo.ContentId = contentID;
                                            inlineLogo.ContentDisposition.Inline = true;
                                            inlineLogo.ContentDisposition.DispositionType = DispositionTypeNames.Inline;

                                            mail.IsBodyHtml = true;
                                            mail.Attachments.Add(inlineLogo);
                                            mail.Body = "<htm><body> <img src=\"cid:" + contentID + "\"> </body></html>";

                                           */


                                            //-----------------------------------------------------
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
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : There is no Estatement has generate on that statement date." });
                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : There is no Estatement has generate on that statement date.");

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
            bool eStatementChecked = rbtneStatement.Checked;
            bool StatementChecked = rbtnStatement.Checked;
            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
            string reply = string.Empty;
            MsgLogWriter objLW = new MsgLogWriter();

            DataTable dtCardbdt = new DataTable();


            //  dtCardbdt = objProvider.ReturnData("select * from (select distinct EMAIL,CONTRACTNO,Statementno,accountno,pan,idclient,client,StDate from Qry_Card_Account where CONTRACTNO in (select Distinct(CONTRACTNO) from statement_DUAL)) as t1 order by CAST( Statementno as int)", ref reply).Tables[0];// where Curr='BDT'
            dtCardbdt = objProvider.ReturnData("select * from Qry_Card_Account where Curr='BDT' order by CAST( Statementno as int) ASC", ref reply).Tables[0];// where Curr='BDT'


            if (dtCardbdt.Rows.Count > 0)
            {
                if (eStatementChecked)
                {
                    txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Estatement." });//Processing Estatement BDT
                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Estatement.");//Processing Estatement BDT.

                    //Process pdf for BDT
                    ProcessStatementBDT(dtCardbdt);
                }
                else if (StatementChecked)
                {

                    try
                    {
                        if (Directory.Exists(_QRPath))
                            System.IO.Directory.Delete(_QRPath, true);
                    }
                    catch (Exception ex)
                    {

                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                        //objLW.logTrace(_StatementLogPath, "Statement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + ex.Message);

                    }
                    try
                    {
                        if (Directory.Exists(_StatementResultPath))
                            System.IO.Directory.Delete(_StatementResultPath, true);

                    }
                    catch (Exception ex)
                    {

                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                        //objLW.logTrace(_StatementLogPath, "Statement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + ex.Message);

                    }
                    txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Statement." });//Processing Statement BDT
                    objLW.logTrace(_StatementLogPath, "Statement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Processing Statement.");//Processing Statement BDT.

                    //Process pdf for BDT
                    ProcessStatementBDT(dtCardbdt);
                }
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

        public static void MergePDFs(string targetPath, params string[] pdfs)
        {
            using (PdfSharp.Pdf.PdfDocument targetDoc = new PdfSharp.Pdf.PdfDocument())
            {
                foreach (string pdf in pdfs)
                {
                    using (PdfSharp.Pdf.PdfDocument pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
                    {
                        for (int i = 0; i < pdfDoc.PageCount; i++)
                        {
                            targetDoc.AddPage(pdfDoc.Pages[i]);
                        }
                    }
                }
                targetDoc.Save(targetPath);
            }
        }
        //Process pdf for BDT
        private void ProcessStatementBDT(DataTable dtCards)
        {
            DataSet ds = new DataSet();
            DataTable stmdt = new DataTable();

            string reply = string.Empty;
            string filePath = string.Empty;
            string filePathQR = string.Empty;
            string filePathforwithoutEmail = string.Empty;
            string fileName = string.Empty;
            string fileNameQRImage = string.Empty;
            //string fileName[]=new sting();
            string[] fileNameArray = { };
            bool eStatementChecked = rbtneStatement.Checked;
            bool StatementChecked = rbtnStatement.Checked;
            string sql = string.Empty;
            int count = 0;



            ConStr = new ConnectionStringBuilder(1);
            objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);
            ds = objProvider.ReturnData("select * from statement_DUAL ORDER BY  CAST( STATEMENTNO as int) ASC", ref reply);

            MsgLogWriter objLW = new MsgLogWriter();
            #region eStatementChecked
            if (eStatementChecked)
            {
                if (ds != null)
                {
                    if (ds.Tables.Count > 0)
                    {
                        if (ds.Tables[0].Rows.Count > 0)
                        {

                            DataTable dtAllRows = ds.Tables[0];

                            FileInfo objFile = new FileInfo(_EmailResultPath);

                            if (!Directory.Exists(_EmailResultPath))
                                Directory.CreateDirectory(_EmailResultPath);


                            filePath = _EmailResultPath + "\\EStatement of " + System.DateTime.Now.ToString("ddMMyyyy") + "_WithEmail";
                            filePathforwithoutEmail = _EmailResultPath + "\\EStatement of " + System.DateTime.Now.ToString("ddMMyyyy") + "_WithoutEmail";

                            if (!Directory.Exists(filePath))
                            {
                                Directory.CreateDirectory(filePath);
                            }

                            if (!Directory.Exists(filePathforwithoutEmail))
                            {
                                Directory.CreateDirectory(filePathforwithoutEmail);
                            }

                            // DataRow dr;

                            txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement." });
                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + "Total " + dtCards.Rows.Count.ToString() + " record has been found to process Estatement.");

                            #region Forloop
                            for (int j = 0; j < dtCards.Rows.Count; j++)//dtCards.Rows.Count
                            {

                                if (dtCards.Rows[j]["EMAIL"].ToString().Trim() != "")
                                {
                                    if (IsValid(dtCards.Rows[j]["EMAIL"].ToString().Trim()))
                                    {
                                        #region try
                                        try
                                        {
                                            pdfCount = pdfCount + 1;
                                            stmdt = new DataTable();
                                            stmdt = objProvider.ReturnData("select * from statement_DUAL where CONTRACTNO='" + dtCards.Rows[j]["CONTRACTNO"].ToString() + "'  ORDER BY SL,[AutoID] ASC", ref reply).Tables[0];

                                            if (stmdt.Rows.Count > 0)
                                            {



                                                EStatement objst = new EStatement();
                                                objst.SetDataSource(stmdt);



                                                /*  string Bin = dtCards.Rows[j]["pan"].ToString().Substring(0, 6);
                                                  fileName = _fiid + "_" + Bin + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + "_" + pdfCount + ".pdf";
                                                  */

                                                string acc_no = dtCards.Rows[j]["ACCOUNTNO"].ToString();
                                                fileName = _fiid + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["CONTRACTNO"].ToString().Substring(0, 4) + dtCards.Rows[j]["ACCOUNTNO"].ToString().Substring(acc_no.Length - 5, 5) + "_" + dtCards.Rows[j]["idclient"].ToString() + ".PDF";


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

                                                // Set font for paging  
                                                XFont font = new XFont("Verdana", 9);
                                                XBrush brush = XBrushes.Black;
                                                // Create variable that store page count  
                                                string TPages = document.Pages.Count.ToString();
                                                string x = string.Empty;
                                                // Set for loop of document page count and set page number using DrawString function of PdfSharp  


                                                for (int i = 0; i < document.Pages.Count; ++i)
                                                {
                                                    PdfPage page = document.Pages[i];
                                                    // Make a layout rectangle.  
                                                    // XRect layoutRectangle = new XRect(240 /*X*/ , page.Height - font.Height - 10 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                                                    XRect layoutRectangle = new XRect(220 /*X*/ , 295 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                                                    using (XGraphics gfx = XGraphics.FromPdfPage(page))
                                                    {
                                                        //gfx.DrawString("Page " + (i + 1).ToString() + " of " + noPages, font, brush, layoutRectangle, XStringFormats.Center);

                                                        //gfx.DrawString("Page " + (i + 1).ToString() +"of"+ noPages + font, brush, layoutRectangle, XStringFormats.Center);
                                                        gfx.DrawString("Page " + (i + 1).ToString() + "of" + TPages, font, brush, layoutRectangle, XStringFormats.Center);

                                                    }


                                                }

                                                document.Options.CompressContentStreams = true;
                                                document.Options.NoCompression = false;

                                                // Save the document...
                                                document.Save(filePath + "\\" + fileName);

                                                //objst.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);


                                                EStatementInfo objEst = new EStatementInfo();
                                                objEst.BANK_CODE = stmdt.Rows[0]["bank_code"].ToString();
                                                objEst.CLIENT = stmdt.Rows[0]["CLIENT"].ToString();
                                                objEst.IDCLIENT = stmdt.Rows[0]["IDCLIENT"].ToString();
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


                                                sql = "INSERT INTO EMAIL_ADDRESS(MAILADDRESS) VALUES('" + objEst.MAILADDRESS + "')";

                                                reply = objProvider.RunQuery(sql);

                                                objEst.FILE_LOCATION = filePath + "\\" + fileName;
                                                objEst.MAILSUBJECT = txtEmailSubject.Text.Replace("'", "''");
                                                objEst.MAILBODY = "";
                                                objEst.STATUS = "1";

                                                reply = EStatementManager.Instance().AddEStatement(objEst, ref reply);

                                                if (reply == "Success")
                                                {
                                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4) });
                                                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + "  : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4));
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
                                        #endregion

                                    }
                                    else
                                    {
                                        #region try
                                        try
                                        {
                                            pdfCount = pdfCount + 1;
                                            stmdt = new DataTable();
                                            stmdt = objProvider.ReturnData("select * from statement_DUAL where CONTRACTNO='" + dtCards.Rows[j]["CONTRACTNO"].ToString() + "'   ORDER BY SL,[AutoID] ASC", ref reply).Tables[0];

                                            if (stmdt.Rows.Count > 0)
                                            {



                                                EStatement objst = new EStatement();
                                                objst.SetDataSource(stmdt);

                                                /*  string Bin = dtCards.Rows[j]["pan"].ToString().Substring(0, 6);
                                                  fileName = _fiid + "_" + Bin + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + "_" + pdfCount + ".pdf";
                                                  */
                                                string acc_no = dtCards.Rows[j]["ACCOUNTNO"].ToString();
                                                fileName = _fiid + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["CONTRACTNO"].ToString().Substring(0, 4) + dtCards.Rows[j]["ACCOUNTNO"].ToString().Substring(acc_no.Length - 5, 5) + "_" + dtCards.Rows[j]["idclient"].ToString() + ".PDF";


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

                                                // Set font for paging  
                                                XFont font = new XFont("Verdana", 9);
                                                XBrush brush = XBrushes.Black;
                                                // Create variable that store page count  
                                                string TPages = document.Pages.Count.ToString();
                                                string x = string.Empty;
                                                // Set for loop of document page count and set page number using DrawString function of PdfSharp  


                                                for (int i = 0; i < document.Pages.Count; ++i)
                                                {
                                                    PdfPage page = document.Pages[i];
                                                    // Make a layout rectangle.  
                                                    // XRect layoutRectangle = new XRect(240 /*X*/ , page.Height - font.Height - 10 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                                                    XRect layoutRectangle = new XRect(220 /*X*/ , 295 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                                                    using (XGraphics gfx = XGraphics.FromPdfPage(page))
                                                    {
                                                        //gfx.DrawString("Page " + (i + 1).ToString() + " of " + noPages, font, brush, layoutRectangle, XStringFormats.Center);

                                                        //gfx.DrawString("Page " + (i + 1).ToString() +"of"+ noPages + font, brush, layoutRectangle, XStringFormats.Center);
                                                        gfx.DrawString("Page " + (i + 1).ToString() + "of" + TPages, font, brush, layoutRectangle, XStringFormats.Center);

                                                    }


                                                }

                                                document.Options.CompressContentStreams = true;
                                                document.Options.NoCompression = false;

                                                // Save the document...
                                                document.Save(filePathforwithoutEmail + "\\" + fileName);

                                                //objst.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);


                                                EStatementInfo objEst = new EStatementInfo();
                                                objEst.BANK_CODE = stmdt.Rows[0]["bank_code"].ToString();
                                                objEst.CLIENT = stmdt.Rows[0]["CLIENT"].ToString();
                                                objEst.IDCLIENT = stmdt.Rows[0]["IDCLIENT"].ToString();
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

                                                objEst.FILE_LOCATION = filePathforwithoutEmail + "\\" + fileName;
                                                objEst.MAILSUBJECT = txtEmailSubject.Text.Replace("'", "''");
                                                objEst.MAILBODY = "";// txtEmailBody.Text.Replace("'", "''");
                                                objEst.STATUS = "1";

                                                reply = EStatementManager.Instance().AddEStatement(objEst, ref reply);

                                                if (reply == "Success")
                                                {
                                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Invalid Email Address present : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4) });
                                                    objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + "  : Invalid Email Address present : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4));
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
                                        #endregion
                                    }

                                }

                                else
                                {

                                    #region try
                                    try
                                    {
                                        pdfCount = pdfCount + 1;
                                        stmdt = new DataTable();
                                        stmdt = objProvider.ReturnData("select * from statement_DUAL where CONTRACTNO='" + dtCards.Rows[j]["CONTRACTNO"].ToString() + "'   ORDER BY SL,[AutoID] ASC", ref reply).Tables[0];

                                        if (stmdt.Rows.Count > 0)
                                        {



                                            EStatement objst = new EStatement();
                                            objst.SetDataSource(stmdt);

                                            /*     string Bin = dtCards.Rows[j]["pan"].ToString().Substring(0, 6);
                                                 fileName = _fiid + "_" + Bin + "_" + stmdt.Rows[0]["Statement_Date"].ToString().Replace('/', '-') + "_" + pdfCount + ".pdf";
                                                
                                                 */

                                            string acc_no = dtCards.Rows[j]["ACCOUNTNO"].ToString();
                                            fileName = _fiid + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["CONTRACTNO"].ToString().Substring(0, 4) + dtCards.Rows[j]["ACCOUNTNO"].ToString().Substring(acc_no.Length - 5, 5) + "_" + dtCards.Rows[j]["idclient"].ToString() + ".PDF";


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

                                            // Set font for paging  
                                            XFont font = new XFont("Verdana", 9);
                                            XBrush brush = XBrushes.Black;
                                            // Create variable that store page count  
                                            string TPages = document.Pages.Count.ToString();
                                            string x = string.Empty;
                                            // Set for loop of document page count and set page number using DrawString function of PdfSharp  


                                            for (int i = 0; i < document.Pages.Count; ++i)
                                            {
                                                PdfPage page = document.Pages[i];
                                                // Make a layout rectangle.  
                                                // XRect layoutRectangle = new XRect(240 /*X*/ , page.Height - font.Height - 10 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                                                XRect layoutRectangle = new XRect(220 /*X*/ , 295 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                                                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                                                {
                                                    //gfx.DrawString("Page " + (i + 1).ToString() + " of " + noPages, font, brush, layoutRectangle, XStringFormats.Center);

                                                    //gfx.DrawString("Page " + (i + 1).ToString() +"of"+ noPages + font, brush, layoutRectangle, XStringFormats.Center);
                                                    gfx.DrawString("Page " + (i + 1).ToString() + "of" + TPages, font, brush, layoutRectangle, XStringFormats.Center);

                                                }


                                            }

                                            document.Options.CompressContentStreams = true;
                                            document.Options.NoCompression = false;

                                            // Save the document...
                                            document.Save(filePathforwithoutEmail + "\\" + fileName);

                                            //objst.ExportToDisk(ExportFormatType.PortableDocFormat, filePath + "\\" + fileName);


                                            EStatementInfo objEst = new EStatementInfo();
                                            objEst.BANK_CODE = stmdt.Rows[0]["bank_code"].ToString();
                                            objEst.CLIENT = stmdt.Rows[0]["CLIENT"].ToString();
                                            objEst.IDCLIENT = stmdt.Rows[0]["IDCLIENT"].ToString();
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

                                            objEst.FILE_LOCATION = filePathforwithoutEmail + "\\" + fileName;
                                            objEst.MAILSUBJECT = txtEmailSubject.Text.Replace("'", "''");
                                            objEst.MAILBODY = "";// txtEmailBody.Text.Replace("'", "''");
                                            objEst.STATUS = "1";

                                            reply = EStatementManager.Instance().AddEStatement(objEst, ref reply);

                                            if (reply == "Success")
                                            {

                                                txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + "  : No Email Address present !!!\n : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4) });
                                                objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : : No Email Address present !!!\n : Estatement has been created for Card# " + objEst.PAN_NUMBER.Substring(0, 6) + "******" + objEst.PAN_NUMBER.Substring(12, 4));
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
                                    #endregion

                                }
                            }
                            #endregion end forloop


                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Estatement has processed out of " + dtCards.Rows.Count + "." });
                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Estatement has processed" + dtCards.Rows.Count + ".");


                        }
                    }


                }
            }
            #endregion

            #region StatementChecked
            else if (StatementChecked)
            {

                if (!Directory.Exists(_QRPath))
                    System.IO.Directory.CreateDirectory(_QRPath);


                if (ds != null)
                {
                    if (ds.Tables.Count > 0)
                    {
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            DataTable dtAllRows = ds.Tables[0];

                            FileInfo objFile = new FileInfo(_StatementResultPath);

                            if (!Directory.Exists(_StatementResultPath))
                                Directory.CreateDirectory(_StatementResultPath);

                            //  filePath = _StatementProcessedPath + "\\Statement of " + System.DateTime.Now.ToString("ddMMyyyy");
                            filePath = _StatementResultPath;// +"\\Statement of " + System.DateTime.Now.ToString("ddMMyyyy");

                            if (!Directory.Exists(filePath))
                                Directory.CreateDirectory(filePath);

                            txtAnalyzer.Invoke(_addText, new object[] { "\n" + System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + dtCards.Rows.Count.ToString() + " record has been found to process statement." });
                            objLW.logTrace(_StatementLogPath, "Statement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + "Total " + dtCards.Rows.Count.ToString() + " record has been found to process statement.");
                            int docCount = 0;
                            DataTable dtStatementRegister = new DataTable();
                            dtStatementRegister.Columns.Add("FileName");
                            dtStatementRegister.Columns.Add("StatementNo");
                            dtStatementRegister.Columns.Add("ClientId");
                            dtStatementRegister.Columns.Add("Name");
                            dtStatementRegister.Columns.Add("Pan");
                            dtStatementRegister.Columns.Add("StartPage");
                            dtStatementRegister.Columns.Add("EndPage");
                            dtStatementRegister.Columns.Add("SL");
                            dtStatementRegister.Columns.Add("NumberOfPage");
                            dtStatementRegister.Columns.Add("StatementDate");
                            dtStatementRegister.Columns.Add("RefNo");
                            dtStatementRegister.Columns.Add("Address");
                            dtStatementRegister.Columns.Add("CONTRACTNO");



                            //  int TotalPage = 0;

                            for (int j = 0; j < dtCards.Rows.Count; j++)//dtCards.Rows.Count
                            {
                                try
                                {
                                    stmdt = new DataTable();
                                    //  stmdt = objProvider.ReturnData("select * from statement_DUAL ", ref reply).Tables[0];
                                    stmdt = objProvider.ReturnData("select * from statement_DUAL where CONTRACTNO='" + dtCards.Rows[j]["CONTRACTNO"].ToString() + "'ORDER BY SL,[AutoID] ASC", ref reply).Tables[0];

                                    if (stmdt.Rows.Count > 0)
                                    {
                                        // string sd = dtCards.Rows[j]["StDate"].ToString().Replace("/","").Replace("-","");

                                        QRGenetetor(dtCards.Rows[j]["idclient"].ToString() + dtCards.Rows[j]["pan"].ToString().Substring(dtCards.Rows[j]["pan"].ToString().Length - 4, 4) + dtCards.Rows[j]["StDate"].ToString().Replace("/", "").Replace("-", ""), dtCards.Rows[j]["CONTRACTNO"].ToString());
                                        // wholeVIN.Substring(wholeVIN.Length - 4, 4);
                                        DataRow dr1 = dtStatementRegister.NewRow();
                                        dtStatementRegister.Rows.InsertAt(dr1, docCount);
                                        dtStatementRegister.Rows[docCount][7] = docCount + 1;
                                        dtStatementRegister.Rows[docCount][1] = dtCards.Rows[j]["StatementNo"].ToString();
                                        dtStatementRegister.Rows[docCount][2] = dtCards.Rows[j]["idclient"].ToString();
                                        dtStatementRegister.Rows[docCount][3] = dtCards.Rows[j]["client"].ToString();
                                        dtStatementRegister.Rows[docCount][4] = dtCards.Rows[j]["PAN"].ToString();
                                        dtStatementRegister.Rows[docCount][9] = dtCards.Rows[j]["StDate"].ToString();

                                        if (!string.IsNullOrEmpty(dtCards.Rows[j]["idclient"].ToString() + dtCards.Rows[j]["pan"].ToString().Substring(dtCards.Rows[j]["pan"].ToString().Length - 4, 4) + dtCards.Rows[j]["StDate"].ToString().Replace("/", "").Replace("-", "")))
                                        {
                                            dtStatementRegister.Rows[docCount][10] = dtCards.Rows[j]["idclient"].ToString() + dtCards.Rows[j]["pan"].ToString().Substring(dtCards.Rows[j]["pan"].ToString().Length - 4, 4) + dtCards.Rows[j]["StDate"].ToString().Replace("/", "").Replace("-", "");
                                        }
                                        else
                                        {
                                            dtStatementRegister.Rows[docCount][10] = "";
                                        }


                                        if (!string.IsNullOrEmpty(stmdt.Rows[0]["ADDRESS"].ToString()))
                                        {
                                            if (!string.IsNullOrEmpty(stmdt.Rows[0]["PROMOTIONALTEXT"].ToString()))
                                            //if (stmdt.Rows[0]["ADDRESS"].ToString().Length > 120)
                                            {
                                                string companyName = stmdt.Rows[0]["PROMOTIONALTEXT"].ToString();
                                                string jobTitle = stmdt.Rows[0]["JOBTITLE"].ToString();
                                                string promo = jobTitle + " , " + companyName;
                                                dtStatementRegister.Rows[docCount][11] = promo + " , " + stmdt.Rows[0]["ADDRESS"].ToString().Replace("*", " ");
                                                // dtStatementRegister.Rows[docCount][11] = stmdt.Rows[0]["ADDRESS"].ToString().Replace("*", " ").Substring(0, 120);
                                            }

                                            else
                                            {

                                                dtStatementRegister.Rows[docCount][11] = stmdt.Rows[0]["ADDRESS"].ToString().Replace("*", " ");

                                            }
                                            //else if (stmdt.Rows[0]["ADDRESS"].ToString().Length <=120)
                                            //{
                                            //    string companyname = stmdt.Rows[0]["PROMOTIONALTEXT"].ToString();

                                            //    string jobTitle = stmdt.Rows[0]["JOBTITLE"].ToString();

                                            //    string promo = companyname +" , "+ jobTitle;
                                            //    dtStatementRegister.Rows[docCount][11] = promo +" , "+ stmdt.Rows[0]["ADDRESS"].ToString().Replace("*", " ");
                                            //}
                                        }
                                        else
                                        {
                                            dtStatementRegister.Rows[docCount][11] = "";
                                        }

                                        dtStatementRegister.Rows[docCount][12] = dtCards.Rows[j]["CONTRACTNO"].ToString();
                                        QCash.EStatement.NBL.Reports.Statement objst = new QCash.EStatement.NBL.Reports.Statement();

                                        objst.SetDataSource(stmdt);


                                        //  fileName = _fiid + "_VISA_Statement" + "_" + dtCards.Rows[j]["client"].ToString() + "_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["Statementno"].ToString().PadLeft(6, '0') + ".pdf";


                                        string acc_no = dtCards.Rows[j]["ACCOUNTNO"].ToString();
                                        fileName = _fiid + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["CONTRACTNO"].ToString().Substring(0, 4) + dtCards.Rows[j]["ACCOUNTNO"].ToString().Substring(acc_no.Length - 5, 5) + "_" + dtCards.Rows[j]["idclient"].ToString() + ".PDF";


                                        System.IO.Stream st = objst.ExportToStream(ExportFormatType.PortableDocFormat);

                                        PdfSharp.Pdf.PdfDocument document = PdfReader.Open(st);
                                        int noPages = document.Pages.Count;
                                        dtStatementRegister.Rows[docCount]["FileName"] = fileName;
                                        dtStatementRegister.Rows[docCount]["NumberOfPage"] = noPages;

                                        // Set font for paging  
                                        XFont font = new XFont("Verdana", 9);
                                        XBrush brush = XBrushes.Black;
                                        // Create variable that store page count  
                                        string Pages = document.Pages.Count.ToString();
                                        string x = string.Empty;
                                        // Set for loop of document page count and set page number using DrawString function of PdfSharp  


                                        for (int i = 0; i < document.Pages.Count; ++i)
                                        {
                                            PdfPage page = document.Pages[i];
                                            // Make a layout rectangle.  
                                            XRect layoutRectangle = new XRect(520 /*X*/ , 325 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                                            //  XRect layoutRectangle = new XRect(2400 /*X*/ , page.Height - font.Height - 10 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                                            using (XGraphics gfx = XGraphics.FromPdfPage(page))
                                            {

                                                gfx.DrawString("Page " + (i + 1).ToString() + "of " + Pages, font, brush, layoutRectangle, XStringFormats.TopLeft);

                                            }


                                        }

                                        //try
                                        //{
                                        //    dtStatementRegister.Rows[docCount]["StartPage"] = (int.Parse(dtStatementRegister.Rows[docCount-1]["EndPage"].ToString()) + 1).ToString();
                                        //}
                                        //catch
                                        //{
                                        //    dtStatementRegister.Rows[docCount]["StartPage"] = "1";
                                        //}
                                        //try
                                        //{
                                        //    dtStatementRegister.Rows[docCount]["EndPage"] = (int.Parse(dtStatementRegister.Rows[docCount - 1]["EndPage"].ToString()) + document.Pages.Count).ToString();
                                        //}
                                        //catch
                                        //{
                                        //    dtStatementRegister.Rows[docCount]["EndPage"] = document.Pages.Count.ToString();
                                        //}

                                        document.Options.CompressContentStreams = true;
                                        document.Options.NoCompression = false;

                                        document.Save(filePath + "\\" + fileName);

                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Statement has been created for Card# " + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) });
                                        objLW.logTrace(_StatementLogPath, "Statement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Statement has been created for Card# " + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "******" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4));
                                        count++;




                                        if (count % 10 == 0)
                                        {
                                            objst.Dispose();
                                            GC.Collect();
                                            Thread.Sleep(200);
                                        }
                                        docCount++;
                                    }


                                }
                                catch (Exception ex)
                                {
                                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                                    objLW.logTrace(_StatementLogPath, "Statement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + ex.Message);
                                }

                            }  // end for loop



                            //string[] pdfs = Directory.GetFiles(_StatementProcessedPath);
                            string[] pdfs = new string[dtStatementRegister.Rows.Count];
                            //for (int i = 0; i < pdfs.Length; i++)
                            //{
                            for (int j = 0; j < dtStatementRegister.Rows.Count; j++)
                                pdfs[j] = filePath + "\\" + dtStatementRegister.Rows[j]["FileName"];

                            //    for (int j = 0; j < dtStatementRegister.Rows.Count; j++)
                            //    {
                            //        if (pdfs[i].Contains(dtStatementRegister.Rows[j]["FileName"].ToString()))
                            //        {
                            //            TotalPage = TotalPage + 1;
                            //            dtStatementRegister.Rows[j]["StartPage"] = TotalPage;
                            //            TotalPage = TotalPage + Convert.ToInt32(dtStatementRegister.Rows[j]["NumberOfPage"].ToString());
                            //            TotalPage = TotalPage - 1;
                            //            dtStatementRegister.Rows[j]["EndPage"] = TotalPage;
                            //            break;
                            //        }
                            //    }

                            //}
                            GetStatementRegister_Info(dtStatementRegister);

                            string CombineStatementFileName = _xmlName + ".pdf";

                            MergeMultiplePDFIntoSinglePDF(_CombineStatementPath, pdfs, dtStatementRegister, CombineStatementFileName);
                            // MergeMultiplePDFIntoSinglePDF(_CombineStatementPath, pdfs, dtStatementRegister);


                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Statement has processed out of " + dtCards.Rows.Count + "." });
                            objLW.logTrace(_StatementLogPath, "Statement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + count.ToString() + " Statement has processed" + dtCards.Rows.Count + ".");
                        } // end if

                    } // end if
                } // end if


            }  // end else if
            #endregion


        } // end function

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

                        FileInfo objFile = new FileInfo(_EmailResultPath);

                        if (!Directory.Exists(_EmailResultPath))
                            Directory.CreateDirectory(_EmailResultPath);

                        filePath = _EmailResultPath + "\\EStatement of " + System.DateTime.Now.ToString("ddMMyyyy");

                        if (!Directory.Exists(filePath))
                            Directory.CreateDirectory(filePath);

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
                                        stmdt = objProvider.ReturnData("select * from statement_USD where CONTRACTNO='" + dtCards.Rows[j]["CONTRACTNO"].ToString() + "'", ref reply).Tables[0];
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

                                            fileName = _fiid + "_VISA_EStatement_" + dtCards.Rows[j]["idclient"].ToString() + "_" + dtCards.Rows[j]["pan"].ToString().Substring(0, 6) + "_" + dtCards.Rows[j]["pan"].ToString().Substring(12, 4) + "_USD.pdf";

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
                                            objEst.MAILBODY = "";
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

        void ReportViewer_Load(object sender, EventArgs e)
        {
            mailProgress.Visible = false;

            // E-Statement Path

            _XMLProcessedPath = ConfigurationManager.AppSettings["EmailProcessPath"].ToString();
            _XMLSourcePath = ConfigurationManager.AppSettings["EmailSourcePath"].ToString();
            _EmailResultPath = ConfigurationManager.AppSettings["EmailResultPath"].ToString();
            _LogPath = ConfigurationManager.AppSettings["EmailLogPath"].ToString();
            _AdditionalAttachment = ConfigurationManager.AppSettings["AdditionalAttachment"].ToString();



            // Statement Path
            _StatementSourcePath = ConfigurationManager.AppSettings["StatementSourcePath"].ToString();
            _StatementResultPath = ConfigurationManager.AppSettings["StatementResultPath"].ToString();
            _StatementLogPath = ConfigurationManager.AppSettings["StatementLogPath"].ToString();
            _CombineStatementPath = ConfigurationManager.AppSettings["CombineStatementPath"].ToString();
            _statementProcessPath = ConfigurationManager.AppSettings["statementProcessPath"].ToString();

            // QR Code Path

            _QRPath = ConfigurationManager.AppSettings["QRPath"].ToString(); ;
            _QRSize = int.Parse(ConfigurationManager.AppSettings["QRSize"].ToString());
            _QRPositionX = int.Parse(ConfigurationManager.AppSettings["QRPositionX"].ToString());
            _QRPositionY = int.Parse(ConfigurationManager.AppSettings["QRPositionY"].ToString());
            _QRPage = int.Parse(ConfigurationManager.AppSettings["QRPage"].ToString());
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
            bool eStatementChecked = rbtneStatement.Checked;
            bool StatementChecked = rbtnStatement.Checked;

            #region Folder Searching in File name path





            #region E-Statement

            if (eStatementChecked)
            {
                DirectoryInfo di = new DirectoryInfo(_XMLSourcePath);
                DirectoryInfo[] dia = di.GetDirectories();

                for (int fcount = 0; fcount < dia.Length; fcount++)
                {
                    if (dia[fcount].FullName.Contains("NBL"))
                    {
                        _bankName = "NBL";
                        _bankCode = "3";
                        _XMLSourcePath = dia[fcount].FullName;
                        //
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
            }
            #endregion

            #region Statement



            if (StatementChecked)
            {
                DirectoryInfo dis = new DirectoryInfo(_StatementSourcePath);
                DirectoryInfo[] dias = dis.GetDirectories();

                for (int fcount = 0; fcount < dias.Length; fcount++)
                {
                    if (dias[fcount].FullName.Contains("NBL"))
                    {
                        _bankName = "NBL";
                        _bankCode = "3";
                        _StatementSourcePath = dias[fcount].FullName;
                        //
                        ProcessFolderFiles(_StatementSourcePath, _bankCode, _bankName, ref _reply);

                    }


                    else
                    {
                        MsgLogWriter objLW = new MsgLogWriter();
                        objLW.logTrace(_StatementLogPath, "Statement.log", "Not an XML data !!!");
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + "Not an XML data !!!" });
                    }

                    Bank_Code = _bankName;

                }


            }
            #endregion

            #endregion

        }

        private void ProcessFolderFiles(string _SourcePath, string BankCode, string BankName, ref string _reply)
        {
            bool eStatementChecked = rbtneStatement.Checked;
            bool StatementChecked = rbtnStatement.Checked;


            #region E-Statement

            if (eStatementChecked)
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
                                objProvider.RunQuery("Delete from  AccumIntAcc");
                                //Clear Previous BonusContrAcc Data
                                objProvider.RunQuery("Delete from  BonusContrAcc");

                                objProvider.RunQuery("insert into statement_info_arc select * from statement_info");

                                objProvider.RunQuery("insert into statement_details_arc select * from statement_details");

                                objProvider.RunQuery("Truncate table  statement_details");
                                objProvider.RunQuery("Delete from  statement_info");

                                for (int i = 0; i < dsXML.Tables.Count; i++)
                                {
                                    if (dsXML.Tables[i].TableName == "Statement")
                                    {
                                        GetCardHolderPersonalInfo(dsXML.Tables[i], BankName, ref reply);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Personal Info data Saved from XML. " + reply });
                                        objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Personal Info data Saved from XML. " + reply);
                                    }
                                    else if (dsXML.Tables[i].TableName == "Operation")
                                    {
                                        reply = GetCardHolderTransactionInfo(dsXML.Tables[i]);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Transaction Info data Saved from XML. " + reply });
                                        objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Transaction Info data Saved from XML. " + reply);
                                    }
                                    else if (dsXML.Tables[i].TableName == "Account")
                                    {
                                        reply = GetCardHolderAccountInfo(dsXML.Tables[i]);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Account Info data Saved from XML. " + reply });
                                        objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Account Info data Saved from XML. " + reply);
                                    }
                                    else if (dsXML.Tables[i].TableName == "Card")
                                    {
                                        reply = GetCardHolderCardInfo(dsXML.Tables[i]);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Card Info data Saved from XML. " + reply });
                                        objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Card Info data Saved from XML. " + reply);
                                    }
                                    else if (dsXML.Tables[i].TableName == "BonusContrAcc")
                                    {
                                        reply = GetBonusContrAccInfo(dsXML.Tables[i]);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Card Info data Saved from XML. " + reply });
                                        objLW.logTrace(_LogPath, "EStatement.log", " : CardHolder Card Info data Saved from XML. " + reply);
                                    }
                                    else if (dsXML.Tables[i].TableName == "AccumIntAcc")
                                    {
                                        reply = GetAccumIntAcc(dsXML.Tables[i]);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Card Info data Saved from XML. " + reply });
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
                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + dsXML.Tables["Card"].Rows.Count.ToString() + " Card record has been found to process.." });
                            objLW.logTrace(_LogPath, "EStatement.log", " : Total " + dsXML.Tables["Card"].Rows.Count.ToString() + " Card record has been found to process..");
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

            #endregion

            #region Statement

            else if (StatementChecked)
            {

                #region Files of a Directory
                string reply = string.Empty;

                try
                {
                    MsgLogWriter objLW = new MsgLogWriter();


                    DirectoryInfo dir = new DirectoryInfo(_SourcePath);
                    FileInfo[] fi = dir.GetFiles();

                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + fi.Length.ToString() + " files found to process.." });
                    objLW.logTrace(_StatementLogPath, "Statement.log", " : Total " + fi.Length.ToString() + " files found to process..");

                    for (int f = 0; f < fi.Length; f++)
                    {
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + fi[f].Name + " on process.." });
                        objLW.logTrace(_StatementLogPath, "Statement.log", " : " + fi[f].Name + " on process..");

                        DataSet dsXML = getDataFromXML(fi[f].FullName);
                        string xmlname = fi[f].Name;
                        _xmlName = xmlname.ToString().Replace(".xml", "");

                        #region Operation On Data
                        if (dsXML != null)
                        {
                            if (dsXML.Tables.Count > 0)
                            {
                                ConStr = new ConnectionStringBuilder(1);
                                objProvider = new SqlDbProvider(ConStr.ConnectionString_DBConfig);

                                string sql = string.Empty;

                                //Clear Previous AccumIntAcc Data
                                objProvider.RunQuery("Delete from  AccumIntAcc");
                                //Clear Previous BonusContrAcc Data
                                objProvider.RunQuery("Delete from  BonusContrAcc");

                                objProvider.RunQuery("insert into statement_info_arc select * from statement_info");

                                objProvider.RunQuery("insert into statement_details_arc select * from statement_details");

                                objProvider.RunQuery("Truncate table  statement_details");
                                objProvider.RunQuery("Delete from  statement_info");


                                for (int i = 0; i < dsXML.Tables.Count; i++)
                                {
                                    if (dsXML.Tables[i].TableName == "Statement")
                                    {
                                        GetCardHolderPersonalInfo(dsXML.Tables[i], BankName, ref reply);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Personal Info data Saved from XML. " + reply });
                                        objLW.logTrace(_StatementLogPath, "Statement.log", " : CardHolder Personal Info data Saved from XML. " + reply);
                                    }
                                    else if (dsXML.Tables[i].TableName == "Operation")
                                    {
                                        reply = GetCardHolderTransactionInfo(dsXML.Tables[i]);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Transaction Info data Saved from XML. " + reply });
                                        objLW.logTrace(_StatementLogPath, "Statement.log", " : CardHolder Transaction Info data Saved from XML. " + reply);
                                    }
                                    else if (dsXML.Tables[i].TableName == "Account")
                                    {
                                        reply = GetCardHolderAccountInfo(dsXML.Tables[i]);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Account Info data Saved from XML. " + reply });
                                        objLW.logTrace(_StatementLogPath, "Statement.log", " : CardHolder Account Info data Saved from XML. " + reply);
                                    }
                                    else if (dsXML.Tables[i].TableName == "Card")
                                    {
                                        reply = GetCardHolderCardInfo(dsXML.Tables[i]);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Card Info data Saved from XML. " + reply });
                                        objLW.logTrace(_StatementLogPath, "Statement.log", " : CardHolder Card Info data Saved from XML. " + reply);
                                    }
                                    else if (dsXML.Tables[i].TableName == "BonusContrAcc")
                                    {
                                        reply = GetBonusContrAccInfo(dsXML.Tables[i]);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Card Info data Saved from XML. " + reply });
                                        objLW.logTrace(_StatementLogPath, "Statement.log", " : CardHolder Card Info data Saved from XML. " + reply);
                                    }
                                    else if (dsXML.Tables[i].TableName == "AccumIntAcc")
                                    {
                                        reply = GetAccumIntAcc(dsXML.Tables[i]);
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : CardHolder Card Info data Saved from XML. " + reply });
                                        objLW.logTrace(_StatementLogPath, "Statement.log", " : CardHolder Card Info data Saved from XML. " + reply);
                                    }
                                }
                                if (reply == "Success")
                                {
                                    if (dsXML.Tables.Contains("Operation"))
                                        GenerateStatementInfo(dsXML, BankName, ref reply);
                                    //for (int i = 0; i < dsXML.Tables.Count; i++)
                                    //{
                                    //    if (dsXML.Tables[i].TableName == "Operation")
                                    //        GenerateStatementInfo(dsXML, BankName, ref reply);
                                    //}
                                }
                            }
                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Total " + dsXML.Tables["Card"].Rows.Count.ToString() + " Card record has been found to process.." });
                            objLW.logTrace(_StatementLogPath, "Statement.log", " : Total " + dsXML.Tables["Card"].Rows.Count.ToString() + " Card record has been found to process..");
                        }
                        #endregion

                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + fi[f].Name + " process complete.." });
                        objLW.logTrace(_StatementLogPath, "Statement.log", " : " + fi[f].Name + " process complete..");
                        txtAnalyzer.Invoke(_addText, new object[] { "\n" });

                        //if (Directory.Exists(_SourcePath))
                        //{
                        //    try
                        //    {
                        //        Directory.Move(dir.FullName, _statementProcessPath + "\\" + dir.Name);

                        //    }
                        //    catch (IOException ex)
                        //    {
                        //        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Source Directory moving Error. Error: " + ex.Message });
                        //        objLW = new MsgLogWriter();
                        //        objLW.logTrace(_LogPath, "Statement.log", "Source Directory moving Error. " + ex.Message);
                        //    }
                        //}
                        btnGenerate_Click(null, null);
                    }
                    if (Directory.Exists(_SourcePath))
                    {
                        try
                        {
                            Directory.Move(dir.FullName, _statementProcessPath + "\\" + dir.Name);

                        }
                        catch (IOException ex)
                        {
                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Source Directory moving Error. Error: " + ex.Message });
                            objLW = new MsgLogWriter();
                            objLW.logTrace(_StatementLogPath, "Statement.log", "Source Directory moving Error. " + ex.Message);
                        }
                    }
                    // return true;
                }
                catch (Exception ex)
                {
                    _reply = ex.StackTrace;
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + ex.Message });
                    MsgLogWriter objLW = new MsgLogWriter();
                    objLW.logTrace(_StatementLogPath, "Statement.log", ex.Message);
                }

                #endregion

            }



            #endregion
        }

        private DataSet getDataFromXML(string _filename)
        {
            bool eStatementChecked = rbtneStatement.Checked;
            bool StatementChecked = rbtnStatement.Checked;
            #region E-Statement

            if (eStatementChecked)
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
            #endregion

            #region Statement

            else if (StatementChecked)
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
                    objLW.logTrace(_StatementLogPath, "Statement.log", ex.Message);
                    return null;
                }
            }
            return null;


            #endregion

        }

        private StatementList GetCardHolderPersonalInfo(DataTable dtStatement, string BankCode, ref string errMsg)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            Statement objSt = null;
            StatementList objStList = new StatementList();

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
                            objSt.EMAIL = dtStatement.Rows[k][j].ToString().Replace("'", "");
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
                            objSt.TELEPHONE = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "Title")
                        {
                            objSt.TITLE = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }

                        else if (dtStatement.Columns[j].ColumnName == "ZIP")
                        {
                            objSt.ZIP = dtStatement.Rows[k][j].ToString().Replace("'", "");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "JobTitle")
                        {
                            objSt.JOBTITLE = dtStatement.Rows[k][j].ToString().Replace("'", "''");
                        }
                        else if (dtStatement.Columns[j].ColumnName == "PromotionalText")
                        {
                            objSt.PROMOTIONALTEXT = dtStatement.Rows[k][j].ToString().Replace("'", "''");

                            //string value = objSt.PROMOTIONALTEXT;
                            //string[] lines = value.Split(new char[] { '|' });

                            //    if (!string.IsNullOrEmpty(lines[0]))
                            //    {
                            //        objSt.INDICATOR = lines[0];
                            //    }
                            //    if (!string.IsNullOrEmpty(lines[1]))
                            //    {
                            //        objSt.COMPANYNAME = lines[1];
                            //    }
                            //    if (!string.IsNullOrEmpty(lines[2]))
                            //    {
                            //        objSt.JOBTITLE = lines[2];
                            //    }
                            //    if (!string.IsNullOrEmpty(lines[3]))
                            //    {
                            //        objSt.COMPANYADDRESS1 = lines[3];
                            //    }
                            //    if (!string.IsNullOrEmpty(lines[4]))
                            //    {
                            //        objSt.COMPANYADDRESS2 = lines[4];
                            //    }
                            //    if (!string.IsNullOrEmpty(lines[5]))
                            //    {
                            //        objSt.CITYN = lines[5];
                            //    }


                            //    objSt.CITY = objSt.CITYN.Substring(0, objSt.CITYN.IndexOf("-"));

                            //    objSt.ZIP = objSt.CITYN.Substring((objSt.CITY).Length+1, objSt.CITYN.IndexOf("-"));

                            //    objSt.COMPANYADDRESS = objSt.COMPANYADDRESS1 + " " + objSt.COMPANYADDRESS2 + " " + objSt.CITY + "-" + objSt.ZIP;


                        }
                        #endregion
                    }
                    //objStList.Add(objSt);

                    sql = "Insert into Statement(BANK_CODE,STATEMENTNO,ADDRESS,CARD_LIST,CITY,COUNTRY,EMAIL," +
                          "STARTDATE,ENDDATE,CLIENT,CONTRACTNO,IDCLIENT,FAX,MAIN_CARD,MOBILE," +
                          "NEXT_STATEMENT_DATE,PAYMENT_DATE,REGION,STATEMENT_DATE,SEX,STREETADDRESS,TELEPHONE,TITLE,ZIP,PROMOTIONALTEXT,JOBTITLE) " +
                          "values('" + objSt.BANK_CODE + "','" + objSt.STATEMENTNO + "','" + objSt.ADDRESS + "','" + objSt.CARD_LIST + "','" + objSt.CITY + "','" + objSt.COUNTRY + "','" + objSt.EMAIL + "'," +
                          "'" + objSt.STARTDATE + "','" + objSt.ENDDATE + "','" + objSt.CLIENT + "','" + objSt.CONTRACTNO + "','" + objSt.IDCLIENT + "','" + objSt.FAX + "','" + objSt.MAIN_CARD + "','" + objSt.MOBILE + "'," +
                          "'" + objSt.NEXT_STATEMENT_DATE + "','" + objSt.PAYMENT_DATE + "','" + objSt.REGION + "','" + objSt.STATEMENT_DATE + "','" + objSt.SEX + "','" + objSt.STREETADDRESS + "'," +
                          "'" + objSt.TELEPHONE + "','" + objSt.TITLE + "','" + objSt.ZIP + "','" + objSt.PROMOTIONALTEXT + "','" + objSt.JOBTITLE + "')";

                    reply = objProvider.RunQuery(sql);
                    //if (!reply.Contains("Success"))
                    errMsg = reply;
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

        private bool IsDuplicateFound(string contactNo)
        {
            string reply = string.Empty;
            bool isDuplicateDataFound = false;
            string query = "select CONTRACTNO from Statement where CONTRACTNO = " + "'" + contactNo + "'";
            DataSet dataSet = objProvider.ReturnData(query, ref reply);

            if (dataSet != null)
                if (dataSet.Tables[0] != null && dataSet.Tables[0].Rows.Count > 0)
                    isDuplicateDataFound = true;

            return isDuplicateDataFound;
        }

        private string GetCardHolderTransactionInfo(DataTable dtOperation)
        {
            #region Operation
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
                        //{
                        #region setting properties values

                        //switch (dtOperation.Columns[j].ColumnName)
                        //{
                        //case "StatementNo":
                        if (dtOperation.Columns.Contains("StatementNo"))
                            objOp.STATEMENTNO = dtOperation.Rows[k]["StatementNo"].ToString().Replace("'", "");
                    //break;
                    //case "O":
                    if (dtOperation.Columns.Contains("O"))
                        objOp.OpID = dtOperation.Rows[k]["O"].ToString().Replace("'", "");
                    //break;
                    //case "OD":
                    if (dtOperation.Columns.Contains("OD"))
                        objOp.OpDate = dtOperation.Rows[k]["OD"].ToString().Replace("'", "");
                    //break;
                    //case "TD":
                    if (dtOperation.Columns.Contains("TD"))
                        objOp.TD = dtOperation.Rows[k]["TD"].ToString().Replace("'", "");
                    //break;
                    //case "A":
                    if (dtOperation.Columns.Contains("A"))
                    {
                        if (string.IsNullOrEmpty(dtOperation.Rows[k]["A"].ToString()))
                            objOp.Amount = "0.00";
                        else
                            objOp.Amount = dtOperation.Rows[k]["A"].ToString().Replace("'", "");
                    }
                    //break;
                    //case "ACURC":
                    if (dtOperation.Columns.Contains("ACURC"))
                        objOp.ACURCode = dtOperation.Rows[k]["ACURC"].ToString().Replace("'", "");
                    //break;
                    //case "ACURN":
                    if (dtOperation.Columns.Contains("ACURN"))
                        objOp.ACURName = dtOperation.Rows[k]["ACURN"].ToString().Replace("'", "");
                    //break;
                    //case "D":
                    if (dtOperation.Columns.Contains("D"))
                        objOp.D = dtOperation.Rows[k]["D"].ToString().Replace("'", "''");
                    //break;
                    //case "DE":
                    if (dtOperation.Columns.Contains("DE"))
                        objOp.DE = dtOperation.Rows[k]["DE"].ToString().Replace("'", "''");
                    //break;
                    //case "CF":
                    if (dtOperation.Columns.Contains("CF"))
                        objOp.CF = dtOperation.Rows[k]["CF"].ToString().Replace("'", "''");
                    //break;
                    //case "DOCNO":
                    if (dtOperation.Columns.Contains("DOCNO"))
                        objOp.DOCNO = dtOperation.Rows[k]["DOCNO"].ToString().Replace("'", "");
                    //break;
                    //case "NO":
                    if (dtOperation.Columns.Contains("NO"))
                        objOp.NO = dtOperation.Rows[k]["NO"].ToString().Replace("'", "");
                    //break;
                    //case "ACCOUNT":
                    if (dtOperation.Columns.Contains("ACCOUNT"))
                        objOp.ACCOUNT = dtOperation.Rows[k]["ACCOUNT"].ToString().Replace("'", "");
                    //break;
                    //case "ACC":
                    if (dtOperation.Columns.Contains("ACC"))
                        objOp.ACC = dtOperation.Rows[k]["ACC"].ToString().Replace("'", "");
                    //break;
                    //case "FR":
                    if (dtOperation.Columns.Contains("FR"))
                        objOp.FR = dtOperation.Rows[k]["FR"].ToString().Replace("'", "''");
                    //break;
                    //case "APPROVAL":
                    if (dtOperation.Columns.Contains("APPROVAL"))
                        objOp.APPROVAL = dtOperation.Rows[k]["APPROVAL"].ToString().Replace("'", "");
                    //break;
                    //case "MN":
                    if (dtOperation.Columns.Contains("MN"))
                        objOp.MN = dtOperation.Rows[k]["MN"].ToString().Replace("'", "''");
                    //break;
                    //case "S":
                    if (dtOperation.Columns.Contains("S"))
                        objOp.S = dtOperation.Rows[k]["S"].ToString().Replace("'", "");
                    //break;
                    //case "TERMN":
                    if (dtOperation.Columns.Contains("TERMN"))
                        objOp.TERMN = dtOperation.Rows[k]["TERMN"].ToString().Replace("'", "''");
                    //break;
                    //case "TL":
                    if (dtOperation.Columns.Contains("TL"))
                        objOp.TL = dtOperation.Rows[k]["TL"].ToString().Replace("'", "''");
                    //break;
                    //case "P":
                    if (dtOperation.Columns.Contains("P"))
                        objOp.P = dtOperation.Rows[k]["P"].ToString().Replace("'", "");
                    //break;
                    //case "SERIALNO":
                    if (dtOperation.Columns.Contains("SERIALNO"))
                        objOp.SERIALNO = dtOperation.Rows[k]["SERIALNO"].ToString().Replace("'", "");
                    //break;
                    //case "OCC":
                    if (dtOperation.Columns.Contains("OCC"))
                        objOp.OCCode = dtOperation.Rows[k]["OCC"].ToString().Replace("'", "");
                    //break;
                    //case "OC":
                    if (dtOperation.Columns.Contains("OC"))
                        objOp.OCName = dtOperation.Rows[k]["OC"].ToString().Replace("'", "");
                    //break;
                    //case "AMOUNTSIGN":
                    if (dtOperation.Columns.Contains("AMOUNTSIGN"))
                        objOp.AMOUNTSIGN = dtOperation.Rows[k]["AMOUNTSIGN"].ToString().Replace("'", "");
                    //break;
                    //case "OA":
                    if (dtOperation.Columns.Contains("OA"))
                    {
                        if (string.IsNullOrEmpty(dtOperation.Rows[k]["OA"].ToString()))
                            objOp.OA = "0.00";
                        else
                            objOp.OA = dtOperation.Rows[k]["OA"].ToString().Replace("'", "");
                    }
                    //break;                            
                    //}


                        #endregion
                    //}
                    //objOpList.Add(objOp);

                    sql = "Insert into Operation(STATEMENTNO,O,OD,TD,A,ACURC,ACURN,D,DE,P,OA,OCC,OC,TL,TERMN,CF,S,MN,DOCNO,NO,ACCOUNT,ACC,FR,APPROVAL,AMOUNTSIGN,SERIALNO) " +
                    "Values('" + objOp.STATEMENTNO + "','" + objOp.OpID + "','" + objOp.OpDate + "','" + objOp.TD + "','" + objOp.Amount + "'," +
                    "'" + objOp.ACURCode + "','" + objOp.ACURName + "','" + objOp.D + "','" + objOp.DE + "','" + objOp.P + "','" + objOp.OA + "'," +
                    "'" + objOp.OCCode + "','" + objOp.OCName + "','" + objOp.TL + "','" + objOp.TERMN + "','" + objOp.CF + "','" + objOp.S + "'," +
                    "'" + objOp.MN + "','" + objOp.DOCNO + "','" + objOp.NO + "','" + objOp.ACCOUNT + "','" + objOp.ACC + "','" + objOp.FR + "','" + objOp.APPROVAL + "','" + objOp.AMOUNTSIGN + "','" + objOp.SERIALNO + "') ";

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
            #endregion Operation
        }

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
                    #region setting properties values
                    for (int j = 0; j < dtBonusContrAcc.Columns.Count; j++)
                    {
                        switch (dtBonusContrAcc.Columns[j].ColumnName)
                        {
                            case "StatementNo":
                                objOp.STATEMENTNO = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "SUM_CREDIT":
                                objOp.SUM_CREDIT = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "SUM_DEBIT":
                                objOp.SUM_DEBIT = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "EBALANCE":
                                objOp.EBALANCE = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "ACCOUNT_NO":
                                objOp.ACCOUNT_NO = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "ACURN":
                                objOp.ACURN = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "ACURC":
                                objOp.ACURC = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "SBALANCE":
                                objOp.SBALANCE = dtBonusContrAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
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

        private string GetStatementRegister_Info(DataTable dtStatementRegister_Info)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            StatementRegisterInfo objOp = null;

            dtStatementRegister_Info.TableName = "StatementRegister";
            try
            {
                //Clear Previous Data
                //objProvider.RunQuery("insert into StatementRegister_ARC		select * from StatementRegister");

                //objProvider.RunQuery("Delete from " + dtStatementRegister_Info.TableName);

                for (int k = 0; k < dtStatementRegister_Info.Rows.Count; k++)
                {
                    objOp = new StatementRegisterInfo();



                    #region setting properties values

                    objOp.SL = dtStatementRegister_Info.Rows[k]["SL"].ToString().Replace("'", "");

                    objOp.StatementNo = dtStatementRegister_Info.Rows[k]["StatementNo"].ToString().Replace("'", "");

                    objOp.ClientID = dtStatementRegister_Info.Rows[k]["ClientId"].ToString().Replace("'", "");

                    objOp.ClientName = dtStatementRegister_Info.Rows[k]["Name"].ToString().Replace("'", "''");

                    objOp.PAN = dtStatementRegister_Info.Rows[k]["Pan"].ToString().Replace("'", "");

                    objOp.StartPage = dtStatementRegister_Info.Rows[k]["StartPage"].ToString();

                    objOp.EndPage = dtStatementRegister_Info.Rows[k]["EndPage"].ToString();

                    objOp.FileName = dtStatementRegister_Info.Rows[k]["FileName"].ToString().Replace("'", "''");
                    objOp.CONTRACTNO = dtStatementRegister_Info.Rows[k]["CONTRACTNO"].ToString().Replace("'", "");

                    objOp.TotalPage = dtStatementRegister_Info.Rows[k]["NumberOfPage"].ToString();
                    objOp.StatementDate = dtStatementRegister_Info.Rows[k]["StatementDate"].ToString().Replace("'", "");
                    objOp.RefNo = dtStatementRegister_Info.Rows[k]["RefNo"].ToString().Replace("'", "''");
                    objOp.Address = dtStatementRegister_Info.Rows[k]["Address"].ToString().Replace("'", "''");

                    // objOp.StatementDate = DateTime.ParseExact(dtStatementRegister_Info.Rows[k]["StatementDate"].ToString(), "dd/MM/yyyy", null);


                    #endregion

                    sql = "Insert into StatementRegister(SL,StatementNo,ClientName,ClientID,PAN,StartPage,EndPage,FileName,StatementDate,TotalPage,RefNo,Address,CONTRACTNO) " +
                    "Values('" + objOp.SL + "','" + objOp.StatementNo + "','" + objOp.ClientName + "','" + objOp.ClientID + "','" + objOp.PAN + "','" +
                    objOp.StartPage + "','" + objOp.EndPage + "','" + objOp.FileName + "','" + objOp.StatementDate + "','" + objOp.TotalPage + "','" + objOp.RefNo + "','" + objOp.Address + "','" + objOp.CONTRACTNO + "') ";

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
                objLW.logTrace(_LogPath, "Statement.log", ex.Message);
                return "Error: " + ex.StackTrace;
            }
        }

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

                #region setting properties values
                for (int k = 0; k < dtGetAccumIntAcc.Rows.Count; k++)
                {
                    objOp = new AccumIntAcc();

                    for (int j = 0; j < dtGetAccumIntAcc.Columns.Count; j++)
                    {
                        switch (dtGetAccumIntAcc.Columns[j].ColumnName)
                        {
                            case "StatementNo":
                                objOp.STATEMENTNO = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "ACCUM_INT_RRELEASE":
                                objOp.ACCUM_INT_RRELEASE = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "ACCUM_INT_EBALANCE":
                                objOp.ACCUM_INT_EBALANCE = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "ACCUM_INT_AMOUNT":
                                objOp.ACCUM_INT_AMOUNT = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "ACCUM_INT_SBALANCE":
                                objOp.ACCUM_INT_SBALANCE = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "ACCOUNT_NO":
                                objOp.ACCOUNT_NO = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                            case "AutoID":
                                objOp.AutoID = dtGetAccumIntAcc.Rows[k][j].ToString().Replace("'", "");
                                break;
                        }




                #endregion
                    }


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

                    #region setting properties values


                    for (int j = 0; j < dtAccount.Columns.Count; j++)
                    {
                        switch (dtAccount.Columns[j].ColumnName)
                        {
                            case "StatementNo":
                                objAc.STATEMENTNO = dtAccount.Rows[k][j].ToString();
                                break;
                            case "ACCOUNTNO":
                                objAc.ACCOUNTNO = dtAccount.Rows[k][j].ToString();
                                break;
                            case "ACURN":
                                objAc.ACURN = dtAccount.Rows[k][j].ToString();
                                break;
                            case "SBALANCE":
                                objAc.SBALANCE = dtAccount.Rows[k][j].ToString();
                                break;
                            case "ACURC":
                                objAc.ACURC = dtAccount.Rows[k][j].ToString();
                                break;
                            case "EBALANCE":
                                objAc.EBALANCE = dtAccount.Rows[k][j].ToString();
                                break;
                            case "AVAIL_CRD_LIMIT":
                                objAc.AVAIL_CRD_LIMIT = dtAccount.Rows[k][j].ToString();
                                break;
                            case "AVAIL_CASH_LIMIT":
                                objAc.AVAIL_CASH_LIMIT = dtAccount.Rows[k][j].ToString();
                                break;
                            case "SUM_WITHDRAWAL":
                                objAc.SUM_WITHDRAWAL = dtAccount.Rows[k][j].ToString();
                                break;
                            case "SUM_INTEREST":
                                objAc.SUM_INTEREST = dtAccount.Rows[k][j].ToString();
                                break;
                            case "OVLFEE_AMOUNT":
                                objAc.OVLFEE_AMOUNT = dtAccount.Rows[k][j].ToString();
                                break;
                            case "OVDFEE_AMOUNT":
                                objAc.OVDFEE_AMOUNT = dtAccount.Rows[k][j].ToString();
                                break;
                            case "SUM_REVERSE":
                                objAc.SUM_REVERSE = dtAccount.Rows[k][j].ToString();
                                break;
                            case "SUM_CREDIT":
                                objAc.SUM_CREDIT = dtAccount.Rows[k][j].ToString();
                                break;
                            case "SUM_OTHER":
                                objAc.SUM_OTHER = dtAccount.Rows[k][j].ToString();
                                break;
                            case "SUM_PURCHASE":
                                objAc.SUM_PURCHASE = dtAccount.Rows[k][j].ToString();
                                break;
                            case "MIN_AMOUNT_DUE":
                                objAc.MIN_AMOUNT_DUE = dtAccount.Rows[k][j].ToString();
                                break;
                            case "CASH_LIMIT":
                                objAc.CASH_LIMIT = dtAccount.Rows[k][j].ToString();
                                break;
                            case "CRD_LIMIT":
                                objAc.CRD_LIMIT = dtAccount.Rows[k][j].ToString();
                                break;
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

        private string GetCardHolderCardInfo(DataTable dtCard)
        {
            string reply = string.Empty;
            string sql = string.Empty;
            Card objCard = null;
            int @vSL = 0;
            CardList objCardList = new CardList();

            try
            {
                //Clear Previous Data
                objProvider.RunQuery("Delete from " + dtCard.TableName);

                for (int k = 0; k < dtCard.Rows.Count; k++)
                {
                    objCard = new Card();
                    @vSL = @vSL + 1;

                    for (int j = 0; j < dtCard.Columns.Count; j++)
                    {
                        #region setting properties values

                        switch (dtCard.Columns[j].ColumnName)
                        {
                            case "StatementNo":
                                objCard.STATEMENTNO = dtCard.Rows[k][j].ToString();
                                break;
                            case "PAN":
                                objCard.PAN = dtCard.Rows[k][j].ToString();
                                break;
                            case "MBR":
                                objCard.MBR = dtCard.Rows[k][j].ToString();
                                break;
                            case "CLIENTNAME":
                                objCard.CLIENTNAME = dtCard.Rows[k][j].ToString().Replace("'", "''");
                                break;

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

                try
                {
                    objSt = new StatementInfo();

                    objSt.BANK_CODE = BankName;

                    //for (int j = 0; j < dtStatement.Columns.Count; j++)
                    //{
                    #region setting properties values

                    if (dtStatement.Columns.Contains("STATEMENTNO"))
                    {
                        objSt.STATEMENTNO = dtStatement.Rows[k]["STATEMENTNO"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("CONTRACTNO"))
                    {
                        objSt.CONTRACTNO = dtStatement.Rows[k]["CONTRACTNO"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("IDCLIENT"))
                    {
                        objSt.IDCLIENT = dtStatement.Rows[k]["IDCLIENT"].ToString().Replace("'", "");
                    }

                    #region ADDRESS

                    if (dtStatement.Columns.Contains("ADDRESS"))
                    {
                        objSt.PROMOTIONALTEXT = dtStatement.Rows[k]["PromotionalText"].ToString().Replace("'", "''");

                        string value = objSt.PROMOTIONALTEXT;
                        string[] lines = value.Split(new char[] { '|' });
                        if (!string.IsNullOrEmpty(lines[0]))
                        {
                            objSt.INDICATOR = lines[0];
                        }
                        if (!string.IsNullOrEmpty(lines[3]))
                        {
                            objSt.COMPANYADDRESS1 = lines[3];
                        }
                        if (!string.IsNullOrEmpty(lines[4]))
                        {
                            objSt.COMPANYADDRESS2 = lines[4];
                        }
                        if (!string.IsNullOrEmpty(lines[5]))
                        {
                            objSt.CITYN = lines[5] +' '+lines[6];
                        }
                        //objSt.CITY = objSt.CITYN.Substring(0, objSt.CITYN.IndexOf("-"));

                        //objSt.ZIP = objSt.CITYN.Substring((objSt.CITY).Length + 1, objSt.CITYN.IndexOf("-"));
                        if (objSt.INDICATOR == "C")
                        {
                            objSt.ADDRESS = objSt.COMPANYADDRESS1 + ' ' + objSt.COMPANYADDRESS2 + ' ' + objSt.CITYN;
                        }
                        else
                        {
                            objSt.ADDRESS = dtStatement.Rows[k]["ADDRESS"].ToString().Replace("'", "''");
                        }

                    }

                    #endregion

                    if (dtStatement.Rows[k]["PAN"].ToString().Length >= 16)
                        objSt.PAN = dtStatement.Rows[k]["PAN"].ToString().Substring(0, 16);
                    else
                    {
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Card Not fount for the Contract " + objSt.CONTRACTNO });
                        MsgLogWriter objLW = new MsgLogWriter();
                        objLW.logTrace(_LogPath, "EStatement.log", "Card Not fount for the Contract " + objSt.CONTRACTNO);
                        continue;
                    }
                    if (dtStatement.Columns.Contains("REGION"))
                    {
                        objSt.CITY = dtStatement.Rows[k]["REGION"].ToString().Replace("'", "''");
                    }
                    if (dtStatement.Columns.Contains("ZIP"))
                    {
                        objSt.ZIP = dtStatement.Rows[k]["ZIP"].ToString().Replace("'", "''");
                    }
                    if (dtStatement.Columns.Contains("CITY"))
                    {
                        objSt.CITY = dtStatement.Rows[k]["City"].ToString().Replace("'", "''"); ;
                    }
                    if (dtStatement.Columns.Contains("COUNTRY"))
                    {
                        objSt.COUNTRY = dtStatement.Rows[k]["COUNTRY"].ToString().Replace("'", "''");
                    }
                    if (dtStatement.Columns.Contains("EMAIL"))
                    {
                        objSt.EMAIL = dtStatement.Rows[k]["EMAIL"].ToString().Replace("'", "''");
                    }
                    if (dtStatement.Columns.Contains("MOBILE"))
                    {
                        objSt.MOBILE = dtStatement.Rows[k]["MOBILE"].ToString().Replace("(", "").Replace(")", "").Replace("8800", "880");
                    }
                    if (dtStatement.Columns.Contains("TITLE"))
                    {
                        objSt.TITLE = dtStatement.Rows[k]["TITLE"].ToString().Replace("'", "''");
                    }

                    #region  JOBTITLE Update
                    if (dtStatement.Columns.Contains("JOBTITLE"))
                    {
                        objSt.PROMOTIONALTEXT = dtStatement.Rows[k]["PromotionalText"].ToString().Replace("'", "''");

                        string value = objSt.PROMOTIONALTEXT;
                        string[] lines = value.Split(new char[] { '|' });

                        if (!string.IsNullOrEmpty(lines[0]))
                        {
                            objSt.INDICATOR = lines[0];
                        }
                        if (objSt.INDICATOR == "C")
                        {
                            objSt.JOBTITLE = dtStatement.Rows[k]["JOBTITLE"].ToString().Replace("'", "''");
                        }
                        else
                        {
                            objSt.JOBTITLE = null;
                        }
                    }

                    #endregion


                    //#region  JOBTITLE Present
                    //if (dtStatement.Columns.Contains("JOBTITLE"))
                    //{
                    //    objSt.PROMOTIONALTEXT = dtStatement.Rows[k]["PromotionalText"].ToString().Replace("'", "''");

                    //    string value = objSt.PROMOTIONALTEXT;
                    //    string[] lines = value.Split(new char[] { '|' });

                    //    if (!string.IsNullOrEmpty(lines[0]))
                    //    {
                    //        objSt.INDICATOR = lines[0];
                    //    }
                    //    if (objSt.INDICATOR == "C")
                    //    {
                    //        if (!string.IsNullOrEmpty(lines[2]))
                    //        {
                    //            objSt.JOBTITLE = lines[2].TrimEnd(',');
                    //        }
                    //        else
                    //            objSt.JOBTITLE = null;
                    //    }
                    //    else
                    //    {
                    //        objSt.JOBTITLE = null;
                    //    }
                    //}

                    //#endregion

                    if (dtStatement.Columns.Contains("CLIENT"))
                    {
                        objSt.CLIENTNAME = dtStatement.Rows[k]["CLIENT"].ToString().Replace("'", "''");
                    }
                    if (dtStatement.Columns.Contains("ACCOUNTNO"))
                    {
                        objSt.ACCOUNTNO = dtStatement.Rows[k]["ACCOUNTNO"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("CURR"))
                    {
                        objSt.ACURN = dtStatement.Rows[k]["CURR"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("PBAL"))
                    {
                        objSt.SBALANCE = dtStatement.Rows[k]["PBAL"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("TOTINTEREST"))
                    {
                        objSt.SUM_INTEREST = dtStatement.Rows[k]["TOTINTEREST"].ToString();
                    }
                    if (dtStatement.Columns.Contains("STARTDATE"))
                    {
                        objSt.STARTDATE = dtStatement.Rows[k]["STARTDATE"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("ENDDATE"))
                    {
                        objSt.ENDDATE = dtStatement.Rows[k]["ENDDATE"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("NEXT_STATEMENT_DATE"))
                    {
                        objSt.NEXT_STATEMENT_DATE = dtStatement.Rows[k]["NEXT_STATEMENT_DATE"].ToString();
                    }
                    if (dtStatement.Columns.Contains("PAYDATE"))
                    {
                        objSt.PAYMENT_DATE = dtStatement.Rows[k]["PAYDATE"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("STDATE"))
                    {
                        objSt.STATEMENT_DATE = dtStatement.Rows[k]["STDATE"].ToString();
                    }
                    if (dtStatement.Columns.Contains("STDATE"))
                    {
                        objSt.STATEMENTID = dtStatement.Rows[k]["STDATE"].ToString().Replace("/", ""); ;
                    }
                    if (dtStatement.Columns.Contains("ACURC"))
                    {
                        objSt.ACURC = dtStatement.Rows[k]["ACURC"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("OVLFEE_AMOUNT"))
                    {
                        objSt.OVLFEE_AMOUNT = dtStatement.Rows[k]["OVLFEE_AMOUNT"].ToString().Replace("-", "");
                    }
                    if (dtStatement.Columns.Contains("ODAMOUNT"))
                    {
                        objSt.OVDFEE_AMOUNT = dtStatement.Rows[k]["ODAMOUNT"].ToString().Replace("-", "");
                    }
                    if (dtStatement.Columns.Contains("MINPAY"))
                    {
                        objSt.MIN_AMOUNT_DUE = dtStatement.Rows[k]["MINPAY"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("TOTLIMIT"))
                    {
                        objSt.CRD_LIMIT = dtStatement.Rows[k]["TOTLIMIT"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("TOTPURCHASE"))
                    {
                        objSt.SUM_PURCHASE = dtStatement.Rows[k]["TOTPURCHASE"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("SUM_REVERSE"))
                    {
                        objSt.SUM_REVERSE = dtStatement.Rows[k]["SUM_REVERSE"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("SUM_CREDIT"))
                    {
                        objSt.SUM_CREDIT = dtStatement.Rows[k]["SUM_CREDIT"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("SUM_OTHER"))
                    {
                        objSt.SUM_OTHER = dtStatement.Rows[k]["SUM_OTHER"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("CASHADV"))
                    {
                        objSt.SUM_WITHDRAWAL = dtStatement.Rows[k]["CASHADV"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("AVLIMIT"))
                    {
                        objSt.AVAIL_CRD_LIMIT = dtStatement.Rows[k]["AVLIMIT"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("AVCASHLIMIT"))
                    {
                        objSt.AVAIL_CASH_LIMIT = dtStatement.Rows[k]["AVCASHLIMIT"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("LASTBAL"))
                    {
                        objSt.EBALANCE = dtStatement.Rows[k]["LASTBAL"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("CASH_LIMIT"))
                    {
                        objSt.CASH_LIMIT = dtStatement.Rows[k]["CASH_LIMIT"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("PromotionalText"))
                    {
                        objSt.PROMOTIONALTEXT = dtStatement.Rows[k]["PromotionalText"].ToString().Replace("'", "''");
                    }

                    #region PromotionalText

                    if (dtStatement.Columns.Contains("PromotionalText"))
                    {
                        objSt.PROMOTIONALTEXT = dtStatement.Rows[k]["PromotionalText"].ToString().Replace("'", "''");


                        //objSt.PROMOTIONALTEXT = dtStatement.Rows[k][j].ToString().Replace("'", "");

                        string value = objSt.PROMOTIONALTEXT;
                        string[] lines = value.Split(new char[] { '|' });

                        if (!string.IsNullOrEmpty(lines[0]))
                        {
                            objSt.INDICATOR = lines[0];
                        }
                        if (objSt.INDICATOR == "C")
                        {
                            if (!string.IsNullOrEmpty(lines[1]))
                            {
                                objSt.PROMOTIONALTEXT = lines[1].TrimEnd(',');
                            }
                        }
                        else
                        {
                            objSt.PROMOTIONALTEXT = null;
                        }
                    }

                    #endregion




                    #endregion
                    // }

                    objSt.STM_MSG = "";//txtStmMsg.Text;
                    objSt.STATUS = "1";

                    sql = "Insert into STATEMENT_INFO(STATEMENTID,BANK_CODE,CONTRACTNO,IDCLIENT,PAN,TITLE,CLIENTNAME,JOBTITLE,STATEMENTNO,ADDRESS,CITY,ZIP,COUNTRY," +
                        "EMAIL,MOBILE,STARTDATE,ENDDATE,NEXT_STATEMENT_DATE,PAYMENT_DATE,STATEMENT_DATE,ACCOUNTNO,ACURN,SBALANCE,ACURC,EBALANCE,AVAIL_CRD_LIMIT," +
                        "AVAIL_CASH_LIMIT,SUM_WITHDRAWAL,SUM_INTEREST,OVLFEE_AMOUNT,OVDFEE_AMOUNT,SUM_REVERSE,SUM_CREDIT,SUM_OTHER,SUM_PURCHASE," +
                        "MIN_AMOUNT_DUE,CASH_LIMIT,CRD_LIMIT,STM_MSG,STATUS,PROMOTIONALTEXT) VALUES('" + objSt.STATEMENTID + "'," +
                        "'" + objSt.BANK_CODE + "','" + objSt.CONTRACTNO + "','" + objSt.IDCLIENT + "','" + objSt.PAN + "','" + objSt.TITLE + "','" + objSt.CLIENTNAME + "','" + objSt.JOBTITLE + "','" + objSt.STATEMENTNO + "'," +
                        "'" + objSt.ADDRESS + "','" + objSt.CITY + "','" + objSt.ZIP + "','" + objSt.COUNTRY + "','" + objSt.EMAIL + "','" + objSt.MOBILE + "','" + objSt.STARTDATE + "','" + objSt.ENDDATE + "'," +
                        "'" + objSt.NEXT_STATEMENT_DATE + "','" + objSt.PAYMENT_DATE + "','" + objSt.STATEMENT_DATE + "','" + objSt.ACCOUNTNO + "','" + objSt.ACURN + "'," +
                        "'" + objSt.SBALANCE + "','" + objSt.ACURC + "','" + objSt.EBALANCE + "','" + objSt.AVAIL_CRD_LIMIT + "','" + objSt.AVAIL_CASH_LIMIT + "'," +
                        "'" + objSt.SUM_WITHDRAWAL + "','" + objSt.SUM_INTEREST + "','" + objSt.OVLFEE_AMOUNT + "','" + objSt.OVDFEE_AMOUNT + "','" + objSt.SUM_REVERSE + "'," +
                        "'" + objSt.SUM_CREDIT + "','" + objSt.SUM_OTHER + "','" + objSt.SUM_PURCHASE + "','" + objSt.MIN_AMOUNT_DUE + "','" + objSt.CASH_LIMIT + "'," +
                        "'" + objSt.CRD_LIMIT + "','" + objSt.STM_MSG + "','" + objSt.STATUS + "','" + objSt.PROMOTIONALTEXT + "')";

                    reply = objProvider.RunQuery(sql);
                    //DataTable dtOperation = dsStatement.Tables["Operation"];

                    if (dtOperation != null && dtOperation.Columns.Contains("ACCOUNT"))

                    // if (dtOperation != null &&   !string.IsNullOrEmpty(dtOperation.Columns.Contains("ACCOUNT")))
                    {
                        if (dtOperation.Rows.Count > 0)
                        {

                            DataRow[] dr = dtOperation.Select("STATEMENTNO='" + objSt.STATEMENTNO + "' AND ACCOUNT='" + objSt.ACCOUNTNO + "'");
                            if (dr.Length > 0)
                            {

                                string trn_Date = string.Empty;

                                for (int l = 0; l < dr.Length; l++)
                                {
                                    #region setting properties values
                                    List<string> INTlist = new List<string>() { "INTEREST ON FEES & CHARGES","INTEREST ON OTHERS","INTEREST ON INTEREST","INTEREST ON ATM TRANSACTION", "INTEREST ON POS TRANSACTION", "INTEREST ON CARD CHEQUE","CHARGE INTEREST FOR 0", "CHARGE INTEREST FOR 1", "CHARGE INTEREST FOR 2", "CHARGE INTEREST FOR 3", "CHARGE INTEREST FOR 4", "CHARGE INTEREST FOR 5", "CHARGE INTEREST FOR 6", "CHARGE INTEREST FOR 7", "CHARGE INTEREST FOR 8", "CHARGE INTEREST FOR 9", "CHARGE INTEREST FOR 10", "CHARGE INTEREST FOR 11", "CHARGE INTEREST FOR 0 OPERATIONS GROUP", "CHARGE INTEREST FOR 1 OPERATIONS GROUP", "CHARGE INTEREST FOR 2 OPERATIONS GROUP", "CHARGE INTEREST FOR 3 OPERATIONS GROUP", "CHARGE INTEREST FOR 4 OPERATIONS GROUP", "CHARGE INTEREST FOR 5 OPERATIONS GROUP", "CHARGE INTEREST FOR 6 OPERATIONS GROUP", "CHARGE INTEREST FOR 7 OPERATIONS GROUP", "INTEREST ON FUND TRANSFER", "INTEREST ON BALANCE TRANSFER", "INTEREST ON EMI", "INTEREST ON FT", "INTEREST ON BT", "INTEREST ON BANK POS TRANSACTION",
                                    "INTEREST ON BPOS TRANSACTION","CHARGE INTEREST FOR INTEREST OPERATIONS", "CHARGE INTEREST FOR POS OPERATIONS", "CHARGE INTEREST FOR ATM OPERATIONS", "LATE PAYMENT CHARGE FOR GROUP 1", "LATE PAYMENT CHARGE FOR GROUP 2", "LATE PAYMENT CHARGE FOR GROUP 3", "CHARGE OF A DEBT FOR CREDIT OVERDRAFTING" ,"INTEREST ON SERVICE FEE","INTEREST ON PREVIOUS BALANCE","REVOLVING INTEREST CHARGE"};
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
                                            //objSTD.TRNDATE = dr[l]["TD"].ToString();
                                        objSTD.TRNDATE = FormatDate(dr[l]["TD"].ToString());

                                        if (dr[l].Table.Columns.Contains("OD"))
                                          //  objSTD.POSTDATE = dr[l]["OD"].ToString();
                                        objSTD.POSTDATE = FormatDate(dr[l]["OD"].ToString());

                                        /*     if (dr[l].Table.Columns.Contains("TD"))
                                             {
                                                 if (dr[l]["TD"].ToString() != "")
                                                 {
                                                     objSTD.TRNDATE = dr[l]["TD"].ToString();
                                                     objSTD.POSTDATE = dr[l]["OD"].ToString();
                                                 }

                                                 else
                                                 {

                                                     objSTD.POSTDATE = dr[l]["OD"].ToString();
                                                     objSTD.TRNDATE = dr[l]["OD"].ToString();  // UPDATE

                                                 }

                                             }

                                             else
                                             {

                                                 objSTD.POSTDATE = dr[l]["OD"].ToString();
                                                 objSTD.TRNDATE = dr[l]["OD"].ToString();  // UPDATE

                                             }  */

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
                                            objSTD.AMOUNTSIGN = dr[l]["AMOUNTSIGN"].ToString().Replace("'", "");
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
                                            if (dr[l]["OA"].ToString() == "" || dr[l]["OA"].ToString() == null)
                                                objSTD.ORGAMOUNT = "0.00";
                                            else
                                                objSTD.ORGAMOUNT = dr[l]["OA"].ToString();
                                        }
                                        else objSTD.ORGAMOUNT = "0.00";

                                        //Remmove Terminal Name when Fee and VAT Impose
                                        //Sum Charges amount with Fees & Charges. 

                                        #region

                                        if ((dr[l]["D"].ToString().ToUpper().Trim() == "CREDIT SHIELD PREMIUM") || (dr[l]["DE"].ToString().ToUpper().Contains("FEE")) || (dr[l]["D"].ToString().ToUpper().Trim() == "CHARGE INTEREST FOR INSTALLMENT") || (dr[l]["D"].ToString().ToUpper().Trim() == "MONTHLY INSTALLMENT"))
                                        {
                                            objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                            //feesnCharges = feesnCharges + Convert.ToDouble(dr[l]["A"].ToString().Replace("-", ""));
                                            if (dr[l].Table.Columns.Contains("OD"))
                                            {
                                                //objSTD.TRNDATE = dr[l]["OD"].ToString();
                                                objSTD.TRNDATE = FormatDate(dr[l]["OD"].ToString());
                                            }

                                        }

                                        else
                                        {
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
                                                    //  objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''"); // modify

                                            }
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
                                            else
                                            {
                                                if (dr[l].Table.Columns.Contains("TL"))
                                                {
                                                    if (dr[l]["FR"].ToString().ToUpper().Contains("A 10") || dr[l]["FR"].ToString().ToUpper().Contains("A 64") || dr[l]["FR"].ToString().ToUpper().Contains("P 14") || dr[l]["FR"].ToString().ToUpper().Contains("P 32") || dr[l]["FR"].ToString().ToUpper().Contains("P 33") || dr[l]["FR"].ToString().ToUpper().Contains("F 29") || dr[l]["FR"].ToString().ToUpper().Contains("P 13"))
                                                    {
                                                        objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");

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
                                                      }  */


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

                                        }


                                        if (objSTD.TRNDESC.ToUpper().Contains("CREDIT CASH DEPOSIT"))
                                        {

                                            if (dr[l]["FR"].ToString() == "" || dr[l]["FR"].ToString() == null)
                                            {
                                                objSTD.TRNDESC = "PAYMENT RECEIVED [THANK YOU]";
                                                //objSTD.TRNDATE = dr[l]["OD"].ToString();
                                                objSTD.TRNDATE = FormatDate(dr[l]["OD"].ToString());
                                            }
                                            else
                                            {
                                                objSTD.TRNDESC = dr[l]["FR"].ToString().Replace("'", "''");
                                               // objSTD.TRNDATE = dr[l]["OD"].ToString();
                                                objSTD.TRNDATE = FormatDate(dr[l]["OD"].ToString());
                                            }

                                        }

                                        if (dr[l].Table.Columns.Contains("APPROVAL"))
                                        {
                                            objSTD.APPROVAL = dr[l]["APPROVAL"].ToString().Replace("'", "");

                                            if (dr[l]["APPROVAL"].ToString() != "" && objSTD.TRNDATE == "")
                                            {
                                               // objSTD.TRNDATE = dr[l]["OD"].ToString();
                                                objSTD.TRNDATE = FormatDate(dr[l]["OD"].ToString());
                                            }
                                        }

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
                                            //objSTD.TRNDATE = dr[l]["TD"].ToString();
                                        objSTD.TRNDATE = FormatDate(dr[l]["TD"].ToString());

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

                                    #endregion
                                }

                                //New View add
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
                                    // objSTD.TRNDATE = objSt.STATEMENT_DATE;  // only for NBL
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

                                }
                                // end IF for SUM_interest
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
                                                            // objSTD.TRNDATE = objSTD.TRNDATE; only for NBL
                                                            objSTD.POSTDATE = objSTD.POSTDATE;
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
                                }   // END ELSE 



                                //New View add

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
                                                        // objSTD.TRNDATE = objSTD.TRNDATE; only for NBL
                                                        objSTD.POSTDATE = objSTD.POSTDATE;
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
                            }   // END ELSE 





                        }
                    }


                    else
                    {

                        if (dtOperation.Rows.Count > 0)
                        {
                            DataRow[] dr = dtOperation.Select("STATEMENTNO='" + objSt.STATEMENTNO + "'");
                            if (dr.Length > 0)
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
                                                        // objSTD.TRNDATE = objSTD.TRNDATE; only for NBL
                                                        objSTD.POSTDATE = objSTD.POSTDATE;
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
                    }   // END ELSE 

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
                    objSt = new StatementInfo();

                    objSt.BANK_CODE = BankName;

                    //for (int j = 0; j < dtStatement.Columns.Count; j++)
                    //{
                    #region setting properties values


                    //objSt.STATEMENTNO = dtStatement.Rows[k]["STATEMENTNO"].ToString().Replace("'", "");

                    //objSt.CONTRACTNO = dtStatement.Rows[k]["CONTRACTNO"].ToString();

                    //objSt.IDCLIENT = dtStatement.Rows[k]["IDCLIENT"].ToString().Replace("'", "");

                    //objSt.ADDRESS = dtStatement.Rows[k]["ADDRESS"].ToString().Replace("'", "");
                    if (dtStatement.Columns.Contains("STATEMENTNO"))
                    {
                        objSt.STATEMENTNO = dtStatement.Rows[k]["STATEMENTNO"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("CONTRACTNO"))
                    {
                        objSt.CONTRACTNO = dtStatement.Rows[k]["CONTRACTNO"].ToString();
                    }
                    if (dtStatement.Columns.Contains("IDCLIENT"))
                    {
                        objSt.IDCLIENT = dtStatement.Rows[k]["IDCLIENT"].ToString().Replace("'", "");
                    }

                    #region ADDRESS

                    if (dtStatement.Columns.Contains("ADDRESS"))
                    {
                        objSt.PROMOTIONALTEXT = dtStatement.Rows[k]["PromotionalText"].ToString().Replace("'", "''");

                        string value = objSt.PROMOTIONALTEXT;
                        string[] lines = value.Split(new char[] { '|' });
                        if (!string.IsNullOrEmpty(lines[0]))
                        {
                            objSt.INDICATOR = lines[0];
                        }
                        if (!string.IsNullOrEmpty(lines[3]))
                        {
                            objSt.COMPANYADDRESS1 = lines[3];
                        }
                        if (!string.IsNullOrEmpty(lines[4]))
                        {
                            objSt.COMPANYADDRESS2 = lines[4];
                        }
                        if (!string.IsNullOrEmpty(lines[5]))
                        {
                            objSt.CITYN = lines[5] + ' ' + lines[6];
                        }
                        //objSt.CITY = objSt.CITYN.Substring(0, objSt.CITYN.IndexOf("-"));

                        //objSt.ZIP = objSt.CITYN.Substring((objSt.CITY).Length + 1, objSt.CITYN.IndexOf("-"));
                        if (objSt.INDICATOR == "C")
                        {
                            objSt.ADDRESS = objSt.COMPANYADDRESS1 + ' ' + objSt.COMPANYADDRESS2 + ' ' + objSt.CITYN;
                        }
                        else
                        {
                            objSt.ADDRESS = dtStatement.Rows[k]["ADDRESS"].ToString().Replace("'", "''");
                        }

                    }
                    #endregion

                    if (dtStatement.Rows[k]["PAN"].ToString().Length >= 16)
                        objSt.PAN = dtStatement.Rows[k]["PAN"].ToString().Substring(0, 16);
                    else
                    {
                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Card Not fount for the Contract " + objSt.CONTRACTNO });
                        MsgLogWriter objLW = new MsgLogWriter();
                        objLW.logTrace(_LogPath, "EStatement.log", "Card Not fount for the Contract " + objSt.CONTRACTNO);
                        continue;
                    }
                    if (dtStatement.Columns.Contains("REGION"))
                    {
                        objSt.CITY = dtStatement.Rows[k]["REGION"].ToString().Replace("'", "''");
                    }
                    if (dtStatement.Columns.Contains("ZIP"))
                    {
                        objSt.ZIP = dtStatement.Rows[k]["ZIP"].ToString().Replace("'", "''");
                    }
                    if (dtStatement.Columns.Contains("CITY"))
                    {
                        objSt.CITY = dtStatement.Rows[k]["City"].ToString().Replace("'", "''"); ;
                    }
                    if (dtStatement.Columns.Contains("COUNTRY"))
                    {
                        objSt.COUNTRY = dtStatement.Rows[k]["COUNTRY"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("EMAIL"))
                    {
                        objSt.EMAIL = dtStatement.Rows[k]["EMAIL"].ToString().Replace("'", "''");
                    }
                    if (dtStatement.Columns.Contains("MOBILE"))
                    {
                        objSt.MOBILE = dtStatement.Rows[k]["MOBILE"].ToString().Replace("(", "").Replace(")", "").Replace("8800", "880");
                    }
                    if (dtStatement.Columns.Contains("TITLE"))
                    {
                        objSt.TITLE = dtStatement.Rows[k]["TITLE"].ToString().Replace("'", "''");
                    }

                    #region JOBTITLE Update

                    if (dtStatement.Columns.Contains("JOBTITLE"))
                    {
                        objSt.PROMOTIONALTEXT = dtStatement.Rows[k]["PromotionalText"].ToString().Replace("'", "''");

                        string value = objSt.PROMOTIONALTEXT;
                        string[] lines = value.Split(new char[] { '|' });

                        if (!string.IsNullOrEmpty(lines[0]))
                        {
                            objSt.INDICATOR = lines[0];
                        }
                        if (objSt.INDICATOR == "C")
                        {

                            objSt.JOBTITLE = dtStatement.Rows[k]["JOBTITLE"].ToString().Replace("'", "''");
                        }
                        else
                        {
                            objSt.JOBTITLE = null;
                        }
                    }

                    #endregion

                    //#region JOBTITLE Present
                    //if (dtStatement.Columns.Contains("JOBTITLE"))
                    //{
                    //    objSt.PROMOTIONALTEXT = dtStatement.Rows[k]["PromotionalText"].ToString().Replace("'", "''");

                    //    string value = objSt.PROMOTIONALTEXT;
                    //    string[] lines = value.Split(new char[] { '|' });

                    //    if (!string.IsNullOrEmpty(lines[0]))
                    //    {
                    //        objSt.INDICATOR = lines[0];
                    //    }
                    //    if (objSt.INDICATOR == "C")
                    //    {
                    //        if (!string.IsNullOrEmpty(lines[2]))
                    //        {
                    //            objSt.JOBTITLE = lines[2].TrimEnd(',');
                    //        }
                    //        else
                    //            objSt.JOBTITLE = null;
                    //    }
                    //    else
                    //    {
                    //        objSt.JOBTITLE = null;
                    //    }
                    //}

                    //#endregion
                    if (dtStatement.Columns.Contains("CLIENT"))
                    {
                        objSt.CLIENTNAME = dtStatement.Rows[k]["CLIENT"].ToString().Replace("'", "''");
                    }
                    if (dtStatement.Columns.Contains("ACCOUNTNO"))
                    {
                        objSt.ACCOUNTNO = dtStatement.Rows[k]["ACCOUNTNO"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("CURR"))
                    {
                        objSt.ACURN = dtStatement.Rows[k]["CURR"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("PBAL"))
                    {
                        objSt.SBALANCE = dtStatement.Rows[k]["PBAL"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("TOTINTEREST"))
                    {
                        objSt.SUM_INTEREST = dtStatement.Rows[k]["TOTINTEREST"].ToString();
                    }
                    if (dtStatement.Columns.Contains("STARTDATE"))
                    {
                        objSt.STARTDATE = dtStatement.Rows[k]["STARTDATE"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("ENDDATE"))
                    {
                        objSt.ENDDATE = dtStatement.Rows[k]["ENDDATE"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("NEXT_STATEMENT_DATE"))
                    {
                        objSt.NEXT_STATEMENT_DATE = dtStatement.Rows[k]["NEXT_STATEMENT_DATE"].ToString();
                    }
                    if (dtStatement.Columns.Contains("PAYDATE"))
                    {
                        objSt.PAYMENT_DATE = dtStatement.Rows[k]["PAYDATE"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("STDATE"))
                    {
                        objSt.STATEMENT_DATE = dtStatement.Rows[k]["STDATE"].ToString();
                    }
                    if (dtStatement.Columns.Contains("STDATE"))
                    {
                        objSt.STATEMENTID = dtStatement.Rows[k]["STDATE"].ToString().Replace("/", ""); ;
                    }
                    if (dtStatement.Columns.Contains("ACURC"))
                    {
                        objSt.ACURC = dtStatement.Rows[k]["ACURC"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("OVLFEE_AMOUNT"))
                    {
                        objSt.OVLFEE_AMOUNT = dtStatement.Rows[k]["OVLFEE_AMOUNT"].ToString().Replace("-", "");
                    }
                    if (dtStatement.Columns.Contains("ODAMOUNT"))
                    {
                        objSt.OVDFEE_AMOUNT = dtStatement.Rows[k]["ODAMOUNT"].ToString().Replace("-", "");
                    }
                    if (dtStatement.Columns.Contains("MINPAY"))
                    {
                        objSt.MIN_AMOUNT_DUE = dtStatement.Rows[k]["MINPAY"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("TOTLIMIT"))
                    {
                        objSt.CRD_LIMIT = dtStatement.Rows[k]["TOTLIMIT"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("TOTPURCHASE"))
                    {
                        objSt.SUM_PURCHASE = dtStatement.Rows[k]["TOTPURCHASE"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("SUM_REVERSE"))
                    {
                        objSt.SUM_REVERSE = dtStatement.Rows[k]["SUM_REVERSE"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("SUM_CREDIT"))
                    {
                        objSt.SUM_CREDIT = dtStatement.Rows[k]["SUM_CREDIT"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("SUM_OTHER"))
                    {
                        objSt.SUM_OTHER = dtStatement.Rows[k]["SUM_OTHER"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("CASHADV"))
                    {
                        objSt.SUM_WITHDRAWAL = dtStatement.Rows[k]["CASHADV"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("AVLIMIT"))
                    {
                        objSt.AVAIL_CRD_LIMIT = dtStatement.Rows[k]["AVLIMIT"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("AVCASHLIMIT"))
                    {
                        objSt.AVAIL_CASH_LIMIT = dtStatement.Rows[k]["AVCASHLIMIT"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("LASTBAL"))
                    {
                        objSt.EBALANCE = dtStatement.Rows[k]["LASTBAL"].ToString().Replace("'", "");
                    }
                    if (dtStatement.Columns.Contains("CASH_LIMIT"))
                    {
                        objSt.CASH_LIMIT = dtStatement.Rows[k]["CASH_LIMIT"].ToString().Replace("'", "");
                    }

                    #region PromotionalText
                    if (dtStatement.Columns.Contains("PromotionalText"))
                    {
                        objSt.PROMOTIONALTEXT = dtStatement.Rows[k]["PromotionalText"].ToString().Replace("'", "''");


                        //objSt.PROMOTIONALTEXT = dtStatement.Rows[k][j].ToString().Replace("'", "");

                        string value = objSt.PROMOTIONALTEXT;
                        string[] lines = value.Split(new char[] { '|' });

                        if (!string.IsNullOrEmpty(lines[0]))
                        {
                            objSt.INDICATOR = lines[0];
                        }
                        if (objSt.INDICATOR == "C")
                        {
                            if (!string.IsNullOrEmpty(lines[1]))
                            {
                                objSt.PROMOTIONALTEXT = lines[1];
                            }
                        }
                        else
                        {
                            objSt.PROMOTIONALTEXT = null;
                        }
                    }

                    #endregion




                    #endregion
                    // }

                    objSt.STM_MSG = "";//txtStmMsg.Text;
                    objSt.STATUS = "1";

                    sql = "Insert into STATEMENT_INFO(STATEMENTID,BANK_CODE,CONTRACTNO,IDCLIENT,PAN,TITLE,CLIENTNAME,JOBTITLE,STATEMENTNO,ADDRESS,CITY,ZIP,COUNTRY," +
                        "EMAIL,MOBILE,STARTDATE,ENDDATE,NEXT_STATEMENT_DATE,PAYMENT_DATE,STATEMENT_DATE,ACCOUNTNO,ACURN,SBALANCE,ACURC,EBALANCE,AVAIL_CRD_LIMIT," +
                        "AVAIL_CASH_LIMIT,SUM_WITHDRAWAL,SUM_INTEREST,OVLFEE_AMOUNT,OVDFEE_AMOUNT,SUM_REVERSE,SUM_CREDIT,SUM_OTHER,SUM_PURCHASE," +
                        "MIN_AMOUNT_DUE,CASH_LIMIT,CRD_LIMIT,STM_MSG,STATUS,PROMOTIONALTEXT) VALUES('" + objSt.STATEMENTID + "'," +
                        "'" + objSt.BANK_CODE + "','" + objSt.CONTRACTNO + "','" + objSt.IDCLIENT + "','" + objSt.PAN + "','" + objSt.TITLE + "','" + objSt.CLIENTNAME + "','" + objSt.JOBTITLE + "','" + objSt.STATEMENTNO + "'," +
                        "'" + objSt.ADDRESS + "','" + objSt.CITY + "','" + objSt.ZIP + "','" + objSt.COUNTRY + "','" + objSt.EMAIL + "','" + objSt.MOBILE + "','" + objSt.STARTDATE + "','" + objSt.ENDDATE + "'," +
                        "'" + objSt.NEXT_STATEMENT_DATE + "','" + objSt.PAYMENT_DATE + "','" + objSt.STATEMENT_DATE + "','" + objSt.ACCOUNTNO + "','" + objSt.ACURN + "'," +
                        "'" + objSt.SBALANCE + "','" + objSt.ACURC + "','" + objSt.EBALANCE + "','" + objSt.AVAIL_CRD_LIMIT + "','" + objSt.AVAIL_CASH_LIMIT + "'," +
                        "'" + objSt.SUM_WITHDRAWAL + "','" + objSt.SUM_INTEREST + "','" + objSt.OVLFEE_AMOUNT + "','" + objSt.OVDFEE_AMOUNT + "','" + objSt.SUM_REVERSE + "'," +
                        "'" + objSt.SUM_CREDIT + "','" + objSt.SUM_OTHER + "','" + objSt.SUM_PURCHASE + "','" + objSt.MIN_AMOUNT_DUE + "','" + objSt.CASH_LIMIT + "'," +
                        "'" + objSt.CRD_LIMIT + "','" + objSt.STM_MSG + "','" + objSt.STATUS + "','" + objSt.PROMOTIONALTEXT + "')";

                    reply = objProvider.RunQuery(sql);
                    //DataTable dtOperation = dsStatement.Tables["Operation"];

                    if (dtOperation != null && dtOperation.Columns.Contains("ACCOUNT"))
                    {
                        if (dtOperation.Rows.Count > 0)
                        {

                            DataRow[] dr = dtOperation.Select("STATEMENTNO='" + objSt.STATEMENTNO + "' AND ACCOUNT='" + objSt.ACCOUNTNO + "'");
                            if (dr.Length > 0)
                            {

                                string trn_Date = string.Empty;

                                for (int l = 0; l < dr.Length; l++)
                                {
                                    #region setting properties values
                                    List<string> INTlist = new List<string>() { "INTEREST ON FEES & CHARGES","INTEREST ON OTHERS","INTEREST ON INTEREST","INTEREST ON ATM TRANSACTION", "INTEREST ON POS TRANSACTION", "INTEREST ON CARD CHEQUE","CHARGE INTEREST FOR 0", "CHARGE INTEREST FOR 1", "CHARGE INTEREST FOR 2", "CHARGE INTEREST FOR 3", "CHARGE INTEREST FOR 4", "CHARGE INTEREST FOR 5", "CHARGE INTEREST FOR 6", "CHARGE INTEREST FOR 7", "CHARGE INTEREST FOR 8", "CHARGE INTEREST FOR 9", "CHARGE INTEREST FOR 10", "CHARGE INTEREST FOR 11", "CHARGE INTEREST FOR 0 OPERATIONS GROUP", "CHARGE INTEREST FOR 1 OPERATIONS GROUP", "CHARGE INTEREST FOR 2 OPERATIONS GROUP", "CHARGE INTEREST FOR 3 OPERATIONS GROUP", "CHARGE INTEREST FOR 4 OPERATIONS GROUP", "CHARGE INTEREST FOR 5 OPERATIONS GROUP", "CHARGE INTEREST FOR 6 OPERATIONS GROUP", "CHARGE INTEREST FOR 7 OPERATIONS GROUP", "INTEREST ON FUND TRANSFER", "INTEREST ON BALANCE TRANSFER", "INTEREST ON EMI", "INTEREST ON FT", "INTEREST ON BT", "INTEREST ON BANK POS TRANSACTION",
                                    "INTEREST ON BPOS TRANSACTION","CHARGE INTEREST FOR INTEREST OPERATIONS", "CHARGE INTEREST FOR POS OPERATIONS", "CHARGE INTEREST FOR ATM OPERATIONS", "LATE PAYMENT CHARGE FOR GROUP 1", "LATE PAYMENT CHARGE FOR GROUP 2", "LATE PAYMENT CHARGE FOR GROUP 3", "CHARGE OF A DEBT FOR CREDIT OVERDRAFTING" ,"INTEREST ON SERVICE FEE","INTEREST ON PREVIOUS BALANCE","REVOLVING INTEREST CHARGE"};
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
                                            //objSTD.TRNDATE = dr[l]["TD"].ToString();
                                              objSTD.TRNDATE = FormatDate(dr[l]["TD"].ToString());

                                        if (dr[l].Table.Columns.Contains("OD"))
                                          //  objSTD.POSTDATE = dr[l]["OD"].ToString();
                                              objSTD.POSTDATE = FormatDate(dr[l]["OD"].ToString());

                                        /*    if (dr[l].Table.Columns.Contains("TD"))
                                            {
                                                if (dr[l]["TD"].ToString() != "")
                                                {
                                                    objSTD.TRNDATE = dr[l]["TD"].ToString();
                                                    objSTD.POSTDATE = dr[l]["OD"].ToString();
                                                }

                                                else
                                                {

                                                    objSTD.POSTDATE = dr[l]["OD"].ToString();
                                                    objSTD.TRNDATE = dr[l]["OD"].ToString();  // UPDATE

                                                }

                                            }

                                            else
                                            {

                                                objSTD.POSTDATE = dr[l]["OD"].ToString();
                                                objSTD.TRNDATE = dr[l]["OD"].ToString();  // UPDATE

                                            }  */

                                        if (dr[l].Table.Columns.Contains("ACURN"))
                                            objSTD.ACURN = dr[l]["ACURN"].ToString();

                                        if (dr[l].Table.Columns.Contains("FR"))
                                            objSTD.FR = dr[l]["FR"].ToString().Replace("'", "''");

                                        if (dr[l].Table.Columns.Contains("DE"))
                                            objSTD.DE = dr[l]["DE"].ToString().Replace("'", "''");

                                        if (dr[l].Table.Columns.Contains("SERIALNO"))
                                            objSTD.SERIALNO = dr[l]["SERIALNO"].ToString().Replace("'", "");

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

                                        #region

                                        if ((dr[l]["D"].ToString().ToUpper().Trim() == "CREDIT SHIELD PREMIUM") || (dr[l]["DE"].ToString().ToUpper().Contains("FEE")) || (dr[l]["D"].ToString().ToUpper().Trim() == "CHARGE INTEREST FOR INSTALLMENT") || (dr[l]["D"].ToString().ToUpper().Trim() == "MONTHLY INSTALLMENT"))
                                        {
                                            objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''");
                                            // feesnCharges = feesnCharges + Convert.ToDouble(dr[l]["A"].ToString().Replace("-", ""));
                                            if (dr[l].Table.Columns.Contains("OD"))
                                            {
                                                //objSTD.TRNDATE = dr[l]["OD"].ToString();
                                                  objSTD.TRNDATE = FormatDate(dr[l]["OD"].ToString());
                                            }

                                        }

                                        else
                                        {
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
                                                    //objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");
                                                    objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''"); // modify

                                            }
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
                                            else
                                            {
                                                if (dr[l].Table.Columns.Contains("TL"))
                                                {
                                                    if (dr[l]["FR"].ToString().ToUpper().Contains("P 10") || dr[l]["FR"].ToString().ToUpper().Contains("A 10") || dr[l]["FR"].ToString().ToUpper().Contains("A 64") || dr[l]["FR"].ToString().ToUpper().Contains("P 14") || dr[l]["FR"].ToString().ToUpper().Contains("P 96") || dr[l]["FR"].ToString().ToUpper().Contains("P 97") || dr[l]["FR"].ToString().ToUpper().Contains("P 32") || dr[l]["FR"].ToString().ToUpper().Contains("P 33") || dr[l]["FR"].ToString().ToUpper().Contains("F 29") || dr[l]["FR"].ToString().ToUpper().Contains("P 13"))
                                                    {
                                                        objSTD.TRNDESC = dr[l]["D"].ToString().Replace("'", "''") + " " + dr[l]["TL"].ToString().Replace("'", "''");

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

                                        }

                                        if (objSTD.TRNDESC.ToUpper().Contains("CREDIT CASH DEPOSIT"))
                                        {

                                            if (dr[l]["FR"].ToString() == "" || dr[l]["FR"].ToString() == null)
                                            {
                                                objSTD.TRNDESC = "PAYMENT RECEIVED [THANK YOU]";
                                               // objSTD.TRNDATE = dr[l]["OD"].ToString();
                                                  objSTD.TRNDATE = FormatDate(dr[l]["OD"].ToString());
                                            }
                                            else
                                            {
                                                objSTD.TRNDESC = dr[l]["FR"].ToString().Replace("'", "''");
                                                //objSTD.TRNDATE = dr[l]["OD"].ToString();
                                                  objSTD.TRNDATE = FormatDate(dr[l]["OD"].ToString());
                                            }

                                        }
                                        if (dr[l].Table.Columns.Contains("APPROVAL"))
                                        {
                                            objSTD.APPROVAL = dr[l]["APPROVAL"].ToString().Replace("'", "");

                                            if (dr[l]["APPROVAL"].ToString() != "" && objSTD.TRNDATE == "")
                                            {
                                                //objSTD.TRNDATE = dr[l]["OD"].ToString();
                                                  objSTD.TRNDATE = FormatDate(dr[l]["OD"].ToString());
                                            }
                                        }

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
                                          //  objSTD.TRNDATE = dr[l]["TD"].ToString();
                                              objSTD.TRNDATE = FormatDate(dr[l]["TD"].ToString());

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

                                    #endregion
                                }

                                //New View add
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
                                    // objSTD.TRNDATE = objSt.STATEMENT_DATE;  // only for NBL
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


                                }
                                // end IF for SUM_interest
                                else
                                {


                                    //New View add
                                    DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW WHERE ACURN='" + objSt.ACURN + "'", ref reply);

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
                                                            // objSTD.TRNDATE = objSTD.TRNDATE; only for NBL
                                                            objSTD.POSTDATE = objSTD.POSTDATE;
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
                                }  // END ELSE
                                //New View add

                            }

                            else
                            {


                                //New View add
                                DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW WHERE ACURN='" + objSt.ACURN + "'", ref reply);

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
                                                        // objSTD.TRNDATE = objSTD.TRNDATE; only for NBL
                                                        objSTD.POSTDATE = objSTD.POSTDATE;
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
                            }  // END ELSE

                        }
                    }

                    else
                    {
                        if (dtOperation.Rows.Count > 0)
                        {

                            DataRow[] dr = dtOperation.Select("STATEMENTNO='" + objSt.STATEMENTNO + "'");
                            if (dr.Length > 0)
                            {

                                //New View add
                                DataSet dsAcI = objProvider.ReturnData("select * from ACCUM_BODY_VW WHERE ACURN='" + objSt.ACURN + "'", ref reply);

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
                                                        // objSTD.TRNDATE = objSTD.TRNDATE; only for NBL
                                                        objSTD.POSTDATE = objSTD.POSTDATE;
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
                    }  // END ELSE

                }
                catch (Exception ex)
                {
                    errMsg = "Error: " + ex.Message;
                }
            }
        }
            #endregion USD

        #region Privat Functions

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
                    txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                    MsgLogWriter objLW = new MsgLogWriter();
                    objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private void QRGenetetor(String data, string CONTRACTNO)
        {
            try
            {
                QRCodeEncoder qrCodeEncoder = new QRCodeEncoder();

                qrCodeEncoder.QRCodeEncodeMode = QRCodeEncoder.ENCODE_MODE.BYTE;

                int scale = 3;  // size
                qrCodeEncoder.QRCodeScale = scale;

                int version = 3;  //version 1,2,3,4,5,6,7,8,9,10,11,12,13,14....40|
                qrCodeEncoder.QRCodeVersion = version;

                qrCodeEncoder.QRCodeErrorCorrect = QRCodeEncoder.ERROR_CORRECTION.L; // CorrectionLevel --L,M,Q,H

                Image image;

                image = qrCodeEncoder.Encode(data);

                string filename = CONTRACTNO + ".jpg";
                string flocation = _QRPath; //@"D:\XML_For_Email\QR";
                string pathstring = System.IO.Path.Combine(flocation, filename);
                image.Save(pathstring, System.Drawing.Imaging.ImageFormat.Jpeg);

            }
            catch (Exception ex)
            {
                ;
            }

        }
        #endregion

        private void rbtneStatement_CheckedChanged(object sender, EventArgs e)
        {
            dtpStmDate.Enabled = true;
            txtEmailSubject.Enabled = true;
        }

        private void rbtnStatement_CheckedChanged(object sender, EventArgs e)
        {
            dtpStmDate.Enabled = false;
            txtEmailSubject.Enabled = false;
        }
        // Helper method to format dates
        string FormatDate(string dateValue)
        {
            if (!string.IsNullOrEmpty(dateValue))
            {
                string[] dateParts = dateValue.Split('/');
                if (dateParts.Length >= 3) // Ensure there are at least 3 parts
                {
                    return dateParts[0] + "/" + dateParts[1] + "/" + dateParts[2].Substring(0, 4);
                }
            }
            return string.Empty; // Return empty if the input is invalid
        }


    }
}
