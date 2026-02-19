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
using CrystalDecisions.Shared;
using FlexiStar.Utilities;
using FlexiStar.Utilities.EncryptionEngine;
using QCash.EStatement.NBL.App_Code;
using QCash.EStatement.NBL.Reports;
using StatementGenerator.App_Code;
using System.Net.Mime;

namespace QCash.EStatement.NBL.Forms
{
    public partial class BulkEmailSender : Form
    {
        #region Declaration
        private ConnectionStringBuilder ConStr = null;
        private SqlDbProvider objProvider = null;
        //
        delegate void SetTextCallback(string text);
        private SetTextCallback _addText = null;
        //
        private string Bank_Code = string.Empty;


        private string _XMLProcessedPath = string.Empty;
        private string _XMLSourcePath = string.Empty;
        private string _LogPath = string.Empty;
        private string _EmailResultPath = string.Empty;
        private string _AdditionalAttachment = string.Empty;
        private string _Mail = string.Empty;
        private string StmDate = string.Empty;
        


        private string stmMessage = string.Empty;
        private string _xmlName = string.Empty;

        Thread tdSendMail = null;

        private string _fiid = string.Empty;
        int pdfCount = 0;

        #endregion

        #region Constructer
        public BulkEmailSender(string fiid)
        {
            InitializeComponent();

            _addText = new SetTextCallback(Output);

          //  this.Load += new EventHandler(ReportViewer_Load);
            this.btnSendMail.Click += new EventHandler(btnSendMail_Click);
            this.btnClose.Click += new EventHandler(btnClose_Click);
            _AdditionalAttachment = ConfigurationManager.AppSettings["AdditionalAttachment"].ToString();

            _fiid = fiid;
        }

        #endregion

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


        void btnClose_Click(object sender, EventArgs e)
        {

            if (tdSendMail != null)
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

      

        void btnSendMail_Click(object sender, EventArgs e)
        {
            if (txtEmailSubject.Text.Length > 100)
            {
                MessageBox.Show("Email subject should be within 100 character...");
            }
            else
            {
                btnSendMail.Enabled = false;

                tdSendMail = new Thread(new ThreadStart(SendMail));
                tdSendMail.IsBackground = true;
                tdSendMail.Start();
            }
        }

        private void SendMail()
        {
            string reply = string.Empty;
            try
            {
               
                MsgLogWriter objLW = new MsgLogWriter();

                EStatementList objESList = EStatementManager.Instance().GetAllEmail(ref reply);
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
                                            mail.Subject = txtEmailSubject.Text;
                                            mail.Body = objESList[i].MAILBODY;
                                            mail.To.Add(objESList[i].MAILADDRESS.Trim());
                                            
                                            

                                           
                                            //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=

                                            //   **** imange in email body code ****
                                             var contentID = "Image";
                                             var inlineLogo = new Attachment(@"D:\XML_For_Email\EmailBodyImage\BodyImage.jpg");  //change_here
                                            inlineLogo.ContentId = contentID;
                                            inlineLogo.ContentDisposition.Inline = true;
                                            inlineLogo.ContentDisposition.DispositionType = DispositionTypeNames.Inline;

                                            mail.IsBodyHtml = true;
                                            mail.Attachments.Add(inlineLogo);
                                            mail.Body = "<htm><body> <img src=\"cid:" + contentID + "\"> </body></html>";

                                           


                                            //-----------------------------------------------------
                                            string[] filePaths = Directory.GetFiles(_AdditionalAttachment);
                                            if (filePaths.Length != 0)
                                            {
                                                for (int x = 0; x < filePaths.Length; x++)
                                                {
                                                    System.Net.Mail.Attachment attachment;
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
                                          
                                            count++;
                                        }
                                        catch (Exception ex)
                                        {
                                            txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message });
                                            objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : Error: " + ex.Message);

                                            objESList[i].STATUS = "2"; // Mail is not Sent
                                           
                                        }
                                    }
                                    else
                                    {
                                        txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + "No Mail Address Found to send the Estatement " + objESList[i].FILE_LOCATION }); ;
                                        objLW.logTrace(_LogPath, "EStatement.log", System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : No Mail Address Found to send the Estatement " + objESList[i].FILE_LOCATION);

                                        objESList[i].STATUS = "8";   //  No Mail Address Found
                                       
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


    }
       
    }

      