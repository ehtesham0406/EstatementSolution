using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using StatementGenerator.App_Code;
using System.Configuration;
using FlexiStar.Utilities.EncryptionEngine;

namespace StatementGenerator
{
    public partial class SMTPConfiguration : Form
    {
        private bool flag = true;
        private string _Fid = string.Empty;

        public SMTPConfiguration()
        {
            InitializeComponent();

            this.btnNew.Click += new EventHandler(btnNew_Click);
            this.btnUpdate.Click += new EventHandler(btnUpdate_Click);
            this.btnSave.Click += new EventHandler(btnSave_Click);
            this.btnClose.Click += new EventHandler(btnClose_Click);
            this.btnCancel.Click += new EventHandler(btnCancel_Click);
            this.Load += new EventHandler(SMTPConfiguration_Load);
        }

        void SMTPConfiguration_Load(object sender, EventArgs e)
        {
            SetButton(true);
            _Fid = ConfigurationManager.AppSettings["FIID"].ToString();
            txtFIID.Text = _Fid;
            GetSMTPInfo();
        }
        private void SetButton(bool st) 
        {
            this.btnNew.Visible = st;
            this.btnUpdate.Visible = st;
            this.btnSave.Visible = !st;
            this.btnClose.Visible = st;
            this.btnCancel.Visible = !st;
        }

        void btnNew_Click(object sender, EventArgs e)
        {
            flag = true;
            SetButton(false);
        }
        void btnUpdate_Click(object sender, EventArgs e)
        {
            flag = false;
            SetButton(false);
        }
        void btnSave_Click(object sender, EventArgs e)
        {
            if (txtSmtpServer.Text != null)
                if (txtFromUser.Text != null)
                    if (txtPassword.Text != null)
                        if (txtFromAccount.Text != null)
                        {
                            Encryption objEnc = new Encryption();
                            string _password = string.Empty;

                            SmtpConfigurationManager objSmtpMan = new SmtpConfigurationManager();
                            SmtpConfigurationInfo objSmtpInfo = new SmtpConfigurationInfo();
                            objSmtpInfo.FIID = txtFIID.Text;
                            objSmtpInfo.Smtp_Server = txtSmtpServer.Text;
                            objSmtpInfo.Smtp_Port = Convert.ToInt32(txtSmtpPort.Text);
                            objSmtpInfo.EnableSSL = Convert.ToInt32(chkSSL.Checked);
                            objSmtpInfo.From_Address = txtFromAccount.Text;
                            objSmtpInfo.From_User = txtFromUser.Text;
                            _password = objEnc.EncryptWord(txtPassword.Text);
                            objSmtpInfo.From_Password = _password;
                            objSmtpInfo.Status = Convert.ToInt32(chkStatus.Checked);

                            if (flag == true)
                                objSmtpMan.SaveSmtpConfiguration(objSmtpInfo);
                            else if (flag == false)
                                objSmtpMan.UpdateSmtpConfiguration(objSmtpInfo);

                            SetButton(true);
                        }
                        else
                        { }
        }
        void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        void btnCancel_Click(object sender, EventArgs e)
        {
            SetButton(true);
        }



        private void GetSMTPInfo()
        {
            SmtpConfigurationManager objSmtpMan = new SmtpConfigurationManager();
            SmtpConfigurationList objSmtpList = new SmtpConfigurationList();

            Encryption objEnc = new Encryption();

            objSmtpList = objSmtpMan.GetSmtpConfiguration(txtFIID.Text, 1);

            if (objSmtpList != null)
            {
                if (objSmtpList.Count > 0)
                {
                    txtFIID.Text = objSmtpList[0].FIID  ;
                    txtSmtpServer.Text = objSmtpList[0].Smtp_Server  ;
                    txtSmtpPort.Text = objSmtpList[0].Smtp_Port.ToString();
                    chkSSL.Checked = Convert.ToBoolean(Convert.ToInt32(objSmtpList[0].EnableSSL));
                    txtFromAccount.Text = objSmtpList[0].From_Address;
                    txtFromUser.Text =objSmtpList[0].From_User;
                    txtPassword.Text = objEnc.DecryptWord(objSmtpList[0].From_Password);
                    chkStatus.Checked = Convert.ToBoolean(Convert.ToInt32(objSmtpList[0].Status));
                }
            }
        }

       

    }
}
