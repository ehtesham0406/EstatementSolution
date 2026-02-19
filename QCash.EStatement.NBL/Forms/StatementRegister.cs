using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Common;
using System.Connection;
//using CrystalDecisions.CrystalReports.Engine;
//using CrystalDecisions.Shared;
using Infragistics.Documents.Report;
using System.Configuration;
using FlexiStar.Utilities;
using StatementGenerator.App_Code;
using QCash.EStatement.NBL.App_Code;
using CrystalDecisions.CrystalReports.Engine;

namespace StatementGenerator
{
    public partial class StatementRegister : Form
    {
        private string StmDate = string.Empty;
        private string _ClientPageInfoPath = string.Empty;
        private string _fiid = string.Empty;

        public StatementRegister(string fiid)
        {
            InitializeComponent();
            this.Load += new EventHandler(crystalReportViewer1_Load);
            _fiid = fiid;
        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {
            string reply = string.Empty;
      
            try
            {
              
               // ReportDocument rd = new ReportDocument();
               
               //string  _strServer = ConfigurationManager.AppSettings["ServerName"].ToString();
               //string   _strDatabase = ConfigurationManager.AppSettings["DataBaseName"].ToString();
               //string   _strUserID = ConfigurationManager.AppSettings["UserId"].ToString();
               //string   _strPwd = ConfigurationManager.AppSettings["Password"].ToString();
                _ClientPageInfoPath = ConfigurationManager.AppSettings["ClientPageInfoPath"].ToString();
        

             //  rd.Load(_ClientPageInfoPath);
               
              //// rd.SetDatabaseLogon(strUserID, strPwd); 
               // rd.DataSourceConnections[0].SetConnection(_strServer, _strDatabase, _strUserID, _strPwd);
               // crystalReportViewer1.ReportSource = rd;
              

            }
            catch (Exception ex)
            {
                ;
            }
        }

        void btnSearch_Click(object sender, EventArgs e)
        {
            string reply = string.Empty;
            try
            {
                if (StmDate == "")
                    StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy");
                else StmDate = dtpStmDate.Value.ToString("dd/MM/yyyy");


               ReportDocument rd = new ReportDocument();
                //private myDataSet ds;

               // string _strServer = ConfigurationManager.AppSettings["ServerName"].ToString();
               // string _strDatabase = ConfigurationManager.AppSettings["DataBaseName"].ToString();
              //  string _strUserID = ConfigurationManager.AppSettings["UserId"].ToString();
               // string _strPwd = ConfigurationManager.AppSettings["Password"].ToString();
             //   _ClientPageInfoPath = ConfigurationManager.AppSettings["ClientPageInfoPath"].ToString();


               // rd.Load(_ClientPageInfoPath);

                // rd.SetDatabaseLogon(strUserID, strPwd);
              //  rd.DataSourceConnections[0].SetConnection(_strServer, _strDatabase, _strUserID, _strPwd);
               // crystalReportViewer1.ReportSource = rd;

                MsgLogWriter objLW = new MsgLogWriter();

                StatementRegisterList objESList = EStatementManager.Instance().GetStatementRegister(StmDate, ref reply);
                if (objESList != null)
                {
                    if (objESList.Count > 0)
                    {
                       // crystalReportViewer1.ReportSource = objESList;
                        rd = new ReportDocument();
                        rd.Load(_ClientPageInfoPath);
                        rd.SetDataSource(objESList);
                        crystalReportViewer1.ReportSource = rd;
                    


                    }
                }
            }
            catch (Exception ex)
            {
                //MsgLogWriter objLW = new MsgLogWriter();
               // objLW.logTrace(_LogPath, "EStatement.log", ex.Message);
                //txtAnalyzer.Invoke(_addText, new object[] { System.DateTime.Now.ToString("MMMM dd, yyyy h:mm:tt") + " : " + ex.Message });
            }
        }

        private void Close_Popup(object sender, PopupEventArgs e)
        {
            this.Close();
        }

        private void btnclose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

       

    }
}
