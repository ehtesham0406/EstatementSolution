namespace StatementGenerator
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.uTabMdiManager = new Infragistics.Win.UltraWinTabbedMdi.UltraTabbedMdiManager(this.components);
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tsmFile = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmExit = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmConfiguration = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmSMTP = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmDatabase = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmEStatement = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmProcess = new System.Windows.Forms.ToolStripMenuItem();
            this.reportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmclientPageReport = new System.Windows.Forms.ToolStripMenuItem();
            this.emailToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmAddemail = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmemailUpload = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmSentStatus = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmBulkEmailSender = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            ((System.ComponentModel.ISupportInitialize)(this.uTabMdiManager)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // uTabMdiManager
            // 
            this.uTabMdiManager.MdiParent = this;
            this.uTabMdiManager.ViewStyle = Infragistics.Win.UltraWinTabbedMdi.ViewStyle.Office2007;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmFile,
            this.tsmConfiguration,
            this.tsmEStatement,
            this.reportToolStripMenuItem,
            this.emailToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(873, 24);
            this.menuStrip1.TabIndex = 5;
            this.menuStrip1.Text = "Main Menu";
            // 
            // tsmFile
            // 
            this.tsmFile.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmExit});
            this.tsmFile.Name = "tsmFile";
            this.tsmFile.Size = new System.Drawing.Size(37, 20);
            this.tsmFile.Text = "File";
            // 
            // tsmExit
            // 
            this.tsmExit.Name = "tsmExit";
            this.tsmExit.Size = new System.Drawing.Size(93, 22);
            this.tsmExit.Text = "Exit";
            // 
            // tsmConfiguration
            // 
            this.tsmConfiguration.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmSMTP,
            this.tsmDatabase});
            this.tsmConfiguration.Name = "tsmConfiguration";
            this.tsmConfiguration.Size = new System.Drawing.Size(93, 20);
            this.tsmConfiguration.Text = "Configuration";
            // 
            // tsmSMTP
            // 
            this.tsmSMTP.Name = "tsmSMTP";
            this.tsmSMTP.Size = new System.Drawing.Size(132, 22);
            this.tsmSMTP.Text = "Mail Server";
            // 
            // tsmDatabase
            // 
            this.tsmDatabase.Name = "tsmDatabase";
            this.tsmDatabase.Size = new System.Drawing.Size(132, 22);
            this.tsmDatabase.Text = "Database";
            // 
            // tsmEStatement
            // 
            this.tsmEStatement.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmProcess});
            this.tsmEStatement.Name = "tsmEStatement";
            this.tsmEStatement.Size = new System.Drawing.Size(72, 20);
            this.tsmEStatement.Text = "Operation";
            // 
            // tsmProcess
            // 
            this.tsmProcess.Name = "tsmProcess";
            this.tsmProcess.Size = new System.Drawing.Size(114, 22);
            this.tsmProcess.Text = "Process";
            // 
            // reportToolStripMenuItem
            // 
            this.reportToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmclientPageReport});
            this.reportToolStripMenuItem.Name = "reportToolStripMenuItem";
            this.reportToolStripMenuItem.Size = new System.Drawing.Size(54, 20);
            this.reportToolStripMenuItem.Text = "Report";
            // 
            // tsmclientPageReport
            // 
            this.tsmclientPageReport.Name = "tsmclientPageReport";
            this.tsmclientPageReport.Size = new System.Drawing.Size(173, 22);
            this.tsmclientPageReport.Text = "Statement Register";
            // 
            // emailToolStripMenuItem
            // 
            this.emailToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmAddemail,
            this.tsmemailUpload,
            this.tsmSentStatus,
            this.tsmBulkEmailSender});
            this.emailToolStripMenuItem.Name = "emailToolStripMenuItem";
            this.emailToolStripMenuItem.Size = new System.Drawing.Size(48, 20);
            this.emailToolStripMenuItem.Text = "Email";
            // 
            // tsmAddemail
            // 
            this.tsmAddemail.Name = "tsmAddemail";
            this.tsmAddemail.Size = new System.Drawing.Size(212, 22);
            this.tsmAddemail.Text = "Add Email";
            this.tsmAddemail.Click += new System.EventHandler(this.tsmAddemail_Click);
            // 
            // tsmemailUpload
            // 
            this.tsmemailUpload.Name = "tsmemailUpload";
            this.tsmemailUpload.Size = new System.Drawing.Size(212, 22);
            this.tsmemailUpload.Text = "Bulk Email Upload";
            this.tsmemailUpload.Click += new System.EventHandler(this.tsmemailUpload_Click);
            // 
            // tsmSentStatus
            // 
            this.tsmSentStatus.Name = "tsmSentStatus";
            this.tsmSentStatus.Size = new System.Drawing.Size(212, 22);
            this.tsmSentStatus.Text = "Statement Email ReSender";
            this.tsmSentStatus.Click += new System.EventHandler(this.tsmSentStatus_Click);
            // 
            // tsmBulkEmailSender
            // 
            this.tsmBulkEmailSender.Name = "tsmBulkEmailSender";
            this.tsmBulkEmailSender.Size = new System.Drawing.Size(212, 22);
            this.tsmBulkEmailSender.Text = "Bulk Email Sender";
            this.tsmBulkEmailSender.Click += new System.EventHandler(this.tsmBulkEmailSender_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 679);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(873, 22);
            this.statusStrip1.TabIndex = 9;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(193, 17);
            this.toolStripStatusLabel1.Text = "Powered by IT Consultants Limited.";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(873, 701);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.Text = "NBL Statement  Engine V3.0.8";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.uTabMdiManager)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Infragistics.Win.UltraWinTabbedMdi.UltraTabbedMdiManager uTabMdiManager;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tsmFile;
        private System.Windows.Forms.ToolStripMenuItem tsmExit;
        private System.Windows.Forms.ToolStripMenuItem tsmConfiguration;
        private System.Windows.Forms.ToolStripMenuItem tsmSMTP;
        private System.Windows.Forms.ToolStripMenuItem tsmDatabase;
        private System.Windows.Forms.ToolStripMenuItem tsmEStatement;
        private System.Windows.Forms.ToolStripMenuItem tsmProcess;
        private System.Windows.Forms.ToolStripMenuItem reportToolStripMenuItem;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripMenuItem tsmclientPageReport;
        private System.Windows.Forms.ToolStripMenuItem emailToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tsmAddemail;
        private System.Windows.Forms.ToolStripMenuItem tsmemailUpload;
        private System.Windows.Forms.ToolStripMenuItem tsmSentStatus;
        private System.Windows.Forms.ToolStripMenuItem tsmBulkEmailSender;
  
       
    }
}

