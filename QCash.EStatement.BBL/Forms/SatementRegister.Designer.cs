namespace QCash.EStatement.NBL.Forms
{
    partial class SatementRegister
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnExpRegReport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(-2, 64);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1088, 480);
            this.dataGridView1.TabIndex = 0;
            // 
            // btnExpRegReport
            // 
            this.btnExpRegReport.Location = new System.Drawing.Point(443, 13);
            this.btnExpRegReport.Name = "btnExpRegReport";
            this.btnExpRegReport.Size = new System.Drawing.Size(144, 45);
            this.btnExpRegReport.TabIndex = 1;
            this.btnExpRegReport.Text = "Export Report";
            this.btnExpRegReport.UseVisualStyleBackColor = true;
            this.btnExpRegReport.Click += new System.EventHandler(this.btnExpRegReport_Click);
            // 
            // SatementRegister
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1088, 514);
            this.Controls.Add(this.btnExpRegReport);
            this.Controls.Add(this.dataGridView1);
            this.Name = "SatementRegister";
            this.Text = "SatementRegister";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnExpRegReport;
    }
}