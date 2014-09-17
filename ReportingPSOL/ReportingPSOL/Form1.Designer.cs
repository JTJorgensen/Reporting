namespace ReportingPSOL
{
    partial class ReportingMainForm
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
            this.button1 = new System.Windows.Forms.Button();
            this.cmbAccount = new System.Windows.Forms.ComboBox();
            this.grpReportType = new System.Windows.Forms.GroupBox();
            this.rdoTechTickets = new System.Windows.Forms.RadioButton();
            this.rdoTicketing = new System.Windows.Forms.RadioButton();
            this.rdoBilling = new System.Windows.Forms.RadioButton();
            this.startDatePicker = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.endDatePicker = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.bw = new System.ComponentModel.BackgroundWorker();
            this.button2 = new System.Windows.Forms.Button();
            this.grpReportType.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(20, 117);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // cmbAccount
            // 
            this.cmbAccount.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbAccount.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbAccount.FormattingEnabled = true;
            this.cmbAccount.Items.AddRange(new object[] {
            "Sparhawk",
            "Rettler",
            "Mullins",
            "CCCW",
            "AllOther"});
            this.cmbAccount.Location = new System.Drawing.Point(212, 80);
            this.cmbAccount.Name = "cmbAccount";
            this.cmbAccount.Size = new System.Drawing.Size(200, 21);
            this.cmbAccount.TabIndex = 1;
            // 
            // grpReportType
            // 
            this.grpReportType.Controls.Add(this.rdoTechTickets);
            this.grpReportType.Controls.Add(this.rdoTicketing);
            this.grpReportType.Controls.Add(this.rdoBilling);
            this.grpReportType.Location = new System.Drawing.Point(13, 13);
            this.grpReportType.Name = "grpReportType";
            this.grpReportType.Size = new System.Drawing.Size(128, 98);
            this.grpReportType.TabIndex = 2;
            this.grpReportType.TabStop = false;
            this.grpReportType.Text = "Report Type";
            // 
            // rdoTechTickets
            // 
            this.rdoTechTickets.AutoSize = true;
            this.rdoTechTickets.Location = new System.Drawing.Point(7, 68);
            this.rdoTechTickets.Name = "rdoTechTickets";
            this.rdoTechTickets.Size = new System.Drawing.Size(116, 17);
            this.rdoTechTickets.TabIndex = 2;
            this.rdoTechTickets.TabStop = true;
            this.rdoTechTickets.Text = "Technician Tickets";
            this.rdoTechTickets.UseVisualStyleBackColor = true;
            this.rdoTechTickets.CheckedChanged += new System.EventHandler(this.rdoTechTickets_CheckedChanged);
            // 
            // rdoTicketing
            // 
            this.rdoTicketing.AutoSize = true;
            this.rdoTicketing.Location = new System.Drawing.Point(7, 44);
            this.rdoTicketing.Name = "rdoTicketing";
            this.rdoTicketing.Size = new System.Drawing.Size(109, 17);
            this.rdoTicketing.TabIndex = 1;
            this.rdoTicketing.TabStop = true;
            this.rdoTicketing.Text = "Ticketing Reports";
            this.rdoTicketing.UseVisualStyleBackColor = true;
            this.rdoTicketing.CheckedChanged += new System.EventHandler(this.rdoTicketing_CheckedChanged);
            // 
            // rdoBilling
            // 
            this.rdoBilling.AutoSize = true;
            this.rdoBilling.Location = new System.Drawing.Point(7, 20);
            this.rdoBilling.Name = "rdoBilling";
            this.rdoBilling.Size = new System.Drawing.Size(92, 17);
            this.rdoBilling.TabIndex = 0;
            this.rdoBilling.TabStop = true;
            this.rdoBilling.Text = "Billing Reports";
            this.rdoBilling.UseVisualStyleBackColor = true;
            this.rdoBilling.CheckedChanged += new System.EventHandler(this.rdoBilling_CheckedChanged);
            // 
            // startDatePicker
            // 
            this.startDatePicker.Location = new System.Drawing.Point(212, 30);
            this.startDatePicker.Name = "startDatePicker";
            this.startDatePicker.Size = new System.Drawing.Size(200, 20);
            this.startDatePicker.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(148, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Start Date:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(151, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "End Date:";
            // 
            // endDatePicker
            // 
            this.endDatePicker.Location = new System.Drawing.Point(213, 57);
            this.endDatePicker.Name = "endDatePicker";
            this.endDatePicker.Size = new System.Drawing.Size(200, 20);
            this.endDatePicker.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(156, 83);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(50, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Account:";
            // 
            // bw
            // 
            this.bw.WorkerReportsProgress = true;
            this.bw.WorkerSupportsCancellation = true;
            this.bw.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bw_DoWork);
            this.bw.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bw_ProgressChanged);
            this.bw.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bw_RunWorkerCompleted);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(130, 116);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 8;
            this.button2.Text = "Cancel Report";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // ReportingMainForm
            // 
            this.ClientSize = new System.Drawing.Size(425, 151);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.endDatePicker);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.startDatePicker);
            this.Controls.Add(this.grpReportType);
            this.Controls.Add(this.cmbAccount);
            this.Controls.Add(this.button1);
            this.Name = "ReportingMainForm";
            this.grpReportType.ResumeLayout(false);
            this.grpReportType.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox cmbAccount;
        private System.Windows.Forms.GroupBox grpReportType;
        private System.Windows.Forms.RadioButton rdoTicketing;
        private System.Windows.Forms.RadioButton rdoBilling;
        private System.Windows.Forms.RadioButton rdoTechTickets;
        private System.Windows.Forms.DateTimePicker startDatePicker;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker endDatePicker;
        private System.Windows.Forms.Label label3;
        private System.ComponentModel.BackgroundWorker bw;
        private System.Windows.Forms.Button button2;
    }
}

