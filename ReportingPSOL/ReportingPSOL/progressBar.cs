using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReportingPSOL
{
    public partial class progressBar : Form
    {
        public progressBar()
        {
            InitializeComponent();
        }

        private int progPercent;

        public int ProgPercent
        {
            get { return progPercent; }
            set { progPercent = value; }
        }
        

        public void showProgress()
        {
            progressLabel.Text = ProgPercent.ToString() + "% Completed";
            progressBar1.Value = ProgPercent;
        }

        public void UpdateProgress(int percent, int row, int count)
        {
            row = row - 1;
            count = count - 1;

            progressLabel.Text = percent.ToString() + "% Completed";
            progressBar1.Value = percent;
            recordWritten.Text = "(" + row.ToString() + " of " + count.ToString() + " records written)";
        }

        public void WindowName()
        {
            this.Text = threadVars.Account + " " + threadVars.ReportType.Replace(".", "");
        }
    }
}
