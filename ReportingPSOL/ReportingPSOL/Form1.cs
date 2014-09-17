using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Odbc;
using System.Threading;

namespace ReportingPSOL
{
    public partial class ReportingMainForm : Form
    {
        //readWrite rw = new readWrite();
        queryBuilder qb = new queryBuilder();
        checkValid cv = new checkValid();
        threadTestReadWrite tt = new threadTestReadWrite();
        //threadVars threadVars;

        public ReportingMainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!checkThreads())
            {
                tt.cancelClick = false;
                tt.activeXl = new List<object>();

                if (cv.checkAllForValid(rdoBilling, rdoTicketing, rdoTechTickets, cmbAccount.SelectedItem))
                {
                    processIfValid();
                }
                else
                {
                    MessageBox.Show(cv.error);
                }
            }
            else
            {
                MessageBox.Show("Please wait for current report to complete.");
            }
        }//end button1_click()

        private void processIfValid()
        {//TODO
            //String query;
            String startDate = String.Format("{0:MM/dd/yyyy}", startDatePicker.Value);
            String endDate = String.Format("{0:MM/dd/yyyy}", endDatePicker.Value);
            //String reportType;
            String accName;
            String accCard;

            if (cmbAccount.SelectedItem != null)
            {
                accName = cmbAccount.SelectedItem.ToString();
                accCard = qb.accountSelect(accName);
                threadVars.Account = accName;
            }
            else
            {
                accName = "";
                accCard = "";
            }

            if (rdoBilling.Checked)
            {
                if (accName != "AllOther")
                {
                    //reportType = "Billing";
                    //query = qb.billingReportByAccount(startDate, endDate, accCard);
                    threadVars.ReportType = "Billing";
                    threadVars.Query = qb.billingReportByAccount(startDate, endDate, accCard);
                }
                else
                {
                    //reportType = ".Billing";
                    //query = qb.billingReportAllOther(startDate, endDate);
                    threadVars.ReportType = ".Billing";
                    threadVars.Query = qb.billingReportAllOther(startDate, endDate);
                }
            }
            else if (rdoTicketing.Checked)
            {
                if (accName != "AllOther")
                {
                    //reportType = "Tickets";
                    //query = qb.ticketReportByAccount(startDate, endDate, accCard);
                    threadVars.ReportType = "Tickets";
                    threadVars.Query = qb.ticketReportByAccount(startDate, endDate, accCard);
                }
                else
                {
                    //reportType = ".Tickets";
                    //query = qb.ticketReportAllOther(startDate, endDate);
                    threadVars.ReportType = ".Tickets";
                    threadVars.Query = qb.ticketReportAllOther(startDate, endDate);
                }
            }
            else if (rdoTechTickets.Checked)
            {
                //reportType = "Technician";
                //query = qb.technicianTickets();
                threadVars.ReportType = "Technician";
                threadVars.Query = qb.technicianTickets();
            }
            else
            {
                //reportType = "";
                //query = "";
                threadVars.Query = "";
                threadVars.ReportType = "";
                MessageBox.Show("An unexpected error has occurred. \r\n Please contact application support.");
            }

            //rw.readFromDbWriteToXlsx(query, reportType);
            tt.readFromDbWriteToXlsx();
        }//end processIfValid()

        private void rdoBilling_CheckedChanged(object sender, EventArgs e)
        {
            label1.Visible = true;
            startDatePicker.Visible = true;
            label2.Visible = true;
            endDatePicker.Visible = true;
            label3.Visible = true;
            cmbAccount.Visible = true;
            //this.Width = 441;
            this.Size = new Size(441, 189);
        }//end rdoBillingCheckedChanged()

        private void rdoTicketing_CheckedChanged(object sender, EventArgs e)
        {
            label1.Visible = true;
            startDatePicker.Visible = true;
            label2.Visible = true;
            endDatePicker.Visible = true;
            label3.Visible = true;
            cmbAccount.Visible = true;
            //this.Width = 441;
            this.Size = new Size(441, 189);
        }//end rdoTicketing_CheckedChaned()

        private void rdoTechTickets_CheckedChanged(object sender, EventArgs e)
        {
            label1.Visible = false;
            startDatePicker.Visible = false;
            label2.Visible = false;
            endDatePicker.Visible = false;
            label3.Visible = false;
            cmbAccount.Visible = false;
            //this.Width = 172;
            //this.Size = new Size(172, 189);
            this.Size = new Size(242, 189);
        }//end rdoTechTickets_CheckedChanged()

        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            System.ComponentModel.BackgroundWorker worker;
            worker = (System.ComponentModel.BackgroundWorker)sender;
        }//end bw_DoWork()

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }//end bw_ProgressChanged()

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show("Error: " + e.Error.Message);
            }
            else if (e.Cancelled)
            {
                MessageBox.Show("Report Generation Cancelled");
            }
            else
            {
                MessageBox.Show("Report Generated Successfully");
            }
        }//end bw_RunWorkerCompleted()

        private void button2_Click(object sender, EventArgs e)
        {
            tt.cancelClick = true;
            tt.cancelActiveThreads();
        }//end button2_Click()

        private bool checkThreads()
        {
            bool result = false;

            foreach (Thread t in threadVars.ActiveThreads)
            {
                if (t.IsAlive)
                {
                    result = true;
                }
            }

            return result;
        }//end checkThreads()
    }//end class
}//end namespace
