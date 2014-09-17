using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Odbc;
using System.Drawing;
using System.Threading;

namespace ReportingPSOL
{
    class threadTestReadWrite
    {
        public List<object> activeXl;// = new List<object>();
        public bool cancelClick = false;

        //Thread t;
        //Thread t2;
        //Thread t3;

        //Thread tbl;
        //Thread tbx;
        //Thread tdbl;
        //Thread tdbx;

        //Thread ttl;
        //Thread ttx;
        //Thread tdtl;
        //Thread tdtx;

        //Thread tTL;
        //Thread tTx;

        //int count;
        //double percent;
        //int progPercent;

        //static object misValue = System.Reflection.Missing.Value; //part of test

        //static Excel.Application xlApp = new Excel.ApplicationClass(); //part of test
        //static Excel.Workbook xlWorkbook = xlApp.Workbooks.Add(misValue); //part of test
        //static Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1); //part of test
        //Excel.Range rng; //part of test

        public void readFromDbWriteToXlsx()
        {
            switch (threadVars.ReportType)
            {
                case "Billing":
                    //List<BillingRecord> billing = new List<BillingRecord>();

                    Thread tbl = new Thread(writeBillingToList);
                    threadVars.ActiveThreads.Add(tbl);
                    tbl.Start();
                    tbl.Join();
                    threadVars.ActiveThreads.Remove(tbl);

                    if (threadVars.BillingRecords != null)
                    {
                        //count = threadVars.BillingRecords.Count + 1;

                        Thread tbx = new Thread(writeBillingListToXlsx);
                        threadVars.ActiveThreads.Add(tbx);
                        tbx.SetApartmentState(ApartmentState.STA);
                        tbx.IsBackground = true;
                        tbx.Priority = ThreadPriority.AboveNormal;
                        tbx.Start();
                        //tbx.Join();
                        //threadVars.ActiveThreads.Remove(tbx);
                    }
                    break;

                case ".Billing":
                    //List<AOBillingRecord> aoBilling = new List<AOBillingRecord>();

                    Thread tdbl = new Thread(writeAOBillingToList);
                    threadVars.ActiveThreads.Add(tdbl);
                    tdbl.Start();
                    tdbl.Join();
                    threadVars.ActiveThreads.Remove(tdbl);

                    if (threadVars.AOBillingRecords != null)
                    {
                        Thread tdbx = new Thread(writeAOBillingListToXlsx);
                        threadVars.ActiveThreads.Add(tdbx);
                        tdbx.SetApartmentState(ApartmentState.STA);
                        tdbx.IsBackground = true;
                        tdbx.Priority = ThreadPriority.AboveNormal;
                        tdbx.Start();
                        //tdbx.Join();
                        //threadVars.ActiveThreads.Remove(tdbx);
                    }
                    break;

                case "Tickets":
                    //List<TicketRecord> ticketing = new List<TicketRecord>();

                    Thread ttl = new Thread(writeTicketToList);
                    threadVars.ActiveThreads.Add(ttl);
                    ttl.Priority = ThreadPriority.AboveNormal;
                    ttl.Start();
                    ttl.Join();
                    threadVars.ActiveThreads.Remove(ttl);

                    if (threadVars.TicketRecords != null)
                    {
                        Thread ttx = new Thread(writeTicketingListToXlsx);
                        threadVars.ActiveThreads.Add(ttx);
                        ttx.SetApartmentState(ApartmentState.STA);
                        ttx.IsBackground = true;
                        ttx.Priority = ThreadPriority.AboveNormal;
                        ttx.Start();
                        //ttx.Join();
                        //threadVars.ActiveThreads.Remove(ttx);
                    }
                    break;

                case ".Tickets":
                    //List<AOTicketRecord> aoTicketing = new List<AOTicketRecord>();

                    Thread tdtl = new Thread(writeAOTicketToList);
                    threadVars.ActiveThreads.Add(tdtl);
                    tdtl.Start();
                    tdtl.Join();
                    threadVars.ActiveThreads.Remove(tdtl);

                    if (threadVars.AOTicketRecords != null)
                    {
                        Thread tdtx = new Thread(writeAOTicketingListToXlsx);
                        threadVars.ActiveThreads.Add(tdtx);
                        tdtx.SetApartmentState(ApartmentState.STA);
                        tdtx.IsBackground = true;
                        tdtx.Priority = ThreadPriority.AboveNormal;
                        tdtx.Start();
                        //tdtx.Join();
                        //threadVars.ActiveThreads.Remove(tdtx);
                    }
                    break;

                case "Technician":
                    //List<TechTixRecord> techTix = new List<TechTixRecord>();

                    Thread tTl = new Thread(writeTechTixToList);
                    threadVars.ActiveThreads.Add(tTl);
                    tTl.Start();
                    tTl.Join();
                    threadVars.ActiveThreads.Remove(tTl);

                    if (threadVars.TechTixRecords != null)
                    {
                        Thread tTx = new Thread(writeTechTixListToXlsx);
                        threadVars.ActiveThreads.Add(tTx);
                        tTx.SetApartmentState(ApartmentState.STA);
                        tTx.IsBackground = true;
                        tTx.Priority = ThreadPriority.AboveNormal;
                        tTx.Start();
                        //tTx.Join();
                        //threadVars.ActiveThreads.Remove(tTx);
                    }
                    break;

                default:
                    break;
            }
        }//end readFromDbWriteToXlsx()

        private void writeBillingToList()
        {
            //List<BillingRecord> billingRecords = new List<BillingRecord>();
            threadVars.BillingRecords = new List<BillingRecord>();

            using (OdbcConnection conn = new OdbcConnection("DSN=CommitSystemODBC"))
            {
                OdbcCommand cmd = new OdbcCommand(threadVars.Query, conn);

                try
                {
                    conn.Open();
                    OdbcDataReader reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        BillingRecord record = new BillingRecord();

                        record.TicketNo = reader[0];
                        record.Summary = reader[1];
                        record.Resolution = reader[2];
                        record.CreatedAt = reader[3];
                        record.ClosedAt = reader[4];
                        record.TotalHours = reader[5];
                        record.TotalLabor = reader[6];
                        record.Purchases = reader[7];
                        record.Expenses = reader[8];
                        record.GrandTotal = reader[9];

                        threadVars.BillingRecords.Add(record);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    threadVars.BillingRecords = null;
                }
            }
        }//end writeBillingToList()

        private void writeTicketToList()
        {
            //List<TicketRecord> ticketRecords = new List<TicketRecord>();
            threadVars.TicketRecords = new List<TicketRecord>();

            using (OdbcConnection conn = new OdbcConnection("DSN=CommitSystemODBC"))
            {
                OdbcCommand cmd = new OdbcCommand(threadVars.Query, conn);

                try
                {
                    conn.Open();
                    OdbcDataReader reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        TicketRecord record = new TicketRecord();

                        record.TicketNo = reader[0];
                        record.Summary = reader[1];
                        record.Status = reader[2];
                        record.Created = reader[3];
                        record.Closed = reader[4];
                        record.Tech = reader[5];

                        threadVars.TicketRecords.Add(record);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    threadVars.TicketRecords = null; ;
                }
            }
        }//end writeTicketToList()

        private void writeAOBillingToList()
        {
            //List<AOBillingRecord> aoBillingRecords = new List<AOBillingRecord>();
            threadVars.AOBillingRecords = new List<AOBillingRecord>();

            using (OdbcConnection conn = new OdbcConnection("DSN=CommitSystemODBC"))
            {
                OdbcCommand cmd = new OdbcCommand(threadVars.Query, conn);

                try
                {
                    conn.Open();
                    OdbcDataReader reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        AOBillingRecord record = new AOBillingRecord();

                        record.TicketNo = reader[0];
                        record.Account = reader[1];
                        record.Summary = reader[2];
                        record.Resolution = reader[3];
                        record.Created = reader[4];
                        record.Closed = reader[5];
                        record.Hours = reader[6];
                        record.Labor = reader[7];
                        record.Purchases = reader[8];
                        record.Expenses = reader[9];
                        record.Total = reader[10];

                        threadVars.AOBillingRecords.Add(record);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    threadVars.AOBillingRecords = null; ;
                }
            }
        }//end writeAOBillingToList()

        private void writeAOTicketToList()
        {
            //List<AOTicketRecord> aoTicketRecords = new List<AOTicketRecord>();
            threadVars.AOTicketRecords = new List<AOTicketRecord>();

            using (OdbcConnection conn = new OdbcConnection("DSN=CommitSystemODBC"))
            {
                OdbcCommand cmd = new OdbcCommand(threadVars.Query, conn);

                try
                {
                    conn.Open();
                    OdbcDataReader reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        AOTicketRecord record = new AOTicketRecord();

                        record.TicketNo = reader[0];
                        record.Account = reader[1];
                        record.Summary = reader[2];
                        record.Status = reader[3];
                        record.Created = reader[4];
                        record.Closed = reader[5];
                        record.Tech = reader[6];

                        threadVars.AOTicketRecords.Add(record);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    threadVars.AOTicketRecords = null; ;
                }
            }
        }//end writeAOTicketToList()

        private void writeTechTixToList()
        {
            //List<TechTixRecord> techTixRecords = new List<TechTixRecord>();
            threadVars.TechTixRecords = new List<TechTixRecord>();

            using (OdbcConnection conn = new OdbcConnection("DSN=CommitSystemODBC"))
            {
                OdbcCommand cmd = new OdbcCommand(threadVars.Query, conn);

                try
                {
                    conn.Open();
                    OdbcDataReader reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        TechTixRecord record = new TechTixRecord();

                        record.TicketNo = reader[0];
                        record.Summary = reader[1];
                        record.Tech = reader[2];
                        record.Account = reader[3];
                        record.Type = reader[4];
                        record.Status = reader[5];
                        record.Notes = reader[6];

                        threadVars.TechTixRecords.Add(record);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    threadVars.TechTixRecords = null;
                }
            }
        }//end writeTechTixToList()

        private void writeBillingListToXlsx()
        {
            //String saveAs;
            int percent;
            int count = threadVars.BillingRecords.Count + 1;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range rng;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            activeXl.Add(xlApp);

            xlWorkbook = xlApp.Workbooks.Add(misValue);
            activeXl.Add(xlWorkbook);

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            activeXl.Add(xlWorksheet);

            xlWorksheet.Cells[1, 1] = "Ticket No";
            xlWorksheet.Cells[1, 2] = "Summary";
            xlWorksheet.Cells[1, 3] = "Resolution";
            xlWorksheet.Cells[1, 4] = "Created";
            xlWorksheet.Cells[1, 5] = "Closed";
            xlWorksheet.Cells[1, 6] = "Hours";
            xlWorksheet.Cells[1, 7] = "Labor";
            xlWorksheet.Cells[1, 8] = "Purchases";
            xlWorksheet.Cells[1, 9] = "Expenses";
            xlWorksheet.Cells[1, 10] = "Total";

            rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[1, 1], xlWorksheet.Cells[1, 10]);
            rng.Interior.Color = ColorTranslator.ToOle(Color.GreenYellow);

            try
            {
                progressBar progressForm = new progressBar();
                progressForm.Show();

                int row = 1;
                foreach (BillingRecord record in threadVars.BillingRecords)
                {
                    row++;

                    xlWorksheet.Cells[row, 1] = record.TicketNo;
                    xlWorksheet.Cells[row, 2] = record.Summary;
                    xlWorksheet.Cells[row, 3] = record.Resolution;
                    xlWorksheet.Cells[row, 4] = record.CreatedAt;
                    xlWorksheet.Cells[row, 5] = record.ClosedAt;
                    xlWorksheet.Cells[row, 6] = record.TotalHours;
                    xlWorksheet.Cells[row, 7] = record.TotalLabor;
                    xlWorksheet.Cells[row, 8] = record.Purchases;
                    xlWorksheet.Cells[row, 9] = record.Expenses;
                    xlWorksheet.Cells[row, 10] = record.GrandTotal;

                    rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[row, 1], xlWorksheet.Cells[row, 10]);
                    if (row % 2 == 0)
                    {
                        rng.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                    }
                    else
                    {
                        rng.Interior.Color = ColorTranslator.ToOle(Color.White);
                    }

                    xlWorksheet.Rows.RowHeight = 16.5;

                    //progress = (double)(row/count);
                    //percent = (double)findPercent(row, count);
                    //percent = percent * 100;
                    //progPercent = (int)(percent);
                    percent = findPercent(row, count);
                    //showProgress(progressForm, percent);
                    //progressForm.showProgress();
                    progressForm.WindowName();
                    progressForm.UpdateProgress(percent, row, count);

                }
                //t2 = new Thread(writeBillToXlEven);
                //t2.SetApartmentState(ApartmentState.STA);
                //t2.IsBackground = true;
                //t2.Start();

                //t3 = new Thread(writeBillToXlOdd);
                //t3.SetApartmentState(ApartmentState.STA);
                //t3.IsBackground = true;
                //t3.Start();

                //t2.Join();
                //t3.Join();

                //SaveFileDialog attempt = new SaveFileDialog();
                //attempt.Filter = "Excel Files | *.xlsx";
                //attempt.DefaultExt = "xlsx";
                //attempt.ShowDialog();
                //saveAs = attempt.FileName;

                ////xlWorkbook.SaveAs(saveAs, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                //xlWorkbook.SaveAs(saveAs);
                saveXL(xlWorkbook);

                //MessageBox.Show("Your report was succesfully created!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //MessageBox.Show("There was an issue writing your report to Excel.");
            }
            finally
            {
                if (!cancelClick)
                {
                    xlWorkbook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorksheet);
                    releaseObject(xlWorkbook);
                    releaseObject(xlApp);
                }
            }
        }//end writeListToXlsx(billing)

        private void writeTicketingListToXlsx()
        {
            //String saveAs;
            int percent;
            int count = threadVars.TicketRecords.Count + 1;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range rng;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            activeXl.Add(xlApp);

            xlWorkbook = xlApp.Workbooks.Add(misValue);
            activeXl.Add(xlWorkbook);

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            activeXl.Add(xlWorksheet);

            xlWorksheet.Cells[1, 1] = "Ticket No";
            xlWorksheet.Cells[1, 2] = "Summary";
            xlWorksheet.Cells[1, 3] = "Status";
            xlWorksheet.Cells[1, 4] = "Created";
            xlWorksheet.Cells[1, 5] = "Closed";
            xlWorksheet.Cells[1, 6] = "Tech";
            xlWorksheet.Cells[1, 7] = "Till Closed";

            rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[1, 1], xlWorksheet.Cells[1, 7]);
            rng.Interior.Color = ColorTranslator.ToOle(Color.GreenYellow);

            try
            {
                progressBar pf = new progressBar();
                pf.Show();

                int row = 1;
                foreach (TicketRecord record in threadVars.TicketRecords)
                {
                    row++;

                    xlWorksheet.Cells[row, 1] = record.TicketNo;
                    xlWorksheet.Cells[row, 2] = record.Summary;
                    xlWorksheet.Cells[row, 3] = record.Status;
                    xlWorksheet.Cells[row, 4] = record.Created;
                    xlWorksheet.Cells[row, 5] = record.Closed;
                    xlWorksheet.Cells[row, 6] = record.Tech;

                    if (record.Closed.ToString() != "")
                    {
                        xlWorksheet.Cells[row, 7] = "=SUM(E" + row + ", - D" + row + ")";
                    }

                    rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[row, 1], xlWorksheet.Cells[row, 7]);
                    if (row % 2 == 0)
                    {
                        rng.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                    }
                    else
                    {
                        rng.Interior.Color = ColorTranslator.ToOle(Color.White);
                    }

                    xlWorksheet.Rows.RowHeight = 16.5;

                    percent = findPercent(row, count);
                    pf.WindowName();
                    pf.UpdateProgress(percent, row, count);
                }
                //SaveFileDialog attempt = new SaveFileDialog();
                //attempt.Filter = "Excel Files | *.xlsx";
                //attempt.DefaultExt = "xlsx";
                //attempt.ShowDialog();
                //saveAs = attempt.FileName;

                //xlWorkbook.SaveAs(saveAs, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                //MessageBox.Show("Your report was succesfully created!");
                saveXL(xlWorkbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show("There was an issue writing your report to Excel.");
            }
            finally
            {
                if (!cancelClick)
                {
                    xlWorkbook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorksheet);
                    releaseObject(xlWorkbook);
                    releaseObject(xlApp);
                }
            }
        }//end writeListToXlsx(Ticketing)

        private void writeAOBillingListToXlsx()
        {
            //String saveAs;
            int percent;
            int count = threadVars.AOBillingRecords.Count + 1;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range rng;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            activeXl.Add(xlApp);

            xlWorkbook = xlApp.Workbooks.Add(misValue);
            activeXl.Add(xlWorkbook);

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            activeXl.Add(xlWorksheet);

            xlWorksheet.Cells[1, 1] = "Ticket No";
            xlWorksheet.Cells[1, 2] = "Account";
            xlWorksheet.Cells[1, 3] = "Summary";
            xlWorksheet.Cells[1, 4] = "Resolution";
            xlWorksheet.Cells[1, 5] = "Created";
            xlWorksheet.Cells[1, 6] = "Closed";
            xlWorksheet.Cells[1, 7] = "Hours";
            xlWorksheet.Cells[1, 8] = "Labor";
            xlWorksheet.Cells[1, 9] = "Purchases";
            xlWorksheet.Cells[1, 10] = "Expenses";
            xlWorksheet.Cells[1, 11] = "Total";

            rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[1, 1], xlWorksheet.Cells[1, 11]);
            rng.Interior.Color = ColorTranslator.ToOle(Color.GreenYellow);

            try
            {
                progressBar pf = new progressBar();
                pf.Show();

                int row = 1;
                foreach (AOBillingRecord record in threadVars.AOBillingRecords)
                {
                    row++;

                    xlWorksheet.Cells[row, 1] = record.TicketNo;
                    xlWorksheet.Cells[row, 2] = record.Account;
                    xlWorksheet.Cells[row, 3] = record.Summary;
                    xlWorksheet.Cells[row, 4] = record.Resolution;
                    xlWorksheet.Cells[row, 5] = record.Created;
                    xlWorksheet.Cells[row, 6] = record.Closed;
                    xlWorksheet.Cells[row, 7] = record.Hours;
                    xlWorksheet.Cells[row, 8] = record.Labor;
                    xlWorksheet.Cells[row, 9] = record.Purchases;
                    xlWorksheet.Cells[row, 10] = record.Expenses;
                    xlWorksheet.Cells[row, 11] = record.Total;

                    rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[row, 1], xlWorksheet.Cells[row, 11]);
                    if (row % 2 == 0)
                    {
                        rng.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                    }
                    else
                    {
                        rng.Interior.Color = ColorTranslator.ToOle(Color.White);
                    }

                    xlWorksheet.Rows.RowHeight = 16.5;

                    percent = findPercent(row, count);
                    pf.WindowName();
                    pf.UpdateProgress(percent, row, count);
                }

                //SaveFileDialog attempt = new SaveFileDialog();
                //attempt.Filter = "Excel Files | *.xlsx";
                //attempt.DefaultExt = "xlsx";
                //attempt.ShowDialog();
                //saveAs = attempt.FileName;

                //xlWorkbook.SaveAs(saveAs, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                //xlWorkbook.SaveAs(saveAs);

                //MessageBox.Show("Your report was succesfully created!");
                saveXL(xlWorkbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show("There was an issue writing your report to Excel.");
            }
            finally
            {
                if (!cancelClick)
                {
                    xlWorkbook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorksheet);
                    releaseObject(xlWorkbook);
                    releaseObject(xlApp);
                }
            }
        }//end writeListToExcel(AOBilling)

        private void writeAOTicketingListToXlsx()
        {
            //String saveAs;
            int percent;
            int count = threadVars.AOTicketRecords.Count + 1;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range rng;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            activeXl.Add(xlApp);

            xlWorkbook = xlApp.Workbooks.Add(misValue);
            activeXl.Add(xlWorkbook);

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            activeXl.Add(xlWorksheet);

            xlWorksheet.Cells[1, 1] = "Ticket No";
            xlWorksheet.Cells[1, 2] = "Account";
            xlWorksheet.Cells[1, 3] = "Summary";
            xlWorksheet.Cells[1, 4] = "Status";
            xlWorksheet.Cells[1, 5] = "Created";
            xlWorksheet.Cells[1, 6] = "Closed";
            xlWorksheet.Cells[1, 7] = "Tech";
            xlWorksheet.Cells[1, 8] = "Till Closed";

            rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[1, 1], xlWorksheet.Cells[1, 8]);
            rng.Interior.Color = ColorTranslator.ToOle(Color.GreenYellow);

            try
            {
                progressBar pf = new progressBar();
                pf.Show();

                int row = 1;
                foreach (AOTicketRecord record in threadVars.AOTicketRecords)
                {
                    row++;

                    xlWorksheet.Cells[row, 1] = record.TicketNo;
                    xlWorksheet.Cells[row, 2] = record.Account;
                    xlWorksheet.Cells[row, 3] = record.Summary;
                    xlWorksheet.Cells[row, 4] = record.Status;
                    xlWorksheet.Cells[row, 5] = record.Created;
                    xlWorksheet.Cells[row, 6] = record.Closed;
                    xlWorksheet.Cells[row, 7] = record.Tech;

                    if (record.Closed.ToString() != "")
                    {
                        xlWorksheet.Cells[row, 8] = "=SUM(E" + row + ", - D" + row + ")";
                    }

                    rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[row, 1], xlWorksheet.Cells[row, 8]);
                    if (row % 2 == 0)
                    {
                        rng.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                    }
                    else
                    {
                        rng.Interior.Color = ColorTranslator.ToOle(Color.White);
                    }

                    xlWorksheet.Rows.RowHeight = 16.5;

                    percent = findPercent(row, count);
                    pf.WindowName();
                    pf.UpdateProgress(percent, row, count);
                }
                //SaveFileDialog attempt = new SaveFileDialog();
                //attempt.Filter = "Excel Files | *.xlsx";
                //attempt.DefaultExt = "xlsx";
                //attempt.ShowDialog();
                //saveAs = attempt.FileName;

                //xlWorkbook.SaveAs(saveAs, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                //MessageBox.Show("Your report was succesfully created!");

                saveXL(xlWorkbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show("There was an issue writing your report to Excel.");
            }
            finally
            {
                if (!cancelClick)
                {
                    xlWorkbook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorksheet);
                    releaseObject(xlWorkbook);
                    releaseObject(xlApp);
                }
            }
        }//end writeListToXlsx(Ticketing)

        private void writeTechTixListToXlsx()
        {
            //String saveAs;
            int percent;
            int count = threadVars.TechTixRecords.Count + 1;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range rng;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            activeXl.Add(xlApp);

            xlWorkbook = xlApp.Workbooks.Add(misValue);
            activeXl.Add(xlWorkbook);

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            activeXl.Add(xlWorksheet);

            xlWorksheet.Cells[1, 1] = "Ticket No";
            xlWorksheet.Cells[1, 2] = "Summary";
            xlWorksheet.Cells[1, 3] = "Tech";
            xlWorksheet.Cells[1, 4] = "Account";
            xlWorksheet.Cells[1, 5] = "Type";
            xlWorksheet.Cells[1, 6] = "Status";
            xlWorksheet.Cells[1, 7] = "Notes";

            rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[1, 1], xlWorksheet.Cells[1, 7]);
            rng.Interior.Color = ColorTranslator.ToOle(Color.GreenYellow);

            try
            {
                progressBar pf = new progressBar();
                pf.Show();

                int row = 1;
                foreach (TechTixRecord record in threadVars.TechTixRecords)
                {
                    row++;

                    xlWorksheet.Cells[row, 1] = record.TicketNo;
                    xlWorksheet.Cells[row, 2] = record.Summary;
                    xlWorksheet.Cells[row, 3] = record.Tech;
                    xlWorksheet.Cells[row, 4] = record.Account;
                    xlWorksheet.Cells[row, 5] = record.Type;
                    xlWorksheet.Cells[row, 6] = record.Status;
                    xlWorksheet.Cells[row, 7] = record.Notes;

                    rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[row, 1], xlWorksheet.Cells[row, 7]);
                    if (row % 2 == 0)
                    {
                        rng.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                    }
                    else
                    {
                        rng.Interior.Color = ColorTranslator.ToOle(Color.White);
                    }

                    xlWorksheet.Rows.RowHeight = 16.5;

                    percent = findPercent(row, count);
                    pf.WindowName();
                    pf.UpdateProgress(percent, row, count);
                }
                //SaveFileDialog attempt = new SaveFileDialog();
                //attempt.Filter = "Excel Files | *.xlsx";
                //attempt.DefaultExt = "xlsx";
                //attempt.ShowDialog();
                //saveAs = attempt.FileName;

                //xlWorkbook.SaveAs(saveAs, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                //MessageBox.Show("Your report was succesfully created!");

                saveXL(xlWorkbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show("There was an issue writing your report to Excel.");
            }
            finally
            {
                if (!cancelClick)
                {
                    xlWorkbook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorksheet);
                    releaseObject(xlWorkbook);
                    releaseObject(xlApp);
                }
            }
        }//end writeListToXlsx(TechTix)

        //private void writeBillToXlEven()
        //{
        //    threadVars.EvenComplete = false;

        //    int row = 1;
        //    foreach (BillingRecord record in threadVars.BillingRecords)
        //    {
        //        row++;

        //        if (row % 2 == 0)
        //        {
        //            xlWorksheet.Cells[row, 1] = record.TicketNo;
        //            xlWorksheet.Cells[row, 2] = record.Summary;
        //            xlWorksheet.Cells[row, 3] = record.Resolution;
        //            xlWorksheet.Cells[row, 4] = record.CreatedAt;
        //            xlWorksheet.Cells[row, 5] = record.ClosedAt;
        //            xlWorksheet.Cells[row, 6] = record.TotalHours;
        //            xlWorksheet.Cells[row, 7] = record.TotalLabor;
        //            xlWorksheet.Cells[row, 8] = record.Purchases;
        //            xlWorksheet.Cells[row, 9] = record.Expenses;
        //            xlWorksheet.Cells[row, 10] = record.GrandTotal;

        //            rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[row, 1], xlWorksheet.Cells[row, 10]);
        //            rng.Interior.Color = ColorTranslator.ToOle(Color.LightGray);

        //            xlWorksheet.Rows.RowHeight = 16.5;
        //        }
        //    }

        //    threadVars.EvenComplete = true;
        //    WriteToXlComplete();
        //}//end writeBillToXlEven()

        //private void writeBillToXlOdd()
        //{
        //    threadVars.OddComplete = false;

        //    int row = 1;
        //    foreach (BillingRecord record in threadVars.BillingRecords)
        //    {
        //        row++;

        //        if (row % 2 != 0)
        //        {
        //            xlWorksheet.Cells[row, 1] = record.TicketNo;
        //            xlWorksheet.Cells[row, 2] = record.Summary;
        //            xlWorksheet.Cells[row, 3] = record.Resolution;
        //            xlWorksheet.Cells[row, 4] = record.CreatedAt;
        //            xlWorksheet.Cells[row, 5] = record.ClosedAt;
        //            xlWorksheet.Cells[row, 6] = record.TotalHours;
        //            xlWorksheet.Cells[row, 7] = record.TotalLabor;
        //            xlWorksheet.Cells[row, 8] = record.Purchases;
        //            xlWorksheet.Cells[row, 9] = record.Expenses;
        //            xlWorksheet.Cells[row, 10] = record.GrandTotal;

        //            rng = (Excel.Range)xlWorksheet.get_Range(xlWorksheet.Cells[row, 1], xlWorksheet.Cells[row, 10]);
        //            rng.Interior.Color = ColorTranslator.ToOle(Color.White);

        //            xlWorksheet.Rows.RowHeight = 16.5;
        //        }
        //    }

        //    threadVars.OddComplete = true;
        //    WriteToXlComplete();
        //}//end writeBillToXlOdd()

        //public void WriteToXlComplete()
        //{
        //    if (threadVars.EvenComplete && threadVars.OddComplete)
        //    {
        //        String saveAs;

        //        SaveFileDialog attempt = new SaveFileDialog();
        //        attempt.Filter = "Excel Files | *.xlsx";
        //        attempt.DefaultExt = "xlsx";
        //        attempt.ShowDialog();
        //        saveAs = attempt.FileName;

        //        xlWorkbook.SaveAs(saveAs);

        //        MessageBox.Show("Your report was succesfully created!");

        //        xlWorkbook.Close(true, misValue, misValue);
        //        xlApp.Quit();

        //        releaseObject(xlWorksheet);
        //        releaseObject(xlWorkbook);
        //        releaseObject(xlApp);
        //    }
        //}

        public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occurred while releasing object" + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }//end releaseObject()

        public void cancelActiveThreads()
        {
            if (threadVars.ActiveThreads != null)
            {
                foreach (object xl in activeXl)
                {
                    String test = xl.ToString();

                    if (xl.ToString() == "Microsoft.Office.Interop.Excel.ApplicationClass")
                    {
                        Excel.Application xlApp = (Excel.Application)xl;
                        xlApp.DisplayAlerts = false;
                        xlApp.Quit();
                    }

                    releaseObject(xl);
                }

                foreach (Thread t in threadVars.ActiveThreads)
                {
                    if (t.IsAlive)
                    {
                        t.Abort();
                    }
                }
            }
            else
            {
                MessageBox.Show("There are currently no reports running");
            }
        }//end cancelActiveThreads()

        private int findPercent(int row, int count)
        {
            //decimal answer;
            int answer;
            decimal findDecimal;

            //answer = row / count;
            //answer = decimal.Divide(row, count);
            findDecimal = decimal.Divide(row, count);
            findDecimal = findDecimal * 100;
            answer = (int)findDecimal;

            return answer;
        }//end findPercent()

        private void showProgress(progressBar pf, int percent)
        {
            pf.WindowName();
            pf.ProgPercent = percent;
        }//end showProgress()

        private void saveXL(Excel.Workbook wb)
        {
            String saveAs;
            SaveFileDialog attempt = new SaveFileDialog();
            attempt.FileName = threadVars.Account + threadVars.ReportType;
            attempt.Filter = "Excel Files | *.xlsx";
            attempt.DefaultExt = "xlsx";
            attempt.ShowDialog();
            saveAs = attempt.FileName;

            wb.SaveAs(saveAs);

            MessageBox.Show("Your report was successfully created");
        }//end saveXL()
    }//end class
}//end namespace
