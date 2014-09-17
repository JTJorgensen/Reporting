using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Odbc;
using System.Drawing;

namespace ReportingPSOL
{
    class readWrite
    {
        public void readFromDbWriteToXlsx(String query, String reportType)
        {
            switch (reportType)
            {
                case "Billing":
                    List<BillingRecord> billing = new List<BillingRecord>();

                    billing = writeBillingToList(query);

                    if (billing != null)
                    {
                        writeListToXlsx(billing);
                    }
                    break;

                case ".Billing":
                    List<AOBillingRecord> aoBilling = new List<AOBillingRecord>();

                    aoBilling = writeAOBillingToList(query);

                    if (aoBilling != null)
                    {
                        writeListToXlsx(aoBilling);
                    }
                    break;

                case "Tickets":
                    List<TicketRecord> ticketing = new List<TicketRecord>();

                    ticketing = writeTicketToList(query);

                    if (ticketing != null)
                    {
                        writeListToXlsx(ticketing);
                    }
                    break;

                case ".Tickets":
                    List<AOTicketRecord> aoTicketing = new List<AOTicketRecord>();

                    aoTicketing = writeAOTicketToList(query);

                    if (aoTicketing != null)
                    {
                        writeListToXlsx(aoTicketing);
                    }
                    break;

                case "Technician":
                    List<TechTixRecord> techTix = new List<TechTixRecord>();

                    techTix = writeTechTixToList(query);

                    if (techTix != null)
                    {
                        writeListToXlsx(techTix);
                    }
                    break;

                default:
                    break;
            }
        }//end readFromDbWriteToXlsx()

        private List<BillingRecord> writeBillingToList(String query)
        {
            List<BillingRecord> billingRecords = new List<BillingRecord>();

            using (OdbcConnection conn = new OdbcConnection("DSN=CommitSystemODBC"))
            {
                OdbcCommand cmd = new OdbcCommand(query, conn);

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

                        billingRecords.Add(record);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    billingRecords = null;;
                }
            }
            return billingRecords;
        }//end writeBillingToList()

        private List<TicketRecord> writeTicketToList(String query)
        {
            List<TicketRecord> ticketRecords = new List<TicketRecord>();

            using (OdbcConnection conn = new OdbcConnection("DSN=CommitSystemODBC"))
            {
                OdbcCommand cmd = new OdbcCommand(query, conn);

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

                        ticketRecords.Add(record);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    ticketRecords = null; ;
                }
            }
            return ticketRecords;
        }//end writeTicketToList()

        private List<AOBillingRecord> writeAOBillingToList(String query)
        {
            List<AOBillingRecord> aoBillingRecords = new List<AOBillingRecord>();

            using (OdbcConnection conn = new OdbcConnection("DSN=CommitSystemODBC"))
            {
                OdbcCommand cmd = new OdbcCommand(query, conn);

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

                        aoBillingRecords.Add(record);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    aoBillingRecords = null; ;
                }
            }
            return aoBillingRecords;
        }//end writeAOBillingToList()

        private List<AOTicketRecord> writeAOTicketToList(String query)
        {
            List<AOTicketRecord> aoTicketRecords = new List<AOTicketRecord>();

            using (OdbcConnection conn = new OdbcConnection("DSN=CommitSystemODBC"))
            {
                OdbcCommand cmd = new OdbcCommand(query, conn);

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

                        aoTicketRecords.Add(record);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    aoTicketRecords = null; ;
                }
            }
            return aoTicketRecords;
        }//end writeAOTicketToList()

        private List<TechTixRecord> writeTechTixToList(String query)
        {
            List<TechTixRecord> techTixRecords = new List<TechTixRecord>();

            using (OdbcConnection conn = new OdbcConnection("DSN=CommitSystemODBC"))
            {
                OdbcCommand cmd = new OdbcCommand(query, conn);

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

                        techTixRecords.Add(record);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    techTixRecords = null;
                }
            }
            return techTixRecords;
        }//end writeTechTixToList()

        private void writeListToXlsx(List<BillingRecord> myList)
        {
            String saveAs;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range rng;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlWorkbook = xlApp.Workbooks.Add(misValue);
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

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
                int row = 1;
                foreach(BillingRecord record in myList)
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
                }

                SaveFileDialog attempt = new SaveFileDialog();
                attempt.Filter = "Excel Files | *.xlsx";
                attempt.DefaultExt = "xlsx";
                attempt.ShowDialog();
                saveAs = attempt.FileName;

                //xlWorkbook.SaveAs(saveAs, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkbook.SaveAs(saveAs);

                MessageBox.Show("Your report was succesfully created!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show("There was an issue writing your report to Excel.");
            }
            finally
            {
                xlWorkbook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorksheet);
                releaseObject(xlWorkbook);
                releaseObject(xlApp);
            }
        }//end writeListToXlsx(billing)

        private void writeListToXlsx(List<TicketRecord> myList)
        {
            String saveAs;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range rng;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlWorkbook = xlApp.Workbooks.Add(misValue);
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

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
                int row = 1;
                foreach (TicketRecord record in myList)
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
                }
                SaveFileDialog attempt = new SaveFileDialog();
                attempt.Filter = "Excel Files | *.xlsx";
                attempt.DefaultExt = "xlsx";
                attempt.ShowDialog();
                saveAs = attempt.FileName;

                xlWorkbook.SaveAs(saveAs, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                MessageBox.Show("Your report was succesfully created!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show("There was an issue writing your report to Excel.");
            }
            finally
            {
                xlWorkbook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorksheet);
                releaseObject(xlWorkbook);
                releaseObject(xlApp);
            }
        }//end writeListToXlsx(Ticketing)

        private void writeListToXlsx(List<AOBillingRecord> myList)
        {
            String saveAs;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range rng;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlWorkbook = xlApp.Workbooks.Add(misValue);
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

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
                int row = 1;
                foreach (AOBillingRecord record in myList)
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
                }

                SaveFileDialog attempt = new SaveFileDialog();
                attempt.Filter = "Excel Files | *.xlsx";
                attempt.DefaultExt = "xlsx";
                attempt.ShowDialog();
                saveAs = attempt.FileName;

                xlWorkbook.SaveAs(saveAs, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                MessageBox.Show("Your report was succesfully created!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show("There was an issue writing your report to Excel.");
            }
            finally
            {
                xlWorkbook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorksheet);
                releaseObject(xlWorkbook);
                releaseObject(xlApp);
            }
        }//end writeListToExcel(AOBilling)

        private void writeListToXlsx(List<AOTicketRecord> myList)
        {
            String saveAs;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range rng;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlWorkbook = xlApp.Workbooks.Add(misValue);
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

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
                int row = 1;
                foreach (AOTicketRecord record in myList)
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
                }
                SaveFileDialog attempt = new SaveFileDialog();
                attempt.Filter = "Excel Files | *.xlsx";
                attempt.DefaultExt = "xlsx";
                attempt.ShowDialog();
                saveAs = attempt.FileName;

                xlWorkbook.SaveAs(saveAs, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                MessageBox.Show("Your report was succesfully created!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show("There was an issue writing your report to Excel.");
            }
            finally
            {
                xlWorkbook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorksheet);
                releaseObject(xlWorkbook);
                releaseObject(xlApp);
            }
        }//end writeListToXlsx(Ticketing)

        private void writeListToXlsx(List<TechTixRecord> myList)
        {
            String saveAs;

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range rng;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlWorkbook = xlApp.Workbooks.Add(misValue);
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

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
                int row = 1;
                foreach (TechTixRecord record in myList)
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
                }
                SaveFileDialog attempt = new SaveFileDialog();
                attempt.Filter = "Excel Files | *.xlsx";
                attempt.DefaultExt = "xlsx";
                attempt.ShowDialog();
                saveAs = attempt.FileName;

                xlWorkbook.SaveAs(saveAs, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                MessageBox.Show("Your report was succesfully created!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show("There was an issue writing your report to Excel.");
            }
            finally
            {
                xlWorkbook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorksheet);
                releaseObject(xlWorkbook);
                releaseObject(xlApp);
            }
        }//end writeListToXlsx(TechTix)

        private void releaseObject(object obj)
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
    }//end readWrite class
}//end namespace
