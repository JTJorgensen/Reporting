using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace ReportingPSOL
{
    public static class threadVars
    {
        private static String reportType;
        private static String query;
        private static String account;
        private static List<BillingRecord> billingRecords = new List<BillingRecord>();
        private static List<TicketRecord> ticketRecords = new List<TicketRecord>();
        private static List<AOBillingRecord> aoBillingRecords = new List<AOBillingRecord>();
        private static List<AOTicketRecord> aoTicketRecords = new List<AOTicketRecord>();
        private static List<TechTixRecord> techTixRecords = new List<TechTixRecord>();
        private static List<Thread> activeThreads = new List<Thread>();
        private static bool oddComplete;
        private static bool evenComplete;

        public static bool EvenComplete
        {
            get { return evenComplete; }
            set { evenComplete = value; }
        }
        

        public static bool OddComplete
        {
            get { return oddComplete; }
            set { oddComplete = value; }
        }


        public static List<Thread> ActiveThreads
        {
            get { return activeThreads; }
            set { activeThreads = value; }
        }


        public static List<TechTixRecord> TechTixRecords
        {
            get { return techTixRecords; }
            set { techTixRecords = value; }
        }
        

        public static List<AOTicketRecord> AOTicketRecords
        {
            get { return aoTicketRecords; }
            set { aoTicketRecords = value; }
        }
        

        public static List<AOBillingRecord> AOBillingRecords
        {
            get { return aoBillingRecords; }
            set { aoBillingRecords = value; }
        }
        

        public static List<TicketRecord> TicketRecords
        {
            get { return ticketRecords; }
            set { ticketRecords = value; }
        }
        

        public static List<BillingRecord> BillingRecords
        {
            get { return billingRecords; }
            set { billingRecords = value; }
        }
        

        public static String Account
        {
            get { return account; }
            set { account = value; }
        }


        public static String Query
        {
            get { return query; }
            set { query = value; }
        }
        

        public static String ReportType
        {
            get { return reportType; }
            set { reportType = value; }
        }
        
    }
}
