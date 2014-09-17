using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportingPSOL
{
    public class BillingRecord
    {
        private object ticketNo;
        private object summary;
        private object resolution;
        private object createdAt;
        private object closedAt;
        private object totalHours;
        private object totalLabor;
        private object purchases;
        private object expenses;
        private object grandTotal;

        public object GrandTotal
        {
            get { return grandTotal; }
            set { grandTotal = value; }
        }
        

        public object Expenses
        {
            get { return expenses; }
            set { expenses = value; }
        }
        

        public object Purchases
        {
            get { return purchases; }
            set { purchases = value; }
        }
        

        public object TotalLabor
        {
            get { return totalLabor; }
            set { totalLabor = value; }
        }
        

        public object TotalHours
        {
            get { return totalHours; }
            set { totalHours = value; }
        }
        

        public object ClosedAt
        {
            get { return closedAt; }
            set { closedAt = value; }
        }
        

        public object CreatedAt
        {
            get { return createdAt; }
            set { createdAt = value; }
        }
        

        public object Resolution
        {
            get { return resolution; }
            set { resolution = value; }
        }
        

        public object Summary
        {
            get { return summary; }
            set { summary = value; }
        }
        

        public object TicketNo
        {
            get { return ticketNo; }
            set { ticketNo = value; }
        }
        
        
    }
}
