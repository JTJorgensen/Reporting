using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportingPSOL
{
    public class AOBillingRecord
    {
        private object ticketNo;
        private object account;
        private object summary;
        private object resolution;
        private object created;
        private object closed;
        private object hours;
        private object labor;
        private object purchases;
        private object expenses;
        private object total;

        public object Total
        {
            get { return total; }
            set { total = value; }
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
        

        public object Labor
        {
            get { return labor; }
            set { labor = value; }
        }
        

        public object Hours
        {
            get { return hours; }
            set { hours = value; }
        }
        

        public object Closed
        {
            get { return closed; }
            set { closed = value; }
        }
        

        public object Created
        {
            get { return created; }
            set { created = value; }
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
        

        public object Account
        {
            get { return account; }
            set { account = value; }
        }
        

        public object TicketNo
        {
            get { return ticketNo; }
            set { ticketNo = value; }
        }
        
    }
}
