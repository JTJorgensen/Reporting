using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportingPSOL
{
    public class TechTixRecord
    {
        private object ticketNo;
        private object summary;
        private object tech;
        private object account;
        private object type;
        private object status;
        private object notes;

        public object Notes
        {
            get { return notes; }
            set { notes = value; }
        }
        

        public object Status
        {
            get { return status; }
            set { status = value; }
        }
        

        public object Type
        {
            get { return type; }
            set { type = value; }
        }
        

        public object Account
        {
            get { return account; }
            set { account = value; }
        }
        

        public object Tech
        {
            get { return tech; }
            set { tech = value; }
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
