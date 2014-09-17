using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportingPSOL
{
    public class TicketRecord
    {
        private object ticketNo;
        private object summary;
        private object status;
        private object created;
        private object closed;
        private object tech;

        public object Tech
        {
            get { return tech; }
            set { tech = value; }
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
        

        public object Status
        {
            get { return status; }
            set { status = value; }
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
