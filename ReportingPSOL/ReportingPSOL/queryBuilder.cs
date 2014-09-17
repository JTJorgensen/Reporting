using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportingPSOL
{
    class queryBuilder
    {
        public String billingReportByAccount(String startDate, String endDate, String accCard)
        {
            String result;

            #region Billing Query No HardCoded Dates/Accounts
            result = "SELECT  (REPLACE(t.TICKETNO, '0500-000000', '')) As \"Ticket No\", " +
                 "t.PROBLEM As \"Summary\", " +
                 "t.SOLUTION As \"Resolution\", " +
                 "t.OPENDATETIME As \"Created\", " +
                 "CASE " +
                    "WHEN t.CLOSEDATETIME = CAST('12/30/1899 00:00:00' as SQL_TIMESTAMP) THEN NULL " +
                    "ELSE t.CLOSEDATETIME " +
                "END AS \"Closed\", " +
                "(SELECT SUM(HOURSAMOUNT) " +
                    "FROM slips s " +
                    "WHERE s.TICKETID = t.RECID " +
                    "AND s.SLIPDATE BETWEEN CAST('" + startDate + " 00:00:00' as SQL_TIMESTAMP) AND CAST('" + endDate + " 00:00:00' as SQL_TIMESTAMP)) AS \"Hours\", " +
                "(SELECT SUM(TOTAL) " +
                    "FROM slips s " +
                    "WHERE s.TICKETID = t.RECID " +
                    "AND s.ITEMID = 'ITM1Q3GUI05ANBQGVY8D' " +
                    "AND s.SLIPDATE BETWEEN CAST('" + startDate + " 00:00:00' as SQL_TIMESTAMP) AND CAST('" + endDate + " 00:00:00' as SQL_TIMESTAMP)) As \"Labor\", " +
                "(SELECT SUM(TOTAL) " +
                    "FROM slips s " +
                    "WHERE s.TICKETID = t.RECID " +
                    "AND s.ITEMID = 'ITMR4H2IFX8RLJKE9O0X' " +
                    "AND s.SLIPDATE BETWEEN CAST('" + startDate + " 00:00:00' as SQL_TIMESTAMP) AND CAST('" + endDate + " 00:00:00' as SQL_TIMESTAMP)) As \"Purchases\", " +
                "(SELECT SUM(TOTAL) " +
                    "FROM slips s " +
                    "WHERE s.TICKETID = t.RECID " +
                    "AND s.ITEMID = 'ITMKVB8R5PD9I5W6HYST' " +
                    "AND s.SLIPDATE BETWEEN CAST('" + startDate + " 00:00:00' as SQL_TIMESTAMP) AND CAST('" + endDate + " 00:00:00' as SQL_TIMESTAMP)) As \"Expenses\", " +
                "(SELECT SUM(TOTAL) " +
                    "FROM slips s " +
                    "WHERE s.TICKETID = t.RECID " +
                    "AND s.SLIPDATE BETWEEN CAST('" + startDate + " 00:00:00' as SQL_TIMESTAMP) AND CAST('" + endDate + " 00:00:00' as SQL_TIMESTAMP)) As \"Total\" " +
                "FROM tickets t " +
                "WHERE t.CARDID = '" + accCard + "' " +
                "AND (t.RECID IN (SELECT TICKETID " +
                                    "FROM slips s " +
                                    "WHERE SLIPDATE BETWEEN CAST('" + startDate + " 00:00:00' as SQL_TIMESTAMP) AND CAST('" + endDate + " 00:00:00' as SQL_TIMESTAMP)) " +
                    "OR (t.OPENDATETIME BETWEEN CAST('" + startDate + " 00:00:00' as SQL_TIMESTAMP) AND CAST('" + endDate + " 00:00:00' as SQL_TIMESTAMP))) " +
                "ORDER BY t.TICKETNO";
            #endregion

            return result;
        }//end billingReportsByAccount()

        public String billingReportAllOther(String startDate, String endDate)
        {
            String result;

            #region AO billing query no hardcoded dates
            result = "SELECT (REPLACE(t.TICKETNO, '0500-000000', '')) AS \"Ticket No\", " +
                "(SELECT c.FULLNAME " +
                    "FROM cards c " +
                    "WHERE t.CARDID = c.RECID) AS \"Account\", " +
                "t.PROBLEM AS \"Summary\", " +
                "t.SOLUTION AS \"Resolution\", " +
                "t.OPENDATETIME AS \"Created\", " +
                "CASE " +
                    "WHEN t.CLOSEDATETIME = CAST('12/30/1899 00:00:00' AS SQL_TIMESTAMP) THEN NULL " +
                    "ELSE t.CLOSEDATETIME " +
                "END AS \"Closed\", " +
                "(SELECT SUM(HOURSAMOUNT) " +
                    "FROM slips s " +
                    "WHERE s.TICKETID = t.RECID " +
                    "AND s.SLIPDATE >= CAST('" + startDate + " 00:00:00' AS SQL_TIMESTAMP) " +
                    "AND s.SLIPDATE <= CAST('" + endDate + " 00:00:00' AS SQL_TIMESTAMP)) AS \"Hours\", " +
                "(SELECT SUM(TOTAL) " +
                    "FROM slips s " +
                    "WHERE s.TICKETID = t.RECID " +
                    "AND s.ITEMID = 'ITM1Q3GUI05ANBQGVY8D' " +
                    "AND s.SLIPDATE >= CAST('" + startDate + " 00:00:00' AS SQL_TIMESTAMP) " +
                    "AND s.SLIPDATE <= CAST('" + endDate + " 00:00:00' AS SQL_TIMESTAMP)) AS \"Labor\", " +
                "(SELECT SUM(TOTAL) " +
                    "FROM slips s " +
                    "WHERE s.TICKETID = t.RECID " +
                    "AND s.ITEMID = 'ITMR4H2IFX8RLJKE9O0X' " +
                    "AND s.SLIPDATE >= CAST('" + startDate + " 00:00:00' AS SQL_TIMESTAMP) " +
                    "AND s.SLIPDATE <= CAST('" + endDate + " 00:00:00' AS SQL_TIMESTAMP)) AS \"Purchases\", " +
                "(SELECT SUM(TOTAL) " +
                    "FROM slips s " +
                    "WHERE s.TICKETID = t.RECID " +
                    "AND s.ITEMID = 'ITMKVB8R5PD9I5W6HYST' " +
                    "AND s.SLIPDATE >= CAST('" + startDate + " 00:00:00' AS SQL_TIMESTAMP) " +
                    "AND s.SLIPDATE <= CAST('" + endDate + " 00:00:00' AS SQL_TIMESTAMP)) AS \"Expenses\", " +
                "(SELECT SUM(TOTAL) " +
                    "FROM slips s " +
                    "WHERE s.TICKETID = t.RECID " +
                    "AND s.SLIPDATE >= CAST('" + startDate + " 00:00:00' AS SQL_TIMESTAMP) " +
                    "AND s.SLIPDATE <= CAST('" + endDate + " 00:00:00' AS SQL_TIMESTAMP)) AS \"Total\" " +
                "FROM tickets t " +
                "WHERE t.CARDID <> 'CRDC1H8Z5YZAZ1XQR41R' " +
                "AND t.CARDID <> 'CRDW7AZIXTOR48TLSHEV' " +
                "AND t.CARDID <> 'CRDW8SST36IWP006EQA1' " +
                "AND t.CARDID <> 'CRDGLQ0FDVKTBB6JD7DZ' " +
                "AND (t.RECID IN (SELECT TICKETID " +
                        "FROM slips " +
                        "WHERE SLIPDATE >= CAST('" + startDate + " 00:00:00' AS SQL_TIMESTAMP) " +
                        "AND SLIPDATE <= CAST('" + endDate + " 00:00:00' AS SQL_TIMESTAMP)) " +
                    "OR (t.OPENDATETIME >= CAST('" + startDate + " 00:00:00' AS SQL_TIMESTAMP) " +
                    "AND t.OPENDATETIME <= CAST('" + endDate + " 00:00:00' AS SQL_TIMESTAMP))) " +
                "ORDER BY t.TICKETNO ";
            #endregion

            return result;
        }//end billingReportAllOther()

        public String ticketReportByAccount(String startDate, String endDate, String accCard)
        {
            String result;

            #region Ticket Query No HardCoded Dates/Account
            result = "SELECT (REPLACE(t.TICKETNO, '0500-000000', '')) AS \"Ticket No\", " +
                "t.PROBLEM AS \"Summary\", " +
                "(SELECT TOP 1 " +
                    "CASE " +
                        "WHEN t.STATUS = 100 THEN 'New' " +
                        "WHEN t.STATUS = 200 THEN 'Pending' " +
                        "WHEN t.STATUS = 300 THEN 'Scheduled' " +
                        "WHEN t.STATUS = 400 THEN 'In-House Service' " +
                        "WHEN t.STATUS = 500 THEN 'On-Site Service' " +
                        "WHEN t.STATUS = 600 THEN 'Follow-Up' " +
                        "WHEN t.STATUS = 700 THEN 'Hold' " +
                        "WHEN t.STATUS = 800 THEN 'Other' " +
                        "WHEN t.STATUS = 900 THEN 'Cancelled' " +
                        "WHEN t.STATUS = 1000 THEN 'Completed' " +
                    "END " +
                "FROM tickets " +
                "WHERE t.CARDID = tickets.CARDID " +
                "ORDER BY UPDATEDATE) AS \"Status\", " +
                "t.OPENDATETIME AS \"Created\", " +
                "(CASE " +
                    "WHEN t.CLOSEDATETIME = CAST('12/30/1899 00:00:00' AS SQL_TIMESTAMP) THEN NULL " +
                    "ELSE t.CLOSEDATETIME " +
                "END) AS \"Closed\", " +
                "(SELECT FULLNAME " +
                    "FROM cards c " +
                    "WHERE c.RECID = t.WORKERID) AS \"Tech\" " +
                "FROM tickets t " +
                "WHERE t.CARDID = '" + accCard + "' " +
                "AND (t.OPENDATETIME BETWEEN CAST('" + startDate + " 00:00:00' AS SQL_TIMESTAMP) AND CAST('" + endDate + " 00:00:00' AS SQL_TIMESTAMP) " +
                    "OR t.CLOSEDATETIME BETWEEN CAST('" + startDate + " 00:00:00' AS SQL_TIMESTAMP) AND CAST('" + endDate + " 00:00:00' AS SQL_TIMESTAMP)) " +
                "ORDER BY t.TICKETNO";
            #endregion

            return result;
        }//end ticketReportByAccount()

        public String ticketReportAllOther(String startDate, String endDate)
        {
            String result;

            #region AO ticket report no hard coded dates
            result = "SELECT (REPLACE(t.TICKETNO, '0500-000000', '')) AS \"Ticket No\", " +
                "(SELECT c.FULLNAME " +
                    "FROM cards c " +
                    "WHERE t.CARDID = c.RECID) AS \"Account\", " +
                "t.PROBLEM AS \"Summary\", " +
                "(SELECT TOP 1 " +
                    "CASE " +
                        "WHEN t.STATUS = 100 THEN 'New' " +
                        "WHEN t.STATUS = 200 THEN 'Pending' " +
                        "WHEN t.STATUS = 300 THEN 'Scheduled' " +
                        "WHEN t.STATUS = 400 THEN 'In-House Service' " +
                        "WHEN t.STATUS = 500 THEN 'On-Site Service' " +
                        "WHEN t.STATUS = 600 THEN 'Follow-Up' " +
                        "WHEN t.STATUS = 700 THEN 'Hold' " +
                        "WHEN t.STATUS = 800 THEN 'Other' " +
                        "WHEN t.STATUS = 900 THEN 'Cancelled' " +
                        "WHEN t.STATUS = 1000 THEN 'Completed' " +
                    "END " +
                "FROM tickets " +
                "WHERE t.CARDID = tickets.CARDID " +
                "ORDER BY UPDATEDATE) AS \"Status\", " +
                "t.OPENDATETIME AS \"Created\", " +
                "(CASE " +
                    "WHEN t.CLOSEDATETIME = CAST('12/30/1899 00:00:00' AS SQL_TIMESTAMP) THEN NULL " +
                    "ELSE t.CLOSEDATETIME " +
                "END) AS \"Closed\", " +
                "(SELECT FULLNAME " +
                    "FROM cards c " +
                    "WHERE c.RECID = t.WORKERID) AS \"Tech\" " +
                "FROM tickets t " +
                "WHERE (t.CARDID <> 'CRDC1H8Z5YZAZ1XQR41R' " +
                    "AND t.CARDID <> 'CRDW7AZIXTOR48TLSHEV' " +
                    "AND t.CARDID <> 'CRDW8SST36IWP006EQA1' " +
                    "AND t.CARDID <> 'CRDGLQ0FDVKTBB6JD7DZ') " +
                "AND (t.OPENDATETIME BETWEEN CAST('" + startDate + " 00:00:00' AS SQL_TIMESTAMP) AND CAST('" + endDate + " 00:00:00' AS SQL_TIMESTAMP) " +
                    "OR t.CLOSEDATETIME BETWEEN CAST('" + startDate + " 00:00:00' AS SQL_TIMESTAMP) AND CAST('" + endDate + " 00:00:00' AS SQL_TIMESTAMP)) " +
                "ORDER BY t.TICKETNO ";
            #endregion

            return result;
        }//end ticketReportAllOther()

        public String technicianTickets()
        {
            String result;

            #region Technician Open Ticket Report
            result = "SELECT (REPLACE(t.TICKETNO, '0500-000000', '')) AS \"Ticket No\", " +
                    "t.PROBLEM AS \"Summary\", " +
                    "(SELECT c.FULLNAME " +
                        "FROM cards c " +
                        "WHERE c.RECID = t.WORKERID) AS \"Tech\", " +
                    "(SELECT c.FULLNAME " +
                        "FROM cards c " +
                        "WHERE t.CARDID = c.RECID) AS \"Account\", " +
                    "t.KIND AS \"Type\", " +
                    "(SELECT TOP 1 " +
                        "CASE " +
                            "WHEN t.STATUS = 100 THEN 'New' " +
                            "WHEN t.STATUS = 200 THEN 'Pending' " +
                            "WHEN t.STATUS = 300 THEN 'Scheduled' " +
                            "WHEN t.STATUS = 400 THEN 'In-House Service' " +
                            "WHEN t.STATUS = 500 THEN 'On-Site Service' " +
                            "WHEN t.STATUS = 600 THEN 'Follow-Up' " +
                            "WHEN t.STATUS = 700 THEN 'Hold' " +
                            "WHEN t.STATUS = 800 THEN 'Other' " +
                            "WHEN t.STATUS = 900 THEN 'Cancelled' " +
                            "WHEN t.STATUS = 1000 THEN 'Completed' " +
                        "END " +
                    "FROM tickets " +
                    "WHERE t.CARDID = tickets.CARDID " +
                    "ORDER BY UPDATEDATE) AS \"Status\", " +
                    "t.NOTES AS \"Notes\" " +
                "FROM tickets t " +
                "WHERE t.STATUS <> 900 " +
                "AND t.STATUS <> 1000 " +
                "ORDER BY t.WORKERID ";
            #endregion

            return result;
        }//end technicianTickets()

        public String accountSelect(String accountName)
        {
            String card;

            switch (accountName)
            {
                case "Sparhawk":
                    card = "CRDGLQ0FDVKTBB6JD7DZ";
                    break;
                case "Rettler":
                    card = "CRDW8SST36IWP006EQA1";
                    break;
                case "Mullins":
                    card = "CRDW7AZIXTOR48TLSHEV";
                    break;
                case "CCCW":
                    card = "CRDC1H8Z5YZAZ1XQR41R";
                    break;
                default:
                    card = "";
                    break;
            }

            return card;
        }//end accountSelect()
    }//end class
}//end namespace
