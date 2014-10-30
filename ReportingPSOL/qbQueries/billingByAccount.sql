SELECT 
	(REPLACE(REPLACE(t.TicketNo, '0500-000000', ''), '0500-00000', '')) AS "TicketNo",
	t.Problem AS "Summary",
	t.Solution AS "Resolution",
	t.OpenDateTime AS "Created", 
	CASE
		WHEN t.CloseDateTime = '1899-12-30 00:00:00.000' THEN NULL
		WHEN t.CloseDateTime = '1900-01-01 00:00:00.000' THEN NULL
		ELSE t.CloseDateTime
	END AS "Closed",
	(SELECT SUM(HoursAmount)
		FROM CommitReport..slips s
		WHERE s.TicketId = t.RecId
		AND s.SlipDate BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999') AS "Hours",
	(SELECT SUM(Total)
		FROM CommitReport..slips s
		WHERE s.TicketId = t.RecId
		AND s.ItemId = 'ITM1Q3GUI05ANBQGVY8D'
		AND s.SlipDate BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999') AS "Labor",
	(SELECT SUM(Total)
		FROM CommitReport..slips s
		WHERE s.TicketId = t.RecId
		AND s.ItemId = 'ITMR4H2IFX8RLJKE9O0X'
		AND s.SlipDate BETWEEN '2014-10-01 00:00:00.000'AND '2014-10-28 23:59:59.999') AS "Purchases",
	(SELECT SUM(Total)
		FROM CommitReport..slips s
		WHERE s.TicketId = t.RecId
		AND s.ItemId = ''
		AND s.SlipDate BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999') AS "Expenses",
	(SELECT SUM(Total)
		FROM CommitReport..slips s 
		WHERE s.TicketId = t.RecId
		AND s.SlipDate BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999') AS "Total"
FROM CommitReport..tickets t
WHERE t.CardId = ''
AND (t.RecId IN (SELECT TicketId
					FROM CommitReport..slips s
					WHERE s.SlipDate BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999')
	OR (t.OpenDateTime BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999'))
ORDER BY t.TicketNo