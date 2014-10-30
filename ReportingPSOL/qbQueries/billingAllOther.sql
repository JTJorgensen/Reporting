SELECT 
	(REPLACE(REPLACE(t.TicketNo, '0500-000000', ''), '0500-00000', '')) AS "TicketNo",
	(SELECT RTRIM(c.FullName)
		FROM CommitReport..cards c
		WHERE t.CardId = c.RecId) AS "Account",
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
		AND s.SlipDate BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999') AS "Purchases",
	(SELECT SUM(Total)
		FROM CommitReport..slips s
		WHERE s.TicketId = t.RecId
		AND s.ItemId = 'ITMKVB8R5PD9I5W6HYST'
		AND s.SlipDate BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999') AS "Expenses",
	(SELECT SUM(Total)
		FROM CommitReport..slips s
		WHERE s.TicketId = t.RecId
		AND s.SlipDate BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999') AS "Total"
FROM CommitReport..tickets t
WHERE t.CardId != 'CRDC1H8Z5YZAZ1XQR41R'
AND t.CardId != 'CRDW7AZIXTOR48TLSHEV'
AND t.CardId != 'CRDW8SST36IWP006EQA1'
AND t.CardId != 'CRDGLQ0FDVKTBB6JD7DZ'
AND (t.RecId IN (SELECT TicketId
					FROM CommitReport..slips
					WHERE SlipDate BETWEEN '2014-10-01 00:00:00.00' AND '2014-10-28 23:59:59.999')
		OR (t.OpenDateTime BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999'))
ORDER BY t.TicketNo