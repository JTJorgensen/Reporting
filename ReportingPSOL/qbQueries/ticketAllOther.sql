SELECT
	(REPLACE(REPLACE(t.TicketNo, '0500-000000', ''), '0500-00000', '')) AS "TicketNo",
	(SELECT RTRIM(c.FullName)
		FROM CommitReport..cards c
		WHERE t.CardId = c.RecId) AS "Account",
	(SUBSTRING(t.Problem, 1, 200)) AS "Summary",
	(SELECT TOP 1
		CASE
			WHEN t.Status = 100 THEN 'New'
			WHEN t.Status = 200 THEN 'Pending'
			WHEN t.Status = 300 THEN 'Scheduled'
			WHEN t.Status = 400 THEN 'In-House Service'
			WHEN t.Status = 500 THEN 'On-Site Service'
			WHEN t.Status = 600 THEN 'Follow-Up'
			WHEN t.Status = 700 THEN 'Hold'
			WHEN t.Status = 800 THEN 'In Progress'
			WHEN t.Status = 900 THEN 'Cancelled'
			WHEN t.Status = 1000 THEN 'Completed'
		END
	FROM CommitReport..tickets
	WHERE t.CardId = tickets.CardId
	ORDER BY UpdateDate) AS "Status",
	t.OpenDateTime AS "Created",
	(CASE
		WHEN t.CloseDateTime = '1899-12-30 00:00:00.000' THEN NULL
		WHEN t.CloseDateTime = '1900-01-01 00:00:00.000' THEN NULL
		ELSE t.CloseDateTime
	END) AS "Closed",
	(SELECT RTRIM(c.FullName)
		FROM CommitReport..cards c
		WHERE c.RecId = t.WorkerId) AS "Tech"
FROM CommitReport..tickets t
WHERE (t.CardId != 'CRDC1H8Z5YZAZ1XQR41R'
	AND t.CardId != 'CRDW7AZIXTOR48TLSHEV'
	AND t.CardId != 'CRDW8SST36IWP006EQA1'
	AND t.CardId != 'CRDGLQ0FDVKTBB6JD7DZ')
AND (t.OpenDateTime BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999'
	OR t.CloseDateTime BETWEEN '2014-10-01 00:00:00.000' AND '2014-10-28 23:59:59.999')
ORDER BY t.TicketNo