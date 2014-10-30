SELECT 
	(REPLACE(REPLACE(t.TicketNo, '0500-000000', ''), '0500-00000', '')) AS "TicketNo",
	(SUBSTRING(t.Problem, 1, 200)) AS "Summary",
	(SELECT RTRIM(c.FullName)
		FROM CommitReport..cards c
		WHERE c.RecId = t.WorkerId) AS "Tech",
	(SELECT RTRIM(c.FullName)
		FROM CommitReport..cards c
		WHERE c.RecId = t.Cardid) AS "Account",
	(RTRIM(t.Kind)) AS "Type",
	(SELECT TOP 1
		CASE
			WHEN t.Priority = 10 THEN 'Immediate'
			WHEN t.Priority = 20 THEN 'High'
			WHEN t.Priority = 30 THEN 'Normal'
			WHEN t.Priority = 40 THEN 'Low'
			WHEN t.Priority > 40 THEN 'Not Applicable'
		END
	FROM CommitReport..tickets
	WHERE t.CardId = tickets.CardId
	ORDER BY UpdateDate) AS "Priority",
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
	t.Notes AS "Notes"
FROM CommitReport..tickets t
WHERE t.Status != 900
AND t.Status != 1000
ORDER BY t.WorkerId