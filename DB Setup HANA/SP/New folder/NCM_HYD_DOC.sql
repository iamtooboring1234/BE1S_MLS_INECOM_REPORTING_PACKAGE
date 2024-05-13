delete from [@NCM_QUERY] where U_Type = 'NCM_HYD_DOC'

insert into [@NCM_QUERY] ([code],[name],[u_type],[u_query])
values ('00005555','00005555','NCM_HYD_DOC',
'
SELECT 	T1.DocEntry, ''IGE1'' AS DocType, T1.Quantity 
FROM 	[IGE1] T1 LEFT OUTER JOIN [OIGE] T2 ON T1.DocEntry = T2.DocEntry
WHERE 	T1.ItemCode = ''<<ITEMCODE>>'' AND T1.BaseType = ''202'' AND T2.DocDate >= DATEADD(MONTH, -12, GETDATE())
UNION ALL
SELECT 	T1.DocEntry, ''DLN1'' AS DocType, T1.Quantity 
FROM 	[DLN1] T1 LEFT OUTER JOIN [ODLN] T2 ON T1.DocEntry = T2.DocEntry
WHERE 	T1.ItemCode = ''<<ITEMCODE>>'' AND (T1.LineStatus = ''O'' OR (T1.LineStatus =''C'' AND T1.TargetType = ''16'')) AND T2.DocDate >= DATEADD(MONTH, -12, GETDATE())
UNION ALL
SELECT 	T1.DocEntry, ''RDN1'' AS DocType, T1.Quantity 
FROM 	[RDN1 ]T1 LEFT OUTER JOIN [ORDN] T2 ON T1.DocEntry = T2.DocEntry
WHERE 	T1.ItemCode = ''<<ITEMCODE>>'' AND (T1.LineStatus = ''O'' OR (T1.LineStatus =''C'' AND T1.BaseType = ''15'')) AND T2.DocDate >= DATEADD(MONTH, -12, GETDATE())
UNION ALL
SELECT 	T1.DocEntry, ''INV1'' AS DocType, T1.Quantity 
FROM 	[INV1] T1 LEFT OUTER JOIN [OINV] T2 ON T1.DocEntry = T2.DocEntry
WHERE 	T1.ItemCode = ''<<ITEMCODE>>'' AND T2.DocDate >= DATEADD(MONTH, -12, GETDATE())
UNION ALL
SELECT 	T1.DocEntry, ''RIN1'' AS DocType,T1. Quantity 
FROM 	[RIN1] T1 LEFT OUTER JOIN [ORIN] T2 ON T1.DocEntry = T2.DocEntry
WHERE 	T1.ItemCode = ''<<ITEMCODE>>'' AND T2.DocDate >= DATEADD(MONTH, -12, GETDATE())
')

