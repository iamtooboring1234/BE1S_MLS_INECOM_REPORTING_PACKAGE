delete from [@NCM_QUERY] where U_Type = 'NCM_HYD_WAR'


insert into [@NCM_QUERY] ([code],[name],[u_type],[u_query])
values ('00006666','00006666','NCM_HYD_WAR',
'
select A.ItemCode, A.Itemname, 
(
SELECT ISNULL(SUM(A1.Qty),0) AS Quantity1 FROM
(
select ISNULL(SUM(B.Quantity),0) AS Qty from [IGE1] B 			
LEFT OUTER JOIN [OIGE] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND B.BaseType = ''202'' AND C.DocDate >= DATEADD(MONTH, -18, GETDATE()) AND C.DocDate < DATEADD(MONTH, -12, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) AS Qty from [DLN1] B 			
LEFT OUTER JOIN [ODLN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND (B.LineStatus = ''O'' OR (B.LineStatus = ''C'' AND B.TargetType =''16'')) AND C.DocDate >= DATEADD(MONTH, -18, GETDATE()) AND C.DocDate < DATEADD(MONTH, -12, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) * -1 AS Qty from [RDN1] B 			
LEFT OUTER JOIN [ORDN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -18, GETDATE()) AND C.DocDate < DATEADD(MONTH, -12, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) AS Qty from [INV1] B 			
LEFT OUTER JOIN [OINV] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -18, GETDATE()) AND C.DocDate < DATEADD(MONTH, -12, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) * -1 AS Qty from [RIN1] B 			
LEFT OUTER JOIN [ORIN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -18, GETDATE()) AND C.DocDate < DATEADD(MONTH, -12, GETDATE())
) A1) AS Quantity1,
(
SELECT ISNULL(SUM(A1.Qty),0) AS Quantity2 FROM
(
select ISNULL(SUM(B.Quantity),0) AS Qty from [IGE1] B 			
LEFT OUTER JOIN [OIGE] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -12, GETDATE()) AND C.DocDate < DATEADD(MONTH, -6, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) AS Qty from [DLN1] B 			
LEFT OUTER JOIN [ODLN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND (B.LineStatus = ''O'' OR (B.LineStatus = ''C'' AND B.TargetType =''16'')) AND C.DocDate >= DATEADD(MONTH, -12, GETDATE()) AND C.DocDate < DATEADD(MONTH, -6, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) * -1 AS Qty from [RDN1] B 			
LEFT OUTER JOIN [ORDN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -12, GETDATE()) AND C.DocDate < DATEADD(MONTH, -6, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) AS Qty from [INV1] B 			
LEFT OUTER JOIN [OINV] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -12, GETDATE()) AND C.DocDate < DATEADD(MONTH, -6, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) * -1 AS Qty from [RIN1] B 			
LEFT OUTER JOIN [ORIN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -12, GETDATE()) AND C.DocDate < DATEADD(MONTH, -6, GETDATE())
) A1) AS Quantity2,
(
SELECT ISNULL(SUM(A1.Qty),0) AS Quantity3 FROM
(
select ISNULL(SUM(B.Quantity),0) AS Qty from [IGE1] B 			
LEFT OUTER JOIN [OIGE] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -6, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) AS Qty from [DLN1] B 			
LEFT OUTER JOIN [ODLN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND (B.LineStatus = ''O'' OR (B.LineStatus = ''C'' AND B.TargetType =''16'')) AND C.DocDate >= DATEADD(MONTH, -6, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) * -1 AS Qty from [RDN1] B 			
LEFT OUTER JOIN [ORDN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -6, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) AS Qty from [INV1] B 			
LEFT OUTER JOIN [OINV] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -6, GETDATE())
UNION ALL
select ISNULL(SUM(B.Quantity),0) * -1 AS Qty from [RIN1] B 			
LEFT OUTER JOIN [ORIN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -6, GETDATE())
) A1) AS Quantity3
From [OITM] A LEFT OUTER JOIN [@NCM_ITEMCAT] C ON ISNULL(A.U_ItemCat,'''') = ISNULL(C.Code,'''') 
')


