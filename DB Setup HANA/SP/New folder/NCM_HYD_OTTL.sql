delete from [@NCM_QUERY] where U_Type = 'NCM_HYD_OTTL'

insert into [@NCM_QUERY] ([code],[name],[u_type],[u_query])
values ('00002222','00002222','NCM_HYD_OTTL',
'
select SUM((X.IGEQTy + X.DLNQty + X.RDNQty + X.INVQty + X.RINQty) * ISNULL(D.AvgPrice,0)) AS TotalValue
FROM (
select A.ItemCode, A.Itemname, A.PrcrmntMtd, 
(select ISNULL(SUM(B.Quantity),0) from [IGE1] B LEFT OUTER JOIN [OIGE] C ON B.DOCENTRY = C.DOCENTRY where A.ItemCode = B.ItemCode AND B.BaseType = ''202'' AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) AS IGEQty,
(select ISNULL(SUM(B.Quantity),0) from [INV1] B	LEFT OUTER JOIN [OINV] C ON B.DOCENTRY = C.DOCENTRY where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) AS INVQty,
(select ISNULL(SUM(B.Quantity),0) * -1 from [RIN1] B LEFT OUTER JOIN [ORIN] C ON B.DOCENTRY = C.DOCENTRY where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) AS RINQty,
(select ISNULL(SUM(B.Quantity),0) from [DLN1] B	LEFT OUTER JOIN [ODLN] C ON B.DOCENTRY = C.DOCENTRY where A.ItemCode = B.ItemCode AND (B.LineStatus = ''O'' OR (B.LineStatus = ''C'' AND B.TargetType =''16'')) AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) AS DLNQty,
(select ISNULL(SUM(B.Quantity),0) * -1 from [RDN1] B LEFT OUTER JOIN [ORDN] C ON B.DOCENTRY = C.DOCENTRY where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) AS RDNQty
From [OITM] A 
) X  
LEFT OUTER JOIN [OITW] D ON X.ItemCode = D.ItemCode
WHERE X.PrcrmntMtd = ''B''
')