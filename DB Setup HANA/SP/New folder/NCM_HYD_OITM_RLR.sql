delete from [@NCM_QUERY] where U_Type = 'NCM_HYD_OITM_RLR'

insert into [@NCM_QUERY] ([code],[name],[u_type],[u_query])
values ('00003333','00003333','NCM_HYD_OITM_RLR',
'
select A.ItemCode, A.Itemname, ISNULL(D.AvgPrice,0) As AvgPrice, ISNULL(A.U_ItemCat,'''') AS ItemCat, 
ISNULL(C.[Name],'''') AS Class, A.MinLevel, ''0'' AS SmoothQty,
''0'' 	AS TOTALVALUE,
''0''	AS VALUE_A,
''0'' 	AS VALUE_B,
''0'' 	AS VALUE_C,

(select ISNULL(SUM(B.Quantity),0) from [IGE1] B 			
LEFT OUTER JOIN [OIGE] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND B.BaseType = ''202'' 
AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) AS IGEQty,
(select ISNULL(SUM(B.Quantity),0) from [INV1] B	 		
LEFT OUTER JOIN [OINV] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) AS INVQty,
(select ISNULL(SUM(B.Quantity),0) * -1 from [RIN1] B	
LEFT OUTER JOIN [ORIN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) AS RINQty,

(select ISNULL(SUM(B.Quantity),0) from [DLN1] B	 		
LEFT OUTER JOIN [ODLN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode 
AND (B.LineStatus = ''O'' OR (B.LineStatus = ''C'' AND B.TargetType =''16'')) 
AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) AS DLNQty,
(select ISNULL(SUM(B.Quantity),0) * -1 from [RDN1] B	
LEFT OUTER JOIN [ORDN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode 
AND (B.LineStatus = ''O'' OR (B.LineStatus = ''C'' AND B.BaseType =''15'')) 
AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) AS RDNQty,

(select COUNT (DISTINCT B.DOCENTRY) from [IGE1] B		
LEFT OUTER JOIN [OIGE] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND B.BaseType = ''202'' 
AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) As IGESum,
(select COUNT (DISTINCT B.DOCENTRY) from [INV1] B		
LEFT OUTER JOIN [OINV] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) As INVSum,
(select COUNT (DISTINCT B.DOCENTRY) * -1 from [RIN1] B	
LEFT OUTER JOIN [ORIN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) As RINSum,
(select COUNT (DISTINCT B.DOCENTRY) from [DLN1] B		
LEFT OUTER JOIN [ODLN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode 
AND (B.LineStatus = ''O'' OR (B.LineStatus = ''C'' AND B.TargetType =''16'')) 
AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) As DLNSum,

(select COUNT (DISTINCT B.DOCENTRY) * -1 from [RDN1] B	
LEFT OUTER JOIN [ORDN] C ON B.DOCENTRY = C.DOCENTRY 
where A.ItemCode = B.ItemCode 
AND (B.LineStatus = ''O'' OR (B.LineStatus = ''C'' AND B.BaseType =''15'')) 
AND C.DocDate >= DATEADD(MONTH, -12, GETDATE())) As RDNSum
From [OITM] A 
LEFT OUTER JOIN [OITW] D ON A.ItemCode = D.ItemCode 
LEFT OUTER JOIN [@NCM_ITEMCAT] C ON ISNULL(A.U_ItemCat,'''') = ISNULL(C.Code,'''') 
')


