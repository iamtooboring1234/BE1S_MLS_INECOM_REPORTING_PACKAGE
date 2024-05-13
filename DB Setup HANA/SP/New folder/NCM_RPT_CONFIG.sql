DECLARE @SQLString NVARCHAR(1500)

/* Build the SQL string once.*/
SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM sys.objects WHERE OBJECT_ID = OBJECT_ID(N''@NCM_RPT_CONFIG'') AND type = (N''U''))
	CREATE TABLE [dbo].[@NCM_RPT_CONFIG] (
	[RptType]  [nvarchar](10)	NOT NULL,
	[RptCode]  [nvarchar](254)	NOT NULL PRIMARY KEY,
	[RptName]  [nvarchar](254)	NULL,
	[Included] [nvarchar](1)	NULL,
	[FilePath] [nvarchar](254)	NULL,
	[UpdDate]  DATETIME
)'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM sys.objects WHERE OBJECT_ID = OBJECT_ID(N''@NCM_NEW_SETTING'') AND type = (N''U''))
	CREATE TABLE [dbo].[@NCM_NEW_SETTING] (
	[U_GSTCurr]  		[nvarchar](10)	NULL,
	[U_InvDetail]  		[nvarchar](1)	NULL,
	[U_TaxDate]  		[nvarchar](1)	NULL,
	[U_IRAInvDetail] 	[nvarchar](1)	NULL,
	[U_IRATaxDate] 		[nvarchar](1)	NULL
)'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_NEW_SETTING] WHERE U_InvDetail IN (''N'',''Y''))
	insert into [@NCM_NEW_SETTING] values (''SGD'',''N'',''D'',''N'',''D'')'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''OFFICIAL_RECEIPT'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''OFFICIAL_RECEIPT'',''Official Receipt'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AP_AGEING_SUMMARY'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AP_AGEING_SUMMARY'',''AP Ageing Summary'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString	

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AP_AGEING_DETAILS'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AP_AGEING_DETAILS'',''AP Ageing Details'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_AGEING_SUMMARY'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AR_AGEING_SUMMARY'',''AR Ageing Summary'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_AGEING_DETAILS'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AR_AGEING_DETAILS'',''AR Ageing Details'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_AGEING_DETAILS_CRM'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AR_AGEING_DETAILS_CRM'',''AR Ageing Details with CRM Notes'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_AGEING_6B_SUMMARY'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AR_AGEING_6B_SUMMARY'',''AR Ageing 6 Buckets Summary'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_AGEING_6B_DETAILS'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AR_AGEING_6B_DETAILS'',''AR Ageing 6 Buckets Details'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_AGEING_7B_SUMMARY'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AR_AGEING_7B_SUMMARY'',''AR Ageing 7 Buckets Summary'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_AGEING_7B_DETAILS'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AR_AGEING_7B_DETAILS'',''AR Ageing 7 Buckets Details'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''SAR_FIFO_DETAILS'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''SAR_FIFO_DETAILS'',''Stock Ageing FIFO Details'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''SAR_FIFO_SUMMARY'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''SAR_FIFO_SUMMARY'',''Stock Ageing FIFO Summary'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString


SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''SAR_FIFO_DETAILS'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''SAR_FIFO_DETAILS'',''Stock Ageing FIFO Details'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString


SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''SAR_MOVAVG_SUMMARY'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''SAR_MOVAVG_SUMMARY'',''Stock Ageing Moving Average Summary'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''SAR_MOVAVG_DETAILS'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''SAR_MOVAVG_DETAILS'',''Stock Ageing Moving Average Details'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''SAR_TM_V1'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''SAR_TM_V1'',''Stock Ageing Original'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''SAR_TM_V2'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''SAR_TM_V2'',''Stock Ageing Sort By Warehouse'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''SAR_TM_V3'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''SAR_TM_V3'',''Stock Ageing Sort By Item'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_PAYMENT'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AR_PAYMENT'',''AR Payment Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AP_PAYMENT'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AP_PAYMENT'',''AP Payment Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_SOA'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AR_SOA'',''AR Statement of Accounts'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString


SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_SOA_LANDSCAPE'')
	insert into [@NCM_RPT_CONFIG] values (''ADV'',''AR_SOA_LANDSCAPE'',''AR Statement of Accounts (Landscape)'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString


SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AP_SOA'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AP_SOA'',''AP Statement of Accounts'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_SOA_EMAIL'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AR_SOA_EMAIL'',''AR SOA Email'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_SOA_PROJECT'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''AR_SOA_PROJECT'',''AR SOA By Project'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''DRAFT_PAYMENT_VOUCHER'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''DRAFT_PAYMENT_VOUCHER'',''Draft Payment Voucher'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''GL_LISTING'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''GL_LISTING'',''GL Listing Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''PAYMENT_VOUCHER'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''PAYMENT_VOUCHER'',''Payment Voucher'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''PAYMENT_VOUCHER_RANGE'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''PAYMENT_VOUCHER_RANGE'',''Payment Voucher Range'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''REMITTANCE_ADVICE'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''REMITTANCE_ADVICE'',''Remittance Advice'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''RECEIPT_PAYMENT'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''RECEIPT_PAYMENT'',''Receipt and Payment Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''RPT_GPA'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''RPT_GPA'',''Gross Profit Analysis Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''RPT_GST'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''RPT_GST'',''GST Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''SAR_MOVAVG_AUDIT_ENQUIRY'')
	insert into [@NCM_RPT_CONFIG] values (''BSC'',''SAR_MOVAVG_AUDIT_ENQUIRY'',''Stock Audit Enquiry Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''U_SAR_MOVAVG_AUDIT_ENQUIRY'')
	UPDATE [@NCM_RPT_CONFIG] SET RptCode = ''SAR_MOVAVG_AUDIT_ENQUIRY'' WHERE RptCode = ''U_SAR_MOVAVG_AUDIT_ENQUIRY'''

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''RPT_IAR'')
	insert into [@NCM_RPT_CONFIG] values (''ADV'',''RPT_IAR'',''Items ABC Analysis Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''RPT_RLR'')
	insert into [@NCM_RPT_CONFIG] values (''ADV'',''RPT_RLR'',''Re-Order Level Recommendation Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''RPT_WAR'')
	insert into [@NCM_RPT_CONFIG] values (''ADV'',''RPT_WAR'',''Weighted Average Demand Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''MRP_RPT'')
	insert into [@NCM_RPT_CONFIG] values (''ADV'',''MRP_RPT'',''MRP Supply Demand Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AP_AGEING_PROJ_DETAILS'')
	insert into [@NCM_RPT_CONFIG] values (''ADV'',''AP_AGEING_PROJ_DETAILS'',''AP Ageing Details with Project Code'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_AGEING_PROJ_DETAILS'')
	insert into [@NCM_RPT_CONFIG] values (''ADV'',''AR_AGEING_PROJ_DETAILS'',''AR Ageing Details with Project Code'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AP_AGEING_PROJ_SUMMARY'')
	insert into [@NCM_RPT_CONFIG] values (''ADV'',''AP_AGEING_PROJ_SUMMARY'',''AP Ageing Summary with Project Code'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''AR_AGEING_PROJ_SUMMARY'')
	insert into [@NCM_RPT_CONFIG] values (''ADV'',''AR_AGEING_PROJ_SUMMARY'',''AR Ageing Summary with Project Code'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''PO_DETAILS_BY_CUSTOMER'')
	insert into [@NCM_RPT_CONFIG] values (''ADV'',''PO_DETAILS_BY_CUSTOMER'',''PO Details by Vendor Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString

SET @SQLString =
     N'IF NOT EXISTS(SELECT 1 FROM [@NCM_RPT_CONFIG] WHERE RptCode = ''SO_DETAILS_BY_CUSTOMER'')
	insert into [@NCM_RPT_CONFIG] values (''ADV'',''SO_DETAILS_BY_CUSTOMER'',''SO Details by Customer Report'',''N'','''',getdate())'

EXECUTE sp_executesql @SQLString
