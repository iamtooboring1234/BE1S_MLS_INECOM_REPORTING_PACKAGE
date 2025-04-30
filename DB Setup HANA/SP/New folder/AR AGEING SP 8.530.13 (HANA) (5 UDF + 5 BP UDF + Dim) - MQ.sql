CREATE PROCEDURE SP_AR_AGEING
-- STORED PROCEDURE : AR AGEING SP (HANA)TO_VARCHAR (TO_DATE('2009-12-31'), 'YYYY/MM/DD')

-- Dated 2024-04-24 Ticket 7679 - to include OcrCode and OcrCode2. This will be Dimension1 and Dimension2 in AR Ageing table.
-- Dated 2023-05-18 Customised for MEDQUEST - HANA Conversion
--					In SQL, the CR has a view to retrieve 3 UDF fields (U_MQ_PatientName, U_MQ_PatientID, U_MQ_SurgeryDate) from OINV, ORIN and ORCT.
--					In HANA, there are UDF U_IRPField1 to U_IRPField5 to be used in Ageing and GST.
--					For MEDQUEST, U_IRPField1 to U_IRPField3 will be used for (U_MQ_PatientName, U_MQ_PatientID, U_MQ_SurgeryDate).
--					Purpose is to use the generic AR Ageing CR.
-- Dated 2021-03-01 Enhancement to include 5 UDf fields in BP Master, U_BPField1 to U_BPField5 (alphanumeric 254)
-- Dated 2020-09-07 Enhancement to include 5 UDf fields in marketing documents, U_IRPField1, U_IRPField2, U_IRPField3, U_IRPField4, U_IRPField5 (alphanumeric 254)



-- Version 8.530.13
-- to include 5 UDf fields in BP Master
-- Version 8.530.12
-- Need to include Prefix field (BeginStr) from Document Series. Requested by Avi-Tech
-- Version 8.530.11
-- Need to include FatherType checking on all marketing documents.
-- Version 8.530.10
-- To include Payment Terms, Credit Limit of BP, Doc Series. Requested by Avi-Tech
-- Version 8.530.9
-- To take care of Payment on Account where BP has Father Code for Payment Consolidation. Changes is in section Incoming and Outgoing Payment.
-- Version 8.530.8
-- Add 3 new fields PaymentCur, PaymentAmt, PaymentDate to retrieve last payment information for BP. Requested by AEM
-- It does not include scenarios to show last payment information of each BP based on currencies.

-- Version 8.530.7
-- Add in the missing selection based on TransRowid in the 3rd part section for Manual reconciliation of ClosePaid and ClosePaidFC
-- In AR Invoices, change the Due Days calculation as T6.DueDate.
-- In AR Invoices, change the drawn amount from DP to be based on BaseNet and BaseVat from INV11. Due to exchange differences between DP and Invoice


-- Version 8.530.6
-- By SY dated 2016-Mar-15
-- Ticket 20160839 - JE cancelled with '5' reconciled type - Make changes to JE section

(	IN vUserName	NVARCHAR(50),
	IN vBPCodeFr	NVARCHAR(30),
	IN vBPCodeTo	NVARCHAR(30),
	IN vBPGpFr		NVARCHAR(30),
	IN vBPGpTo		NVARCHAR(30),
	IN vSlpNameFr	NVARCHAR(50),
	IN vSlpNameTo	NVARCHAR(50),
	IN vBPCode		NVARCHAR(30),
	IN vPeriodTo	DATE
) 
LANGUAGE SQLSCRIPT
AS
-- DECLARE SP VARIABLES --------------------------------------------
	vCurrSchema		VARCHAR(50);
	vIsCreateTable	INTEGER;
	vTempColCount	INTEGER;
	vTempTableCount	INTEGER;
	
	vFilSlpFr		NVARCHAR(1);
	vCurr			NVARCHAR(5);
	vDirectRate		NVARCHAR(1);

    vPrvAgeDay		DATE;
    vPrvAgeTime     TIME;
    vPrvAgeRun		NVARCHAR(50);	-- username parameter consists of datetime and username
	
	vOrigTrnsId		INTEGER;
	vRUTransId		INTEGER;
	vOrigObjTyp		INTEGER;
	vCardCode		NVARCHAR(20);
	vOrigTransCurr	NVARCHAR(3);
	vTrnsTtlAmt		DECIMAL(19,6);
	vTrnsTtlFc		DECIMAL(19,6);
	vClosePaid		DECIMAL(19,6);
	vClosePaidFc	DECIMAL(19,6);
	vRecordCount	INTEGER;
	vNCM_TransId	INTEGER;
	vNCM_CardCode	NVARCHAR(20);
	vNCM_TransCurr	NVARCHAR(3);
	vNCM_Instl_Id	INTEGER;
	vNCM_BaseTransId	INTEGER;
	vNCM_ClosePaid	DECIMAL(19,6);
	vNCM_ClosePaidFc	DECIMAL(19,6);
	vNCM_RecordCount	INTEGER;

BEGIN -- START PROCEDURE

-- CREATE TEMPORARY TABLES
	CREATE LOCAL TEMPORARY TABLE #NCM_AR_AGEING_CASE1	-- (CASE1 LOOP)
	(	OrigTrnsId		BIGINT,
		RUTransId		BIGINT,
		OrigObjTyp		INTEGER,
		CardCode		NVARCHAR(20),
		OrigTransCurr	NVARCHAR(3), 
		TrnsTtlAmt		DECIMAL(19,6),
		TrnsTtlFc		DECIMAL(19,6)	);

	CREATE LOCAL TEMPORARY TABLE #NCM_AR_AGEING_NCMCASE1	-- (@NCM_CASE1 LOOP)
	(	NCM_TransId		INTEGER,
		NCM_CardCode	NVARCHAR(20),
		NCM_TransCurr	NVARCHAR(3), 
		NCM_Instl_Id	INTEGER,
		NCM_BaseTransId	INTEGER,
		NCM_AmtApplied	DECIMAL(19,6),
		NCM_AmtAppliedFC	DECIMAL(19,6)
	);

	-- Username parameter consists of datetime and username
    vPrvAgeTime := ADD_SECONDS(SUBSTRING(vUSERNAME,1,24),-3600); 
    vPrvAgeDay := SUBSTRING(vUserName,1,10);
    vPrvAgeRun := vPrvAgeDay||' '||vPrvAgeTime;	
	
	DELETE FROM "@NCM_AR_AGEING"
	WHERE  TRIM(SUBSTRING("USERNAME",25,26)) = TRIM(SUBSTRING(vUserName,25,26))
	AND     SUBSTRING("USERNAME",1,19) < vPrvAgeRun;

-- (START) DELETE OLD RECORDS for vUserName

-- (START) ASSIGNMENT TO PARAMETER RANGE
	SELECT CASE WHEN LENGTH(IFNULL(vBPCodeFr,'')) = 0 THEN MIN("CardCode") ELSE vBPCodeFr END,
		   CASE WHEN LENGTH(IFNULL(vBPCodeTo,'')) = 0 THEN MAX("CardCode") ELSE vBPCodeTo END,
		   CASE WHEN LENGTH(IFNULL(vBPGpFr,'')) = 0   THEN TO_ALPHANUM(MIN(IFNULL("GroupCode",0))) ELSE vBPGpFr END,
		   CASE WHEN LENGTH(IFNULL(vBPGpTo,'')) = 0   THEN TO_ALPHANUM(MAX(IFNULL("GroupCode",0))) ELSE vBPGpTo END,
		   CASE WHEN LENGTH(IFNULL(vBPCode,'')) = 0   THEN '%' ELSE '%' || vBPCode || '%' END
	INTO vBPCodeFr, vBPCodeTo, vBPGpFr, vBPGpTo, vBPCode
	FROM "OCRD"
	WHERE "CardType"='C';	

	SELECT CASE WHEN LENGTH(IFNULL(vSlpNameFr,'')) = 0 THEN MIN("SlpName") ELSE vSlpNameFr END,
		   CASE WHEN LENGTH(IFNULL(vSlpNameTo,'')) = 0 THEN MAX("SlpName") ELSE vSlpNameTo END
	INTO vSlpNameFr, vSlpNameTo
	FROM "OSLP";
	
-- (END) ASSIGNMENT TO PARAMETER RANGE

-- (START) DEFINE GENERAL VARIABLES
	SELECT TOP 1 "MainCurncy", "DirectRate"
	INTO vCurr, vDirectRate
	FROM "OADM";
	
	SELECT TOP 1 IFNULL("U_FilSlpFr",'D') 
	INTO vFilSlpFr
	FROM "@NCM_SETTING";
	
	IF vFilSlpFr = '' THEN
		vFilSlpFr := 'D';
	END	IF;
	
-- (END) DEFINE GENERAL VARIABLES

-- (START) INSERT AR INVOICES
	INSERT INTO "@NCM_AR_AGEING"
	SELECT DISTINCT :vUserName, 
				'01 AR Invoices', 
/*CardCode*/	(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END),
/*ChildCC*/		T0."CardCode", 
				T0."DocNum", T0."TransId", 
				T6."InstlmntID",
				(CASE WHEN T0."Installmnt">1 THEN TO_ALPHANUM(T6."InstlmntID")||'/'||TO_ALPHANUM(T0."Installmnt")
						ELSE '' END),
				T0."DocDate", T0."TaxDate", T6."DueDate",
				T0."NumAtCard", T0."Ref1", T0."Ref2", J."Project", T0."SlpCode", 
				T0."DocCur", 
/*DocRate*/		(CASE WHEN T0."DocCur"=vCurr THEN 1 ELSE T0."DocRate" END),
/*DocTotal*/	(CASE WHEN T6."InstlmntID"=1 THEN 
						T6."InsTotal"+IFNULL((SELECT SUM(T9."BaseNet"+T9."BaseVat") FROM "INV11" T9
											WHERE T9."DocEntry"=T0."DocEntry"
											AND T9."BaseAbs" IN (SELECT DP."DocEntry" FROM "ODPI" DP WHERE DP."TransId" IS NULL)),0)
				ELSE T6."InsTotal" END),
/*DocTotalFC*/(CASE WHEN T0."DocCur"=vCurr THEN
					(CASE WHEN T6."InstlmntID"=1 THEN 
							T6."InsTotal"+IFNULL((SELECT SUM(T9."BaseNet"+T9."BaseVat") FROM "INV11" T9
												WHERE T9."DocEntry"=T0."DocEntry"
												AND T9."BaseAbs" IN (SELECT DP."DocEntry" FROM "ODPI" DP WHERE DP."TransId" IS NULL)),0)
					ELSE T6."InsTotal" END)
				ELSE
					(CASE WHEN T6."InstlmntID"=1 THEN 
							T6."InsTotalFC"+IFNULL((SELECT SUM(T9."BaseNetFc"+T9."BaseVatFc") FROM "INV11" T9
													WHERE T9."DocEntry"=T0."DocEntry"
													AND T9."BaseAbs" IN (SELECT DP."DocEntry" FROM "ODPI" DP WHERE DP."TransId" IS NULL)),0)
					ELSE T6."InsTotalFC" END)
			END),
/*DocType*/'IN',
/*ClosePaid*/
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='N'
					AND RC."ReconType" IN (0,1,3,4)
					AND RC."ReconDate"<=vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
							AND RC."CancelAbs"=CC."CancelAbs"	
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (3,4)
					AND C1."ReconType"=5
					AND C1."ReconDate">vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
							AND RC."SrcObjAbs"=CC."SrcObjAbs"
							AND RC."SrcObjTyp"=CC."SrcObjTyp"
							AND RC."TransId"=CC."TransId"
							AND	RC."TransRowId" = CC."TransRowId"
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (0,1)
					AND CC."ReconType"=7
					AND CC."ReconDate">vPeriodTo),0)),
/*ClosePaidFC*/
			(IFNULL((SELECT CASE WHEN T0."DocCur"=vCurr THEN
									SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
								ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) 
								END
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
				AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
				AND RC."Canceled"='N'
				AND RC."ReconType" IN (0,1,3,4)
				AND RC."ReconDate"<=vPeriodTo),0))
			+
			(IFNULL((SELECT CASE WHEN T0."DocCur"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
							AND RC."CancelAbs"=CC."CancelAbs"	
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (3,4)
					AND C1."ReconType"=5
					AND C1."ReconDate">vPeriodTo),0))
			+
			(IFNULL((SELECT CASE WHEN T0."DocCur"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
							AND RC."SrcObjAbs"=CC."SrcObjAbs"
							AND RC."SrcObjTyp"=CC."SrcObjTyp"
							AND RC."TransId"=CC."TransId"
							AND	RC."TransRowId" = CC."TransRowId"
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (0,1)
					AND CC."ReconType"=7
					AND CC."ReconDate">vPeriodTo),0)),
/*OpenAmt*/0,
/*OpenAmtFC*/0,
		J."IntrnMatch",
/*PosAgeDays*/DAYS_BETWEEN(T0."DocDate",vPeriodTo),
/*PosAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DocDate")*12)+MONTH(T0."DocDate")),
/*DocAgeDays*/DAYS_BETWEEN(T0."TaxDate",vPeriodTo),
/*DocAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."TaxDate")*12)+MONTH(T0."TaxDate")),
/*DueAgeDays*/DAYS_BETWEEN(T6."DueDate",vPeriodTo),
/*DueAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T6."DueDate")*12)+MONTH(T6."DueDate")),
		T0."DocStatus", T0."CANCELED",
		'', 0, '',
		PT."PymntGroup", BP."CreditLine", 
		T0."Series", IFNULL(N1."SeriesName",''), IFNULL(N1."Remark",'') ,
		IFNULL(T0."U_MQ_PatientID",''),
		IFNULL(T0."U_MQ_PatientName",''),
		IFNULL(TO_VARCHAR(T0."U_MQ_SurgeryDate",'DD/MM/YYYY'),''),
		IFNULL(T0."U_IRPField4",''),
		IFNULL(T0."U_IRPField5",''),
		IFNULL(BP."U_IRPBPField1",''),
		IFNULL(BP."U_IRPBPField2",''),
		IFNULL(BP."U_IRPBPField3",''),
		IFNULL(BP."U_IRPBPField4",''),
		IFNULL(BP."U_IRPBPField5",''),
		IFNULL(DIM1."OcrCode",''),
		IFNULL(DIM2."OcrCode2",'')
	FROM "OINV" T0
		INNER JOIN "OCRD" BP ON (CASE WHEN IFNULL(T0."FatherCard",'')='' THEN T0."CardCode" 
									ELSE T0."FatherCard" END)=BP."CardCode" 
			AND BP."CardType"='C'
		INNER JOIN "JDT1" J ON T0."TransId"=J."TransId"
			AND (CASE WHEN IFNULL(T0."FatherCard",'')='' THEN T0."CardCode" 
					ELSE T0."FatherCard" END)=J."ShortName"
		INNER JOIN "INV6" T6 ON T0."DocEntry"=T6."DocEntry" 
			AND T6."InstlmntID"=IFNULL(J."SourceLine",1)
		LEFT OUTER JOIN "NNM1" N1 ON T0."Series"=N1."Series"
		INNER JOIN "OCTG" PT ON T0."GroupNum"= PT."GroupNum"
		-- Dimension 1
		LEFT JOIN (SELECT Z0."DocEntry", STRING_AGG(Z0."OcrCode",',') as "OcrCode"
					FROM (SELECT DISTINCT "DocEntry","OcrCode"
							FROM INV1
							WHERE IFNULL("OcrCode" ,'') <> ''
							-- ORDER BY "LineNum"
							) Z0
					GROUP BY Z0."DocEntry"
					) DIM1 ON T0."DocEntry" = DIM1."DocEntry"
		-- Dimension 2
		LEFT JOIN (SELECT Z0."DocEntry", STRING_AGG(Z0."OcrCode2",',') as "OcrCode2"
					FROM (SELECT DISTINCT "DocEntry","OcrCode2"
							FROM INV1
							WHERE IFNULL("OcrCode2" ,'') <> ''
							-- ORDER BY "LineNum"
							) Z0
					GROUP BY Z0."DocEntry"
					) DIM2 ON T0."DocEntry" = DIM2."DocEntry"

	WHERE T0."DocDate"<=vPeriodTo
	AND (CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END) LIKE vBPCode
	AND (CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END) BETWEEN vBPCodeFr AND vBPCodeTo
	AND	BP."GroupCode" BETWEEN vBPGpFr AND vBPGpTo
	AND	((T0."SlpCode" IN (SELECT SLP."SlpCode" FROM "OSLP" SLP WHERE SLP."SlpName" BETWEEN vSlpNameFr AND vSlpNameTo) AND vFilSlpFr='D') 
			OR (BP."SlpCode" IN (SELECT SLP."SlpCode" FROM "OSLP" SLP WHERE SLP."SlpName" BETWEEN vSlpNameFr AND vSlpNameTo) AND vFilSlpFr='B'));
-- (END) INSERT AR INVOICES

-- (START) INSERT AR DOWNPAYMENT INVOICES
	INSERT INTO "@NCM_AR_AGEING"
	SELECT DISTINCT :vUserName, 
				'02 AR DP Invoices', 
/*CardCode*/	(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END),
/*ChildCC*/		T0."CardCode", 
				T0."DocNum", T0."TransId",
				1,
				'',
				T0."DocDate", T0."TaxDate", T0."DocDueDate",
				T0."NumAtCard", T0."Ref1", T0."Ref2", J."Project", T0."SlpCode", 
				T0."DocCur", 
/*DocRate*/		(CASE WHEN T0."DocCur"=vCurr THEN 1 ELSE T0."DocRate" END),
/*DocTotal*/	T0."DocTotal",
/*DocTotalFC*/	(CASE WHEN T0."DocCur"=vCurr THEN T0."DocTotal" ELSE T0."DocTotalFC" END),
/*DocType*/	'DP',
/*ClosePaid*/
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='N'
					AND RC."ReconType" IN (0,1,3,4)
					AND RC."ReconDate"<=vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
							AND RC."CancelAbs"=CC."CancelAbs"	
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (3,4)
					AND C1."ReconType"=5
					AND C1."ReconDate">vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
							AND RC."SrcObjAbs"=CC."SrcObjAbs"
							AND RC."SrcObjTyp"=CC."SrcObjTyp"
							AND RC."TransId"=CC."TransId"
							AND	RC."TransRowId" = CC."TransRowId"
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (0,1)
					AND CC."ReconType"=7
					AND CC."ReconDate">vPeriodTo),0)),
/*ClosePaidFC*/
			(IFNULL((SELECT CASE WHEN T0."DocCur"=vCurr THEN
						SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) END
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='N'
					AND RC."ReconType" IN (0,1,3,4)
					AND RC."ReconDate"<=vPeriodTo),0))
			+
			(IFNULL((SELECT CASE WHEN T0."DocCur"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
							AND RC."CancelAbs"=CC."CancelAbs"	
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (3,4)
					AND	C1."ReconType"=5
					AND C1."ReconDate">vPeriodTo),0))
			+
			(IFNULL((SELECT CASE WHEN T0."DocCur"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
							AND RC."SrcObjAbs"=CC."SrcObjAbs"
							AND RC."SrcObjTyp"=CC."SrcObjTyp"
							AND RC."TransId"=CC."TransId"
							AND	RC."TransRowId" = CC."TransRowId"
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (0,1)
					AND CC."ReconType"=7
					AND CC."ReconDate">vPeriodTo),0)),
/*OpenAmt*/0,
/*OpenAmtFC*/0,
		J."IntrnMatch",
/*PosAgeDays*/DAYS_BETWEEN(T0."DocDate",vPeriodTo),
/*PosAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DocDate")*12)+MONTH(T0."DocDate")),
/*DocAgeDays*/DAYS_BETWEEN(T0."TaxDate",vPeriodTo),
/*DocAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."TaxDate")*12)+MONTH(T0."TaxDate")),
/*DueAgeDays*/DAYS_BETWEEN(T0."DocDueDate",vPeriodTo),
/*DueAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DocDueDate")*12)+MONTH(T0."DocDueDate")),
		T0."DocStatus", T0."CANCELED",
		'', 0, '',
		PT."PymntGroup", BP."CreditLine", 
		T0."Series", IFNULL(N1."SeriesName",''), IFNULL(N1."Remark",'') ,
		IFNULL(T0."U_MQ_PatientID",''),
		IFNULL(T0."U_MQ_PatientName",''),
		IFNULL(TO_VARCHAR(T0."U_MQ_SurgeryDate",'DD/MM/YYYY'),''),
		IFNULL(T0."U_IRPField4",''),
		IFNULL(T0."U_IRPField5",''),
		IFNULL(BP."U_IRPBPField1",''),
		IFNULL(BP."U_IRPBPField2",''),
		IFNULL(BP."U_IRPBPField3",''),
		IFNULL(BP."U_IRPBPField4",''),
		IFNULL(BP."U_IRPBPField5",''),
		IFNULL(DIM1."OcrCode",''),
		IFNULL(DIM2."OcrCode2",'')

	FROM "ODPI" T0
		INNER JOIN "OCRD" BP ON (CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)=BP."CardCode" 
			AND BP."CardType"='C'
		INNER JOIN "JDT1" J ON T0."TransId"=J."TransId"
			AND (CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)=J."ShortName"	
		LEFT OUTER JOIN "NNM1" N1 ON T0."Series"=N1."Series"
		INNER JOIN "OCTG" PT ON T0."GroupNum"= PT."GroupNum"
		-- Dimension 1
		LEFT JOIN (SELECT Z0."DocEntry", STRING_AGG(Z0."OcrCode",',') as "OcrCode"
					FROM (SELECT DISTINCT "DocEntry","OcrCode"
							FROM DPI1
							WHERE IFNULL("OcrCode" ,'') <> ''
							-- ORDER BY "LineNum"
							) Z0
					GROUP BY Z0."DocEntry"
					) DIM1 ON T0."DocEntry" = DIM1."DocEntry"
		-- Dimension 2
		LEFT JOIN (SELECT Z0."DocEntry", STRING_AGG(Z0."OcrCode2",',') as "OcrCode2"
					FROM (SELECT DISTINCT "DocEntry","OcrCode2"
							FROM DPI1
							WHERE IFNULL("OcrCode2" ,'') <> ''
							-- ORDER BY "LineNum"
							) Z0
					GROUP BY Z0."DocEntry"
					) DIM2 ON T0."DocEntry" = DIM2."DocEntry"

	WHERE T0."DocDate"<=vPeriodTo
	AND (CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END) LIKE vBPCode
	AND (CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END) BETWEEN vBPCodeFr AND vBPCodeTo
	AND	BP."GroupCode" BETWEEN vBPGpFr AND vBPGpTo
	AND	((T0."SlpCode" IN (SELECT SLP."SlpCode" FROM "OSLP" SLP WHERE SLP."SlpName" BETWEEN vSlpNameFr AND vSlpNameTo) AND vFilSlpFr='D') 
		OR (BP."SlpCode" IN (SELECT SLP."SlpCode" FROM "OSLP" SLP WHERE SLP."SlpName" BETWEEN vSlpNameFr AND vSlpNameTo) AND vFilSlpFr='B'));
-- (END) INSERT AR DOWNPAYMENT INVOICES

-- (START) INSERT AR CREDIT NOTE
	INSERT INTO "@NCM_AR_AGEING"
	SELECT DISTINCT :vUserName, 
				'03 AR Credit Note', 
/*CardCode*/	(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END),
/*ChildCC*/		T0."CardCode", 
				T0."DocNum", T0."TransId",
				1,
				'',
				T0."DocDate", T0."TaxDate", T0."DocDueDate",
				T0."NumAtCard", T0."Ref1", T0."Ref2", J."Project", T0."SlpCode", 
				T0."DocCur", 
/*DocRate*/		(CASE WHEN T0."DocCur"=vCurr THEN 1 ELSE T0."DocRate" END),
/*DocTotal*/	(T0."DocTotal"+IFNULL((SELECT SUM(T9."DrawnSum"+T9."Vat") FROM "RIN9" T9
										WHERE T9."DocEntry"=T0."DocEntry"
										AND T9."BaseAbs" IN (SELECT DP."DocEntry" FROM "ODPI" DP WHERE DP."TransId" IS NULL)),0))*-1,
/*DocTotalFC*/	(CASE WHEN T0."DocCur"=vCurr THEN
					(T0."DocTotal"+IFNULL((SELECT SUM(T9."DrawnSum"+T9."Vat") FROM "RIN9" T9
											WHERE T9."DocEntry"=T0."DocEntry"
											AND T9."BaseAbs" IN (SELECT DP."DocEntry" FROM "ODPI" DP WHERE DP."TransId" IS NULL)),0))*-1
				ELSE
					(T0."DocTotalFC"+IFNULL((SELECT SUM(T9."DrawnSumFc"+T9."VatFc") FROM "RIN9" T9
											WHERE T9."DocEntry"=T0."DocEntry"
											AND T9."BaseAbs" IN (SELECT DP."DocEntry" FROM "ODPI" DP WHERE DP."TransId" IS NULL)),0))*-1
			END),
/*DocType*/'CN',
/*ClosePaid*/((IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
			FROM "NCM_RECON_DET" RC
			WHERE RC."SrcObjAbs"=T0."DocEntry"
			AND RC."TransId"=T0."TransId"
			AND RC."TransRowId"=J."Line_ID"
			AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
			AND RC."Canceled"='N'
			AND RC."ReconType" IN (0,1,3,4)
			AND RC."ReconDate"<=vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
				FROM "NCM_RECON_DET" RC
					INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
						AND RC."InitObjAbs"=CC."SrcObjAbs"
						AND RC."InitObjTyp"=CC."SrcObjTyp"
						AND RC."CancelAbs"=CC."CancelAbs"	-- v7.4.2
					INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
						AND CC."SrcObjTyp"=C1."SrcObjTyp"
						AND CC."TransId"=C1."TransId"
						AND CC."TransRowId"=C1."TransRowId"
				WHERE RC."SrcObjAbs"=T0."DocEntry"
				AND RC."TransId"=T0."TransId"
				AND RC."TransRowId"=J."Line_ID"
				AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
				AND RC."Canceled"='Y'
				AND RC."ReconDate"<=vPeriodTo
				AND RC."ReconType" IN (3,4)
				AND C1."ReconType"=5
				AND C1."ReconDate">vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
				FROM "NCM_RECON_DET" RC
					INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
						AND RC."SrcObjAbs"=CC."SrcObjAbs"
						AND RC."SrcObjTyp"=CC."SrcObjTyp"
						AND RC."TransId"=CC."TransId"
						AND	RC."TransRowId" = CC."TransRowId"
				WHERE RC."SrcObjAbs"=T0."DocEntry"
				AND RC."TransId"=T0."TransId"
				AND RC."TransRowId"=J."Line_ID"
				AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
				AND RC."Canceled"='Y'
				AND RC."ReconDate"<=vPeriodTo
				AND RC."ReconType" IN (0,1)
				AND CC."ReconType"=7
				AND CC."ReconDate">vPeriodTo),0))),
/*ClosePaidFC*/
			((IFNULL((SELECT CASE WHEN T0."DocCur"=vCurr THEN
							SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
							ELSE SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)*-1 END
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='N'
					AND RC."ReconType" IN (0,1,3,4)
					AND RC."ReconDate"<=vPeriodTo),0))
			+
			(IFNULL((SELECT CASE WHEN T0."DocCur"=vCurr THEN
							SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
						ELSE SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)*-1 END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
							AND RC."CancelAbs"=CC."CancelAbs"	-- v7.4.2
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (3,4)
					AND C1."ReconType"=5
					AND C1."ReconDate">vPeriodTo),0))
			+
			(IFNULL((SELECT CASE WHEN T0."DocCur"=vCurr THEN
							SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
							ELSE SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)*-1 END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
							AND RC."SrcObjAbs"=CC."SrcObjAbs"
							AND RC."SrcObjTyp"=CC."SrcObjTyp"
							AND RC."TransId"=CC."TransId"
							AND	RC."TransRowId" = CC."TransRowId"
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=(CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (0,1)
					AND CC."ReconType"=7
					AND CC."ReconDate">vPeriodTo),0))),
/*OpenAmt*/0,
/*OpenAmtFC*/0,
		J."IntrnMatch",
/*PosAgeDays*/DAYS_BETWEEN(T0."DocDate",vPeriodTo),
/*PosAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DocDate")*12)+MONTH(T0."DocDate")),
/*DocAgeDays*/DAYS_BETWEEN(T0."TaxDate",vPeriodTo),
/*DocAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."TaxDate")*12)+MONTH(T0."TaxDate")),
/*DueAgeDays*/DAYS_BETWEEN(T0."DocDueDate",vPeriodTo),
/*DueAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DocDueDate")*12)+MONTH(T0."DocDueDate")),
		T0."DocStatus", T0."CANCELED",
		'', 0, '',
		PT."PymntGroup", BP."CreditLine", 
		T0."Series", IFNULL(N1."SeriesName",''), IFNULL(N1."Remark",'') ,
		IFNULL(T0."U_MQ_PatientID",''),
		IFNULL(T0."U_MQ_PatientName",''),
		IFNULL(TO_VARCHAR(T0."U_MQ_SurgeryDate",'DD/MM/YYYY'),''),
		IFNULL(T0."U_IRPField4",''),
		IFNULL(T0."U_IRPField5",''),
		IFNULL(BP."U_IRPBPField1",''),
		IFNULL(BP."U_IRPBPField2",''),
		IFNULL(BP."U_IRPBPField3",''),
		IFNULL(BP."U_IRPBPField4",''),
		IFNULL(BP."U_IRPBPField5",''),
		IFNULL(DIM1."OcrCode",''),
		IFNULL(DIM2."OcrCode2",'')

	FROM "ORIN" T0
		INNER JOIN "OCRD" BP ON (CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END)=BP."CardCode" 
			AND BP."CardType"='C'
		INNER JOIN "JDT1" J ON T0."TransId"=J."TransId"	AND (CASE WHEN IFNULL(T0."FatherCard",'')='' THEN T0."CardCode" 
				ELSE T0."FatherCard" END)=J."ShortName"				
		LEFT OUTER JOIN "NNM1" N1 ON T0."Series"=N1."Series"
		INNER JOIN "OCTG" PT ON T0."GroupNum"= PT."GroupNum"
		-- Dimension 1
		LEFT JOIN (SELECT Z0."DocEntry", STRING_AGG(Z0."OcrCode",',') as "OcrCode"
					FROM (SELECT DISTINCT "DocEntry","OcrCode"
							FROM RIN1
							WHERE IFNULL("OcrCode" ,'') <> ''
							-- ORDER BY "LineNum"
							) Z0
					GROUP BY Z0."DocEntry"
					) DIM1 ON T0."DocEntry" = DIM1."DocEntry"
		-- Dimension 2
		LEFT JOIN (SELECT Z0."DocEntry", STRING_AGG(Z0."OcrCode2",',') as "OcrCode2"
					FROM (SELECT DISTINCT "DocEntry","OcrCode2"
							FROM RIN1
							WHERE IFNULL("OcrCode2" ,'') <> ''
							-- ORDER BY "LineNum"
							) Z0
					GROUP BY Z0."DocEntry"
					) DIM2 ON T0."DocEntry" = DIM2."DocEntry"

	WHERE T0."DocDate"<=vPeriodTo
	AND (CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END) LIKE vBPCode
	AND (CASE WHEN T0."FatherType" = 'D' THEN T0."CardCode" ELSE (CASE WHEN IFNULL(T0."FatherCard",'') = '' THEN T0."CardCode" ELSE T0."FatherCard" END) END) BETWEEN vBPCodeFr AND vBPCodeTo
	AND	BP."GroupCode" BETWEEN vBPGpFr AND vBPGpTo
	AND	((T0."SlpCode" IN (SELECT SLP."SlpCode" FROM "OSLP" SLP WHERE SLP."SlpName" BETWEEN vSlpNameFr AND vSlpNameTo) AND vFilSlpFr='D') 
		OR (BP."SlpCode" IN (SELECT SLP."SlpCode" FROM "OSLP" SLP WHERE SLP."SlpName" BETWEEN vSlpNameFr AND vSlpNameTo) AND vFilSlpFr='B'));
-- (END) INSERT AR CREDIT NOTE

-- (START) INSERT JOURNAL ENTRIES
	INSERT INTO "@NCM_AR_AGEING"
	SELECT :vUserName,
			'04 Journal Entries', 
/*CardCode*/T2."CardCode",
			T2."CardCode", 
			T0."TransId", T0."TransId",
			1,
			'',
			T0."RefDate", T0."TaxDate", T0."DueDate",
			T0."TransId", T0."Ref1", T0."Ref2", T1."Project", T2."SlpCode", 
/*DocCurr*/(CASE IFNULL(T1."FCCurrency",vCurr) 
				WHEN '' THEN vCurr
				WHEN vCurr THEN vCurr
				ELSE T1."FCCurrency" END),
/*DocRate*/(CASE IFNULL(T1."FCCurrency",vCurr) 
				WHEN '' THEN 1
				WHEN vCurr THEN 1
				ELSE (CASE WHEN vDIRECTRATE='Y' THEN 
							(T1."Debit"-T1."Credit")/
								(CASE WHEN (T1."FCDebit"-T1."FCCredit")=0 THEN 1 
									ELSE (T1."FCDebit"-T1."FCCredit") END)
						ELSE
							(T1."FCDebit"-T1."FCCredit")/
								(CASE WHEN (T1."Debit"-T1."Credit")=0 THEN 1 
									ELSE (T1."Debit"-T1."Credit") END)
						END) END),
/*DocTotal*/T1."Debit"-T1."Credit", 
/*DocTotalFC*/(CASE IFNULL(T1."FCCurrency",(CASE WHEN T2."Currency" IN (vCurr, '##') THEN vCurr ELSE NULL END))
					WHEN vCurr THEN T1."Debit"-T1."Credit" 
					WHEN '' THEN (T1."FCDebit"-T1."FCCredit")
					WHEN NULL THEN (T1."FCDebit"-T1."FCCredit")
					ELSE (T1."FCDebit"-T1."FCCredit") END),
/*DocType*/'JE',
/*ClosePaid*/
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=T0."TransId"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=T1."Line_ID"
					AND RC."ShortName"=T1."ShortName"
					AND RC."Canceled"='N'
					AND RC."ReconType" IN (0,1,3,4,5)
					AND RC."ReconDate"<=vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
							AND RC."CancelAbs"=CC."CancelAbs"	
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
					WHERE RC."SrcObjAbs"=T0."TransId"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=T1."Line_ID"
					AND RC."ShortName"=T1."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (3,4)
					AND C1."ReconType"=5
					AND C1."ReconDate">vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
							AND RC."SrcObjAbs"=CC."SrcObjAbs"
							AND RC."SrcObjTyp"=CC."SrcObjTyp"
							AND RC."TransId"=CC."TransId"
							AND RC."TransRowId"=CC."TransRowId"
					WHERE RC."SrcObjAbs"=T0."TransId"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=T1."Line_ID"
					AND RC."ShortName"=T1."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (0,1)
					AND CC."ReconType"=7
					AND CC."ReconDate">vPeriodTo),0)),
/*ClosePaidFC*/
			(IFNULL((SELECT CASE WHEN IFNULL(T1."FCCurrency",(CASE WHEN T2."Currency" IN (vCurr,'##') THEN vCurr ELSE NULL END))=vCurr THEN 
									SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
								WHEN LENGTH(T1."FCCurrency")=0 THEN 
									SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
								ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) 
							END
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=T0."TransId"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=T1."Line_ID"
					AND RC."ShortName"=T1."ShortName"
					AND RC."Canceled"='N'
					AND RC."ReconType" IN (0,1,3,4,5)
					AND RC."ReconDate"<=vPeriodTo),0))
			+
			IFNULL((SELECT CASE WHEN IFNULL(T1."FCCurrency",(CASE WHEN T2."Currency" IN (vCurr,'##') THEN vCurr ELSE NULL END))=vCurr THEN 
									SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
								WHEN LENGTH(T1."FCCurrency")=0 THEN 
									SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
								ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) 
					END 
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
							AND RC."CancelAbs"=CC."CancelAbs"	
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
					WHERE RC."SrcObjAbs"=T0."TransId"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=T1."Line_ID"
					AND RC."ShortName"=T1."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (3,4)
					AND C1."ReconType"=5
					AND C1."ReconDate">vPeriodTo),0)
			+
			IFNULL((SELECT CASE WHEN IFNULL(T1."FCCurrency",CASE WHEN T2."Currency" IN (vCurr,'##') THEN vCurr ELSE NULL END)=vCurr THEN 
									SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
								WHEN LENGTH(T1."FCCurrency")=0 THEN 
									SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
								ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) 
							END
				FROM "NCM_RECON_DET" RC
					INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
						AND RC."SrcObjAbs"=CC."SrcObjAbs"
						AND RC."SrcObjTyp"=CC."SrcObjTyp"
						AND RC."TransId"=CC."TransId"
						AND RC."TransRowId"=CC."TransRowId"
				WHERE RC."SrcObjAbs"=T0."TransId"
				AND RC."TransId"=T0."TransId"
				AND RC."TransRowId"=T1."Line_ID"
				AND RC."ShortName"=T1."ShortName"
				AND RC."Canceled"='Y'
				AND RC."ReconDate"<=vPeriodTo
				AND RC."ReconType" IN (0,1)
				AND CC."ReconType"=7
				AND CC."ReconDate">vPeriodTo),0),
/*OpenAmt*/0,
/*OpenAmtFC*/0,
		T1."IntrnMatch",
/*PosAgeDays*/DAYS_BETWEEN(T0."RefDate",vPeriodTo),
/*PosAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."RefDate")*12)+MONTH(T0."RefDate")),
/*DocAgeDays*/DAYS_BETWEEN(T0."TaxDate",vPeriodTo),
/*DocAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."TaxDate")*12)+MONTH(T0."TaxDate")),
/*DueAgeDays*/DAYS_BETWEEN(T0."DueDate",vPeriodTo),
/*DueAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DueDate")*12)+MONTH(T0."DueDate")),
/*DocStatus*/'',
/*Cancelled*/'',
		'', 0, '',
		PT."PymntGroup", T2."CreditLine", 
		T0."Series", IFNULL(N1."SeriesName",''), IFNULL(N1."Remark",'') ,
		'', '', '', '', '',
		IFNULL(T2."U_IRPBPField1",''),
		IFNULL(T2."U_IRPBPField2",''),
		IFNULL(T2."U_IRPBPField3",''),
		IFNULL(T2."U_IRPBPField4",''),
		IFNULL(T2."U_IRPBPField5",''),
		IFNULL(DIM1."OcrCode",''),
		IFNULL(DIM2."OcrCode2",'')

	FROM OJDT T0
		INNER JOIN JDT1 T1 ON T0."TransId"=T1."TransId"
		INNER JOIN OCRD T2 ON T1."ShortName"=T2."CardCode" AND T2."CardType"='C'
		LEFT OUTER JOIN "NNM1" N1 ON T0."Series"=N1."Series"
		INNER JOIN "OCTG" PT ON T2."GroupNum"= PT."GroupNum"
		-- Dimension 1
		LEFT JOIN (SELECT Z0."TransId", STRING_AGG(Z0."ProfitCode",',') as "OcrCode"
					FROM (SELECT DISTINCT "TransId","ProfitCode"
							FROM JDT1
							WHERE IFNULL("ProfitCode" ,'') <> ''
							--ORDER BY "Line_ID"
							) Z0
					GROUP BY Z0."TransId"
					) DIM1 ON T0."TransId" = DIM1."TransId"
		-- Dimension 2
		LEFT JOIN (SELECT Z0."TransId", STRING_AGG(Z0."OcrCode2",',') as "OcrCode2"
					FROM (SELECT DISTINCT "TransId","OcrCode2"
							FROM JDT1
							WHERE IFNULL("OcrCode2" ,'') <> ''
							--ORDER BY "Line_ID"
							) Z0
					GROUP BY Z0."TransId"
					) DIM2 ON T0."TransId" = DIM2."TransId"

	WHERE T0."TransType"=30
	AND T0."RefDate"<=vPeriodTo
	AND T1."ShortName" LIKE vBPCode
	AND T1."ShortName" BETWEEN vBPCodeFr AND vBPCodeTo
	AND T2."GroupCode" BETWEEN vBPGpFr AND vBPGpTo
	AND T2."SlpCode" IN (SELECT T9."SlpCode" FROM "OSLP" T9 WHERE T9."SlpName" BETWEEN vSlpNameFr AND vSlpNameTo);
-- (END) INSERT JOURNAL ENTRIES

-- (START) INSERT INCOMING PAYMENTS (Caters to Transactions other than DownPayment Request Invoices)
	INSERT INTO "@NCM_AR_AGEING"
	SELECT DISTINCT :vUserName, 
			'05 Receipt', 
/*CardCode*/J."ShortName",
			T0."CardCode", T0."DocNum", T0."TransId",
			1,
			'',
			T0."DocDate", T0."TaxDate", T0."DocDueDate",
			T0."CounterRef", T0."Ref1", T0."Ref2", J."Project", BP."SlpCode", 
			T0."DocCurr", 
/*DocRate*/(CASE WHEN T0."DocCurr"=vCurr THEN 1 ELSE T0."DocRate" END),
/*DocTotal*/IFNULL((SELECT SUM(HY."Debit"-HY."Credit")
				FROM "JDT1" HY
				WHERE HY."TransId"=T0."TransId"
				AND HY."ShortName"=J."ShortName"
				AND HY."LineType"=0),0),
/*DocTotalFC*/(CASE WHEN T0."DocCurr"=vCurr THEN 
					IFNULL((SELECT SUM(HY."Debit"-HY."Credit") FROM "JDT1" HY
							WHERE HY."TransId"=T0."TransId"
							AND HY."ShortName"=J."ShortName"
							AND HY."LineType"=0),0)
				ELSE IFNULL((SELECT SUM(HY."FCDebit"-HY."FCCredit") FROM "JDT1" HY
							WHERE HY."TransId"=T0."TransId"
							AND HY."ShortName"=J."ShortName"
							AND HY."LineType"=0),0) END),
/*DocType*/'RC',
/*ClosePaid*/
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
					FROM "NCM_RECON_DET" RC
						INNER JOIN "JDT1" J1 ON RC."TransId"=J1."TransId" AND RC."ShortName"=J1."ShortName"
							AND RC."TransRowId"=J1."Line_ID"
					WHERE RC."TransId"=T0."TransId"
					AND RC."ShortName"=J."ShortName"
					AND RC."ReconType" NOT IN (7,5)
					AND RC."Canceled"='N'
					AND RC."ReconDate"<=vPeriodTo
					AND J1."LineType"=0),0))
			+
			IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum"
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
					INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
						AND CC."SrcObjTyp"=C1."SrcObjTyp"
						AND CC."TransId"=C1."TransId"
						AND CC."TransRowId"=C1."TransRowId"
						AND C1."ReconType"=5
					INNER JOIN "JDT1" J1 ON CC."TransId"=J1."TransId"
						AND CC."ShortName"=J1."ShortName"
						AND CC."TransRowId"=J1."Line_ID"
						AND J1."LineType"=0
				WHERE RC."SrcObjAbs"=T0."DocEntry"
				AND RC."TransId"=T0."TransId"
				AND RC."TransRowId"=J."Line_ID"
				AND RC."ShortName"=J."ShortName"
				AND RC."Canceled"='Y'
				AND RC."ReconDate"<=vPeriodTo
				AND RC."ReconType"=3
				AND C1."ReconDate">vPeriodTo),0)
			+
			IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "JDT1" J1 ON RC."TransId"= J1."TransId"
							AND RC."ShortName"= J1."ShortName"
							AND RC."TransRowId"= J1."Line_ID"
							AND J1."LineType"=0
					WHERE RC."TransId"= T0."TransId"
					AND RC."ShortName"= J."ShortName"
					AND RC."ReconType"= 7
					AND RC."ReconDate">vPeriodTo
					AND RC."CancelAbs" IN (SELECT DISTINCT CC."ReconNum" FROM "NCM_RECON_DET" CC
											WHERE	CC."ShortName"=J."ShortName"
											AND		CC."TransId"=T0."TransId"
											AND		CC."ReconType"=0
											AND		CC."ReconDate"<=vPeriodTo)),0)
			+
			IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
					FROM "NCM_RECON_DET" RC
						INNER JOIN "JDT1" J1 ON RC."TransId"=J1."TransId"
							AND RC."ShortName"=J1."ShortName"
							AND RC."TransRowId"=J1."Line_ID"
							AND J1."LineType"=0
					WHERE RC."TransId"=T0."TransId"
					AND RC."ShortName"=J."ShortName"
					AND RC."ReconType"=5
					AND RC."ReconDate"<=vPeriodTo),0),
/*ClosePaidFC*/
			(IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
							ELSE SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)*-1
							END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "JDT1" J1 ON RC."TransId"=J1."TransId" AND RC."ShortName"=J1."ShortName"
							AND RC."TransRowId"=J1."Line_ID"
					WHERE RC."TransId"=T0."TransId"
					AND RC."ShortName"=J."ShortName"
					AND RC."ReconType" NOT IN (7,5)
					AND RC."Canceled"='N'
					AND RC."ReconDate"<=vPeriodTo
					AND J1."LineType"=0),0))
			+
			IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
							ELSE SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)*-1
							END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum"
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
							AND C1."ReconType"=5
						INNER JOIN "JDT1" J1 ON CC."TransId"=J1."TransId"
							AND CC."ShortName"=J1."ShortName"
							AND CC."TransRowId"=J1."Line_ID"
							AND J1."LineType"=0
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=J."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND	RC."ReconType"=3
					AND C1."ReconDate">vPeriodTo),0)
			+
			IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)
							END
				FROM "NCM_RECON_DET" RC
					INNER JOIN "JDT1" J1 ON RC."TransId"= J1."TransId"
						AND RC."ShortName"= J1."ShortName"
						AND RC."TransRowId"= J1."Line_ID"
						AND J1."LineType"=0
				WHERE RC."TransId"= T0."TransId"
				AND RC."ShortName"= J."ShortName"
				AND RC."ReconType"= 7
				AND RC."ReconDate">vPeriodTo
				AND RC."CancelAbs" IN (SELECT DISTINCT CC."ReconNum" FROM "NCM_RECON_DET" CC
										 WHERE	CC."ShortName"=J."ShortName"
										 AND	CC."TransId"=T0."TransId"
										 AND	CC."ReconType"=0
										 AND	CC."ReconDate"<=vPeriodTo)),0)
			+
			IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
							ELSE SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)*-1
							END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "JDT1" J1 ON RC."TransId"=J1."TransId"
							AND RC."ShortName"=J1."ShortName"
							AND RC."TransRowId"=J1."Line_ID"
							AND J1."LineType"=0
					WHERE RC."TransId"=T0."TransId"
					AND RC."ShortName"=J."ShortName"
					AND RC."ReconType"=5
					AND RC."ReconDate"<=vPeriodTo),0),
/*OpenAmt*/0,
/*OpenAmtFC*/0,
			J."IntrnMatch",
/*PosAgeDays*/DAYS_BETWEEN(T0."DocDate",vPeriodTo),
/*PosAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DocDate")*12)+MONTH(T0."DocDate")),
/*DocAgeDays*/DAYS_BETWEEN(T0."TaxDate",vPeriodTo),
/*DocAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."TaxDate")*12)+MONTH(T0."TaxDate")),
/*DueAgeDays*/DAYS_BETWEEN(T0."DocDueDate",vPeriodTo),
/*DueAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DocDueDate")*12)+MONTH(T0."DocDueDate")),
			'',
			T0."Canceled",
		'', 0, '',
		PT."PymntGroup", BP."CreditLine", 
		T0."Series", IFNULL(N1."SeriesName",''), IFNULL(N1."Remark",'') ,
		'', --IFNULL(T0."U_MQ_PatientID",''),
		'', --IFNULL(T0."U_MQ_PatientName",''), 
		'', --IFNULL(TO_VARCHAR(T0."U_MQ_SurgeryDate",'DD/MM/YYYY'),''), 
		'', '',
		IFNULL(BP."U_IRPBPField1",''),
		IFNULL(BP."U_IRPBPField2",''),
		IFNULL(BP."U_IRPBPField3",''),
		IFNULL(BP."U_IRPBPField4",''),
		IFNULL(BP."U_IRPBPField5",''),
		IFNULL(DIM1."OcrCode",''),
		IFNULL(DIM2."OcrCode2",'')

	FROM "ORCT" T0
		INNER JOIN "OCRD" BP ON BP."CardCode"=T0."CardCode" AND BP."CardType"='C'
		INNER JOIN "JDT1" J ON (J."TransId"=T0."TransId" AND J."ShortName"= CASE WHEN IFNULL(BP."FatherCard",'') = '' then BP."CardCode" else BP."FatherCard" end)
			AND J."LineType"=0 AND IFNULL(J."ContraAct",'')<>''
		LEFT OUTER JOIN "NNM1" N1 ON T0."Series"=N1."Series"
		INNER JOIN "OCTG" PT ON BP."GroupNum"= PT."GroupNum"
		-- Dimension 1
		LEFT JOIN (SELECT Z0."DocNum", STRING_AGG(Z0."OcrCode",',') as "OcrCode"
					FROM (SELECT DISTINCT "DocNum","OcrCode"
							FROM RCT2
							WHERE IFNULL("OcrCode" ,'') <> ''
							--ORDER BY "InvoiceId"
							) Z0
					GROUP BY Z0."DocNum"
					) DIM1 ON T0."DocEntry" = DIM1."DocNum"
		-- Dimension 2
		LEFT JOIN (SELECT Z0."DocNum", STRING_AGG(Z0."OcrCode2",',') as "OcrCode2"
					FROM (SELECT DISTINCT "DocNum","OcrCode2"
							FROM RCT2
							WHERE IFNULL("OcrCode2" ,'') <> ''
							--ORDER BY "InvoiceId"
							) Z0
					GROUP BY Z0."DocNum"
					) DIM2 ON T0."DocEntry" = DIM2."DocNum"

	WHERE J."ShortName" LIKE vBPCode
	AND J."ShortName" BETWEEN vBPCodeFr AND vBPCodeTo
	AND BP."GroupCode" BETWEEN vBPGpFr AND vBPGpTo
	AND BP."SlpCode" IN (SELECT T9."SlpCode" FROM "OSLP" T9 WHERE T9."SlpName" BETWEEN vSlpNameFr AND vSlpNameTo)
	AND T0."DocDate"<=vPeriodTo;
-- (END) INSERT INCOMING PAYMENTS (Caters to Transactions Other DownPayment Request Invoices)

-- (START) INSERT INCOMING PAYMENTS (Caters to DownPayment Request Invoices)
	INSERT INTO "@NCM_AR_AGEING"
	SELECT DISTINCT :vUserName, 
			'06 Receipt-DPRI', 
/*CardCode*/J."ShortName",
			T0."CardCode", T0."DocNum", T0."TransId",
			1,
			'',
			T0."DocDate", T0."TaxDate", T0."DocDueDate",
			T0."CounterRef", T0."Ref1", T0."Ref2", J."Project", BP."SlpCode", 
			T0."DocCurr", 
/*DocRate*/(CASE WHEN T0."DocCurr"=vCurr THEN 1 ELSE T0."DocRate" END),
/*DocTotal*/IFNULL((SELECT SUM(HY."Debit"-HY."Credit")
					FROM "JDT1" HY
					WHERE HY."TransId"=T0."TransId"
					AND HY."ShortName"=J."ShortName"
					AND HY."LineType"=1),0),
/*DocTotalFC*/(CASE WHEN T0."DocCurr"=vCurr THEN 
						IFNULL((SELECT SUM(HY."Debit"-HY."Credit") FROM "JDT1" HY
								WHERE HY."TransId"=T0."TransId"
								AND HY."ShortName"=J."ShortName"
								AND HY."LineType"=1),0)
				ELSE IFNULL((SELECT SUM(HY."FCDebit"-HY."FCCredit") FROM "JDT1" HY
							WHERE HY."TransId"=T0."TransId"
							AND HY."ShortName"=J."ShortName"
							AND HY."LineType"=1),0) 
				END),
/*DocType*/'RC',
/*ClosePaid*/
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
					FROM "NCM_RECON_DET" RC
						INNER JOIN "JDT1" J1 ON RC."TransId"=J1."TransId" AND RC."ShortName"=J1."ShortName"
							AND RC."TransRowId"=J1."Line_ID"
					WHERE RC."TransId"=T0."TransId"
					AND RC."ShortName"=J."ShortName"
					AND RC."ReconType" NOT IN (7,5)
					AND RC."Canceled"='N'
					AND RC."ReconDate"<=vPeriodTo
					AND J1."LineType"=1),0))
			+
			IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum"
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
							AND C1."ReconType"=5
						INNER JOIN "JDT1" J1 ON CC."TransId"=J1."TransId"
							AND CC."ShortName"=J1."ShortName"
							AND CC."TransRowId"=J1."Line_ID"
							AND J1."LineType"=1
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=J."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType"=3
					AND C1."ReconDate">vPeriodTo),0)
			+
			IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "JDT1" J1 ON RC."TransId"= J1."TransId"
							AND RC."ShortName"= J1."ShortName"
							AND RC."TransRowId"= J1."Line_ID"
							AND J1."LineType"=1
					WHERE RC."TransId"= T0."TransId"
					AND RC."ShortName"= J."ShortName"
					AND RC."ReconType"= 7
					AND RC."ReconDate">vPeriodTo
					AND RC."CancelAbs" IN (SELECT DISTINCT CC."ReconNum" FROM "NCM_RECON_DET" CC
											WHERE	CC."ShortName"=J."ShortName"
											 AND	CC."TransId"=T0."TransId"
											 AND	CC."ReconType"=0
											 AND	CC."ReconDate"<=vPeriodTo)),0)
			+
			IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
					FROM "NCM_RECON_DET" RC
						INNER JOIN "JDT1" J1 ON RC."TransId"=J1."TransId"
							AND RC."ShortName"=J1."ShortName"
							AND RC."TransRowId"=J1."Line_ID"
							AND J1."LineType"=1
					WHERE RC."TransId"=T0."TransId"
					AND RC."ShortName"=J."ShortName"
					AND RC."ReconType"=5
					AND RC."ReconDate"<=vPeriodTo),0),
/*ClosePaidFC*/
			(IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
							ELSE SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)*-1
							END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "JDT1" J1 ON RC."TransId"=J1."TransId" AND RC."ShortName"=J1."ShortName"
							AND RC."TransRowId"=J1."Line_ID"
					WHERE RC."TransId"=T0."TransId"
					AND RC."ShortName"=J."ShortName"
					AND RC."ReconType" NOT IN (7,5)
					AND RC."Canceled"='N'
					AND RC."ReconDate"<=vPeriodTo
					AND J1."LineType"=1),0))
			+
			IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
							ELSE SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)*-1
							END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum"
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
							AND C1."ReconType"=5
						INNER JOIN "JDT1" J1 ON CC."TransId"=J1."TransId"
							AND CC."ShortName"=J1."ShortName"
							AND CC."TransRowId"=J1."Line_ID"
							AND J1."LineType"=1
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=J."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType"=3
					AND C1."ReconDate">vPeriodTo),0)
			+
			IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)
							END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "JDT1" J1 ON RC."TransId"= J1."TransId"
							AND RC."ShortName"= J1."ShortName"
							AND RC."TransRowId"= J1."Line_ID"
							AND J1."LineType"=1
					WHERE RC."TransId"= T0."TransId"
					AND RC."ShortName"= J."ShortName"
					AND RC."ReconType"= 7
					AND RC."ReconDate">vPeriodTo
					AND RC."CancelAbs" IN (SELECT DISTINCT CC."ReconNum" FROM "NCM_RECON_DET" CC
											 WHERE	CC."ShortName"=J."ShortName"
											 AND	CC."TransId"=T0."TransId"
											 AND	CC."ReconType"=0
											 AND	CC."ReconDate"<=vPeriodTo)),0)
			+
			IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
							ELSE SUM(CASE WHEN RC."IsCredit"='C' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)*-1
							END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "JDT1" J1 ON RC."TransId"=J1."TransId"
							AND RC."ShortName"=J1."ShortName"
							AND RC."TransRowId"=J1."Line_ID"
							AND J1."LineType"=1
					WHERE RC."TransId"=T0."TransId"
					AND RC."ShortName"=J."ShortName"
					AND RC."ReconType"=5
					AND RC."ReconDate"<=vPeriodTo),0),
/*OpenAmt*/0,
/*OpenAmtFC*/0,
			J."IntrnMatch",
/*PosAgeDays*/DAYS_BETWEEN(T0."DocDate",vPeriodTo),
/*PosAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DocDate")*12)+MONTH(T0."DocDate")),
/*DocAgeDays*/DAYS_BETWEEN(T0."TaxDate",vPeriodTo),
/*DocAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."TaxDate")*12)+MONTH(T0."TaxDate")),
/*DueAgeDays*/DAYS_BETWEEN(T0."DocDueDate",vPeriodTo),
/*DueAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DocDueDate")*12)+MONTH(T0."DocDueDate")),
			'',
			T0."Canceled",
			'', 0, '',
			PT."PymntGroup", BP."CreditLine", 
			T0."Series", IFNULL(N1."SeriesName",''), IFNULL(N1."Remark",'') ,
			'', --IFNULL(T0."U_MQ_PatientID",''),
			'', --IFNULL(T0."U_MQ_PatientName",''), 
			'', --IFNULL(TO_VARCHAR(T0."U_MQ_SurgeryDate",'DD/MM/YYYY'),''), 
			'', '',
			IFNULL(BP."U_IRPBPField1",''),
			IFNULL(BP."U_IRPBPField2",''),
			IFNULL(BP."U_IRPBPField3",''),
			IFNULL(BP."U_IRPBPField4",''),
			IFNULL(BP."U_IRPBPField5",''),
			IFNULL(DIM1."OcrCode",''),
			IFNULL(DIM2."OcrCode2",'')


	FROM "ORCT" T0
			INNER JOIN "OCRD" BP ON BP."CardCode"=T0."CardCode" AND BP."CardType"='C'
			INNER JOIN "JDT1" J ON (J."TransId"=T0."TransId" AND J."ShortName"= CASE WHEN IFNULL(BP."FatherCard",'') = '' then BP."CardCode" else BP."FatherCard" end)
				AND J."LineType"=1 AND IFNULL(J."ContraAct",'')<>''
			LEFT OUTER JOIN "NNM1" N1 ON T0."Series"=N1."Series"
			INNER JOIN "OCTG" PT ON BP."GroupNum"= PT."GroupNum"
		-- Dimension 1
		LEFT JOIN (SELECT Z0."DocNum", STRING_AGG(Z0."OcrCode",',') as "OcrCode"
					FROM (SELECT DISTINCT "DocNum","OcrCode"
							FROM RCT2
							WHERE IFNULL("OcrCode" ,'') <> ''
							--ORDER BY "InvoiceId"
							) Z0
					GROUP BY Z0."DocNum"
					) DIM1 ON T0."DocEntry" = DIM1."DocNum"
		-- Dimension 2
		LEFT JOIN (SELECT Z0."DocNum", STRING_AGG(Z0."OcrCode2",',') as "OcrCode2"
					FROM (SELECT DISTINCT "DocNum","OcrCode2"
							FROM RCT2
							WHERE IFNULL("OcrCode2" ,'') <> ''
							--ORDER BY "InvoiceId"
							) Z0
					GROUP BY Z0."DocNum"
					) DIM2 ON T0."DocEntry" = DIM2."DocNum"

	WHERE J."ShortName" LIKE vBPCode
	AND J."ShortName" BETWEEN vBPCodeFr AND vBPCodeTo
	AND BP."GroupCode" BETWEEN vBPGpFr AND vBPGpTo
	AND BP."SlpCode" IN (SELECT T9."SlpCode" FROM "OSLP" T9 WHERE T9."SlpName" BETWEEN vSlpNameFr AND vSlpNameTo)
	AND T0."DocDate"<=vPeriodTo;
-- (END) INSERT INCOMING PAYMENTS (Caters to DownPayment Request Invoices)

-- (START) INSERT OUTGOING PAYMENTS 
	INSERT INTO "@NCM_AR_AGEING"
	SELECT DISTINCT :vUserName, 
			'07 Payment', 
/*CardCode*/J."ShortName",
			T0."CardCode", T0."DocNum", T0."TransId",
			1,
			'',
			T0."DocDate", T0."TaxDate", T0."DocDueDate",
			T0."CounterRef", T0."Ref1", T0."Ref2", J."Project", BP."SlpCode", 
			T0."DocCurr", 
/*DocRate*/(CASE WHEN T0."DocCurr"=vCurr THEN 1 ELSE T0."DocRate" END),
/*DocTotal*/IFNULL((SELECT SUM(HY."Debit"-HY."Credit")
					FROM "JDT1" HY
					WHERE HY."TransId"=T0."TransId"
					AND HY."ShortName"=J."ShortName"),0),
/*DocTotalFC*/(CASE WHEN T0."DocCurr"=vCurr THEN 
						IFNULL((SELECT SUM(HY."Debit"-HY."Credit") FROM "JDT1" HY
								WHERE HY."TransId"=T0."TransId"
								AND HY."ShortName"=J."ShortName"),0)
				ELSE IFNULL((SELECT SUM(HY."FCDebit"-HY."FCCredit") FROM "JDT1" HY
							WHERE HY."TransId"=T0."TransId"
							AND HY."ShortName"=J."ShortName"),0) END),
/*DocType*/'PY',
/*ClosePaid*/
		(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
				FROM "NCM_RECON_DET" RC
				WHERE RC."TransId"=T0."TransId"
				AND RC."ShortName"=J."ShortName"
				AND RC."ReconType" NOT IN (7,5)
				AND RC."Canceled"='N'
				AND RC."ReconDate"<=vPeriodTo),0))
			+
			IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum"
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
							AND C1."ReconType"=5
						INNER JOIN "JDT1" J1 ON CC."TransId"=J1."TransId"
							AND CC."ShortName"=J1."ShortName"
							AND CC."TransRowId"=J1."Line_ID"
							AND J1."LineType"=0
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=J."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType"=3
					AND C1."ReconDate">vPeriodTo),0)
			+
			IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
					FROM "NCM_RECON_DET" RC
					WHERE RC."TransId"= T0."TransId"
					AND RC."ShortName"= J."ShortName"
					AND RC."ReconType"= 7
					AND RC."ReconDate">vPeriodTo
					AND RC."CancelAbs" IN (SELECT DISTINCT CC."ReconNum" FROM "NCM_RECON_DET" CC
											 WHERE	CC."ShortName"=J."ShortName"
											 AND	CC."TransId"=T0."TransId"
											 AND	CC."ReconType"=0
											 AND	CC."ReconDate"<=vPeriodTo)),0)
			+
			IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
					WHERE RC."TransId"=T0."TransId"
					AND RC."ShortName"=J."ShortName"
					AND RC."ReconType"=5
					AND RC."ReconDate"<=vPeriodTo),0),
/*ClosePaidFC*/
			(IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)
							END
					FROM "NCM_RECON_DET" RC
					WHERE RC."TransId"=T0."TransId"
					AND RC."ShortName"=J."ShortName"
					AND RC."ReconType" NOT IN (7,5)
					AND RC."Canceled"='N'
					AND RC."ReconDate"<=vPeriodTo),0))
			+
			IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)
							END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum"
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
							AND C1."ReconType"=5
						INNER JOIN "JDT1" J1 ON CC."TransId"=J1."TransId"
							AND CC."ShortName"=J1."ShortName"
							AND CC."TransRowId"=J1."Line_ID"
							AND J1."LineType"=0
					WHERE RC."SrcObjAbs"=T0."DocEntry"
					AND RC."TransId"=T0."TransId"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=J."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType"=3
					AND C1."ReconDate">vPeriodTo),0)
			+
			IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)*-1
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)*-1
							END
					FROM "NCM_RECON_DET" RC
					WHERE RC."TransId"= T0."TransId"
					AND RC."ShortName"= J."ShortName"
					AND RC."ReconType"= 7
					AND RC."ReconDate">vPeriodTo
					AND RC."CancelAbs" IN (SELECT DISTINCT CC."ReconNum" FROM "NCM_RECON_DET" CC
											 WHERE	CC."ShortName"=J."ShortName"
											 AND	CC."TransId"=T0."TransId"
											 AND	CC."ReconType"=0
											 AND	CC."ReconDate"<=vPeriodTo)),0)
			+
			IFNULL((SELECT CASE WHEN T0."DocCurr"=vCurr THEN
						SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END)
					END
				FROM "NCM_RECON_DET" RC
				WHERE RC."TransId"=T0."TransId"
				AND RC."ShortName"=J."ShortName"
				AND RC."ReconType"=5
				AND RC."ReconDate"<=vPeriodTo),0),
/*OpenAmt*/0,
/*OpenAmtFC*/0,
		J."IntrnMatch",
/*PosAgeDays*/DAYS_BETWEEN(T0."DocDate",vPeriodTo),
/*PosAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DocDate")*12)+MONTH(T0."DocDate")),
/*DocAgeDays*/DAYS_BETWEEN(T0."TaxDate",vPeriodTo),
/*DocAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."TaxDate")*12)+MONTH(T0."TaxDate")),
/*DueAgeDays*/DAYS_BETWEEN(T0."DocDueDate",vPeriodTo),
/*DueAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DocDueDate")*12)+MONTH(T0."DocDueDate")),
/*DocStatus*/'',
		T0."Canceled",
		'', 0, '',
		PT."PymntGroup", BP."CreditLine", 
		T0."Series", IFNULL(N1."SeriesName",''), IFNULL(N1."Remark",'') ,
		'', '', '', '', '',
		IFNULL(BP."U_IRPBPField1",''),
		IFNULL(BP."U_IRPBPField2",''),
		IFNULL(BP."U_IRPBPField3",''),
		IFNULL(BP."U_IRPBPField4",''),
		IFNULL(BP."U_IRPBPField5",''),
		IFNULL(DIM1."OcrCode",''),
		IFNULL(DIM2."OcrCode2",'')

	FROM "OVPM" T0
		INNER JOIN "OCRD" BP ON BP."CardType"='C'
		INNER JOIN "JDT1" J ON (J."TransId"=T0."TransId" AND J."ShortName"= CASE WHEN IFNULL(BP."FatherCard",'') = '' then BP."CardCode" else BP."FatherCard" end)
			AND IFNULL(J."ContraAct",'')<>''
		LEFT OUTER JOIN "NNM1" N1 ON T0."Series"=N1."Series"
		INNER JOIN "OCTG" PT ON BP."GroupNum"= PT."GroupNum"
		-- Dimension 1
		LEFT JOIN (SELECT Z0."DocNum", STRING_AGG(Z0."OcrCode",',') as "OcrCode"
					FROM (SELECT DISTINCT "DocNum","OcrCode"
							FROM VPM2
							WHERE IFNULL("OcrCode" ,'') <> ''
							-- ORDER BY "InvoiceId"
							) Z0
					GROUP BY Z0."DocNum"
					) DIM1 ON T0."DocEntry" = DIM1."DocNum"
		-- Dimension 2
		LEFT JOIN (SELECT Z0."DocNum", STRING_AGG(Z0."OcrCode2",',') as "OcrCode2"
					FROM (SELECT DISTINCT "DocNum","OcrCode2"
							FROM VPM2
							WHERE IFNULL("OcrCode2" ,'') <> ''
							-- ORDER BY "InvoiceId"
							) Z0
					GROUP BY Z0."DocNum"
					) DIM2 ON T0."DocEntry" = DIM2."DocNum"

	WHERE J."ShortName" LIKE vBPCode
	AND J."ShortName" BETWEEN vBPCodeFr AND vBPCodeTo
	AND BP."GroupCode" BETWEEN vBPGpFr AND vBPGpTo
	AND BP."SlpCode" IN (SELECT T9."SlpCode" FROM "OSLP" T9 WHERE T9."SlpName" BETWEEN vSlpNameFr AND vSlpNameTo)
	AND J."ShortName" IN (SELECT T9."CardCode" FROM "OCRD" T9)
	AND T0."DocDate"<=vPeriodTo;
-- (END) INSERT OUTGOING PAYMENTS 

-- (START) AR CHECK PAYMENT
	INSERT INTO "@NCM_AR_AGEING"
	SELECT DISTINCT :vUserName, 
			'08 Check Payment', 
			T0."VendorCode",
			T0."VendorCode", 
			T0."CheckKey", T0."TransNum",
			1,
			'',
			T0."PmntDate", T0."PmntDate", T0."PmntDate",
			T0."TransRef", '', '', 
			J."Project", BP."SlpCode", 
/*DocCurr*/(CASE IFNULL(J."FCCurrency",vCurr) 
				WHEN '' THEN vCurr
				WHEN vCurr THEN vCurr
				ELSE J."FCCurrency" END),
/*DocRate*/(CASE IFNULL(J."FCCurrency",vCurr) 
					WHEN '' THEN 1
					WHEN vCurr THEN 1
					ELSE (CASE WHEN vDirectRate='Y' THEN 
								(J."Debit"-J."Credit")/
									(CASE WHEN (J."FCDebit"-J."FCCredit")=0 THEN 1 
										ELSE (J."FCDebit"-J."FCCredit") END)
							ELSE
								(J."FCDebit"-J."FCCredit")/
									(CASE WHEN (J."Debit"-J."Credit")=0 THEN 1 
										ELSE (J."Debit"-J."Credit") END)
							END) 
			END),
/*DocTotal*/J."Debit"-J."Credit", 
/*DocTotalFC*/(CASE IFNULL(J."FCCurrency",vCurr)
					WHEN vCurr THEN J."Debit"-J."Credit" 
					WHEN '' THEN (J."FCDebit"-J."FCCredit")
					WHEN NULL THEN (J."FCDebit"-J."FCCredit")
					ELSE (J."FCDebit"-J."FCCredit") END),
/*DocType*/'CP',
/*ClosePaid*/(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=T0."CheckKey"
					AND RC."TransId"=T0."TransNum"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=T0."VendorCode"
					AND RC."Canceled"='N'
					AND RC."ReconType" IN (0,3)
					AND RC."ReconDate"<=vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
							AND RC."CancelAbs"=CC."CancelAbs"	
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
					WHERE RC."SrcObjAbs"=T0."TransNum"
					AND RC."TransId"=T0."TransNum"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=T0."VendorCode"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType"=3
					AND C1."ReconType"=5
					AND C1."ReconDate">vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
							AND RC."SrcObjAbs"=CC."SrcObjAbs"
							AND RC."SrcObjTyp"=CC."SrcObjTyp"
							AND RC."TransId"=CC."TransId"
					WHERE RC."SrcObjAbs"=T0."CheckKey"
					AND RC."TransId"=T0."TransNum"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=T0."VendorCode"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType"=0
					AND CC."ReconType"=7
					AND CC."ReconDate">vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=T0."CheckKey"
					AND RC."TransId"=T0."TransNum"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=T0."VendorCode"
					AND RC."InitObjAbs"=T0."TransNum"
					AND RC."Canceled"='N'
					AND RC."ReconType"=5
					AND RC."ReconDate"<=vPeriodTo),0)),
/*ClosePaidFC*/
			(IFNULL((SELECT CASE WHEN IFNULL(J."FCCurrency",vCurr)=vCurr THEN 
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) 
							END
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=T0."CheckKey"
					AND RC."TransId"=T0."TransNum"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=T0."VendorCode"
					AND RC."Canceled"='N'
					AND RC."ReconType" IN (0,3)
					AND RC."ReconDate"<=vPeriodTo),0))
			+
			IFNULL((SELECT CASE WHEN IFNULL(J."FCCurrency",vCurr)=vCurr THEN 
									SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) 
							END 
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
							AND RC."CancelAbs"=CC."CancelAbs"	-- v7.4.2
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
					WHERE RC."SrcObjAbs"=T0."CheckKey"
					AND RC."TransId"=T0."TransNum"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=T0."VendorCode"
					AND RC."Canceled"='Y'
					AND	RC."ReconDate"<=vPeriodTo
					AND RC."ReconType"=3
					AND C1."ReconType"=5
					AND C1."ReconDate">vPeriodTo),0)
			+
			IFNULL((SELECT CASE WHEN IFNULL(J."FCCurrency",vCurr)=vCurr THEN 
									SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) 
							END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
							AND RC."SrcObjAbs"=CC."SrcObjAbs"
							AND RC."SrcObjTyp"=CC."SrcObjTyp"
							AND RC."TransId"=CC."TransId"
					WHERE RC."SrcObjAbs"=T0."CheckKey"
					AND RC."TransId"=T0."TransNum"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=T0."VendorCode"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType"=0
					AND CC."ReconType"=7
					AND CC."ReconDate">vPeriodTo),0)
			+
			(IFNULL((SELECT CASE WHEN IFNULL(J."FCCurrency",vCurr)=vCurr THEN 
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) 
							END
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=T0."CheckKey"
					AND RC."TransId"=T0."TransNum"
					AND RC."TransRowId"=J."Line_ID"
					AND RC."ShortName"=T0."VendorCode"
					AND RC."InitObjAbs"=T0."TransNum"
					AND RC."Canceled"='N'
					AND RC."ReconType"=5
					AND RC."ReconDate"<=vPeriodTo),0)),
/*OpenAmt*/0,
/*OpenAmtFC*/0,
			J."IntrnMatch",
/*PosAgeDays*/DAYS_BETWEEN(T0."PmntDate",vPeriodTo),
/*PosAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."PmntDate")*12)+MONTH(T0."PmntDate")),
/*DocAgeDays*/DAYS_BETWEEN(T0."PmntDate",vPeriodTo),
/*DocAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."PmntDate")*12)+MONTH(T0."PmntDate")),
/*DueAgeDays*/DAYS_BETWEEN(T0."PmntDate",vPeriodTo),
/*DueAgeMths*/((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."PmntDate")*12)+MONTH(T0."PmntDate")),
			'',
			T0."Canceled",
			'', 0, '',
			PT."PymntGroup", BP."CreditLine", 
			0,'', '', --T0."Series", IFNULL(N1."SeriesName",''),
			'', '', '', '', '',
			IFNULL(BP."U_IRPBPField1",''),
			IFNULL(BP."U_IRPBPField2",''),
			IFNULL(BP."U_IRPBPField3",''),
			IFNULL(BP."U_IRPBPField4",''),
			IFNULL(BP."U_IRPBPField5",''),
			'',
			''

	FROM "OCHO" T0
		INNER JOIN "OCRD" BP ON BP."CardCode"=T0."VendorCode" AND BP."CardType"='C'
		INNER JOIN "JDT1" J ON J."TransId"=T0."TransNum" AND J."ShortName"=T0."VendorCode"
		--LEFT OUTER JOIN "NNM1" N1 ON T0."Series"=N1."Series"
		INNER JOIN "OCTG" PT ON BP."GroupNum"= PT."GroupNum"
	WHERE T0."VendorCode" LIKE vBPCode
	AND T0."VendorCode" BETWEEN vBPCodeFr AND vBPCodeTo
	AND BP."GroupCode" BETWEEN vBPGpFr AND vBPGpTo
	AND BP."SlpCode" IN (SELECT T9."SlpCode" FROM "OSLP" T9 WHERE T9."SlpName" BETWEEN vSlpNameFr AND vSlpNameTo)
	AND T0."PmntDate"<=vPeriodTo;
-- (END) AR CHECK PAYMENT

-- (START) RECONCILIATION WITH CASE1
	INSERT INTO #NCM_AR_AGEING_CASE1
	SELECT DISTINCT T0."OrigTrnsId", T0."TransId", T0."OrigObjTyp", T0."ShortName", 
			IFNULL(JT."TransCurr",vCurr), T0."TrnsTtlAmt", T0."TrnsTtlFc"
	FROM "CASE1" T0
		LEFT OUTER JOIN "OCRD" BP ON T0."ShortName"=BP."CardCode"
		LEFT OUTER JOIN "OJDT" JT ON T0."OrigTrnsId"= JT."TransId"
	WHERE BP."CardType"= 'C'
	AND T0."OrigObjTyp" IN (13,14,203,24,46)
	AND T0."ShortName" LIKE vBPCode
	AND BP."GroupCode" BETWEEN vBPGpFr AND vBPGpTo
	AND BP."SlpCode" IN (SELECT T9."SlpCode" FROM "OSLP" T9 
					WHERE T9."SlpName" BETWEEN vSlpNameFr AND vSlpnameTo);
					
	SELECT IFNULL(MIN("RUTRANSID"),0)
	INTO vRUTransId
	FROM #NCM_AR_AGEING_CASE1;
	
	WHILE vRUTransId<>0 DO

		SELECT T0."ORIGTRNSID", T0."ORIGOBJTYP", T0."CARDCODE", T0."ORIGTRANSCURR", T0."TRNSTTLAMT", T0."TRNSTTLFC"
		INTO vOrigTrnsId, vOrigObjTyp, vCardCode,vOrigTransCurr,vTrnsTtlAmt,vTrnsTtlFc
		FROM #NCM_AR_AGEING_CASE1 T0
		WHERE T0."RUTRANSID"=vRUTransId;

		vClosePaid := 0;
		vClosePaidFc := 0;
		vRecordCount := 0;
	
		SELECT TOP 1 C0."TrnsTtlAmt",
				IFNULL(CASE WHEN IFNULL(vOrigTransCurr,vCurr)=vCurr THEN C0."TrnsTtlAmt"
						ELSE C0."TrnsTtlFc" END,0)
		INTO vClosePaid, vClosePaidFc
		FROM "CASE1" C0 
			LEFT OUTER JOIN "OJDT" JT1 ON C0."TransId"=JT1."TransId"
		WHERE C0."TransId"=vRUTransId
		AND C0."ShortName"=vCardCode
		AND C0."OrigTrnsId"=vOrigTrnsId
		AND C0."TrnsTtlAmt"=vTrnsTtlAmt
		AND C0."TrnsTtlFc"=vTrnsTtlFc			
		AND IFNULL(JT1."RefDate",ADD_DAYS(vPeriodTo,1))<=vPeriodTo;

		SELECT COUNT(*)
		INTO vRecordCount
		FROM "@NCM_AR_AGEING"
		WHERE "USERNAME"=vUserName
		AND "TRANSID"=vOrigTrnsId;

		IF vRecordCount=0 THEN
			vRecordCount := 1;
		END IF;

		UPDATE "@NCM_AR_AGEING"
		SET "CLOSEPAID"="CLOSEPAID"+(vClosePaid/vRecordCount),
			"CLOSEPAIDFC"="CLOSEPAIDFC"+(vClosePaidFc/vRecordCount)
		WHERE "USERNAME"=vUserName
		AND "TRANSID"= vOrigTrnsId
		AND	("DOCTOTALFC"-"CLOSEPAIDFC")<>0;
			
		SELECT IFNULL(MIN("RUTRANSID"),0)
		INTO vRUTransId
		FROM #NCM_AR_AGEING_CASE1
		WHERE "RUTRANSID">vRUTransId;
		
	END WHILE;
-- (END) RECONCILIATION WITH CASE1

-- (START) RECONCILIATION WITH @NCM_CASE1
/*
	INSERT INTO #NCM_AR_AGEING_NCMCASE1
	SELECT DISTINCT T0."U_Transid", T0."U_CardCode", IFNULL(JT."DOCCUR",vCurr), 
		T0."U_InstlmntId", T0."U_BaseTransid",T0."U_AmtApplied", T0."U_AmtAppliedFC"
	FROM "@NCM_CASE1" T0
		LEFT OUTER JOIN "OCRD" BP ON T0."U_CardCode"=BP."CardCode"
		LEFT OUTER JOIN "@NCM_AR_AGEING" JT ON T0."U_Transid"=JT."TRANSID"
	WHERE BP."CardType"= 'C'
	AND T0."U_RptType" IN ('B','A')
	AND T0."U_CardCode" LIKE vBPCode
	AND BP."GroupCode" BETWEEN vBPGpFr AND vBPGpTo
	AND BP."SlpCode" IN (SELECT T9."SlpCode" FROM "OSLP" T9 
					WHERE T9."SlpName" BETWEEN vSlpNameFr AND vSlpnameTo)
	AND IFNULL(T0."U_PosDate",ADD_DAYS(vPeriodTo,1))<=vPeriodTo
	AND JT."USERNAME"=vUserName;

	UPDATE "@NCM_AR_AGEING"
	SET "OPENPAID"="OPENPAID"+#NCM_AR_AGEING_NCMCASE1."NCM_AMTAPPLIED"
	FROM "@NCM_AR_AGEING" INNER JOIN #NCM_AR_AGEING_NCMCASE1
			ON ("@NCM_AR_AGEING"."TRANSID" = #NCM_AR_AGEING_NCMCASE1."NCM_TRANSID"
				and "@NCM_AR_AGEING"."BASETRANSID" = #NCM_AR_AGEING_NCMCASE1."NCM_BASETRANSID"
				and "@NCM_AR_AGEING"."CARDCODE" = #NCM_AR_AGEING_NCMCASE1."NCM_CARDCODE")
	WHERE #NCM_AR_AGEING_NCMCASE1."NCM_TRANSCURR"=vCurr ;

	UPDATE "@NCM_AR_AGEING"
	SET "OPENPAID"="OPENPAID"+#NCM_AR_AGEING_NCMCASE1."NCM_AMTAPPLIEDFC"
	FROM "@NCM_AR_AGEING" INNER JOIN #NCM_AR_AGEING_NCMCASE1
			ON ("@NCM_AR_AGEING"."TRANSID" = #NCM_AR_AGEING_NCMCASE1."NCM_TRANSID"
				and "@NCM_AR_AGEING"."BASETRANSID" = #NCM_AR_AGEING_NCMCASE1."NCM_BASETRANSID"
				and "@NCM_AR_AGEING"."CARDCODE" = #NCM_AR_AGEING_NCMCASE1."NCM_CARDCODE")
	WHERE #NCM_AR_AGEING_NCMCASE1."NCM_TRANSCURR"<>vCurr ;
	*/
-- (END) RECONCILIATION WITH @NCM_CASE1

/* CASE1 TABLE HAS NO DATA FOR ALL 4 COMPANIES
-- (START) INSERT RU TRANSACTIONS
	INSERT INTO "@NCM_AR_AGEING"
	SELECT DISTINCT :vUserName, 
			'09 Recon Upgrade', 
			RU."ShortName",RU."ShortName",
			T0."TransId", T0."TransId",
			1,
			'',
			T0."RefDate", T0."TaxDate", T0."DueDate",
			(T0."TransId"), 
			T0."Ref1", T0."Ref2", T1."Project",
			T2."SlpCode",
			(CASE IFNULL(T1."FCCurrency",vCurr) 
					WHEN '' THEN vCurr
					WHEN vCurr THEN vCurr
					ELSE T1."FCCurrency" END),
			1,
			CASE WHEN RU."CredDeb" = 'D' THEN RU."Amount" ELSE RU."Amount" *-1 END, 
			(CASE WHEN IFNULL(T1."FCCurrency",vCurr)=vCurr THEN 
							CASE WHEN RU."CredDeb"= 'D' THEN RU."Amount" ELSE RU."Amount" *-1 END
						ELSE CASE WHEN RU."CredDeb" = 'D' THEN RU."AmountFC" ELSE RU."AmountFC" *-1 END
						END),
			'RU',
	
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=RU."TransId"
					AND RC."TransId"=RU."TransId"
					AND RC."TransRowId"=T1."Line_ID"
					AND RC."ShortName"=RU."ShortName"
					AND RC."Canceled"='N'
					AND RC."ReconType" IN (0,1,3,4)
					AND RC."ReconDate"<=vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
							AND RC."CancelAbs"=CC."CancelAbs"
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
					WHERE RC."SrcObjAbs"=RU."TransId"
					AND RC."TransId"=RU."TransId"
					AND RC."TransRowId"=T1."Line_ID"
					AND RC."ShortName"=RU."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (3,4)
					AND C1."ReconType"=5
					AND C1."ReconDate">vPeriodTo),0))
			+
			(IFNULL((SELECT SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
							AND RC."SrcObjAbs"=CC."SrcObjAbs"
							AND RC."SrcObjTyp"=CC."SrcObjTyp"
							AND RC."TransId"=CC."TransId"
					WHERE RC."SrcObjAbs"=RU."TransId"
					AND RC."TransId"=RU."TransId"
					AND RC."TransRowId"=T1."Line_ID"
					AND RC."ShortName"=RU."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (0,1)
					AND CC."ReconType"=7
					AND CC."ReconDate">vPeriodTo),0)),
	
			(IFNULL((SELECT CASE WHEN IFNULL(T1."FCCurrency",vCurr)=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) END
					FROM "NCM_RECON_DET" RC
					WHERE RC."SrcObjAbs"=RU."TransId"
					AND RC."TransId"=RU."TransId"
					AND RC."TransRowId"=T1."Line_ID"
					AND RC."ShortName"=RU."ShortName"
					AND RC."Canceled"='N'
					AND RC."ReconType" IN (0,1,3,4)
					AND RC."ReconDate"<=vPeriodTo),0))
			+
			(IFNULL((SELECT CASE WHEN IFNULL(T1."FCCurrency",vCurr)=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."ReconNum"=CC."ReconNum" 
							AND RC."InitObjAbs"=CC."SrcObjAbs"
							AND RC."InitObjTyp"=CC."SrcObjTyp"
							AND RC."CancelAbs"=CC."CancelAbs"	
						INNER JOIN "NCM_RECON_DET" C1 ON CC."SrcObjAbs"=C1."SrcObjAbs"
							AND CC."SrcObjTyp"=C1."SrcObjTyp"
							AND CC."TransId"=C1."TransId"
							AND CC."TransRowId"=C1."TransRowId"
					WHERE RC."SrcObjAbs"=RU."TransId"
					AND RC."TransId"=RU."TransId"
					AND RC."TransRowId"=T1."Line_ID"
					AND RC."ShortName"=RU."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (3,4)
					AND C1."ReconType"=5
					AND C1."ReconDate">vPeriodTo),0))
			+
			(IFNULL((SELECT CASE WHEN IFNULL(T1."FCCurrency",vCurr)=vCurr THEN
								SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSum" ELSE RC."ReconSum"*-1 END)
							ELSE SUM(CASE WHEN RC."IsCredit"='D' THEN RC."ReconSumFC" ELSE RC."ReconSumFC"*-1 END) END
					FROM "NCM_RECON_DET" RC
						INNER JOIN "NCM_RECON_DET" CC ON RC."CancelAbs"=CC."ReconNum"
							AND RC."SrcObjAbs"=CC."SrcObjAbs"
							AND RC."SrcObjTyp"=CC."SrcObjTyp"
							AND RC."TransId"=CC."TransId"
					WHERE RC."SrcObjAbs"=RU."TransId"
					AND RC."TransId"=RU."TransId"
					AND RC."TransRowId"=T1."Line_ID"
					AND RC."ShortName"=RU."ShortName"
					AND RC."Canceled"='Y'
					AND RC."ReconDate"<=vPeriodTo
					AND RC."ReconType" IN (0,1)
					AND CC."ReconType"=7
					AND CC."ReconDate">vPeriodTo),0)),
			0,
			0,
			T1."IntrnMatch",
			DAYS_BETWEEN(T0."RefDate",vPeriodTo),
			((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."RefDate")*12)+MONTH(T0."RefDate")),
			DAYS_BETWEEN(T0."TaxDate",vPeriodTo),
			((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."TaxDate")*12)+MONTH(T0."TaxDate")),
			DAYS_BETWEEN(T0."DueDate",vPeriodTo),
			((YEAR(vPeriodTo)*12)+MONTH(vPeriodTo))-((YEAR(T0."DueDate")*12)+MONTH(T0."DueDate")),
			'',
			'',
			'', 0, '',
			PT."PymntGroup", T2."CreditLine", 0,'', '',
			'', '', '', '', '',
			IFNULL(T2."U_IRPBPField1",''),
			IFNULL(T2."U_IRPBPField2",''),
			IFNULL(T2."U_IRPBPField3",''),
			IFNULL(T2."U_IRPBPField4",''),
			IFNULL(T2."U_IRPBPField5",'')
	FROM CASE1 RU
		LEFT JOIN OJDT T0 ON RU."TransId"=T0."TransId"
		LEFT JOIN JDT1 T1 ON (RU."TransId"=T1."TransId" AND RU."TransLine" = T1."Line_ID")
		LEFT JOIN OCRD T2 ON RU."ShortName"=T2."CardCode" AND T2."CardType"='C'
		INNER JOIN "OCTG" PT ON T2."GroupNum"= PT."GroupNum"
	WHERE T0."RefDate"<=vPeriodTo
	AND RU."ShortName" LIKE vBPCode
	AND RU."ShortName" BETWEEN vBPCodeFr AND vBPCodeTo
	AND T2."GroupCode" BETWEEN vBPGpFr AND vBPGpTo
	AND T2."SlpCode" IN (SELECT "SlpCode" FROM "OSLP" WHERE "SlpName" BETWEEN vSlpNameFr AND vSlpNameTo)
	AND T1."ContraAct" NOT IN (SELECT "AcctCode" FROM OACT WHERE "U_Excl_AC"='Y');
-- (END) INSERT RU TRANSACTIONS
*/

-- (START) UPDATE
	UPDATE "@NCM_AR_AGEING"
	SET "DOCTOTAL"="DOCTOTAL"*-1,
		"DOCTOTALFC"="DOCTOTALFC"*-1
	WHERE "USERNAME"=vUserName
	AND "DOCTYPE" IN ('IN','DP','CN')
	AND "DOCSTATUS"='C'
	AND "CANCELLED"='C';
-- (END) UPDATE 

-- (START) UPDATE 
	UPDATE "@NCM_AR_AGEING"
	SET "OPENAMT"="DOCTOTAL"-"CLOSEPAID",
		"OPENAMTFC"="DOCTOTALFC"-"CLOSEPAIDFC"
	WHERE "USERNAME"=vUserName;
-- (END) UPDATE 

-- (START) UPDATE 
	UPDATE "@NCM_AR_AGEING"
	SET "OPENAMT"=0
	WHERE "USERNAME"=vUserName
	AND "OPENAMTFC"=0
	AND "DOCTYPE" NOT IN ('JE','RU');
-- (END) UPDATE 

-- (START) DELETE WHERE OpenAmt AND OpenAmtFC are zero.
	DELETE FROM "@NCM_AR_AGEING"
	WHERE "USERNAME"=vUserName
	AND "OPENAMT"=0
	AND "DOCTYPE"<>'RU';
-- (END) DELETE WHERE OpenAmt AND OpenAmtFC are zero.

-- (START) DELETE WHERE OpenAmt AND OpenAmtFC are zero.
	DELETE FROM "@NCM_AR_AGEING"
	WHERE "USERNAME"=vUserName
	AND "DOCTYPE"='RC'
	AND	"TRANSID" IN (SELECT "TransId" FROM "ORCT"
					 WHERE IFNULL("CancelDate",ADD_DAYS(vPeriodTo,1))<=vPeriodTo)
	AND "TRANSID" NOT IN (SELECT "TransId" FROM "NCM_RECON_DET" 
							WHERE "SrcObjTyp"='24');
-- (END) DELETE WHERE OpenAmt AND OpenAmtFC are zero.

-- (START) DELETE WHERE OpenAmt AND OpenAmtFC are zero.
	DELETE FROM "@NCM_AR_AGEING"
	WHERE "USERNAME"=vUserName
	AND "DOCTYPE"='RC'
	AND	"TRANSID" IN (SELECT "TransId" FROM "ORCT"
					 WHERE "CancelDate" IS NULL
					 AND "Canceled"='Y');
-- (END) DELETE WHERE OpenAmt AND OpenAmtFC are zero.

-- (START) DELETE WHERE OpenAmt AND OpenAmtFC are zero.
	DELETE FROM "@NCM_AR_AGEING"
	WHERE "USERNAME"=vUserName
	AND "DOCTYPE"='PY'
	AND	"TRANSID" IN (SELECT "TransId" FROM "OVPM"
					 WHERE "CancelDate" IS NULL
					 AND "Canceled"='Y');
-- (END) DELETE WHERE OpenAmt AND OpenAmtFC are zero.

-- (START) DELETE WHERE OpenAmt AND OpenAmtFC are zero.
	DELETE FROM "@NCM_AR_AGEING"
	WHERE "USERNAME"=vUserName
	AND	"TRANSID" IN (SELECT "TRANSID" FROM "@NCM_AR_AGEING"
					 WHERE "USERNAME"=vUserName
					 GROUP BY "TRANSID" 
					 HAVING SUM("OPENAMT")=0)
	AND "DOCTYPE" NOT IN ('JE','RU');
-- (END) DELETE WHERE OpenAmt AND OpenAmtFC are zero.

-- (START) DELETE WHERE OpenAmt AND OpenAmtFC are zero.
	DELETE FROM "@NCM_AR_AGEING"
	WHERE "USERNAME"=vUserName
	AND	"TRANSID" IN (SELECT "TRANSID" FROM "@NCM_AR_AGEING"
					 WHERE "USERNAME"=vUserName
					 GROUP BY "TRANSID" 
					 HAVING SUM("OPENAMT")=0 AND SUM("OPENAMTFC")=0)
	AND "DOCTYPE"='RU';
-- (END) DELETE WHERE OpenAmt AND OpenAmtFC are zero.

-- (START) UPDATE CUSTOMER WITH THE LAST PAYMENT INFORMATION

	UPDATE T0
	SET "PAYMENTCUR" = T1."PaymentCur",
		"PAYMENTAMT" = T1."PaymentAmt",
		"PAYMENTDATE" = T1."PaymentDate"
	FROM "@NCM_AR_AGEING" T0
			INNER JOIN 
			(SELECT RC."CardCode", RC."DocNum", RC."TransId", RC."DocCurr" as "PaymentCur", 
					CASE WHEN RC."DocCurr"=vCurr THEN SUM(IFNULL(J."Debit",0) - IFNULL(J."Credit",0)) ELSE SUM(IFNULL(J."FCDebit",0)-IFNULL(J."FCCredit",0)) END as "PaymentAmt",
					RC."DocDate" as "PaymentDate",
					ROW_NUMBER() OVER (PARTITION BY "CardCode" order by RC."DocDate" DESC, RC."TransId" Desc) as "RowNum"
			FROM	ORCT RC INNER JOIN JDT1 J ON (RC."TransId" = J."TransId" and RC."CardCode" = J."ShortName")
			GROUP BY RC."CardCode", RC."DocNum", RC."TransId",RC."DocDate",RC."DocCurr"
			)  T1 ON T0."CARDCODE" = T1."CardCode"	AND T1."RowNum" = 1
	;
-- (END) UPDATE CUSTOMER WITH THE LAST PAYMENT INFORMATION	
	
	DROP TABLE #NCM_AR_AGEING_CASE1;
	DROP TABLE #NCM_AR_AGEING_NCMCASE1;

-- (START) RECORD LISTING
	SELECT T0.* FROM "@NCM_AR_AGEING" T0 WHERE T0."USERNAME"=vUserName;
-- (END) RECORD LISTING


END;

