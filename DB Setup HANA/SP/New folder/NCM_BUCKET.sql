DECLARE @COUNT INT
DECLARE @RecordCount INT
SET @RecordCount = 0
SELECT @COUNT = MAX(CODE) FROM [@NCM_BUCKET]
SET @COUNT = ISNULL(@COUNT,0) + 1

SELECT @RecordCount = COUNT(*) FROM [@NCM_BUCKET] WHERE U_TYPE = 'NCM_SAR_MOV1'

IF @RecordCount = 0
BEGIN
	INSERT INTO [@NCM_BUCKET] (Code, Name, U_Type, U_Bucket1Txt, U_Bucket2Txt, U_Bucket3Txt, U_Bucket4Txt
		, U_Bucket5Txt, U_Bucket6Txt, U_Bucket7Txt, U_Bucket8Txt, U_Bucket9Txt, U_Bucket1Val
		, U_Bucket2Val, U_Bucket3Val, U_Bucket4Val, U_Bucket5Val, U_Bucket6Val, U_Bucket7Val
		, U_Bucket8Val, U_Bucket9Val) VALUES (@COUNT,@COUNT,'NCM_SAR_MOV1', '0-30 Days', '31-60 Days'
		, '61-90 Days', '91-120 Days', '121-150 Days', '151-180 Days', '181-270 Days', '271-365 Days'
		, '>365 Days', 30, 60, 90, 120, 150, 180, 270, 365, 365)
	SET @COUNT = @COUNT + 1
END

SELECT @RecordCount = COUNT(*) FROM [@NCM_BUCKET] WHERE U_TYPE = 'NCM_AR_AGEING'

IF @RecordCount = 0
BEGIN
	INSERT INTO [@NCM_BUCKET] (Code, Name, U_Type, U_Bucket1Txt, U_Bucket2Txt, U_Bucket3Txt, U_Bucket4Txt
		, U_Bucket5Txt, U_Bucket6Txt, U_Bucket7Txt, U_Bucket8Txt, U_Bucket9Txt, U_Bucket1Val
		, U_Bucket2Val, U_Bucket3Val, U_Bucket4Val, U_Bucket5Val, U_Bucket6Val, U_Bucket7Val
		, U_Bucket8Val, U_Bucket9Val) VALUES (@COUNT,@COUNT,'NCM_AR_AGEING', '0-30', '31-60'
		, '61-90', '91-120', '>120', '', '', ''
		, '>365 Days', 30, 60, 90, 120, 120, 120, 120, 120, 120)
	SET @COUNT = @COUNT + 1
END

SELECT @RecordCount = COUNT(*) FROM [@NCM_BUCKET] WHERE U_TYPE = 'NCM_AP_AGEING'

IF @RecordCount = 0
BEGIN
	INSERT INTO [@NCM_BUCKET] (Code, Name, U_Type, U_Bucket1Txt, U_Bucket2Txt, U_Bucket3Txt, U_Bucket4Txt
		, U_Bucket5Txt, U_Bucket6Txt, U_Bucket7Txt, U_Bucket8Txt, U_Bucket9Txt, U_Bucket1Val
		, U_Bucket2Val, U_Bucket3Val, U_Bucket4Val, U_Bucket5Val, U_Bucket6Val, U_Bucket7Val
		, U_Bucket8Val, U_Bucket9Val) VALUES (@COUNT,@COUNT,'NCM_AP_AGEING', '0-30', '31-60'
		, '61-90', '91-120', '>120', '', '', ''
		, '>365 Days', 30, 60, 90, 120, 120, 120, 120, 120, 120)
	SET @COUNT = @COUNT + 1
END

SELECT @RecordCount = COUNT(*) FROM [@NCM_BUCKET] WHERE U_TYPE = 'NCM_AR_AGINGP_6B'

IF @RecordCount = 0
BEGIN
	INSERT INTO [@NCM_BUCKET] (Code, Name, U_Type, U_Bucket1Txt, U_Bucket2Txt, U_Bucket3Txt, U_Bucket4Txt
		, U_Bucket5Txt, U_Bucket6Txt, U_Bucket7Txt, U_Bucket8Txt, U_Bucket9Txt, U_Bucket1Val
		, U_Bucket2Val, U_Bucket3Val, U_Bucket4Val, U_Bucket5Val, U_Bucket6Val, U_Bucket7Val
		, U_Bucket8Val, U_Bucket9Val) VALUES (@COUNT,@COUNT,'NCM_AR_AGINGP_6B', '0-30', '31-60'
		, '61-90', '91-120', '120-180', '>180', '', ''
		, '>365 Days', 30, 60, 90, 120, 180, 180, 180, 180, 180)
	SET @COUNT = @COUNT + 1
END

SELECT @RecordCount = COUNT(*) FROM [@NCM_BUCKET] WHERE U_TYPE = 'NCM_AR_AGEING_7B'

IF @RecordCount = 0
BEGIN
	INSERT INTO [@NCM_BUCKET] (Code, Name, U_Type, U_Bucket1Txt, U_Bucket2Txt, U_Bucket3Txt, U_Bucket4Txt
		, U_Bucket5Txt, U_Bucket6Txt, U_Bucket7Txt, U_Bucket8Txt, U_Bucket9Txt, U_Bucket1Val
		, U_Bucket2Val, U_Bucket3Val, U_Bucket4Val, U_Bucket5Val, U_Bucket6Val, U_Bucket7Val
		, U_Bucket8Val, U_Bucket9Val) VALUES (@COUNT,@COUNT,'NCM_AR_AGEING_7B' 
		, '0-30', '31-60', '61-90', '91-120', '120-150', '151-180', '>180', '', '>365 Days'
		, 30, 60, 90, 120, 150, 180, 180, 180, 180)

	SET @COUNT = @COUNT + 1
END