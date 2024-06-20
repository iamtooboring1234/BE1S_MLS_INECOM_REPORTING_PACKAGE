DECLARE			@ItemCodeFr_2		NVARCHAR(20)
DECLARE			@ItemCodeTo_2		NVARCHAR(20)
DECLARE			@WhseCodeFr_2		NVARCHAR(20)
DECLARE			@WhseCodeTo_2		NVARCHAR(20)
DECLARE			@ItemDim1Fr_2		NVARCHAR(8)
DECLARE			@ItemDim1To_2		NVARCHAR(8)
DECLARE			@ItemDim2Fr_2		NVARCHAR(8)
DECLARE			@ItemDim2To_2		NVARCHAR(8)
DECLARE			@ItemGrpNamFr_2	NVARCHAR(100)
DECLARE			@ItemGrpNamTo_2	NVARCHAR(100)
DECLARE			@AsAtDate_2		DATETIME
DECLARE			@Username_2		NVARCHAR(100)
DECLARE			@ItemTypesFr_2	NVARCHAR(8)
DECLARE			@ItemTypesTo_2	NVARCHAR(30)

SET @ItemCodeFr_2 = '{0}'
SET @ItemCodeTo_2 = '{1}'
SET @WhseCodeFr_2 = '{2}'
SET @WhseCodeTo_2 = '{3}'
SET @ItemDim1Fr_2 = '{4}'
SET @ItemDim1To_2 = '{5}'
SET @ItemDim2Fr_2 = '{6}'
SET @ItemDim2To_2 = '{7}'
SET @ItemGrpNamFr_2 = '{8}'
SET @ItemGrpNamTo_2 = '{9}'
SET @AsAtDate_2 = '{10}'
SET @Username_2 = '{11}'
SET @ItemTypesFr_2 = '{12}'
SET @ItemTypesTo_2 = '{13}'

EXEC NCM_RPT_TM_SAR
	@ItemCodeFr = @ItemCodeFr_2
	, @ItemCodeTo = @ItemCodeTo_2
	, @WhseCodeFr = @WhseCodeFr_2
	, @WhseCodeTo = @WhseCodeTo_2
	, @ItemDim1Fr = @ItemDim1Fr_2
	, @ItemDim1To = @ItemDim1To_2
	, @ItemDim2Fr = @ItemDim2Fr_2
	, @ItemDim2To = @ItemDim2To_2
	, @ItemGrpNamFr = @ItemGrpNamFr_2
	, @ItemGrpNamTo = @ItemGrpNamTo_2
	, @AsAtDate = @AsAtDate_2
	, @Username = @Username_2
	, @ItemTypesFr = @ItemTypesFr_2
	, @ItemTypesTo = @ItemTypesTo_2