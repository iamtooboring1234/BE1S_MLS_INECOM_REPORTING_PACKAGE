CREATE PROCEDURE NCM_SP_CHG_LOG_AUDIT_COMBINE
(
	IN VUSERCODE VARCHAR(20),
	IN VDTFROM	 VARCHAR(20),
	IN VDTTO	 VARCHAR(20),
	IN VOBJTYPE VARCHAR(20),
	IN VTABLENM VARCHAR(20)
)
LANGUAGE SQLSCRIPT
AS
BEGIN

IF :VOBJTYPE = '2' or :VOBJTYPE = '4' or :VOBJTYPE = '30' THEN
	IF :VOBJTYPE = '2' THEN
		Select T1."CardCode" as "No", "NCM_GETOBJTYPENAME"(T1."ObjType") as "ObjType",T1."LogInstanc" as "Instance", 
		T2."USER_CODE" as "UserCode", "U_NAME" as "UserName", 
		T1."UpdateDate" as "ActDate", IFNULL(T1."U_UpdateTS", T1."U_CreateTS") as "ActTime", 
		CASE
	    --WHEN T1.CreateDate = T1.UpdateDate THEN 'Create ' + dbo.NCM_GetObjTypeName(T1.ObjType) + ' ' + CAST(T1.DocNum as nvarchar)
	    WHEN T1."LogInstanc" = 1 THEN 'Create ' || "NCM_GETOBJTYPENAME"(T1."ObjType") || ' ' || CAST(T1."CardCode" as nvarchar)
	    ELSE 'Update ' || "NCM_GETOBJTYPENAME"(T1."ObjType") || ' ' || CAST(T1."CardCode" as nvarchar)
		END AS "ActDesc"
		from ACRD T1
		--Inner Join OUSR T2 on CASE WHEN T1."UserSign2" is null THEN T1."UserSign" ELSE T1."UserSign2" = T2."USERID"
		Inner Join OUSR T2 on IFNULL("UserSign2", T1."UserSign") = T2."USERID"
		WHERE T1."ObjType" = :VOBJTYPE 
		AND ("USER_CODE" = IFNULL(NULLIF(:VUSERCODE,'ALL'), "USER_CODE") OR "USER_CODE" is null)
		AND T1."UpdateDate" >= CAST(:VDTFROM as Date) AND T1."UpdateDate" <= CAST(:VDTTO as Date);
	END IF;
	IF :VOBJTYPE = '4' THEN
		Select T1."ItemCode" as "No", "NCM_GETOBJTYPENAME"(T1."ObjType") as "ObjType",T1."LogInstanc" as "Instance", 
		T2."USER_CODE" as "UserCode", "U_NAME" as "UserName", 
		T1."UpdateDate" as "ActDate", IFNULL(T1."U_UpdateTS", T1."U_CreateTS") as "ActTime", 
		CASE
	    WHEN T1."LogInstanc" = 1 THEN 'Create ' || "NCM_GETOBJTYPENAME"(T1."ObjType") || ' ' || CAST(T1."ItemCode" as nvarchar)
	    ELSE 'Update ' || "NCM_GETOBJTYPENAME"(T1."ObjType") || ' ' || CAST(T1."ItemCode" as nvarchar)
		END AS "ActDesc"
		from AITM T1
		Inner Join OUSR T2 on IFNULL("UserSign2", T1."UserSign") = T2."USERID"
		WHERE T1."ObjType" = :VOBJTYPE 
		AND ("USER_CODE" = IFNULL(NULLIF(:VUSERCODE,'ALL'), "USER_CODE") OR "USER_CODE" is null)
		AND T1."UpdateDate" >= CAST(:VDTFROM as Date) AND T1."UpdateDate" <= CAST(:VDTTO as Date);
	END IF;
	IF :VOBJTYPE = '30' THEN
		Select T1."TransId" as "No", "NCM_GETOBJTYPENAME"(T1."ObjType") as "ObjType",T1."LogInstanc" as "Instance", 
		T2."USER_CODE" as "UserCode", "U_NAME" as "UserName", 
		T1."UpdateDate" as "ActDate", IFNULL(T1."U_UpdateTS", T1."U_CreateTS") as "ActTime", 
		CASE
	    WHEN T1."LogInstanc" = 1 THEN 'Create ' || "NCM_GETOBJTYPENAME"(T1."ObjType") || ' ' || CAST(T1."Number" as nvarchar)
	    ELSE 'Update ' || "NCM_GETOBJTYPENAME"(T1."ObjType") || ' ' || CAST(T1."Number" as nvarchar)
		END AS "ActDesc"
		from AJDT T1
		Inner Join OUSR T2 on IFNULL("UserSign2", T1."UserSign") = T2."USERID"
		WHERE T1."TransType" = 30 AND T1."ObjType" = :VOBJTYPE 
		AND ("USER_CODE" = IFNULL(NULLIF(:VUSERCODE,'ALL'), "USER_CODE") OR "USER_CODE" is null)
		AND T1."UpdateDate" >= CAST(:VDTFROM as Date) AND T1."UpdateDate" <= CAST(:VDTTO as Date);
	END IF;
ELSE
	Select T1."DocEntry" as "No", "NCM_GETOBJTYPENAME"(T1."ObjType") as "ObjType",T1."LogInstanc" as "Instance", 
	T2."USER_CODE" as "UserCode", "U_NAME" as "UserName", 
	T1."UpdateDate" as "ActDate", T1."UpdateTS" as "ActTime", 
	CASE
    --WHEN T1.CreateDate = T1.UpdateDate THEN 'Create ' + dbo.NCM_GetObjTypeName(T1.ObjType) + ' ' + CAST(T1.DocNum as nvarchar)
    WHEN T1."LogInstanc" = 1 THEN 'Create ' || "NCM_GETOBJTYPENAME"(T1."ObjType") || ' ' || CAST(T1."DocNum" as nvarchar)
    ELSE 'Update ' || "NCM_GETOBJTYPENAME"(T1."ObjType") || ' ' || CAST(T1."DocNum" as nvarchar)
	END AS "ActDesc"
	from ADOC T1
	Inner Join OUSR T2 on T1."UserSign2" = T2."USERID"
	WHERE T1."ObjType" = :VOBJTYPE 
	AND ("USER_CODE" = IFNULL(NULLIF(:VUSERCODE,'ALL'), "USER_CODE") OR "USER_CODE" is null)
	AND T1."UpdateDate" >= CAST(:VDTFROM as Date) AND T1."UpdateDate" <= CAST(:VDTTO as Date);	
END IF;
END;