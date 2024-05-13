CREATE PROCEDURE NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ
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
	IF(:VUSERCODE = 'ALL') THEN
		IF :VOBJTYPE = '2' THEN
			EXECUTE IMMEDIATE 
		   'Select T3."CardCode" as "No", "NCM_GETOBJTYPENAME"(T3."ObjType") as "ObjType", 
		   (Select IFNULL(MAX("LogInstanc"),0) + 1 from ACRD WHERE "ObjType" = T3."ObjType" AND "CardCode" = T3."CardCode") as "Instance",
			T2."USER_CODE"  as "UserCode", "U_NAME" as "UserName", IFNULL(T3."UpdateDate",T3."CreateDate") AS "ActDate" 
			, CASE WHEN (Select COUNT(*) from ACRD WHERE "ObjType" = T3."ObjType" AND "CardCode" = T3."CardCode") = 0 
			THEN T3."U_CreateTS" ELSE T3."U_UpdateTS" END AS "ActTime"
			, CASE WHEN (Select COUNT(*) from ACRD WHERE "ObjType" = T3."ObjType" AND "CardCode" = T3."CardCode") = 0 
			THEN ''Create '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."CardCode" as nvarchar)
			ELSE ''Update '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."CardCode" as nvarchar) END AS "ActDesc"
			from  ' || :VTABLENM || ' T3
			Inner Join OUSR T2 on IFNULL(T3."UserSign2",T3."UserSign") = T2."USERID"
			WHERE T3."UpdateDate" >= CAST(''' || :VDTFROM || ''' as datetime) AND T3."UpdateDate" <=CAST(''' || :VDTTO || ''' as datetime)';
		END IF;
		
		IF :VOBJTYPE = '4' THEN
			EXECUTE IMMEDIATE 
		   'Select T3."ItemCode" as "No", "NCM_GETOBJTYPENAME"(T3."ObjType") as "ObjType", 
		   (Select IFNULL(MAX("LogInstanc"),0) + 1 from AITM WHERE "ObjType" = T3."ObjType" AND "ItemCode" = T3."ItemCode") as "Instance",
			T2."USER_CODE"  as "UserCode", "U_NAME" as "UserName", IFNULL(T3."UpdateDate",T3."CreateDate") AS "ActDate" 
			, CASE WHEN (Select COUNT(*) from AITM WHERE "ObjType" = T3."ObjType" AND "ItemCode" = T3."ItemCode") = 0 
			THEN T3."U_CreateTS" ELSE T3."U_UpdateTS" END AS "ActTime"
			, CASE WHEN (Select COUNT(*) from AITM WHERE "ObjType" = T3."ObjType" AND "ItemCode" = T3."ItemCode") = 0 
			THEN ''Create '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."ItemCode" as nvarchar)
			ELSE ''Update '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."ItemCode" as nvarchar) END AS "ActDesc"
			from  ' || :VTABLENM || ' T3
			Inner Join OUSR T2 on IFNULL(T3."UserSign2",T3."UserSign") = T2."USERID"
			WHERE T3."UpdateDate" >= CAST(''' || :VDTFROM || ''' as datetime) AND T3."UpdateDate" <=CAST(''' || :VDTTO || ''' as datetime)';
		END IF;
		
		IF :VOBJTYPE = '30' THEN
			EXECUTE IMMEDIATE 
		   'Select T3."TransId" as "No", "NCM_GETOBJTYPENAME"(T3."ObjType") as "ObjType", 
		   (Select IFNULL(MAX("LogInstanc"),0) + 1 from AJDT WHERE "ObjType" = T3."ObjType" AND "TransId" = T3."TransId") as "Instance",
			T2."USER_CODE"  as "UserCode", "U_NAME" as "UserName", IFNULL(T3."UpdateDate",T3."CreateDate") AS "ActDate" 
			, CASE WHEN (Select COUNT(*) from AJDT WHERE "ObjType" = T3."ObjType" AND "TransId" = T3."TransId") = 0 
			THEN T3."U_CreateTS" ELSE T3."U_UpdateTS" END AS "ActTime"
			, CASE WHEN (Select COUNT(*) from AJDT WHERE "ObjType" = T3."ObjType" AND "TransId" = T3."TransId") = 0 
			THEN ''Create '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."Number" as nvarchar)
			ELSE ''Update '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."Number" as nvarchar) END AS "ActDesc"
			from  ' || :VTABLENM || ' T3
			Inner Join OUSR T2 on IFNULL(T3."UserSign2",T3."UserSign") = T2."USERID"
			WHERE T3."UpdateDate" >= CAST(''' || :VDTFROM || ''' as datetime) AND T3."UpdateDate" <=CAST(''' || :VDTTO || ''' as datetime)';
		END IF;
	Else
		IF :VOBJTYPE = '2' THEN
			EXECUTE IMMEDIATE 
		   'Select T3."CardCode" as "No", "NCM_GETOBJTYPENAME"(T3."ObjType") as "ObjType", 
		   (Select IFNULL(MAX("LogInstanc"),0) + 1 from ACRD WHERE "ObjType" = T3."ObjType" AND "CardCode" = T3."CardCode") as "Instance",
			T2."USER_CODE"  as "UserCode", "U_NAME" as "UserName", IFNULL(T3."UpdateDate",T3."CreateDate") AS "ActDate" 
			, CASE WHEN (Select COUNT(*) from ACRD WHERE "ObjType" = T3."ObjType" AND "CardCode" = T3."CardCode") = 0 
			THEN T3."U_CreateTS" ELSE T3."U_UpdateTS" END AS "ActTime"
			, CASE WHEN (Select COUNT(*) from ACRD WHERE "ObjType" = T3."ObjType" AND "CardCode" = T3."CardCode") = 0 
			THEN ''Create '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."CardCode" as nvarchar)
			ELSE ''Update '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."CardCode" as nvarchar) END AS "ActDesc"
			from  ' || :VTABLENM || ' T3
			Inner Join OUSR T2 on IFNULL(T3."UserSign2",T3."UserSign") = T2."USERID" 
			WHERE "USER_CODE" = ''' || :VUSERCODE || ''' AND T3."UpdateDate" >= CAST(''' || :VDTFROM || ''' as datetime) AND T3."UpdateDate" <=CAST(''' || :VDTTO || ''' as datetime)';
		END IF;
		
		IF :VOBJTYPE = '4' THEN
			EXECUTE IMMEDIATE 
		   'Select T3."ItemCode" as "No", "NCM_GETOBJTYPENAME"(T3."ObjType") as "ObjType", 
		   (Select IFNULL(MAX("LogInstanc"),0) + 1 from AITM WHERE "ObjType" = T3."ObjType" AND "ItemCode" = T3."ItemCode") as "Instance",
			T2."USER_CODE"  as "UserCode", "U_NAME" as "UserName", IFNULL(T3."UpdateDate",T3."CreateDate") AS "ActDate" 
			, CASE WHEN (Select COUNT(*) from AITM WHERE "ObjType" = T3."ObjType" AND "ItemCode" = T3."ItemCode") = 0 
			THEN T3."U_CreateTS" ELSE T3."U_UpdateTS" END AS "ActTime"
			, CASE WHEN (Select COUNT(*) from AITM WHERE "ObjType" = T3."ObjType" AND "ItemCode" = T3."ItemCode") = 0 
			THEN ''Create '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."ItemCode" as nvarchar)
			ELSE ''Update '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."ItemCode" as nvarchar) END AS "ActDesc"
			from  ' || :VTABLENM || ' T3
			Inner Join OUSR T2 on IFNULL(T3."UserSign2",T3."UserSign") = T2."USERID" 
			WHERE "USER_CODE" = ''' || :VUSERCODE || ''' AND T3."UpdateDate" >= CAST(''' || :VDTFROM || ''' as datetime) AND T3."UpdateDate" <=CAST(''' || :VDTTO || ''' as datetime)';
		END IF;
		
		IF :VOBJTYPE = '30' THEN
			EXECUTE IMMEDIATE 
		   'Select T3."TransId" as "No", "NCM_GETOBJTYPENAME"(T3."ObjType") as "ObjType", 
		   (Select IFNULL(MAX("LogInstanc"),0) + 1 from AJDT WHERE "ObjType" = T3."ObjType" AND "TransId" = T3."TransId") as "Instance",
			T2."USER_CODE"  as "UserCode", "U_NAME" as "UserName", IFNULL(T3."UpdateDate",T3."CreateDate") AS "ActDate" 
			, CASE WHEN (Select COUNT(*) from AJDT WHERE "ObjType" = T3."ObjType" AND "TransId" = T3."TransId") = 0 
			THEN T3."U_CreateTS" ELSE T3."U_UpdateTS" END AS "ActTime"
			, CASE WHEN (Select COUNT(*) from AJDT WHERE "ObjType" = T3."ObjType" AND "TransId" = T3."TransId") = 0 
			THEN ''Create '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."Number" as nvarchar)
			ELSE ''Update '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."Number" as nvarchar) END AS "ActDesc"
			from  ' || :VTABLENM || ' T3
			Inner Join OUSR T2 on IFNULL(T3."UserSign2",T3."UserSign") = T2."USERID"
			WHERE "USER_CODE" = ''' || :VUSERCODE || ''' AND T3."UpdateDate" >= CAST(''' || :VDTFROM || ''' as datetime) AND T3."UpdateDate" <=CAST(''' || :VDTTO || ''' as datetime)';
		END IF;
	END IF;
ELSE --Marking Document
	IF(:VUSERCODE = 'ALL') Then
	    	EXECUTE IMMEDIATE 
		   'Select T3."DocEntry" as "No", "NCM_GETOBJTYPENAME"(T3."ObjType") as "ObjType", 
		   (Select IFNULL(MAX("LogInstanc"),0) + 1 from ADOC WHERE "ObjType" = T3."ObjType" AND "DocEntry" = T3."DocEntry") as "Instance",
			T2."USER_CODE"  as "UserCode", "U_NAME" as "UserName", T3."UpdateDate" as "ActDate", T3."UpdateTS" as "ActTime"
			, CASE WHEN (Select COUNT(*) from ADOC WHERE "ObjType" = T3."ObjType" AND "DocEntry" = T3."DocEntry") = 0 
			THEN ''Create '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."DocNum" as nvarchar)
			ELSE ''Update '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."DocNum" as nvarchar) END AS "ActDesc"
			from  ' || :VTABLENM || ' T3
			Inner Join OUSR T2 on T3."UserSign2" = T2."USERID"
			WHERE T3."UpdateDate" >= CAST(''' || :VDTFROM || ''' as datetime) AND T3."UpdateDate" <=CAST(''' || :VDTTO || ''' as datetime)';
	
	Else
			EXECUTE IMMEDIATE 
		   'Select T3."DocEntry" as "No", "NCM_GETOBJTYPENAME"(T3."ObjType") as "ObjType", 
		   (Select IFNULL(MAX("LogInstanc"),0) + 1 from ADOC WHERE "ObjType" = T3."ObjType" AND "DocEntry" = T3."DocEntry") as "Instance",
			T2."USER_CODE"  as "UserCode", "U_NAME" as "UserName", T3."UpdateDate" as "ActDate", T3."UpdateTS" as "ActTime"
			, CASE WHEN (Select COUNT(*) from ADOC WHERE "ObjType" = T3."ObjType" AND "DocEntry" = T3."DocEntry") = 0 
			THEN ''Create '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."DocNum" as nvarchar)
			ELSE ''Update '' || "NCM_GETOBJTYPENAME"(T3."ObjType") || '' '' || CAST(T3."DocNum" as nvarchar) END AS "ActDesc"
			from  ' || :VTABLENM || ' T3
			Inner Join OUSR T2 on T3."UserSign2" = T2."USERID" 
			WHERE "USER_CODE" = ''' || :VUSERCODE || ''' AND T3."UpdateDate" >= CAST(''' || :VDTFROM || ''' as datetime) AND T3."UpdateDate" <=CAST(''' || :VDTTO || ''' as datetime)';
	END IF;
END IF;
END;