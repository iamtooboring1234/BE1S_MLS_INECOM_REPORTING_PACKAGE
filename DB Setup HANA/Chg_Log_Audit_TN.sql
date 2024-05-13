-- B1 DEPENDS: BEFORE:PT:PROCESS_START

ALTER PROCEDURE SBO_SP_TransactionNotification
(
	in object_type nvarchar(20), 				-- SBO Object Type
	in transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
	in num_of_cols_in_key int,
	in list_of_key_cols_tab_del nvarchar(255),
	in list_of_cols_val_tab_del nvarchar(255)
)
LANGUAGE SQLSCRIPT
AS
-- Return values
error  int;				-- Result (0 for no error)
error_message nvarchar (200); 		-- Error string to be displayed
begin

error := 0;
error_message := N'Ok';

--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE
		
--for change log audit report
	--Item Master Data
	IF :object_type ='4'  THEN
	
		/*			
		UPDATE OITM
		SET [U_UpdateTS] = CONVERT(Char, @Time,100) 
		WHERE ItemCode = @List_Of_Cols_Val_Tab_Del 	
		*/
		UPDATE OITM
		SET "U_UpdateTS" = ((CAST(LEFT(CAST(CURRENT_TIME as nvarchar),2) as int) * 100) + CAST(Right(LEFT(CAST(CURRENT_TIME as nvarchar),5),2) as int))
		WHERE "ItemCode" = :list_of_cols_val_tab_del 
			AND :transaction_type IN ('U');
			
		UPDATE OITM
		SET "U_CreateTS" = ((CAST(LEFT(CAST(CURRENT_TIME as nvarchar),2) as int) * 100) + CAST(Right(LEFT(CAST(CURRENT_TIME as nvarchar),5),2) as int))
		WHERE "ItemCode" = :list_of_cols_val_tab_del 
			AND :transaction_type IN ('A');
	END IF;
	
	
--for change log audit report
	--BP Master Data
	IF :object_type ='2' THEN

		UPDATE OCRD
		SET "U_UpdateTS" = ((CAST(LEFT(CAST(CURRENT_TIME as nvarchar),2) as int) * 100) + CAST(Right(LEFT(CAST(CURRENT_TIME as nvarchar),5),2) as int))
		WHERE "CardCode" = :list_of_cols_val_tab_del 
			AND :transaction_type IN ('U');
			
		UPDATE OCRD
		SET "U_CreateTS" = ((CAST(LEFT(CAST(CURRENT_TIME as nvarchar),2) as int) * 100) + CAST(Right(LEFT(CAST(CURRENT_TIME as nvarchar),5),2) as int))
		WHERE "CardCode" = :list_of_cols_val_tab_del 
			AND :transaction_type IN ('A');
	END IF;

--for change log audit report
	--Journal Entry
	IF :object_type ='30' THEN
		
		UPDATE OJDT
		SET "U_UpdateTS" = ((CAST(LEFT(CAST(CURRENT_TIME as nvarchar),2) as int) * 100) + CAST(Right(LEFT(CAST(CURRENT_TIME as nvarchar),5),2) as int))
		WHERE "TransId" = :list_of_cols_val_tab_del 
			AND :transaction_type IN ('U');
			
		UPDATE OJDT
		SET "U_CreateTS" = ((CAST(LEFT(CAST(CURRENT_TIME as nvarchar),2) as int) * 100) + CAST(Right(LEFT(CAST(CURRENT_TIME as nvarchar),5),2) as int))
		WHERE "TransId" = :list_of_cols_val_tab_del 
			AND :transaction_type IN ('A');

	END IF;	
--------------------------------------------------------------------------------------------------------------------------------

-- Select the return values
select :error, :error_message FROM dummy;

end;
