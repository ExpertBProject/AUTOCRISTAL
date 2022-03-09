CREATE PROCEDURE _98_EXO_CHK_OITB
(
               IN pObject_type NVARCHAR(30),
               IN pTransaction_type NCHAR(1),
               IN pList_of_cols_val_tab_del NVARCHAR(255),
               OUT pError INT,
               OUT pError_message NVARCHAR(200)
)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
       
BEGIN
-- Return values
  	declare vCode NVARCHAR(50);
  	declare vU_EXO_BBDD NVARCHAR(50);  
  	declare vHAYDATO INTEGER ; 
  	declare vDesFam NVARCHAR (100);
  	DECLARE CURSOR c_EMP FOR
	SELECT T0."Code" CODE, T0."U_EXO_BBDD" BBDD FROM "@EXO_OADMINTERL"  T0;
   	IF :pObject_type = '52' THEN
  		IF (:pTransaction_type = 'A' OR :pTransaction_type = 'U') THEN
    		DECLARE EXIT HANDLER FOR SQLEXCEPTION
       		BEGIN
          		IF ::SQL_ERROR_CODE <> 0 THEN
                             pError := 1;
                             pError_message := '(EXO) ' || ::SQL_ERROR_CODE || ' ' || ::SQL_ERROR_MESSAGE;
          		END IF;
           	END;
                               
           	FOR c_row_EMP AS c_EMP DO
            	vCode := c_row_EMP.CODE;  
                vU_EXO_BBDD := c_row_EMP.BBDD;
                SELECT COUNT(*) INTO  vHAYDATO FROM "OITB" T0 WHERE T0."ItmsGrpCod" = :pList_of_cols_val_tab_del;   
              	IF (:vHAYDATO) > 0 THEN
	              	SELECT T0."ItmsGrpNam" INTO  vDesFam FROM "OITB" T0 WHERE T0."ItmsGrpCod" = :pList_of_cols_val_tab_del;
	                SELECT  COUNT(*) INTO vHAYDATO FROM  "REPLICATE" WHERE  "DBNAMEORIG" = vCode AND  "DBNAMEDEST"= vU_EXO_BBDD AND "CODETABLE" =  :pList_of_cols_val_tab_del  
	                AND "TABLENAME"='OITB';
	                IF (:vHAYDATO) = '0' THEN
						INSERT INTO "REPLICATE" ("DBNAMEORIG","DBNAMEDEST","TABLECATEGORY","TABLENAME","CODETABLE","CODETABLE2","CODETABLE3", "CODETABLE4","DATEADD")
						VALUES (vCode,vU_EXO_BBDD,:pObject_type,'OITB',:pList_of_cols_val_tab_del,vDesFam,'','',TO_NVARCHAR(CURRENT_DATE, 'YYYY/MM/DD'));
					ELSE
					 	--DELETE
						DELETE FROM "REPLICATE" WHERE  "DBNAMEORIG" = vCode AND  "DBNAMEDEST"= vU_EXO_BBDD AND "CODETABLE" =  :pList_of_cols_val_tab_del  AND "TABLENAME"='OITB';
						-- INSERT
						INSERT INTO "REPLICATE" ("DBNAMEORIG","DBNAMEDEST","TABLECATEGORY","TABLENAME","CODETABLE","CODETABLE2","CODETABLE3", "CODETABLE4","DATEADD")
						VALUES (vCode,vU_EXO_BBDD,:pObject_type,'OITB',:pList_of_cols_val_tab_del,vDesFam,'','',TO_NVARCHAR(CURRENT_DATE, 'YYYY/MM/DD'));
					END IF;
				END IF;
			END FOR;
                       
      END IF;
   END IF;
END;
