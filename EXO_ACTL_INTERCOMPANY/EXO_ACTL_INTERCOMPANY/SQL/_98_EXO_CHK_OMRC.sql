CREATE PROCEDURE _98_EXO_CHK_OMRC
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
  	declare vDes NVARCHAR (100);
  	DECLARE CURSOR c_EMP FOR
	SELECT T0."Code" CODE, T0."U_EXO_BBDD" BBDD FROM "@EXO_OADMINTERL"  T0;
   	IF :pObject_type = '43' THEN
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
              	SELECT T0."FirmName" INTO  vDes FROM "OMRC" T0 WHERE T0."FirmCode" = :pList_of_cols_val_tab_del;
                SELECT  COUNT(*) INTO vHAYDATO FROM  "REPLICATE" WHERE  "DBNAMEORIG" = vCode AND  "DBNAMEDEST"= vU_EXO_BBDD AND "CODETABLE" =  :pList_of_cols_val_tab_del  
                AND "TABLENAME"='OMRC';
                IF (:vHAYDATO) = '0' THEN
					INSERT INTO "REPLICATE" ("DBNAMEORIG","DBNAMEDEST","TABLECATEGORY","TABLENAME","CODETABLE","CODETABLE2","CODETABLE3", "CODETABLE4","DATEADD")
					VALUES (vCode,vU_EXO_BBDD,:pObject_type,'OMRC',:pList_of_cols_val_tab_del,vDes,'','',TO_NVARCHAR(CURRENT_DATE, 'YYYY/MM/DD'));
				ELSE
				 	--DELETE
					DELETE FROM "REPLICATE" WHERE  "DBNAMEORIG" = vCode AND  "DBNAMEDEST"= vU_EXO_BBDD AND "CODETABLE" =  :pList_of_cols_val_tab_del  AND "TABLENAME"='OMRC';
					-- INSERT
					INSERT INTO "REPLICATE" ("DBNAMEORIG","DBNAMEDEST","TABLECATEGORY","TABLENAME","CODETABLE","CODETABLE2","CODETABLE3", "CODETABLE4","DATEADD")
					VALUES (vCode,vU_EXO_BBDD,:pObject_type,'OMRC',:pList_of_cols_val_tab_del,vDes,'','',TO_NVARCHAR(CURRENT_DATE, 'YYYY/MM/DD'));
				END IF;
			END FOR;
                       
      END IF;
   END IF;
END;
