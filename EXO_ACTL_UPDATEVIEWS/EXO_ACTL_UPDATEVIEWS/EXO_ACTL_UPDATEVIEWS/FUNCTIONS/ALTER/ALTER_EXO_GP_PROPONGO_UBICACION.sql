ALTER FUNCTION "BBDD"."EXO_GP_PROPONGO_UBICACION"
(
 	IN pAlmacen NVARCHAR(8),
    IN pArticulo NVARCHAR(50),
	IN pLote NVARCHAR(36),
	IN pCantidad DECIMAL(19,6),
	IN pTipoOperacion NVARCHAR(1)

) RETURNS pResultado NVARCHAR(200)
LANGUAGE SQLSCRIPT
AS
	-- Declare the return variable here
	vConLote NVARCHAR(1);
	vResultado NVARCHAR(200);
BEGIN

 vResultado := '';
 pResultado := '';



 IF :pTipoOperacion = 'V' THEN
 
 	SELECT tart."ManBtchNum" INTO vConLote FROM "OITM" TArt WHERE TArt."ItemCode" = :pArticulo;
 	
 	IF :vConLote = 'N' THEN
 	
		DECLARE CURSOR c_1 FOR
		SELECT  COALESCE(T4."BinCode",'') BinCode
 		FROM "OITM" T2 
		INNER JOIN "OIBQ" T3 ON t3."ItemCode"=T2."ItemCode"
		INNER  JOIN "OBIN" T4  ON  T3."BinAbs" = T4."AbsEntry" AND T4."Attr2Val"='Picking'
		WHERE T2."ItemCode"=:pArticulo AND t4."WhsCode"=:pAlmacen AND T3."OnHandQty" > 0 
		ORDER BY T3."OnHandQty"
		LIMIT 1; 
		
	 	FOR c_row_1 AS c_1 DO
			vResultado := c_row_1.BinCode;
		END FOR;	
 	
 	ELSE
 	
 		DECLARE CURSOR c_1 FOR
		SELECT COALESCE(T6."BinCode",'') BinCode
 		FROM "OBTN" T4
		INNER JOIN "OBBQ" T5 ON T4."AbsEntry"=T5."SnBMDAbs"
		INNER JOIN "OBIN" T6 ON T5."BinAbs"=T6."AbsEntry"	AND T6."Attr2Val"='Picking'					 
		WHERE "DistNumber" = :pLote AND T4."ItemCode"=:pArticulo AND T6."WhsCode"=:pAlmacen 
		
		LIMIT 1; 
		
	 	FOR c_row_1 AS c_1 DO
			vResultado := c_row_1.BinCode;
		END FOR;	
 	
 	END IF;
 	
 ELSE

		DECLARE CURSOR c_1 FOR
		SELECT MIN( T4."BinCode") BinCode
  	FROM "OITM" T2 
    INNER JOIN "OIBQ" T3 ON T3."ItemCode"=T2."ItemCode" AND  T2."ManBtchNum"='N' 
    INNER JOIN "OBIN" T4  ON  T3."BinAbs" = T4."AbsEntry" AND T4."Attr2Val"='Picking'
	WHERE T2."ItemCode"=:pArticulo AND T4."WhsCode"=:pAlmacen 
	LIMIT 1;
		
	 	FOR c_row_1 AS c_1 DO
			vResultado := c_row_1.BinCode;
		END FOR;	
  
 END IF;

  
  IF :vResultado <> '' THEN
  	pResultado := :vResultado;
  END IF;
     

END;
