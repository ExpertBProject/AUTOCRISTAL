CREATE PROCEDURE "BBDD"."EXO_GP_PROPONGO_LOTE"
(
	IN pAlmacen NVARCHAR(30),
	IN pArticulo NVARCHAR(50),
	IN pCantidadSolicitada DECIMAL(19,6),
	OUT oLote NVARCHAR(36),
  	OUT oCantidad DECIMAL(19,6),
  	OUT oUbicacion NVARCHAR(228)
)
LANGUAGE SQLSCRIPT
AS
-- Return values
BEGIN


	DECLARE CURSOR c_0 FOR
	 		select A0."DistNumber" DistNumber, A0."OnHandQty" OnHandQty, A0."BinCode"  BinCode
	 from
	(
	SELECT 'A' "Orden",T1."DistNumber" , T0."OnHandQty" , T2."BinCode" , T1."InDate",T1."SysNumber"
		FROM "OBBQ" T0
			INNER JOIN "OBTN" T1 ON T0."SnBMDAbs" = T1."AbsEntry" AND T0."ItemCode" = T1."ItemCode"
			INNER JOIN "OBIN" T2 ON T2."AbsEntry" = T0."BinAbs"   and IFNULL(T2."Attr2Val",'Stock') IN ('Picking','Stock')
			INNER JOIN "OITW" T3 ON T3."WhsCode"=T0."WhsCode" and t3."ItemCode"=T0."ItemCode" and T3."DftBinAbs"=T2."AbsEntry"
	WHERE T0."ItemCode" = :pArticulo AND T0."WhsCode" = :pAlmacen AND  T1."Status"=0 --COALESCE(T0."OnHandQty", 0) >= :pCantidadSolicitada AND
		AND  COALESCE(T0."OnHandQty", 0) > 0 
		AND
		CONCAT(CONCAT(T0."ItemCode", '#') , T1."DistNumber") NOT IN (SELECT CONCAT(CONCAT(TLote."ItemCode",'#'), TLote."BatchNum") FROM "OIBT" TLote WHERE TLote."ItemCode" = T0."ItemCode" AND COALESCE(TLote."IsCommited", 0) <> 0)

	UNION 
	 	SELECT 'B' "Orden",T1."DistNumber" , T0."OnHandQty" , T2."BinCode" , T1."InDate",T1."SysNumber"
		FROM "OBBQ" T0
			INNER JOIN "OBTN" T1 ON T0."SnBMDAbs" = T1."AbsEntry" AND T0."ItemCode" = T1."ItemCode"
			INNER JOIN "OBIN" T2 ON T2."AbsEntry" = T0."BinAbs"   and IFNULL(T2."Attr2Val",'Stock') IN ('Picking','Stock')
		WHERE T0."ItemCode" = :pArticulo AND T0."WhsCode" = :pAlmacen AND T1."Status"=0 -- AND COALESCE(T0."OnHandQty", 0) >= :pCantidadSolicitada
		AND  COALESCE(T0."OnHandQty", 0) > 0 AND
		CONCAT(CONCAT(T0."ItemCode", '#') , T1."DistNumber") NOT IN (SELECT CONCAT(CONCAT(TLote."ItemCode",'#'), TLote."BatchNum") FROM "OIBT" TLote WHERE TLote."ItemCode" = T0."ItemCode" AND COALESCE(TLote."IsCommited", 0) <> 0)
) A0 
ORDER BY  A0."Orden" ASC,  A0."InDate" ASC, A0."SysNumber" ASC,COALESCE(A0."OnHandQty", 0),A0."DistNumber" ASC
	LIMIT 1;
		
	oLote := '';
	oCantidad := 0;
	oUbicacion := '';
	
	FOR c_row_0 AS c_0 DO
		oLote := c_row_0.DistNumber;
		oCantidad := c_row_0.OnHandQty;
		oUbicacion := c_row_0.BinCode;
	END FOR;	
	
END;
