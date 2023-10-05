CREATE PROCEDURE "EXO_RECAL_CONTROL_STOCK" (IN pUser NVARCHAR(50)) 
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER AS
--Creamos variables
	vCount INTEGER;
	vPICKING NVARCHAR(50);
	vDOCUMENTO NVARCHAR(50);
	vOBJTYPE NVARCHAR(50);
	vLINEA INTEGER;
	vITEMCODE NVARCHAR(50);
	vPICKINGCANT DECIMAL(21, 6);
	vSTOCK DECIMAL(21, 6);
	vESTADO NVARCHAR(50);
	vITEMCODEACU NVARCHAR(50);
	vPACU DECIMAL(21, 6);
BEGIN
	
	
	
	--Creamos tabla temporal
	CREATE LOCAL TEMPORARY TABLE #tmp (
            PICKING NVARCHAR(50),
            DOCUMENTO NVARCHAR(50),
            OBJTYPE NVARCHAR(50),
			LINEA INTEGER,
			ITEMCODE NVARCHAR(50),
			PICKINGCANT DECIMAL(21, 6),
			STOCK DECIMAL(21, 6),
			PICKINGACU DECIMAL(21, 6),
			ESTADO NVARCHAR(50)
      );
     --Borramos todos los registros que pudiera tener en memoria 
	 TRUNCATE TABLE #tmp;
	 SELECT COALESCE(COUNT(*), 0) INTO vCount from (
							SELECT 'P' as "TIPO",T0."AbsEntry" "ABSENTRY"
							FROM "OPKL"  T0 
							INNER JOIN "PKL1" T1 ON T0."AbsEntry" = T1."AbsEntry"
							INNER JOIN "RDR1" T2 On T1."OrderEntry"=T2."DocEntry" and T1."OrderLine"=T2."LineNum"
							INNER JOIN "OITW" T3 ON T3."ItemCode"=T2."ItemCode" and T3."WhsCode"=T2."WhsCode"
							INNER JOIN "OBIN" T4 ON T4."WhsCode"=T3."WhsCode" and T4."AbsEntry"=T3."DftBinAbs"
							LEFT JOIN "OIBQ" T5 ON  T5."BinAbs" = T4."AbsEntry" and T5."WhsCode"=T4."WhsCode" and T5."ItemCode"=T2."ItemCode"
							WHERE "Status" not in ('Y','C')  and "Canceled"='N' and COALESCE("U_EXO_PPIST",'N')='N'
							and 'Y' = COALESCE((SELECT MAX('Y') from "PKL1" AS T1 
												INNER JOIN "RDR1" T2 ON T1."OrderEntry"=T2."DocEntry" and T1."OrderLine"=T2."LineNum" and t1."BaseObject"=17 
												WHERE T0."AbsEntry"=T1."AbsEntry"   
												and t2."WhsCode" in  ( select "WhsCode" from "VEXO_USUARIO_ALMACENES" 
																		where "USER_CODE"= :pUser 
																	 )
										),'N') 
							
							union all
							SELECT 'T' as "TIPO",T0."AbsEntry" "ABSENTRY"	
							FROM "OPKL"  T0 
							INNER JOIN "PKL1" T1 ON T0."AbsEntry" = T1."AbsEntry"
							INNER JOIN "WTQ1" T2 On T1."OrderEntry"=T2."DocEntry" and T1."OrderLine"=T2."LineNum"
							INNER JOIN "OITW" T3 ON T3."ItemCode"=T2."ItemCode" and T3."WhsCode"=T2."WhsCode"
							INNER JOIN "OBIN" T4 ON T4."WhsCode"=T3."WhsCode" and T4."AbsEntry"=T3."DftBinAbs"
							LEFT JOIN "OIBQ" T5 ON  T5."BinAbs" = T4."AbsEntry" and T5."WhsCode"=T4."WhsCode" and T5."ItemCode"=T2."ItemCode"
							WHERE "Status" not in ('Y','C')  and "Canceled"='N' and COALESCE("U_EXO_PPIST",'N')='N'
							and 'Y' = COALESCE((SELECT MAX('Y') from "SBO_AUTOCRISTAL_PRUEBAS"."PKL1" AS T1 
												INNER JOIN "SBO_AUTOCRISTAL_PRUEBAS"."WTQ1" T2 ON T1."OrderEntry"=T2."DocEntry" and T1."OrderLine"=T2."LineNum" and t1."BaseObject"=1250000001
												WHERE T0."AbsEntry"=T1."AbsEntry"   
												and t2."FromWhsCod" in  ( select "WhsCode" from "VEXO_USUARIO_ALMACENES" 
																			where "USER_CODE"= :pUser 
																	    )
										),'N') 
						);
		
    IF :vCount > 0 THEN
    	DECLARE CURSOR c_EXO_CONTROL_STOCK FOR
    	SELECT * FROM 
    	(
    						SELECT 'P' as "TIPO",T0."AbsEntry" "ABSENTRY",T0."PickDate" "PICKDATE",T0."Remarks" "REMARKS",T2."ItemCode" "ITEMCODE",
							T2."Dscription" "DES",T2."WhsCode" "ALMACEN", T3."OnHand" "STOCK",
							T4."BinCode" "BINCODE", IFNULL(T5."OnHandQty",0) "SUBI",
							T1."RelQtty" "CANTIDAD",T1."OrderEntry" "ORDERENTRY",T1."OrderLine" "ORDERLINE", T1."BaseObject" "OBJTYPE"	
							FROM "OPKL"  T0 
							INNER JOIN "PKL1" T1 ON T0."AbsEntry" = T1."AbsEntry"
							INNER JOIN "RDR1" T2 On T1."OrderEntry"=T2."DocEntry" and T1."OrderLine"=T2."LineNum"
							INNER JOIN "OITW" T3 ON T3."ItemCode"=T2."ItemCode" and T3."WhsCode"=T2."WhsCode"
							INNER JOIN "OBIN" T4 ON T4."WhsCode"=T3."WhsCode" and T4."AbsEntry"=T3."DftBinAbs"
							LEFT JOIN "OIBQ" T5 ON  T5."BinAbs" = T4."AbsEntry" and T5."WhsCode"=T4."WhsCode" and T5."ItemCode"=T2."ItemCode"
							WHERE "Status" not in ('Y','C')  and "Canceled"='N' and COALESCE("U_EXO_PPIST",'N')='N'
							and 'Y' = COALESCE((SELECT MAX('Y') from "PKL1" AS T1 
												INNER JOIN "RDR1" T2 ON T1."OrderEntry"=T2."DocEntry" and T1."OrderLine"=T2."LineNum" and t1."BaseObject"=17 
												WHERE T0."AbsEntry"=T1."AbsEntry"   
												and t2."WhsCode" in  ( select "WhsCode" from "VEXO_USUARIO_ALMACENES" 
																		where "USER_CODE"= :pUser 
																	 )
										),'N') 
							
							union all
							SELECT 'T' as "TIPO",T0."AbsEntry" "ABSENTRY",T0."PickDate" "PICKDATE",T0."Remarks" "REMARKS",T2."ItemCode" "ITEMCODE",
							T2."Dscription" "DES",T2."WhsCode" "ALMACEN", T3."OnHand" "STOCK",
							T4."BinCode" "BINCODE", IFNULL(T5."OnHandQty",0) "SUBI",
							T1."RelQtty" "CANTIDAD",T1."OrderEntry" "ORDERENTRY",T1."OrderLine" "ORDERLINE", T1."BaseObject" "OBJTYPE"		
							FROM "OPKL"  T0 
							INNER JOIN "PKL1" T1 ON T0."AbsEntry" = T1."AbsEntry"
							INNER JOIN "WTQ1" T2 On T1."OrderEntry"=T2."DocEntry" and T1."OrderLine"=T2."LineNum"
							INNER JOIN "OITW" T3 ON T3."ItemCode"=T2."ItemCode" and T3."WhsCode"=T2."WhsCode"
							INNER JOIN "OBIN" T4 ON T4."WhsCode"=T3."WhsCode" and T4."AbsEntry"=T3."DftBinAbs"
							LEFT JOIN "OIBQ" T5 ON  T5."BinAbs" = T4."AbsEntry" and T5."WhsCode"=T4."WhsCode" and T5."ItemCode"=T2."ItemCode"
							WHERE "Status" not in ('Y','C')  and "Canceled"='N' and COALESCE("U_EXO_PPIST",'N')='N'
							and 'Y' = COALESCE((SELECT MAX('Y') from "SBO_AUTOCRISTAL_PRUEBAS"."PKL1" AS T1 
												INNER JOIN "SBO_AUTOCRISTAL_PRUEBAS"."WTQ1" T2 ON T1."OrderEntry"=T2."DocEntry" and T1."OrderLine"=T2."LineNum" and t1."BaseObject"=1250000001
												WHERE T0."AbsEntry"=T1."AbsEntry"   
												and t2."FromWhsCod" in  ( select "WhsCode" from "VEXO_USUARIO_ALMACENES" 
																			where "USER_CODE"= :pUser 
																	    )
										),'N') 
    	)T ORDER BY "ITEMCODE","ABSENTRY" ;
		vPACU:= 0;	
		vITEMCODEACU:= '';
			
		FOR c_row_CONTROL_STOCK AS c_EXO_CONTROL_STOCK DO	
			vPICKING := c_row_CONTROL_STOCK.ABSENTRY;
			vDOCUMENTO := c_row_CONTROL_STOCK.ORDERENTRY;
			vOBJTYPE := c_row_CONTROL_STOCK.OBJTYPE;
			vLINEA := c_row_CONTROL_STOCK.ORDERLINE;
			vITEMCODE := c_row_CONTROL_STOCK.ITEMCODE;
			vPICKINGCANT:= c_row_CONTROL_STOCK.CANTIDAD;
			vSTOCK := c_row_CONTROL_STOCK.SUBI;	
			
			if vITEMCODEACU<>vITEMCODE then
				vITEMCODEACU:= vITEMCODE;
				vPACU:= 0;	
			end if;
			
			
			if vPACU+ vPICKINGCANT<=vSTOCK then
				vESTADO := 'PICKING';
				vPACU := vPACU+ vPICKINGCANT;
			else
				vESTADO := 'STOCK';
			end if;
			
			INSERT INTO #tmp VALUES
                  (:vPICKING, :vDOCUMENTO, :vOBJTYPE, :vLINEA, :vITEMCODE, :vPICKINGCANT, :vSTOCK, :vPACU, :vESTADO );
			
		END FOR;
    END IF;

	--Con la select devolvemos lo que tengamos en la temporal y la borramos
	SELECT * FROM #tmp;     
    DROP TABLE #tmp;
	
END;