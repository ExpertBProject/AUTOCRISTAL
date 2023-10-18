ALTER VIEW "BBDD"."EXO_SOL_DEVOLUCION" ( "T. SALIDA",
	 "DELEGACIÓN",
	 "FECHA CREACION",
	 "FECHA ENTREGA",
	 "Nº INTERNO",
	 "Nº DOCUMENTO",
	 "AUTORIZADO",
	 "COM",
	 "CÓDIGO",
	 "EMPRESA",
	 "CLASE EXP.",
	 "ROT. STOCK",
	 "WhsCode",
	 "Territory",
	 "A",
	 "UBICACIÓN",
	 "ZONA TRANSPORTE",
	 "Sel" ) AS SELECT
	 DISTINCT CAST('SDPROV' as nVARCHAR(50)) "T. SALIDA",
	 CAST(IFNULL(T2."Name",
	 ' ') as nVARCHAR(50)) "DELEGACIÓN",
	 T0."DocDate",
	 T0."DocDueDate",
	 CAST(T0."DocEntry" as nVARCHAR(50)) "Nº INTERNO",
	 CAST(T0."DocNum" as nVARCHAR(50)) "Nº DOCUMENTO",
	 T0."Confirmed" "AUTORIZADO",
	 (CASE WHEN IFNULL(T0."Comments",
	 '') = '' 
	THEN 'N' 
	ELSE 'Y' 
	END) AS "COM",
	 CAST(T0."CardCode" as nVARCHAR(50)) "CÓDIGO",
	 CAST(T0."CardName" as nVARCHAR(150)) "EMPRESA",
	 CAST(T0."TrnspCode" as nVARCHAR(50)) "CLASE EXP.",
	 ifnull(R."ROTURA",
	 'N') "ROT. STOCK",
	 TL."WhsCode",
	 T1."Territory",
	 IFNULL(A."A",
	 'N') "A",
	 CAST(IFNULL((SELECT
	 CASE WHEN COUNT("Situacion")=1 
			then max("Situacion") 
			ELSE 'Ambos' 
			END "Sit" 
			FROM ( SELECT
	 X1."DocEntry",
	 X0."ObjType",
	 X0."DocNum",
	 IFNULL(OBIN."Attr1Val",
	 '') "Situacion" 
				FROM PRR1 X1 
				INNER JOIN OPRR X0 ON X0."DocEntry"=X1."DocEntry" 
				INNER JOIN OITW AL ON AL."WhsCode"=X1."WhsCode" 
				and AL."ItemCode"=X1."ItemCode" 
				LEFT JOIN OBIN ON OBIN."AbsEntry"=AL."DftBinAbs" 
				Group BY X1."DocEntry",
	 X0."ObjType",
	 X0."DocNum",
	 IFNULL(OBIN."Attr1Val",
	 '') )T 
			WHERE T."DocEntry" = T0."DocEntry" 
			Group BY T."DocEntry",
	 T."ObjType",
	 T."DocNum"),
	 'SIN SITUACIÓN') as nVARCHAR(50)) "UBICACIÓN",
	 CAST(TT."descript" as nVARCHAR(50)) "ZONA TRANSPORTE",
	 'N' "Sel" 
FROM OPRR T0 
LEFT JOIN PRR1 TL ON TL."DocEntry"=T0."DocEntry" 
LEFT JOIN OCRD T1 ON T0."CardCode"=T1."CardCode" 
LEFT JOIN OUBR T2 ON T1."U_EXO_DELE"=T2."Code" 
LEFT JOIN "EXO_ROTURA" R ON R."DocEntry"=T0."DocEntry" 
and R."ObjType"=T0."ObjType" 
LEFT JOIN "EXO_SITUACION" S ON S."DocEntry"=T0."DocEntry" 
and S."ObjType"=T0."ObjType" 
LEFT JOIN "EXO_A" A ON A."CardCode"=T0."CardCode" 
and A."WhsCode"=TL."WhsCode" 
LEFT JOIN OTER TT ON T1."Territory"=TT."territryID" JOIN OITM ON TL."ItemCode" = OITM."ItemCode" 
WHERE TL."LineStatus"='O' 
and T0."Confirmed"='Y' 
AND OITM."InvntItem" = 'Y' 
and T0."U_EXO_STATUSP"='P' WITH READ ONLY
