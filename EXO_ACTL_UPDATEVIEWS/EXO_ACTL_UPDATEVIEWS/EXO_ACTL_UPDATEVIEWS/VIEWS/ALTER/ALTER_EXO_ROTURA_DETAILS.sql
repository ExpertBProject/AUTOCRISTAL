ALTER VIEW "BBDD"."EXO_ROTURA_DETAILS" ( "DocEntry",
	 "DocNum",
	 "LineNum",
	 "ObjType",
	 "ItemCode",
	 "ALMACEN",
	 "OpenQty",
	 "U_EXO_DOCPRI",
	 "OnHand",
	 "CreateDate",
	 "CreateTS",
	 "ACUMULADO",
	 "ROTURA" ) AS SELECT
	 "Z"."DocEntry" ,
	 "Z"."DocNum" ,
	 "Z"."LineNum" ,
	 "Z"."ObjType" ,
	 "Z"."ItemCode" ,
	 "Z"."ALMACEN" ,
	 "Z"."OpenQty" ,
	 "Z"."U_EXO_DOCPRI" ,
	 OITW."OnHand",
	 "Z"."CreateDate" ,
	 "Z"."CreateTS" ,
	 "Z"."ACUMULADO" ,
	 "Z"."ROTURA" 
FROM ( SELECT
	 T."DocEntry",
	 T."DocNum",
	 T."LineNum",
	 t."ObjType",
	 T."ItemCode",
	 "ALMACEN",
	 "OpenQty",
	 "U_EXO_DOCPRI",
	 "CreateDate",
	 "CreateTS",
	 sum("OpenQty") over (partition By T."ItemCode",
	 "ALMACEN" 
		ORDER BY T."ItemCode" ASC,
	 "ALMACEN" ASC,
	 "U_EXO_DOCPRI" DESC,
	 "CreateDate" ASC,
	 "CreateTS" ASC) "ACUMULADO" ,
	 case when (sum("OpenQty") over (partition By T."ItemCode",
	 "ALMACEN" 
			ORDER BY T."ItemCode" ASC,
	 "ALMACEN" ASC,
	 "U_EXO_DOCPRI" DESC,
	 "CreateDate" ASC,
	 "CreateTS" ASC)) > AL."OnHand" 
	then 'Y' 
	else 'N' 
	end "ROTURA" 
	FROM ( select
	 T1."DocEntry",
	 T0."DocNum",
	 T1."LineNum",
	 T0."ObjType",
	 T1."ItemCode",
	 T1."OpenQty",
	 T0."U_EXO_DOCPRI",
	 T0."CreateDate",
	 T0."CreateTS",
	 T1."FromWhsCod" "ALMACEN" 
		FROM WTQ1 T1 
		INNER JOIN OWTQ T0 ON T0."DocEntry"=T1. "DocEntry" 
		INNER JOIN OITM T2 ON T1."ItemCode" = T2."ItemCode" 
		Where T1."LineStatus" = 'O' 
		and T0."U_EXO_TIPO"='ITC' 
		AND T2."InvntItem" = 'Y' 
		UNION ALL select
	 T1."DocEntry",
	 T0."DocNum",
	 T1."LineNum",
	 T0."ObjType",
	 T1."ItemCode",
	 T1."OpenQty",
	 T0."U_EXO_DOCPRI",
	 T0."CreateDate",
	 T0."CreateTS",
	 T1."WhsCode" "ALMACEN" 
		FROM rdr1 T1 
		INNER JOIN ORDR T0 ON T0."DocEntry"=T1. "DocEntry" 
		INNER JOIN OITM T2 ON T1."ItemCode" = T2."ItemCode" 
		Where T1."LineStatus" = 'O' 
		and T0."Confirmed"='Y' 
		AND T2."InvntItem" = 'Y' )T 
	INNER JOIN OITW AL ON AL."WhsCode"=T."ALMACEN" 
	and AL."ItemCode"=T."ItemCode" 
	ORDER BY T."ItemCode" ASC,
	 "ALMACEN" ASC,
	 "U_EXO_DOCPRI" DESC,
	 "CreateDate" ASC,
	 "CreateTS" ASC )Z 
INNER JOIN OITW ON Z."ItemCode" = OITW."ItemCode" 
AND Z."ALMACEN" = OITW."WhsCode" 
WHERE "ROTURA"='Y' WITH READ ONLY
