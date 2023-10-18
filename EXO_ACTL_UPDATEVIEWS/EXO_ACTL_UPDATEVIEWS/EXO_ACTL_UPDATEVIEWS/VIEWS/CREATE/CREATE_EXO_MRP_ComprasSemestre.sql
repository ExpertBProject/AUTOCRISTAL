CREATE VIEW "BBDD"."EXO_MRP_ComprasSemestre" ( "ItemCode",
	 "WhsCode",
	 "Compras_Ult_SEM" ) AS Select
	 X."ItemCode",
	 X."WhsCode",
	 X."Quantity" AS "Compras_Ult_SEM" 
from (Select
	 T0."ItemCode",
	 T1."Quantity" ,
	 T2."DocDate" ,
	 T1."WhsCode" 
	from OPCH T2 
	Left join PCH1 t1 ON T1."DocEntry" = T2."DocEntry" 
	Left JOin OITM T0 ON T0."ItemCode" = T1."ItemCode" 
	Where T1."Quantity" is not null 
	and T2."DocType" <> 'S' 
	and T2."DocDate" >= ADD_MONTHS(ADD_DAYS(CURRENT_DATE,
	 -EXTRACT(DAY 
				FROM CURRENT_DATE) + 1),
	 -6) 
	uNION ALL Select
	 T0."ItemCode",
	 - T1."Quantity" ,
	 T2."DocDate" ,
	 T1."WhsCode" 
	from ORPC T2 
	Left join RPC1 t1 ON T1."DocEntry" = T2."DocEntry" 
	Left JOin OITM T0 ON T0."ItemCode" = T1."ItemCode" 
	Where T1."Quantity" is not null 
	and T2."DocType" <> 'S' 
	and T2."DocDate" >= ADD_MONTHS(ADD_DAYS(CURRENT_DATE,
	 -EXTRACT(DAY 
				FROM CURRENT_DATE) + 1),
	 -6) 
	union all Select
	 T0."ItemCode",
	 T0."Quantity" ,
	 T0."DocDate" ,
	 T0."WhsCode" 
	from "HISTORICO_ACS"."COMPRASACS" T0 
	Where T0."Quantity" is not null 
	and T0."DocDate" >= ADD_MONTHS(ADD_DAYS(CURRENT_DATE,
	 -EXTRACT(DAY 
				FROM CURRENT_DATE) + 1),
	 -6) ) as X 
WHERE X."WhsCode" in ('AL0',
	 'AL7',
	 'AL14',
	 'AL8',
	 'AL16') WITH READ ONLY
