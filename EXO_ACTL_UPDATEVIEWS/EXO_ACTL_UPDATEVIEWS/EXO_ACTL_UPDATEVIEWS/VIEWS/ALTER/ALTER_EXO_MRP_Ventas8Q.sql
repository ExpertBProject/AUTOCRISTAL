ALTER VIEW "BBDD"."EXO_MRP_Ventas8Q" ( "ItemCode",
	 "WhsCode",
	 "Ventas_8Q" ) AS Select
	 t0."ItemCode",
	 T0."WhsCode",
	 coalesce(sum( X."Quantity"),
	 0) AS "Ventas_8Q" 
from OITW T0 
Left Join (Select
	 T0."ItemCode",
	 T1."Quantity" ,
	 T2."DocDate" ,
	 T1."WhsCode" 
	from OINV T2 
	Left join INV1 t1 ON T1."DocEntry" = T2."DocEntry" 
	Left JOin OITM T0 ON T0."ItemCode" = T1."ItemCode" 
	Where T1."Quantity" is not null 
	and T2."DocType" <> 'S' 
	and T2."DocDate" >= ADD_MONTHS(ADD_DAYS(CURRENT_DATE,
	 -EXTRACT(DAY 
				FROM CURRENT_DATE) + 1),
	 - 4) 
	uNION ALL Select
	 T0."ItemCode",
	 - T1."Quantity" ,
	 T2."DocDate" ,
	 T1."WhsCode" 
	from ORIN T2 
	Left join RIN1 t1 ON T1."DocEntry" = T2."DocEntry" 
	Left JOin OITM T0 ON T0."ItemCode" = T1."ItemCode" 
	Where T1."Quantity" is not null 
	and T2."DocType" <> 'S' 
	and T2."DocDate" >= ADD_MONTHS(ADD_DAYS(CURRENT_DATE,
	 -EXTRACT(DAY 
				FROM CURRENT_DATE) + 1),
	 -4) 
	and T1."WhsCode" in ('AL0',
	 'AL7',
	 'AL14',
	 'AL8',
	 'AL16') 
	union all Select
	 T0."ItemCode",
	 T0."Quantity" ,
	 T0."DocDate" ,
	 T0."WhsCode" 
	from "HISTORICO_ACS"."VENTASACS" T0 
	Where T0."DocDate" >= ADD_MONTHS(ADD_DAYS(CURRENT_DATE,
	 -EXTRACT(DAY 
				FROM CURRENT_DATE) + 1),
	 -4) 
	and T0."WhsCode" in ('AL0',
	 'AL7',
	 'AL14',
	 'AL8',
	 'AL16') ) as X ON X."ItemCode" = T0."ItemCode" 
and X."WhsCode" = T0."WhsCode" 
group by T0."ItemCode",
	 T0."WhsCode" WITH READ ONLY
