CREATE VIEW "EXO_A" ( "A",
	 "CardCode",
	 "WhsCode" ) AS Select
	 case WHEN Count(X."DocEntry") > 1 
THEN 'Y' 
ELSE 'N' 
END AS "A",
	 X."CardCode" ,
	 X."WhsCode" 
from ( Select
	 DISTINCT T0."DocEntry",
	 T0."CardCode",
	 T1."WhsCode" 
	from ORDR T0 
	left join RDR1 T1 ON T0."DocEntry" = T1."DocEntry" 
	and T1."LineStatus" = 'O' 
	WHERE T0."CANCELED" <> 'Y' 
	AND T0."Confirmed"='Y' 
	AND T0."DocStatus" <> 'C' 
	Union all Select
	 DISTINCT T0."DocEntry",
	 T0."CardCode",
	 T1."WhsCode" 
	from ODLN T0 
	left join DLN1 T1 ON T0."DocEntry" = T1."DocEntry" 
	and T1."LineStatus" = 'O' 
	Left JOIN "@EXO_LSTEMB" T2 ON T2."DocEntry" = T0."U_EXO_LSTEMB" 
	WHERE T0."CANCELED" <> 'Y' 
	AND T0."Confirmed"='Y' 
	AND T0."DocStatus" <> 'C' 
	and T2."Status" = 'O' 
	union all Select
	 DISTINCT T0."DocEntry",
	 T0."CardCode",
	 T1."WhsCode" 
	from OWTQ T0 
	left join WTQ1 T1 ON T0."DocEntry" = T1."DocEntry" 
	and T1."LineStatus" = 'O' 
	Left JOIN "@EXO_LSTEMB" T2 ON T2."DocEntry" = T0."U_EXO_LSTEMB" 
	left join OCRD T3 ON T3."CardCode" = T0."CardCode" 
	WHERE T0."CANCELED" <> 'Y' 
	AND T0."Confirmed"='Y' 
	AND T0."DocStatus" <> 'C' 
	AND T3."CardType" = 'C' ) X 
where X."CardCode" is not null 
Group BY X."CardCode" ,
	 X."WhsCode" 
ORDER BY X."WhsCode" WITH READ ONLY