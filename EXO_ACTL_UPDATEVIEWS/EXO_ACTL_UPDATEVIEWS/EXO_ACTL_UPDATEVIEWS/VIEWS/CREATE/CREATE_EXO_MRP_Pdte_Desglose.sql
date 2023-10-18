CREATE VIEW "BBDD"."EXO_MRP_Pdte_Desglose" ( "ItemCode",
	 "WhsCode",
	 "PDTE" ) AS select
	 T0."ItemCode",
	 T0."WhsCode",
	 COALESCE(coalesce(T0."OnOrder",
	 0) - COALESCE(X."CantidadSolTraInt",
	 0) ,
	 0) as "PDTE" 
from OITW T0 
left join ( Select
	 T1."ItemCode" ,
	 T1."WhsCode" ,
	 coalesce(Sum(T1."OpenQty"),
	 0) as "CantidadSolTraInt" 
	from OWTQ T0 
	LEFT JOIN WTQ1 T1 ON T0."DocEntry" = T1."DocEntry" 
	Where T0."DocStatus" = 'O' 
	and T1."LineStatus" = 'O' 
	and T1."FromWhsCod" = T1."WhsCode" 
	Group by T1."ItemCode",
	t1."WhsCode" ) X On X."ItemCode" = T0."ItemCode" 
and T0."WhsCode" = X."WhsCode" WITH READ ONLY
