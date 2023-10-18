CREATE VIEW "BBDD"."EXO_MRP_Proveedores" ( "ItemCode",
	 "Provee",
	 "Provee_II",
	 "Provee_III",
	 "Mejor" ) AS Select
	 T0."ItemCode",
	 Case WHen T0."QryGroup1" = 'Y' 
then 'STOCK' 
ELSE T0."CardCode" 
END as "Provee" ,
	 Case WHen T0."QryGroup1" = 'Y' 
then T0."CardCode" 
ELSE MAX(T1."VendorCode") 
END as "Provee_II" ,
	 TY."CardCode" as "Provee_III" ,
	 TY."CardCode" || '_' || TY."Price" as "Mejor" 
from OITM t0 
LEFT JOIN ITM2 T1 ON T1."ItemCode" = T0."ItemCode" 
and T1."VendorCode" <> T0."CardCode" 
LEFT join ( Select
	 T0."ItemCode",
	 T0."PriceList" ,
	 T0."Price",
	 T3."CardCode" 
	from ITM1 T0 
	INNER JOIN (Select
	 T0."ItemCode" ,
	 MIn(T0."Price") as "Precio_Min" 
		from ITM1 T0 
		INNER JOIN OPLN T2 ON T2."ListNum" = T0."PriceList" 
		and T2."U_EXO_TARCOM" = 'Si' 
		LEFT JOIN OCRD T1 On T1."ListNum" = T0."PriceList" 
		Where Coalesce(T1."U_EXO_TSUM",
	 0) <= 13 
		and T0."Price" <> 0 
		Group by T0."ItemCode" 
		Order By T0."ItemCode") TX On TX."ItemCode" = T0."ItemCode" 
	and T0."Price" = TX."Precio_Min" 
	LEFT JOIN OCRD T3 On T3."ListNum" = T0."PriceList" 
	LEFT JOIN OPLN T4 on T0."PriceList" = T4."ListNum" 
	Where t0."Price" > 0 
	and T4."U_EXO_TARCOM" = 'Si' ) TY ON TY."ItemCode" = T0."ItemCode" 
Group by T0."ItemCode",
	 T0."QryGroup1" ,
	 T0."CardCode" ,
	 TY."CardCode" ,
	 TY."Price" WITH READ ONLY
