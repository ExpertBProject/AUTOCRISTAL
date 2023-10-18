ALTER VIEW "BBDD"."EXO_MRP_Clasificacion_Artículos" ( "ItemCode",
	 "WhsCode",
	 "Clasificacion" ) AS Select
	 T0."ItemCode" ,
	 T0."WhsCode" ,
	 Case When T1."Ventas_Ult_Año" > 3 
then 'A' When T1."Ventas_Ult_Año" <= 3 
and T2."Ventas_8Q" > 0 
then 'B' When T1."Ventas_Ult_Año" <= 3 
and T2."Ventas_8Q" > 0 
then 'B' When T3."Compras_Ult_SEM" > 0 
ANd T1."Ventas_Ult_Año" = 0 
then 'E' When T1."Ventas_Ult_Año" = 0 
then 'F' 
else 'Z' 
end as "Clasificacion" 
from OITW t0 
Left join "EXO_MRP_Ventas24Q" T1 on T1."ItemCode" = T0."ItemCode" 
and T1."WhsCode" = T0."WhsCode" 
Left Join "EXO_MRP_Ventas8Q" T2 ON T2."ItemCode" = T0."ItemCode" 
and T2."WhsCode" = T0."WhsCode" 
Left Join "EXO_MRP_ComprasSemestre" T3 ON T3."ItemCode" = T0."ItemCode" 
and T3."WhsCode" = T0."WhsCode" WITH READ ONLY
