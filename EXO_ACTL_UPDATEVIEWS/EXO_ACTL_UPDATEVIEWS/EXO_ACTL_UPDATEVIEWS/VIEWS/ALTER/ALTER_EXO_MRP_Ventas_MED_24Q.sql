ALTER VIEW "BBDD"."EXO_MRP_Ventas_MED_24Q" ( "ItemCode",
	 "24M_Q_AL0",
	 "24M_Q_AL14",
	 "24M_Q_AL16",
	 "24M_Q_AL7",
	 "24M_Q_AL8" ) AS Select
	 TX."ItemCode" ,
	 Sum("24M_Q_AL0") as "24M_Q_AL0",
	 Sum("24M_Q_AL14") as "24M_Q_AL14",
	 Sum("24M_Q_AL16") as "24M_Q_AL16",
	 Sum("24M_Q_AL7") as "24M_Q_AL7",
	 Sum("24M_Q_AL8") as "24M_Q_AL8" 
from ( select
	 "ItemCode",
	 CASE when "WhsCode" = 'AL0' 
	then "Ventas_Med_Año" 
	else 0 
	end as "24M_Q_AL0",
	 CASE when "WhsCode" = 'AL14' 
	then "Ventas_Med_Año" 
	else 0 
	end as "24M_Q_AL14",
	 CASE when "WhsCode" = 'AL16' 
	then "Ventas_Med_Año" 
	else 0 
	end as "24M_Q_AL16",
	 CASE when "WhsCode" = 'AL7' 
	then "Ventas_Med_Año" 
	else 0 
	end as "24M_Q_AL7",
	 CASE when "WhsCode" = 'AL8' 
	then "Ventas_Med_Año" 
	else 0 
	end as "24M_Q_AL8" 
	from "EXO_MRP_Ventas24Q" t0) TX 
GROUP BY TX."ItemCode" WITH READ ONLY
