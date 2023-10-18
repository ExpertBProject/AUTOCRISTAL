CREATE VIEW "BBDD"."VEXO_PARRILLA_ESTADO_SALCOMP" ( "ORIGEN",
	 "DocEntry",
	 "Estado",
	 "Estado de Packing" ) AS ((select
	 'ALBVTA' as "ORIGEN",
	 T0."DocEntry",
	 case when coalesce(SUM (T2."U_EXO_CANT") ,
	 0) = 0 
		then 'PP' 
		else case when SUM(T01."Quantity") - coalesce(SUM (T2."U_EXO_CANT") ,
	 0) > 0 
		then 'PC' 
		else case when SUM(T01."Quantity") - coalesce(SUM (T2."U_EXO_CANT") ,
	 0) <= 0 
		Then 'PT' 
		end 
		end 
		end as "Estado",
	 case when coalesce(SUM (T2."U_EXO_CANT") ,
	 0) = 0 
		then 'Packing Pendiente' 
		else case when SUM(T01."Quantity") - coalesce(SUM (T2."U_EXO_CANT") ,
	 0) > 0 
		then 'Packing En Curso' 
		else case when SUM(T01."Quantity") - coalesce(SUM (T2."U_EXO_CANT") ,
	 0) <= 0 
		Then 'Packing Completado' 
		end 
		end 
		end as "Estado de Packing" 
		from ODLN T0 
		LEFT JOIN DLN1 T01 ON T01."DocEntry" = T0."DocEntry" 
		left join "@EXO_LSTEMB" T1 ON T0."U_EXO_LSTEMB" = T1."DocEntry" 
		LEFT JOIN "@EXO_LSTEMBL" T2 ON T2."DocEntry" = T1."DocEntry" 
		and T2."U_EXO_ORIGEN" = 'ALBVTA' 
		AND T2."U_EXO_DOCENTRY" = T0."DocEntry" 
		group by T2."U_EXO_ORIGEN",
	 T0."DocEntry") 
	UNION ALL (select
	'SOLTRA' as "ORIGEN",
	 T0."DocEntry",
	 case when coalesce(SUM (T2."U_EXO_CANT") ,
	 0) = 0 
		then 'PP' 
		else case when SUM(T01."Quantity") - coalesce(SUM (T2."U_EXO_CANT") ,
	 0) > 0 
		then 'PC' 
		else case when SUM(T01."Quantity") - coalesce(SUM (T2."U_EXO_CANT") ,
	 0) <= 0 
		Then 'PT' 
		end 
		end 
		end as "Estado",
	 case when coalesce(SUM (T2."U_EXO_CANT") ,
	 0) = 0 
		then 'Packing Pendiente' 
		else case when SUM(T01."Quantity") - coalesce(SUM (T2."U_EXO_CANT") ,
	 0) > 0 
		then 'Packing En Curso' 
		else case when SUM(T01."Quantity") - coalesce(SUM (T2."U_EXO_CANT") ,
	 0) <= 0 
		Then 'Packing Completado' 
		end 
		end 
		end as "Estado de Packing" 
		from OWTQ T0 
		LEFT JOIN WTQ1 T01 ON T01."DocEntry" = T0."DocEntry" 
		left join "@EXO_LSTEMB" T1 ON T0."U_EXO_LSTEMB" = T1."DocEntry" 
		LEFT JOIN "@EXO_LSTEMBL" T2 ON T2."DocEntry" = T1."DocEntry" 
		and T2."U_EXO_ORIGEN" = 'SOLTRA' 
		AND T2."U_EXO_DOCENTRY" = T0."DocEntry" 
		group by T2."U_EXO_ORIGEN",
	 T0."DocEntry")) WITH READ ONLY
