CREATE VIEW "BBDD"."EXO_MANIFIESTO_TTE" ( "IdExpedicion",
	 "IdEnvioTTE",
	 "U_EXO_ALMACEN",
	 "WhsName",
	 "CardName",
	 "StreetS",
	 "ZipCodeS",
	 "CityS",
	 "Name",
	 "Documentos",
	 "CantBulto",
	 "Peso",
	 "Volumen",
	 "TotalBultosExp",
	 "U_EXO_PORDEB" ) AS Select
	 DISTINCT t0."DocEntry" as "IdExpedicion",
	 T8."DocEntry" as "IdEnvioTTE",
	 T0."U_EXO_ALMACEN",
	 T01."WhsName",
	 CASE WHEN T1."U_EXO_ORIGEN" = 'ALBVTA' 
THEN T0."U_EXO_DESTINO" 
else T51."WhsName" 
end as "CardName",
	 case when T1."U_EXO_ORIGEN" = 'ALBVTA' 
then T3."StreetS" 
else T51."Street" 
end as "StreetS",
	 case when T1."U_EXO_ORIGEN" = 'ALBVTA' 
then T3."ZipCodeS" 
else T51."ZipCode" 
end as "ZipCodeS" ,
	 case when T1."U_EXO_ORIGEN" = 'ALBVTA' 
then T3."CityS" 
else T51."City" 
end as "CityS",
	 case when T1."U_EXO_ORIGEN" = 'ALBVTA' 
then T31."Name" 
else T52."Name" 
end as "Name",
	 T4."Documentos" ,
	 TZ."Resumenbulto" as "CantBulto",
	 TX."Peso" ,
	 TX."Volumen" ,
	 TZ."TotalBultosExp" ,
	 coalesce(T2."U_EXO_PORDEB",
	 'No') as "U_EXO_PORDEB" 
from "@EXO_LSTEMB" T0 
LEFT JOIN "@EXO_ENVTRANS" T8 ON cast(t8."DocEntry" as Nvarchar(15)) = T0."U_EXO_IDENVIO" 
LEFT JOIN "@EXO_LSTEMBL" T1 ON T0."DocEntry" = T1."DocEntry" 
LEFT JOIN "ODLN" T2 ON T1."U_EXO_DOCENTRY" = T2."DocEntry" 
and T1."U_EXO_ORIGEN" = 'ALBVTA' 
LEFT JOIN "DLN12" T3 ON T3."DocEntry" = T2."DocEntry" 
LEFT JOIN "OWTQ" T5 ON T1."U_EXO_DOCENTRY" = T5."DocEntry" 
and T1."U_EXO_ORIGEN" = 'SOLTRA' 
LEFT JOIN "OWHS" t51 ON T51."WhsCode" = T5."ToWhsCode" 
and T1."U_EXO_ORIGEN" = 'SOLTRA' 
LEFT JOIN "OCST" t52 ON t52."Code" = T51."State" 
and T52."Country" = T51."Country" 
LEFT JOIN "OCST" t31 ON t31."Code" = T3."StateS" 
and T31."Country" = T3."CountryS" 
LEFT JOIN "OWHS" t01 ON T01."WhsCode" = T0."U_EXO_ALMACEN" --   Concatena Documentos
 --LEFT JOIN (Select "DocEntry" , "U_EXO_ORIGEN" || '- ' || String_AGG  ("Documentos",', ') as "Documentos" from (
--			Select DISTINCT "DocEntry" , "U_EXO_ORIGEN",   "U_EXO_DOCNUM" as "Documentos"
--			from    "@EXO_LSTEMBL"  )
--			group BY "DocEntry" ,  "U_EXO_ORIGEN")   T4  ON T4."DocEntry" = T0."DocEntry"

LEFT JOIN ( Select
	 DISTINCT "DocEntry" ,
	 "U_EXO_ORIGEN",
	 MIN("U_EXO_DOCNUM") as "Documentos" 
	from "@EXO_LSTEMBL" 
	GROUP BY "DocEntry" ,
	 "U_EXO_ORIGEN") T4 on T4."DocEntry" = T0."DocEntry" ------ Concatena resumen Bultos
 
LEFT JOIN (select
	 Z."DocEntry",
	 String_Agg ("Resumenbulto",
	 ' - ' ) as "Resumenbulto" ,
	 Sum(Z."CanBultos") as "TotalBultosExp" 
	From ( Select
	 Distinct T1."DocEntry" as "DocEntry",
	 T1."U_EXO_TBULTO" || '( ' || count(distinct T1."U_EXO_IDBULTO") || ')  - ' || IFNULL(T1."U_EXO_AGRBUL",
	'') as "Resumenbulto" ,
	 count(distinct T1."U_EXO_IDBULTO") as "CanBultos" 
		from "@EXO_LSTEMBL" T1 
		Left join "@EXO_LSTEMB" T0 ON T0."DocEntry" = T1."DocEntry" 
		group by T1."DocEntry",
	 T1."U_EXO_TBULTO",
	 T0."U_EXO_CEXP" ,
	 T1."U_EXO_TBULTO" ,
	 T1."U_EXO_AGRBUL"
		Order by T1."DocEntry" ) Z 
	group by Z."DocEntry" ) TZ ON TZ."DocEntry" = t0."DocEntry" ------- calcula peso

LEFT JOIN ( Select
	 TY."DocEntry",
	 Sum( TY."Peso") as "Peso" ,
	 Sum(TY."Volumen") as "Volumen" 
	from (select
	 T2."DocEntry" ,
	 T2."ID BULTO" ,
	 T2."PESO" as "Peso",
	 T2."VOLUMEN" as "Volumen" 
		from "EXO_ResumenBultosExpedicion" T2 ) TY 
	group by TY."DocEntry" ) TX ON TX."DocEntry" = T0."DocEntry" 
oRDER BY t0."DocEntry" WITH READ ONLY
