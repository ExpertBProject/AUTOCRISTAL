CREATE VIEW "BBDD"."EXO_DetalleBultosExpediciones" ( "IdExpedición",
	 "U_EXO_IDBULTO",
	 "U_EXO_TBULTO",
	 "Volumen",
	 "Peso" ) AS SELECT
	 Z."DocEntry" AS "IdExpedición",
	 Z."U_EXO_IDBULTO" as "U_EXO_IDBULTO",
	 Z."U_EXO_TBULTO" as "U_EXO_TBULTO",
	 CAse when t4."U_EXO_TIPTAR" = 'Peso' 
Then T22."U_EXO_VOLUMEN" 
else Sum(IFNULL(T2."Volumen",
	 t4."Volume")) 
end as "Volumen",
	 CAse when t4."U_EXO_TIPTAR" = 'Peso' 
Then T22."U_EXO_PESO" 
else Sum(IFNULL(T2."Peso",
	 T4."Weight1")) 
end as "Peso" 
from ( Select
	 Distinct T1."DocEntry" as "DocEntry",
	 T1."U_EXO_IDBULTO" as "U_EXO_IDBULTO",
	 T1."U_EXO_TBULTO" as "U_EXO_TBULTO" ,
	 T0."U_EXO_CEXP" as "U_EXO_CEXP" 
	from "SBO_AUTOCRISTAL_PRUEBAS"."@EXO_LSTEMBL" T1 
	Left join "SBO_AUTOCRISTAL_PRUEBAS"."@EXO_LSTEMB" T0 ON T0."DocEntry" = T1."DocEntry" ) Z 
Left Join OSHP T3 On T3."TrnspCode" = Z."U_EXO_CEXP" 
LEFT JOIN "SBO_AUTOCRISTAL_PRUEBAS"."EXO_PesoBultos_Agencia" T2 ON T2."CardCode" = T3."U_EXO_AGE" 
and Z."U_EXO_TBULTO" = T2."PkgType" 
LEFT JOIN "@EXO_LSTEMBB" T22 On Z."DocEntry" = T22."DocEntry" 
AND Z."U_EXO_IDBULTO" = t22."U_EXO_IDBULTO" 
LEFT JOIN "SBO_AUTOCRISTAL_PRUEBAS"."OPKG" T4 ON Z."U_EXO_TBULTO" = T4."PkgType" 
group by Z."DocEntry" ,
	 Z."U_EXO_IDBULTO",
	 Z."U_EXO_TBULTO" ,
	 t4."U_EXO_TIPTAR",
	 T22."U_EXO_VOLUMEN",
	 T22."U_EXO_PESO" 
Order BY Z."DocEntry" ,
	 Z."U_EXO_IDBULTO" WITH READ ONLY
