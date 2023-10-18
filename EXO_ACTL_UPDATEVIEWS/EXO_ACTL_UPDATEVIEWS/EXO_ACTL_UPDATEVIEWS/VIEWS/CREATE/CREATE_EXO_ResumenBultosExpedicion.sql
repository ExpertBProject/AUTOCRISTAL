CREATE VIEW "BBDD"."EXO_ResumenBultosExpedicion" ( "DocEntry",
	 "ID BULTO",
	 "BULTO",
	 "VOLUMEN",
	 "PESO" ) AS SELECT
	 T0."DocEntry",
	 T0."U_EXO_IDBULTO" as "ID BULTO",
	 T0."U_EXO_TBULTO" as "BULTO",
	 CAse when t4."U_EXO_TIPTAR" = 'Peso' 
Then T22."U_EXO_VOLUMEN" 
else ifnull(T1."Volumen",
	 T4."Volume") 
end as "VOLUMEN",
	 CAse when t4."U_EXO_TIPTAR" = 'Peso' 
Then T22."U_EXO_PESO" 
else ifnull(T1."Peso",
	 T4."Weight1") 
end AS "PESO" 
FROM "@EXO_LSTEMBL" t0 
Left Join "@EXO_LSTEMB" T2 On T0."DocEntry" = T2."DocEntry" 
Left Join "@EXO_LSTEMBB" T22 On T0."DocEntry" = T22."DocEntry" 
and T0."U_EXO_IDBULTO" = T22."U_EXO_IDBULTO" 
Left JOin OSHP T3 ON T2."U_EXO_CEXP" = T3."TrnspCode" 
Left join "EXO_PesoBultos_Agencia" T1 ON T0."U_EXO_TBULTO" = T1."PkgType" 
and T1."CardCode" = T3."U_EXO_AGE" 
LEFT JOIN OPKG T4 ON T4."PkgType" = T0."U_EXO_TBULTO" --LEFT JOIN "@EXO_LSTEMBB" T4 On T0."DocEntry" = T4."DocEntry" AND T0."U_EXO_IDBULTO" = t4."U_EXO_IDBULTO" 
 
group BY T0."DocEntry",
	 T0."U_EXO_IDBULTO" ,
	 T0."U_EXO_TBULTO" ,
	 T1."Volumen",
	 T1."Peso" ,
	 T4."Volume" ,
	 t4."U_EXO_TIPTAR",
	 T4."Weight1" ,
	 T22."U_EXO_VOLUMEN",
	 T22."U_EXO_PESO" 
Order by T0."DocEntry",
	 T0."U_EXO_IDBULTO" ,
	 T0."U_EXO_TBULTO" ,
	 T1."Volumen",
	 T1."Peso" WITH READ ONLY
