CREATE VIEW "BBDD"."EXO_DetalleBultosEnvioTransporte" ( "IdEnvioTTE",
	 "IdExpedicion",
	 "U_EXO_DESTINO",
	 "Direccion",
	 "Resumenbulto",
	 "Peso",
	 "Volumen" ) AS Select
	 T0."DocEntry" as "IdEnvioTTE",
	 T1."DocEntry" as "IdExpedicion",
	 T1."U_EXO_DESTINO" ,
	 T2."Street" || ' ' || T2."City" || '(' || T2."ZipCode" || ') ' || T3."Name" AS "Direccion" ,
	 TZ."Resumenbulto" ,
	 TY."Peso",
	 TY."Volumen" 
from "SBO_AUTOCRISTAL_PRUEBAS"."@EXO_ENVTRANS" T0 
LEFT JOIN "SBO_AUTOCRISTAL_PRUEBAS"."@EXO_LSTEMB" T1 ON cast(t0."DocEntry" as Nvarchar(15)) = T1."U_EXO_IDENVIO" -------- Direccion

LEFT JOIN "SBO_AUTOCRISTAL_PRUEBAS"."CRD1" T2 ON T2."CardCode" = T1."U_EXO_IC" 
and "AdresType" = 'S' 
and T2."Address" = T1."U_EXO_DIR" 
LEFT JOIN "SBO_AUTOCRISTAL_PRUEBAS"."OCST" t3 ON t3."Code" = T2."State" 
and T3."Country" = T2."Country" -------  Bultos

LEFT JOIN (select
	 Z."DocEntry",
	 String_Agg ("Resumenbulto",
	 ' - ' ) as "Resumenbulto" 
	From ( Select
	 Distinct T1."DocEntry" as "DocEntry",
	 T1."U_EXO_TBULTO" || '( ' || count(DISTINCT T1."U_EXO_IDBULTO") || ')' as "Resumenbulto",
	 T0."U_EXO_CEXP",
	 T1."U_EXO_TBULTO" 
		from "SBO_AUTOCRISTAL_PRUEBAS"."@EXO_LSTEMBL" T1 
		Left join "SBO_AUTOCRISTAL_PRUEBAS"."@EXO_LSTEMB" T0 ON T0."DocEntry" = T1."DocEntry" 
		group by T1."DocEntry",
	 T1."U_EXO_TBULTO",
	 T0."U_EXO_CEXP" ,
	 T1."U_EXO_TBULTO" ) Z 
	Left Join "SBO_AUTOCRISTAL_PRUEBAS".OSHP T3 On T3."TrnspCode" = Z."U_EXO_CEXP" 
	LEFT JOIN "SBO_AUTOCRISTAL_PRUEBAS"."EXO_PesoBultos_Agencia" T2 ON T2."CardCode" = T3."U_EXO_AGE" 
	and Z."U_EXO_TBULTO" = T2."PkgType" 
	group by Z."DocEntry" ) TZ ON TZ."DocEntry" = t1."DocEntry" ------- calcula peso

LEFT JOIN ( Select
	 Z."DocEntry",
	 coalesce (Sum( T2."PESO"),
	 Sum( T4."Weight1")) as "Peso" ,
	 coalesce(Sum(T2."VOLUMEN"),
	 Sum(T4."Volume")) as "Volumen" --, Sum( T4."Weight1") as "Peso2" , Sum(T4."Volume") as "Volumen2" 
 
	from ( Select
	 Distinct T1."DocEntry" as "DocEntry",
	 T1."U_EXO_IDBULTO" as "U_EXO_IDBULTO",
	 T1."U_EXO_TBULTO" as "U_EXO_TBULTO" ,
	 T0."U_EXO_CEXP" as "U_EXO_CEXP" 
		from "SBO_AUTOCRISTAL_PRUEBAS"."@EXO_LSTEMBL" T1 
		Left join "SBO_AUTOCRISTAL_PRUEBAS"."@EXO_LSTEMB" T0 ON T0."DocEntry" = T1."DocEntry" ) Z 
	Left Join "SBO_AUTOCRISTAL_PRUEBAS".OSHP T3 On T3."TrnspCode" = Z."U_EXO_CEXP" 
	LEFT JOIN "SBO_AUTOCRISTAL_PRUEBAS"."EXO_ResumenBultosExpedicion" T2 ON Z."DocEntry" = T2."DocEntry" 
	and Z."U_EXO_IDBULTO" = T2."ID BULTO" 
	LEFT JOIN "SBO_AUTOCRISTAL_PRUEBAS"."OPKG" T4 ON Z."U_EXO_TBULTO" = T4."PkgType" 
	group by Z."DocEntry" ) TY ON TY."DocEntry" = t1."DocEntry" WITH READ ONLY
