CREATE VIEW "BBDD"."EXO_EtiquetaBultos" ( "U_EXO_ALMACEN",
	 "U_EXO_NEXT",
	 "DocEntry",
	 "CardCode",
	 "CardName",
	 "TrnspName",
	 "U_EXO_DIR",
	 "Street",
	 "ZipCode",
	 "City",
	 "Name",
	 "Phone1",
	 "Albaranes",
	 "Referencias",
	 "U_EXO_IDBULTO",
	 "BultoMax",
	 "Bultosorden",
	 "Remark",
	 "U_EXO_TBULTO",
	 "CodBar" ) AS SELECT
	 T0."U_EXO_ALMACEN",
	 T0."U_EXO_NEXT",
	 T0."DocEntry",
	 T2."CardCode",
	 T2."CardName",
	 T3."TrnspName",
	 T0."U_EXO_DIR",
	 T4."Street",
	 T4."ZipCode",
	 T4."City",
	 T5."Name",
	 T2."Phone1",
	 Z."Albaranes" ,
	 z1."Referencias" ,
	 T1."U_EXO_IDBULTO" ,
	 Y."BultoMax" ,
	 cast(T1."U_EXO_IDBULTO" as nvarchar(10) ) || ' / ' || cast (Y."BultoMax" as nvarchar(10)) as "Bultosorden",
	 T0."Remark",
	 T1."U_EXO_TBULTO",
	 T0."DocEntry" || ';' || T1."U_EXO_IDBULTO" || ';' || T0."DocNum" || ';' || T0."U_EXO_IDENVIO" as "CodBar" 
FROM "@EXO_LSTEMB" T0 
LEFT JOIN "@EXO_LSTEMBL" T1 ON T1."DocEntry" = T0."DocEntry" 
LEFT JOIN OCRD T2 ON T2."CardCode" = T0."U_EXO_IC" 
LEFT JOIN OSHP T3 ON T0."U_EXO_CEXP" = T3."TrnspCode" 
LEFT JOIN CRD1 T4 ON T4."CardCode" = T2."CardCode" 
and T4."Address" = T0."U_EXO_DIR" 
and T4."AdresType" = 'S' 
LEFT JOIN OCST T5 ON T4."State" = t5."Code" 
and T4."Country" = T5."Country" 
LEFT JOIN (Select
	 string_agg ( X."albaran",
	 ',') as "Albaranes" ,
	 X."DocEntry" ,
	 X."U_EXO_IDBULTO" 
	from ( Select
	 distinct T1."DocNum" as "albaran" ,
	 T0."DocEntry" ,
	 T0."U_EXO_IDBULTO" 
		FROM "@EXO_LSTEMBL" T0 
		LEFT JOIN ODLN T1 ON T1."DocEntry" = T0."U_EXO_DOCENTRY" ) as X 
	group by X."DocEntry",
	 X."U_EXO_IDBULTO" ) Z On Z."DocEntry" = T0."DocEntry" 
and Z."U_EXO_IDBULTO" = T1."U_EXO_IDBULTO" 
LEFT JOIN (Select
	 string_agg ( X1."referencia",
	 ',') as "Referencias" ,
	 X1."DocEntry" 
	from ( Select
	 distinct T1."NumAtCard" as "referencia" ,
	 T0."DocEntry" 
		FROM "@EXO_LSTEMBL" T0 
		LEFT JOIN ODLN T1 ON T1."DocEntry" = T0."U_EXO_DOCENTRY" ) as X1 
	group by X1."DocEntry" ) Z1 On Z1."DocEntry" = T0."DocEntry" 
LEFT JOIN ( Select
	 Max("U_EXO_IDBULTO") as "BultoMax" ,
	 "DocEntry" 
	from "@EXO_LSTEMBL" 
	group by "DocEntry") Y ON Y."DocEntry" = T0."DocEntry" WITH READ ONLY
