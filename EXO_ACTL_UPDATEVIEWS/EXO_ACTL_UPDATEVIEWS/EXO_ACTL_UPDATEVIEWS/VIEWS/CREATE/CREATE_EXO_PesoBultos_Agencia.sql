CREATE VIEW "BBDD"."EXO_PesoBultos_Agencia" ( "CardCode",
	 "PkgCode",
	 "PkgType",
	 "Peso",
	 "Volumen" ) AS SELECT
	 T0."CardCode" ,
	 T1."PkgCode",
	 T1."PkgType" ,
	 IFnull(T2."U_EXO_PESO",
	 T1."Weight1") As "Peso",
	 IFnull(T2."U_EXO_VOLUMEN",
	 T1."Volume") As "Volumen" 
From OCRD T0 
inner join OPKG T1 on 1 = 1 
LEFT JOIN "@EXO_BULTOSAGL" T2 ON T0."CardCode" = T2."Code" 
and T2."U_EXO_BULTO" = T1."PkgCode" 
where T0."QryGroup1" = 'Y' WITH READ ONLY