﻿ALTER VIEW "BBDD"."EXO_PedidoCompraEnvioTransporte" ( "IDENVIO",
	 "CODTRASP",
	 "FECHA",
	 "CLIENTE",
	 "IDEXPEDICION",
	 "CODPOSTAL",
	 "IDBULTO",
	 "BULTO",
	 "U_EXO_PRECIO" ) AS Select
	 T0."DocEntry" AS "IDENVIO",
	 T0."U_EXO_AGTCODE" as "CODTRASP",
	 T0."U_EXO_DOCDATE" aS "FECHA",
	 T1."U_EXO_IC" AS "CLIENTE",
	 T1."DocEntry" as "IDEXPEDICION",
	 T2."ZipCode" as "CODPOSTAL",
	 T3."U_EXO_IDBULTO" AS "IDBULTO",
	 T3."U_EXO_TBULTO" as "BULTO",
	 T15."U_EXO_PRECIO" 
from "@EXO_ENVTRANS" T0 
LEFT JOIN "@EXO_LSTEMB" T1 ON T1."U_EXO_IDENVIO" = cast(T0."DocEntry" as nvarchar(50)) 
LEFT JOIN CRD1 T2 ON T1."U_EXO_IC" = T2."CardCode" 
and T2."Address" = T1."U_EXO_DIR" 
and T2."AdresType" = 'S' 
LEFT JOIN "EXO_DetalleBultosExpediciones" T3 ON T3."IdExpedición" = T1."DocEntry" 
LEFT JOIN OWHS T4 ON T4."WhsCode" = T0."U_EXO_ALMACEN" 
LEFT JOIN OPKG T5 ON T5."PkgType" = T3."U_EXO_TBULTO" 
LEFT JOIN "@EXO_SERVICIOS" T10 ON t10."Code" = T0."U_EXO_AGTCODE" 
LEFT JOIN "@EXO_SERVICIOSL" T11 ON T10."Code" = T11."Code" 
and T11."U_EXO_DEL" = T4."U_EXO_SUCURSAL" 
LEFT JOIN "@EXO_TAGENCIA" T12 ON T12."Code" = T11."U_EXO_TARIFA" 
LEFT JOIN "@EXO_TAGENCIAL" T13 ON T12."Code" = T13."Code" 
and T0."U_EXO_DOCDATE" >= T13."U_EXO_FECHAD" 
and T0."U_EXO_DOCDATE" <= T13."U_EXO_FECHAH" 
LEFT JOIN "@EXO_DTARIFAS" T14 ON T14."Code" = T13."U_EXO_TARIFA" 
LEFT JOIN "@EXO_DTARIFASL" T15 ON T14."Code" = T15."Code" 
and T15."U_EXO_LOCAL" = 
left(T2."ZipCode",
	 2) 
and T15."U_EXO_TBULTO" = T5."PkgCode" WITH READ ONLY