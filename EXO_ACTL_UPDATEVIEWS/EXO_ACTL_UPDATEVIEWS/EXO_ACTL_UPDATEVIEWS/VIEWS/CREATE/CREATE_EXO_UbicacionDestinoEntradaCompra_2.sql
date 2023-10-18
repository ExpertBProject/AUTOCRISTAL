CREATE VIEW "BBDD"."EXO_UbicacionDestinoEntradaCompra_2" ( "Code",
	 "LineId",
	 "UBICADESTINO" ) AS select
	 T0."Code" ,
	 T0."LineId" ,
	 Case when X."TipoBulto" = 'MIXTO' 
then ifnull(T3."BinCode" ,
	 Y."BinCode") 
else ifnull(Y."BinCode",
	 T3."BinCode") 
end AS "UBICADESTINO" 
from "@EXO_PACKINGL" T0 
left Join (Select
	 "Code" ,
	 "U_EXO_IDBULTO" ,
	 case when count("LineId") = 1 
	then 'COMPLETO' 
	ELSE 'MIXTO' 
	END "TipoBulto" 
	from "@EXO_PACKINGL" 
	Group by "Code" ,
	 "U_EXO_IDBULTO" ) X ON X."Code" = T0."Code" 
and X."U_EXO_IDBULTO" = T0."U_EXO_IDBULTO" 
left join OITW T2 ON T2."ItemCode" = T0."U_EXO_CODE" 
and T0."U_EXO_ALM" = T2."WhsCode" 
left join OBIN T3 ON T3."AbsEntry" = T2."DftBinAbs" 
Left join (select
	 count( DISTINCT t0."BinCode") over (PARTITION BY T0."WhsCode" 
		ORDER BY t0."BinCode")as "CONTADOR" ,
	 T0."BinCode",
	 T0."WhsCode" ,
	 IFNULL(Sum(T1."OnHandQty"),
	 0) as "Cantidad" 
	from OBIN T0 
	left join OIBQ T1 ON T1."BinAbs" = T0."AbsEntry" 
	where t0."Attr2Val" in ('Picking',
	 'Stocks') 
	group by t0."WhsCode",
	 T0."AbsEntry",
	 T0."BinCode" having IFNULL(Sum(T1."OnHandQty"),
	 0) = 0 
	order by T0."AbsEntry",
	 "Cantidad" ) Y ON Y."WhsCode" = T0."U_EXO_ALM" 
and Y."CONTADOR" = T0."LineId" + 1 
order by "LineId" WITH READ ONLY
