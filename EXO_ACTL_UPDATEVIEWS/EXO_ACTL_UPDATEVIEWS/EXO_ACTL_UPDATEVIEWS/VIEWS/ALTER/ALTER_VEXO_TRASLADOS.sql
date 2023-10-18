ALTER VIEW "BBDD"."VEXO_TRASLADOS" ( "Cant. T",
	 "OrderEntry",
	 "BaseObject" ) AS SELECT
	 sum(ifnull(WTR1."Quantity",
	 0)) "Cant. T",
	 PKL1."OrderEntry",
	 PKL1."BaseObject" 
FROM PKL1 
INNER JOIN OPKL ON OPKL."AbsEntry"=PKL1."AbsEntry" 
LEFT JOIN OWTR ON OWTR."U_EXO_NUMPIC"=PKL1."AbsEntry" 
and OWTR."U_EXO_LINPIC"=PKL1."PickEntry" 
LEFT JOIN WTR1 ON OWTR."DocEntry"=WTR1."DocEntry" 
WHERE OPKL."Status"<>'C' 
GROUP BY PKL1."OrderEntry",
	 PKL1."BaseObject" 
ORDER BY PKL1."BaseObject",
	 PKL1."OrderEntry" WITH READ ONLY
