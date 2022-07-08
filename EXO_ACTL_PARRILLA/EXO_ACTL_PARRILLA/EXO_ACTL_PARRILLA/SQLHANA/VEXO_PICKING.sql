CREATE VIEW "VEXO_PICKING"  AS 

SELECT sum(IFNULL(PKL1."RelQtty",0)) "Cant.", sum(ifnull(WTR1."Quantity",0)) "Cant. P",PKL1."OrderEntry",PKL1."OrderLine",PKL1."BaseObject" 
FROM PKL1
INNER JOIN OPKL ON OPKL."AbsEntry"=PKL1."AbsEntry"
LEFT JOIN OWTR ON OWTR."U_EXO_NUMPIC"=PKL1."AbsEntry" and OWTR."U_EXO_LINPIC"=PKL1."PickEntry"
LEFT JOIN WTR1 ON OWTR."DocEntry"=WTR1."DocEntry"
GROUP BY PKL1."OrderEntry",PKL1."OrderLine",PKL1."BaseObject" 
ORDER BY PKL1."BaseObject",PKL1."OrderEntry",PKL1."OrderLine"
