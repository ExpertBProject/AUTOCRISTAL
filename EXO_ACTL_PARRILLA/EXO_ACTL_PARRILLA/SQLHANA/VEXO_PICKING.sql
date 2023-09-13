CREATE VIEW "VEXO_PICKING" ( "Picking", "Cant.", "OrderEntry", "BaseObject" ) AS 
SELECT IFNULL(OPKL."AbsEntry",0) "Picking", sum(IFNULL(PKL1."RelQtty",0)) "Cant.",PKL1."OrderEntry",	PKL1."BaseObject" 
FROM PKL1 
INNER JOIN OPKL ON OPKL."AbsEntry"=PKL1."AbsEntry" 
WHERE OPKL."Status"<>'C'
GROUP BY OPKL."AbsEntry", PKL1."OrderEntry",	PKL1."BaseObject" 
ORDER BY PKL1."BaseObject",	PKL1."OrderEntry" 
