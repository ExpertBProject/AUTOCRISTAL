CREATE VIEW "EXO_A"  AS 
SELECT DISTINCT CASE WHEN COUNT(T0."DocEntry")>1 THEN 'Y' ELSE 'N' END "A","CardCode",T1."WhsCode"
FROM rdr1 T1
INNER JOIN ORDR T0 ON T0."DocEntry"=T1. "DocEntry"
Where T1."ItemCode" = 'v' and T1."WhsCode"='AL0' and T1."LineStatus" = 'O' and T0."Confirmed"='Y'
GROUP BY "CardCode",T1."WhsCode"
