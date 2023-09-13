CREATE VIEW "EXO_SITUACION"  AS 

SELECT T."DocEntry",T."ObjType",T."DocNum", CASE WHEN COUNT("Situacion")=1 then max("Situacion") ELSE 'Ambos' END "Sit"
FROM (
SELECT T1."DocEntry", T0."ObjType",T0."DocNum", IFNULL(OBIN."Attr1Val",'') "Situacion"
FROM rdr1 T1
INNER JOIN ORDR T0 ON T0."DocEntry"=T1. "DocEntry"
INNER JOIN OITW AL ON AL."WhsCode"=T1."WhsCode" and AL."ItemCode"=T1."ItemCode"
LEFT JOIN OBIN ON OBIN."AbsEntry"=AL."DftBinAbs"
Where T1."LineStatus" = 'O' and T0."Confirmed"='Y' 
Group BY T1."DocEntry",T0."ObjType", T0."DocNum", IFNULL(OBIN."Attr1Val",'')
)T
Group BY T."DocEntry",T."ObjType",T."DocNum"
