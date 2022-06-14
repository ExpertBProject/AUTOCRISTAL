CREATE VIEW "EXO_PEDIDOS_VENTA"  AS 

SELECT DISTINCT CAST('PEDVTA' as nVARCHAR(50)) "T. SALIDA", CAST(IFNULL(T2."Name",' ') as nVARCHAR(50)) "DELEGACIÓN", CAST(T0."DocEntry" as nVARCHAR(50)) "Nº INTERNO", CAST(T0."DocNum" as nVARCHAR(50)) "Nº DOCUMENTO", 
 T0."Confirmed" "AUTORIZADO", CAST(T0."CardCode" as nVARCHAR(50)) "CÓDIGO",  CAST(T0."CardName" as nVARCHAR(150))	"EMPRESA", CAST(T0."TrnspCode" as nVARCHAR(50)) "CLASE EXP.", 
 ifnull(R."ROTURA",'N') "ROT. STOCK", TL."WhsCode", T1."Territory",
 IFNULL(A."A",'N') "A", CAST(IFNULL(S."Sit",'SIN SITUACIÓN') as nVARCHAR(50)) "UBICACIÓN", CAST(TT."descript" as nVARCHAR(50)) "ZONA TRANSPORTE", 'N' "Sel" 
FROM ORDR T0 
 LEFT JOIN RDR1 TL ON TL."DocEntry"=T0."DocEntry" 
 INNER JOIN OCRD T1 ON T0."CardCode"=T1."CardCode" 
 LEFT JOIN OUBR T2 ON T1."U_EXO_DELE"=T2."Code" 
 LEFT JOIN "EXO_ROTURA" R ON R."DocEntry"=T0."DocEntry" and R."ObjType"=T0."ObjType" 
 LEFT JOIN "EXO_SITUACION" S ON S."DocEntry"=T0."DocEntry" and S."ObjType"=T0."ObjType" 
 LEFT JOIN "EXO_A" A ON A."CardCode"=T0."CardCode" and A."WhsCode"=TL."WhsCode" 
 LEFT JOIN OTER TT ON T1."Territory"=TT."territryID" 
 WHERE TL."LineStatus"='O' and T0."Confirmed"='Y' and T0."U_EXO_STATUSP"='P' 
