CREATE VIEW "EXO_SOL_TRASLADO" ( "T. SALIDA",
     "DELEGACIÓN",
     "FECHA CREACION",
     "FECHA ENTREGA",
     "Nº INTERNO",
     "Nº DOCUMENTO",
     "AUTORIZADO",
     "COM",
     "CÓDIGO",
     "EMPRESA",
     "CLASE EXP.",
     "ROT. STOCK",
     "FromWhsCod",
     "TrnspCode",
     "Territory",
     "A",
     "UBICACIÓN",
     "ZONA TRANSPORTE",
     "Sel" ) AS SELECT
     DISTINCT CAST('SOLTRA' as nVARCHAR(50)) "T. SALIDA",
     CAST(IFNULL(T2."Name",
     ' ') as nVARCHAR(50)) "DELEGACIÓN",
     T0."DocDate",
     T0."DocDueDate",
     CAST(T0."DocEntry" as nVARCHAR(50)) "Nº INTERNO",
     CAST(T0."DocNum" as nVARCHAR(50)) "Nº DOCUMENTO",
     T0."Confirmed" "AUTORIZADO",
     (CASE WHEN IFNULL(T0."Comments",'') = ''     THEN 'N'     ELSE 'Y'     END) AS "COM",
     CAST(T0."CardCode" as nVARCHAR(50)) "CÓDIGO",
     CAST(T0."CardName" as nVARCHAR(150)) "EMPRESA",
     CAST(T0."U_EXO_CLASEE" as nVARCHAR(50)) "CLASE EXP.",
     ifnull(R."ROTURA",     'N') "ROT. STOCK",
     TL."FromWhsCod",
     T0."TrnspCode",
     T1."Territory",
     IFNULL(A."A",     'N') "A",
     CAST(IFNULL((SELECT     CASE WHEN COUNT("Situacion")=1 then max("Situacion") ELSE 'Ambos'     END "Sit" 
            FROM ( SELECT     X1."DocEntry",     X0."ObjType",     X0."DocNum",     IFNULL(OBIN."Attr1Val", '') "Situacion" 
                        FROM WTQ1 X1 
                        INNER JOIN OWTQ X0 ON X0."DocEntry"=X1."DocEntry" 
                        INNER JOIN OITW AL ON AL."WhsCode"=X1."WhsCode" and AL."ItemCode"=X1."ItemCode" 
                         LEFT JOIN OBIN ON OBIN."AbsEntry"=AL."DftBinAbs" 
                         Group BY X1."DocEntry", X0."ObjType", X0."DocNum", IFNULL(OBIN."Attr1Val", '') )T 
    WHERE T."DocEntry" = T0."DocEntry" 
        Group BY T."DocEntry", T."ObjType", T."DocNum")
     , 'SIN SITUACIÓN') as nVARCHAR(50)) "UBICACIÓN",
     CAST(TT."descript" as nVARCHAR(50)) "ZONA TRANSPORTE",
     'N' "Sel"

 

FROM OWTQ T0 
LEFT JOIN WTQ1 TL ON TL."DocEntry"=T0."DocEntry" 
LEFT JOIN OCRD T1 ON T0."CardCode"=T1."CardCode" 
LEFT JOIN OUBR T2 ON T1."U_EXO_DELE"=T2."Code" 
LEFT JOIN "EXO_ROTURA" R ON R."DocEntry"=T0."DocEntry" and R."ObjType"=T0."ObjType" 
LEFT JOIN "EXO_SITUACION" S ON S."DocEntry"=T0."DocEntry" and S."ObjType"=T0."ObjType" 
LEFT JOIN "EXO_A" A ON A."CardCode"=T0."CardCode" and A."WhsCode"=TL."WhsCode" 
LEFT JOIN OTER TT ON T1."Territory"=TT."territryID" 
LEFT JOIN OITM ON TL."ItemCode" = OITM."ItemCode" 
LEFT JOIN ( SELECT
     "OrderEntry",
     "OrderLine",
     "BaseObject",
     SUM("RelQtty") AS "RelQtty" 
    FROM PKL1 INNER JOIN OPKL ON PKL1."AbsEntry"=OPKL."AbsEntry"
    WHERE OPKL."Status"<>'C' and OPKL."U_EXO_CIEMAN"<>'Si'
    GROUP BY "OrderEntry",
     "OrderLine",
     "BaseObject" ) "Pick" ON T0."ObjType" = "Pick"."BaseObject"

 

WHERE TL."LineStatus"='O' AND T0."Confirmed" = 'Y' AND OITM."InvntItem" = 'Y' and T0."U_EXO_STATUSP"='P' AND T0."U_EXO_TIPO" = 'ITC' 
     AND (CASE WHEN IFNULL(T0."U_EXO_LSTEMB", 0) = 0 
    THEN 1 
    ELSE ( SELECT
     COUNT(1) 
        FROM "@EXO_LSTEMBL" X0 JOIN "@EXO_LSTEMB" X1 ON X0."DocEntry" = X1."DocEntry" 
        WHERE X1."Status" = 'C' 
        AND "U_EXO_DOCENTRY" = T0."U_EXO_LSTEMB" ) 
    END) > 0 WITH READ ONLY

