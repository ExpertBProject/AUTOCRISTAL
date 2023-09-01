CREATE VIEW "EXO_PEDIDOS_VENTA" ( "T. SALIDA",

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

     "WhsCode",

     "Territory",

     "A",

     "UBICACIÓN",

     "ZONA TRANSPORTE",

     "Sel" ) AS

     SELECT

     DISTINCT CAST('PEDVTA' as nVARCHAR(50)) "T. SALIDA",

     CAST(IFNULL(T2."Name",

     ' ') as nVARCHAR(50)) "DELEGACIÓN",

     T0."DocDate",

     T0."DocDueDate",

     CAST(T0."DocEntry" as nVARCHAR(50)) "Nº INTERNO",

     CAST(T0."DocNum" as nVARCHAR(50)) "Nº DOCUMENTO",

     T0."Confirmed" "AUTORIZADO",

     (CASE WHEN IFNULL(T0."Comments",

     '') = '' 

    THEN 'N' 

    ELSE 'Y' 

    END) AS "COM",

     CAST(T0."CardCode" as nVARCHAR(50)) "CÓDIGO",

     CAST(T0."CardName" as nVARCHAR(150)) "EMPRESA",

     CAST(TL."TrnsCode" as nVARCHAR(50)) "CLASE EXP.",

     ifnull(R."ROTURA",

     'N') "ROT. STOCK",

     TL."WhsCode",

     T1."Territory",

     (CASE WHEN (SELECT

     COUNT(1) 

        FROM ORDR X0 

        WHERE X0."CardCode" = T0."CardCode" 

        AND X0."DocEntry" <> T0."DocEntry" 

        AND X0."DocStatus" = 'O') > 1 

    OR (SELECT

     COUNT(1) 

        FROM ODLN X0 

        WHERE X0."CardCode" = T0."CardCode" 

        AND X0."DocDate" = CURRENT_DATE) > 1 

    THEN 'Y' 

    ELSE 'N' 

    END) "A",

     CAST(IFNULL((SELECT

     CASE WHEN COUNT("Situacion")=1 

            then max("Situacion") 

            ELSE 'Ambos' 

            END "Sit" 

            FROM ( SELECT

     X1."DocEntry",

     X0."ObjType",

     X0."DocNum",

     IFNULL(OBIN."Attr1Val",

     '') "Situacion" 

                FROM RDR1 X1 

                INNER JOIN ORDR X0 ON X0."DocEntry"=X1."DocEntry" 

                INNER JOIN OITW AL ON AL."WhsCode"=X1."WhsCode" 

                and AL."ItemCode"=X1."ItemCode" 

                LEFT JOIN OBIN ON OBIN."AbsEntry"=AL."DftBinAbs" 

                Group BY X1."DocEntry",

     X0."ObjType",

     X0."DocNum",

     IFNULL(OBIN."Attr1Val",

     '') )T 

            WHERE T."DocEntry" = T0."DocEntry" 

            Group BY T."DocEntry",

     T."ObjType",

     T."DocNum"),

     'SIN SITUACIÓN') as nVARCHAR(50)) "UBICACIÓN",

     CAST(TT."descript" as nVARCHAR(50)) "ZONA TRANSPORTE",

     'N' "Sel" 

FROM ORDR T0 

LEFT JOIN RDR1 TL ON TL."DocEntry"=T0."DocEntry" 

INNER JOIN OCRD T1 ON T0."CardCode"=T1."CardCode" 

LEFT JOIN OUBR T2 ON T1."U_EXO_DELE"=T2."Code" 

LEFT JOIN "EXO_ROTURA" R ON R."DocEntry"=T0."DocEntry" 

and R."ObjType"=T0."ObjType" 

LEFT JOIN "EXO_SITUACION" S ON S."DocEntry"=T0."DocEntry" 

and S."ObjType"=T0."ObjType" 

LEFT JOIN "EXO_A" A ON A."CardCode"=T0."CardCode" 

and A."WhsCode"=TL."WhsCode" 

LEFT JOIN OTER TT ON T1."Territory"=TT."territryID" JOIN OITM ON TL."ItemCode" = OITM."ItemCode" 

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

AND T0."DocEntry" = "Pick"."OrderEntry" 

AND TL."LineNum" = "Pick"."OrderLine" 

WHERE TL."LineStatus"='O' 

AND IFNULL("Pick"."RelQtty",     0)<=TL."Quantity"  


AND T0."Confirmed" = 'Y' 

AND OITM."InvntItem" = 'Y' 

      WITH READ ONLY
