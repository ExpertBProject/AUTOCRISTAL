CREATE VIEW "EXO_MRP_Pdte" ( "ItemCode","Pdte_AL0","Pdte_AL14", "Pdte_AL16","Pdte_AL7", "Pdte_AL8" ) AS Select
      TX."ItemCode" ,
      Coalesce(sum(case when TX."WhsCode" = 'AL0' 
            then "PDTE" 
            end),
      0) AS "Pdte_AL0",
      Coalesce(sum(case when TX."WhsCode" = 'AL14' 
            then "PDTE" 
            end),
      0) AS "Pdte_AL14",
      Coalesce(sum(case when TX."WhsCode" = 'AL16' 
            then "PDTE" 
            end),
      0) AS "Pdte_AL16",
      Coalesce(sum(case when TX."WhsCode" = 'AL7' 
            then "PDTE" 
            end),
      0) AS "Pdte_AL7",
      Coalesce(sum(case when TX."WhsCode" = 'AL8' 
            then "PDTE" 
            end),
      0) AS "Pdte_AL8" 
from ( select
      T0."ItemCode",
      T0."WhsCode",
      COALESCE(T0."OnOrder" - X."CantidadSolTraInt" ,
      0) as "PDTE" 
      from OITW T0 
      left join ( Select
      T1."ItemCode" ,
      T1."WhsCode" ,
      coalesce(Sum(T1."OpenQty"),
      0) as "CantidadSolTraInt" 
            from OWTQ T0 
            LEFT JOIN WTQ1 T1 ON T0."DocEntry" = T1."DocEntry" 
            Where T0."DocStatus" = 'O' 
            and T1."LineStatus" = 'O' 
            and T1."FromWhsCod" = T1."WhsCode" 
            and T1."WhsCode" in ('AL0',
      'AL7',
      'AL14',
      'AL8',
      'AL16') 
            Group by T1."ItemCode",
      t1."WhsCode" ) X On X."ItemCode" = T0."ItemCode" 
      and T0."WhsCode" = X."WhsCode" 
      Where T0."OnOrder" > 0 ) TX 
Group BY TX."ItemCode" WITH READ ONLY
