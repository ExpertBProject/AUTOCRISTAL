CREATE VIEW "EXO_MRP_StocksActuales" ( "ItemCode", "Stock_AL0", "Stock_AL14","Stock_AL8","Stock_AL16", "Stock_AL7" ) AS Select
      Y."ItemCode",
      Y."Stock_AL0",
      Y."Stock_AL14",
      Y."Stock_AL8",
      Y."Stock_AL16",
      Y."Stock_AL7" 
from (Select
      X."ItemCode",
      sum(case when X."WhsCode" = 'AL0' 
            then X."OnHand" 
            end) AS "Stock_AL0",
      sum(case when X."WhsCode" = 'AL14' 
            then X."OnHand" 
            end) AS "Stock_AL14",
      sum(case when X."WhsCode" = 'AL8' 
            then X."OnHand" 
            end) AS "Stock_AL8",
      sum(case when X."WhsCode" = 'AL16' 
            then X."OnHand" 
            end) AS "Stock_AL16",
      sum(case when X."WhsCode" = 'AL7' 
            then X."OnHand" 
            end) AS "Stock_AL7" 
      from ( Select
      T0."ItemCode",
      T0."ItemName" ,
      T1."WhsCode" ,
      T1."OnHand" 
            from OITM T0 
            left join OITW T1 On T0."ItemCode" = T1."ItemCode" 
            Where T1."WhsCode" in ('AL0',
      'AL7',
      'AL14',
      'AL8',
      'AL16') ) X 
      GROUP BY x."ItemCode") Y 
WHERE Y."Stock_AL0" <> 0 
OR Y."Stock_AL14" <> 0 
OR Y."Stock_AL8" <> 0 
OR Y."Stock_AL7" <> 0 
OR Y."Stock_AL16" <> 0 WITH READ ONLY


