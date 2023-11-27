﻿ALTER PROCEDURE "BBDD"."EXO_GET_INTEGRACION_SYNC"(TIPO VARCHAR(50), ESTADO varchar(50))
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN
--Listado de SOURCE_ENTITY

SELECT T_SYNC."ID", T_SYNC."DATE_AND_TIME", T_SYNC."SOURCE_SYSTEM", T_SYNC."TARGET_SYSTEM", T_SYNC."TARGET_SUBSYSTEM", T_SYNC."SOURCE_PK",
	   T_SYNC."TARGET_PAYLOAD", T_SYNC."RESULT_STATUS_CODE", T_SYNC."RESULT_BODY", T_SYNC."RESULT_TARGET_PK", T_SYNC."OB_RESPONSE_ID", 
	   CASE WHEN T_SYNC."SOURCE_ENTITY" = 'TARIF_S' THEN 'Tarifa de Proveedores'
	   WHEN T_SYNC."SOURCE_ENTITY" = 'TARIF_C' THEN 'Tarifa de Clientes'
	   ELSE T_SYNC."SOURCE_ENTITY"
	   END AS "SOURCE_ENTITY"
	   --JSON_QUERY(T_SYNC."RESULT_BODY",'$.error.message') as "ERROR"
FROM (
		    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_ITEMS_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_ITEMS_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE T_SYNC."METHOD" = 'GET' AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_CATALOGS_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_CATALOGS_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE T_SYNC."METHOD" = 'GET' AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_PARTNERS_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_ITEMS_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE T_SYNC."METHOD" = 'GET' AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_STOCKS_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_STOCKS_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE T_SYNC."METHOD" = 'GET' AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_TARIFS_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_TARIFS_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE T_SYNC."METHOD" = 'GET' AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_TARIFS_SUPPLIER_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_TARIFS_SUPPLIER_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE T_SYNC."METHOD" = 'GET' AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_DISTRIBUTION_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_DISTRIBUTION_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE T_SYNC."METHOD" = 'GET' AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_BUSINESS_PARTNERS_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_BUSINESS_PARTNERS_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_CREDIT_NOTES_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_CREDIT_NOTES_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_GOODS_ISSUED_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_GOODS_ISSUED_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_GOODS_RECEIPT_PO_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_GOODS_RECEIPT_PO_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_INVENTORY_POSTING_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_INVENTORY_POSTING_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_INVENTORY_TRANSFER_REQUEST_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_INVENTORY_TRANSFER_REQUEST_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_INVENTORY_TRANSFER_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_INVENTORY_TRANSFER_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_INVENTORY_TRANSFER_TOCLOSE_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_INVENTORY_TRANSFER_TOCLOSE_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_INVOICES_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_INVOICES_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_JOURNAL_ENTRIES_CAJA_BANCO_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_JOURNAL_ENTRIES_CAJA_BANCO_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_JOURNAL_ENTRIES_GASTO_MENOR_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_JOURNAL_ENTRIES_GASTO_MENOR_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_JOURNAL_ENTRIES_TPV_CAJA_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_JOURNAL_ENTRIES_TPV_CAJA_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
	UNION ALL
			    (SELECT T_SYNC.*,
				   CASE WHEN T_SYNC."ID" <= LAST_INT_OK."ID" THEN 'INTEGRADO'
				   WHEN T_SYNC."ID" > COALESCE(LAST_INT_OK."ID",0) AND T_SYNC."VERIFIED_OB_RESPONSE" = '0' THEN 'POR VERIFICAR'
				   ELSE 'ERROR'
				   END AS "ESTADO"
			FROM EXO_INTEGRACION.EXO_MANUFACTURING_SYNC as T_SYNC
			LEFT OUTER JOIN (
					   SELECT T_SYNC."SOURCE_PK", MAX(T_SYNC."ID") as "ID"
					   FROM EXO_INTEGRACION.EXO_MANUFACTURING_SYNC as T_SYNC
					   WHERE COALESCE(T_SYNC."RESULT_TARGET_PK",'') != ''
					   GROUP BY T_SYNC."SOURCE_PK"
					   ) AS LAST_INT_OK ON T_SYNC."SOURCE_PK" = LAST_INT_OK."SOURCE_PK"
			WHERE COALESCE(T_SYNC."RESULT_STATUS_CODE",0) != 0 AND COALESCE(T_SYNC."RESULT_TARGET_PK",'') = ''
			ORDER BY T_SYNC."ID" DESC)
) as T_SYNC
WHERE (T_SYNC."SOURCE_ENTITY" = :TIPO OR :TIPO = '')
 	  AND (T_SYNC."ESTADO" = :ESTADO OR :ESTADO = '')
ORDER BY T_SYNC."DATE_AND_TIME" desc;

END;