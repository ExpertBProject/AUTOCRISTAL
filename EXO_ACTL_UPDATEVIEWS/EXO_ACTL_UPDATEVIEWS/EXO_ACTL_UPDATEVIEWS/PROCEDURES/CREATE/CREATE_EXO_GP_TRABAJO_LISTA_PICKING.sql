CREATE PROCEDURE "BBDD"."EXO_GP_TRABAJO_LISTA_PICKING"
(
	IN ClavePicking NVARCHAR(30),
	
	OUT TablaDatos TABLE
	(
	AbsEntry INT,
	PickEntry INT,
	ItemCode NVARCHAR(50),
	ItemName NVARCHAR(100),
	CantidadTotal DECIMAL(19,6),
	Cantidad DECIMAL(19,6),
	Udm NVARCHAR(20),
	BatchNum NVARCHAR(36),
	BinCode NVARCHAR(228),
	Propuesto NVARCHAR(1),
	EsLote NCHAR(1),    
	SePuedeGestionar NVARCHAR(1),
	NumPerMsr DECIMAL(16,6),
	CantidadPick DECIMAL(19,6),
	BinCodeSitu NVARCHAR(30)
	)
)
LANGUAGE SQLSCRIPT
AS
-- Return values
BEGIN
--esto de abajo sobra 
	
	DECLARE vCantidadAsignada  DECIMAL(19,6) := 0;
	DECLARE vCantidadFaltante  DECIMAL(19,6) := 0;
	DECLARE vLineaPicking  INT := 0;
	
	DECLARE vLotePropuesto  NVARCHAR(36) :='';
  	DECLARE vCantidadPropuesta DECIMAL(19,6) :=0;
  	DECLARE vUbicProp  NVARCHAR(228) :='';
	
	
  --OK Los de lote asignados  
	DECLARE CURSOR c_1 FOR
	SELECT T2."AbsEntry" AbsEntry, T2."PickEntry" LinPicket, T1."ItemCode" ItemCode,T2."PickQtty" PickQtty, T1."Dscription" ItemName,
	sum(case when t5."Quantity">0 then t5."Quantity" else 0 end) as Cantidad, 
			COALESCE(Tart."InvntryUom", '') Udm, T6."DistNumber" BatchNum,
			ifnull(t8."BinCode",'') BinCode,'N'  propuesto, 'Y' EsLote, 'Y' SePuedeGestionar ,
			COALESCE(TArt."NumInSale",1) NumPerMsr,T2."PickQtty" + T2."RelQtty" CantidadPick,
			ifnull(T8."Attr1Val",'')  UbiORSitu
 	FROM "OWTR" T0	   
		INNER JOIN "WTR1" T1 ON T0."DocEntry" = T1."DocEntry"
		INNER JOIN "OITM" TArt on TArt."ItemCode" = T1."ItemCode"
		INNER JOIN "PKL1" T2 ON T2."AbsEntry" = T0."U_EXO_NUMPIC" AND T2."PickEntry" = T0."U_EXO_LINPIC"

		INNER Join "OITL" T4 On T4."DocEntry"=T1."DocEntry" And T4."DocLine"=T1."LineNum" And T4."DocType"=t1."ObjType"
		INNER JOIN "ITL1" T5 ON T5."LogEntry" = T4."LogEntry"
		INNER JOIN "OBTN" T6 ON  T6."SysNumber" = T5."SysNumber" AND T6."ItemCode" = T5."ItemCode" And T6."AbsEntry"=T5."MdAbsEntry"
		INNER JOIN "OITW" T7 ON T1."ItemCode" = T7."ItemCode" and t1."FromWhsCod"=t7."WhsCode"
		left join "OBIN" T8 ON T7."DftBinAbs" = t8."AbsEntry"

	WHERE  T0."U_EXO_NUMPIC" = :ClavePicking AND ( T6."DistNumber" NOT IN (

	SELECT TAsig6."DistNumber"  
	FROM "OITL" TAsig4 
	INNER JOIN "ITL1" TAsig5 ON TAsig5."LogEntry" = TAsig4."LogEntry"
	INNER JOIN "OBTN" TAsig6 ON  TAsig6."SysNumber" = TAsig5."SysNumber" AND TAsig6."ItemCode" = TAsig5."ItemCode" And TAsig6."AbsEntry"=TAsig5."MdAbsEntry"
	AND TAsig4."ApplyType" = T2."BaseObject" AND TAsig4."ApplyEntry" = T2."OrderEntry" AND TAsig4."ApplyLine" = T2."OrderLine"
	) )group by T2."AbsEntry" ,T2."PickEntry" ,T1."ItemCode" , T1."Dscription",
	 T2."PickQtty",T1."Quantity" ,TArt."InvntryUom",T6."DistNumber",
	 TArt."NumInSale",T2."RelQtty",T8."BinCode",T8."Attr1Val";


	   --OK los articulos sin lote  y ya han pasado
	DECLARE CURSOR c_2 FOR
 	SELECT T2."AbsEntry" AbsEntry, T2."PickEntry" LinPicket,
	   		T1."ItemCode" ItemCode, T1."Dscription" ItemName, T2."PickQtty"  PickQtty,T1."Quantity" Cantidad, 
	   		COALESCE(TArt."InvntryUom", '') AS Udm, '' BatchNum, '' AS BinCode,
			'N' propuesto, 'N' EsLote, 'Y' SePuedeGestionar  ,COALESCE(TArt."NumInSale",1) NumPerMsr , T2."PickQtty"+T2."RelQtty" CantidadPick,ifnull(T8."Attr1Val",'')  UbiORSitu
	FROM "OWTR" T0 
		INNER JOIN "WTR1" T1 ON T0."DocEntry" = T1."DocEntry"
		INNER JOIN "OITM" TArt on TArt."ItemCode" = T1."ItemCode"
		INNER JOIN "PKL1" T2 ON T2."AbsEntry" = T0."U_EXO_NUMPIC" AND T2."PickEntry" = T0."U_EXO_LINPIC"
			INNER JOIN "OITW" T7 ON T1."ItemCode" = T7."ItemCode" and t1."FromWhsCod"=t7."WhsCode"
		left join "OBIN" T8 ON T7."DftBinAbs" = t8."AbsEntry"									
	WHERE  TART."ManBtchNum" = 'N' AND T0."U_EXO_NUMPIC" = :ClavePicking;
		
	--OK los que faltan, picking completo, que le iremos restando lo de arriba	
	
		DECLARE CURSOR cPicking_3 FOR
		SELECT T1."AbsEntry" AbsEntry, T1."PickEntry" PickEntry, T2."ItemCode" ItemCode, T2."Dscription" ItemName, 
				T1."RelQtty"  CantLinea, COALESCE(T2."unitMsr", '') AS Udm, 
				T2."WhsCode" Almacen, T3."ManBtchNum" ConLote,COALESCE(T3."NumInSale",1) NumPerMsr,T1."PickQtty"+T1."RelQtty" CantidadPick,ifnull(T8."Attr1Val",'') BinCodeSitu,t8."BinCode" BinCode
		FROM "OPKL" T0
				INNER JOIN "PKL1" T1 ON T0."AbsEntry" = T1."AbsEntry"
			INNER JOIN "RDR1" T2 ON T1."BaseObject" = T2."ObjType" AND T2."DocEntry" = T1."OrderEntry" AND T2."LineNum" = T1."OrderLine"
			INNER JOIN "OITM" T3 ON T3."ItemCode" = T2."ItemCode" 
			INNER JOIN "OITW" T7 ON T2."ItemCode" = T7."ItemCode" and t2."WhsCode"=t7."WhsCode"
			left join "OBIN" T8 ON T7."DftBinAbs" = t8."AbsEntry"
		WHERE T0."AbsEntry" = :ClavePicking	
		UNION ALL
		SELECT T1."AbsEntry" AbsEntry, T1."PickEntry" PickEntry, T2."ItemCode" ItemCode, T2."Dscription" ItemName, 
				T1."RelQtty"  CantLinea, COALESCE(T2."unitMsr", '') AS Udm, 
				T2."FromWhsCod" Almacen, T3."ManBtchNum" ConLote,COALESCE(T3."NumInSale",1) NumPerMsr,T1."PickQtty"+T1."RelQtty" CantidadPick,ifnull(T8."Attr1Val",'') BinCodeSitu,t8."BinCode" BinCode
		FROM "OPKL" T0
				INNER JOIN "PKL1" T1 ON T0."AbsEntry" = T1."AbsEntry"
			INNER JOIN "WTQ1" T2 ON T1."BaseObject" = T2."ObjType" AND T2."DocEntry" = T1."OrderEntry" AND T2."LineNum" = T1."OrderLine"
			INNER JOIN "OITM" T3 ON T3."ItemCode" = T2."ItemCode" 
			INNER JOIN "OITW" T7 ON T2."ItemCode" = T7."ItemCode" and t2."FromWhsCod"=t7."WhsCode"
			left join "OBIN" T8 ON T7."DftBinAbs" = t8."AbsEntry"
		WHERE T0."AbsEntry" = :ClavePicking;	
	
		

	CREATE LOCAL TEMPORARY TABLE #tmp (
		AbsEntry INT,
		PickEntry INT,
		ItemCode NVARCHAR(50),
		ItemName NVARCHAR(100),
		CantidadTotal DECIMAL(19,6),
		Cantidad DECIMAL(19,6),
		Udm NVARCHAR(20),
		BatchNum NVARCHAR(36),
		BinCode NVARCHAR(228),
		Propuesto NVARCHAR(1),
		EsLote NCHAR(1),    
		SePuedeGestionar NVARCHAR(1),
		NumPerMsr DECIMAL(16,6),
		CantidadPick DECIMAL(19,6),
		BinCodeSitu NVARCHAR(30)
	);

	
	--Agregamos todos los datos a la tabla temporals	

	FOR c_row_1 AS c_1 DO
		INSERT INTO #tmp VALUES (c_row_1.AbsEntry, c_row_1.LinPicket,c_row_1.ItemCode,c_row_1.ItemName,c_row_1.PickQtty,c_row_1.Cantidad,
		c_row_1.Udm,c_row_1.BatchNum,c_row_1.BinCode,c_row_1.propuesto,c_row_1.EsLote,c_row_1.SePuedeGestionar,c_row_1.NumPerMsr,c_row_1.CantidadPick,c_row_1.UbiORSitu);
	END FOR;

	FOR c_row_2 AS c_2 DO
		INSERT INTO #tmp VALUES (c_row_2.AbsEntry, c_row_2.LinPicket,c_row_2.ItemCode,c_row_2.ItemName,c_row_2.PickQtty,c_row_2.Cantidad,
		c_row_2.Udm,c_row_2.BatchNum,c_row_2.BinCode,c_row_2.propuesto,c_row_2.EsLote,c_row_2.SePuedeGestionar,c_row_2.NumPerMsr,c_row_2.CantidadPick,c_row_2.UbiORSitu);
	END FOR;

	FOR c_row_3 AS cPicking_3 DO

 		SELECT coalesce(SUM(Cantidad),0) INTO vCantidadAsignada 
 		FROM #tmp  
 		WHERE AbsEntry = :ClavePicking and PickEntry = c_row_3.PickEntry;
 		

 		IF COALESCE(vCantidadAsignada,0) < c_row_3.CantLinea THEN
			vCantidadFaltante :=  c_row_3.CantLinea - vCantidadAsignada;
			vCantidadPropuesta := vCantidadFaltante;

			IF c_row_3.ConLote = 'Y' THEN
				vCantidadFaltante :=  c_row_3.CantLinea - vCantidadAsignada;
				--llamamos a procedimiento propongo lote, que nos devuelve lote,cantidad, ubicacion
				CALL EXO_GP_PROPONGO_LOTE(c_row_3.Almacen,c_row_3.ItemCode,vCantidadFaltante,vLotePropuesto,vCantidadPropuesta,vUbicProp);

			ELSE
				vLotePropuesto :='';
  				vCantidadPropuesta := vCantidadFaltante;
  	 			vUbicProp :='';
				--llamamos a propongo ubicacion para uqe nos devuelva la ubicacion del articulo
				SELECT EXO_GP_PROPONGO_UBICACION(c_row_3.Almacen, c_row_3.ItemCode,'N', vCantidadPropuesta, 'V') INTO vUbicProp FROM DUMMY;
			
			END IF;
			
			--hacer el insert para la tabla
			INSERT INTO #tmp VALUES (:ClavePicking, c_row_3.PickEntry,c_row_3.ItemCode,c_row_3.ItemName,vCantidadFaltante,:vCantidadPropuesta,
			c_row_3.Udm,vLotePropuesto,vUbicProp,'Y',c_row_3.ConLote,'N',c_row_3.NumPerMsr,c_row_3.CantidadPick,c_row_3.BinCodeSitu);
			
 		END IF;
 		
	END FOR;

	TablaDatos = SELECT AbsEntry,PickEntry,ItemCode,ItemName,MAX(CantidadTotal) CantidadTotal,SUM(Cantidad) Cantidad,Udm,BatchNum,BinCode,Propuesto,EsLote, SePuedeGestionar,NumPerMsr  ,CantidadPick,BinCodeSitu
	FROM #tmp 
	GROUP BY AbsEntry,PickEntry,ItemCode,ItemName,Udm,BatchNum,BinCode,Propuesto,EsLote, SePuedeGestionar,NumPerMsr,CantidadPick,BinCodeSitu;
	DROP TABLE #tmp;

END;
