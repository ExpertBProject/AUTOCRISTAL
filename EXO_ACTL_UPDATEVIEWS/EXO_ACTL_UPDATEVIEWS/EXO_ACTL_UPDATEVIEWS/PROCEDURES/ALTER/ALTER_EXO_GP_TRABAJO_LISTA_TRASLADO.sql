ALTER PROCEDURE "BBDD"."EXO_GP_TRABAJO_LISTA_TRASLADO"
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
	BinCodeDestino NVARCHAR(228),
	CodeBars NVARCHAR(228),
	BinCodeSitu NVARCHAR(30),
	BinCodeDestinoSitu NVARCHAR(30),
	Substitute NVARCHAR(100),
	IDBulto NVARCHAR(100),
	TipoBulto NVARCHAR(100),
	Comentario NVARCHAR(100)
	)
)
LANGUAGE SQLSCRIPT
AS
-- Return values
BEGIN
	
	DECLARE vCantidadAsignada  DECIMAL(19,6) := 0;
	DECLARE vCantidadFaltante  DECIMAL(19,6) := 0;
	DECLARE vLineaPicking  INT := 0;
	
	DECLARE vLotePropuesto  NVARCHAR(36) :='';
  	DECLARE vCantidadPropuesta DECIMAL(19,6) :=0;
  	DECLARE vUbicProp  NVARCHAR(228) :='';
	

    --los articulos con lote que no estaban asignados y ya han pasado
	DECLARE CURSOR c_1 FOR
	
		SELECT T7."DocEntry" AbsEntry, T7."LineNum" LinPicket,
	   		T1."ItemCode" ItemCode, T1."Dscription" ItemName, 
	   		case when coalesce(T15."U_EXO_CANT",-1)=-1 THEN T7."Quantity" ELSE T15."U_EXO_CANT" END  PickQtty,
	   		sum(T1."Quantity") Cantidad, 
	   		COALESCE(TArt."InvntryUom", '') Udm, 
	   		T15."U_EXO_LOTE" BatchNum,
	   		T7."U_EXO_UBI_OR" BinCode,
	   		'N'  propuesto, 'Y' EsLote, 'Y' SePuedeGestionar ,COALESCE(TArt."NumInSale",1) NumPerMsr,
	   		case when coalesce(T15."U_EXO_CANT",-1)=-1 THEN T7."Quantity" ELSE T15."U_EXO_CANT" END CantidadPick,
			case when coalesce(T15."U_EXO_UBIDES",'')='' THEN T7."U_EXO_UBI_DE" ELSE T15."U_EXO_UBIDES" end BinCodeDestino,
			tArt."CodeBars" CodeBars, T8."Attr1Val"  UbiORSitu, 
			case when coalesce(T99."Attr1Val",'')='' THEN T9."Attr1Val" ELSE T99."Attr1Val" end  UbiDESitu,
			IFNULL(T12."Substitute",'') Substitute,
			T1."U_EXO_LOT_ID",
			T1."U_EXO_TBULTO", T15."U_EXO_COMENT" Comentario
			FROM "OWTR" T0
				INNER JOIN "WTR1" T1 ON T0."DocEntry" = T1."DocEntry"
				INNER JOIN "OITM" TArt on TArt."ItemCode" = T1."ItemCode"
				
				INNER JOIN "WTQ1" T7 ON  T7."DocEntry" = T1."BaseEntry" AND T7."LineNum" = T1."BaseLine"
				INNER JOIN "OWTQ" T10 ON T7."DocEntry"=T10."DocEntry"
				
				
				INNER Join "OITL" T4 On T4."DocEntry"=T1."DocEntry" And T4."DocLine"=T1."LineNum" And T4."DocType"=t1."ObjType"
				INNER JOIN "ITL1" T5 ON T5."LogEntry" = T4."LogEntry" AND T5."Quantity">0
				INNER JOIN "OBTN" T6 ON  T6."SysNumber" = T5."SysNumber" AND T6."ItemCode" = T5."ItemCode" And T6."AbsEntry"=T5."MdAbsEntry"
				
				
				LEFT JOIN "OBIN" T8 ON T7."U_EXO_UBI_OR"=T8."BinCode"
				LEFT JOIN "OBIN" T9 ON T7."U_EXO_UBI_DE"=T9."BinCode"
				
				LEFT JOIN "@EXO_PACKINGL" T15 ON T10."U_EXO_PACKING"=T15."Code" AND T15."U_EXO_LINEA"=T7."LineNum" and t15."U_EXO_IDBULTO"=T1."U_EXO_LOT_ID"
				LEFT JOIN "OBIN" T99 ON T15."U_EXO_UBIDES"=T99."BinCode"
				LEFT JOIN "OSCN" T12 ON T10."CardCode"=T12."CardCode" and TArt."ItemCode"=T12."ItemCode"
				WHERE T0."U_EXO_NUMPICE" = :ClavePicking and T0."CANCELED"<>'Y'
			group by T7."DocEntry" ,T7."LineNum" ,T1."ItemCode" , T1."Dscription",TArt."InvntryUom",TArt."NumInSale",T7."Quantity",T7."U_EXO_UBI_OR",T7."U_EXO_UBI_DE",tArt."CodeBars",
				T8."Attr1Val",T9."Attr1Val",T12."Substitute",T1."U_EXO_LOT_ID",T1."U_EXO_TBULTO",T15."U_EXO_CANT",T15."U_EXO_UBIDES",T99."Attr1Val",T15."U_EXO_LOTE", T15."U_EXO_COMENT";

		
		
    --los articulos sin lote  y ya han pasado
	--DECLARE CURSOR c_2 FOR
	-- SELECT T2."AbsEntry" AbsEntry, T2."PickEntry" LinPicket,
	--   		T1."ItemCode" ItemCode, T1."Dscription" ItemName, T2."PickQtty" PickQtty,T1."Quantity" Cantidad, 
	--   		COALESCE(TArt."InvntryUom", '') Udm, '' BatchNum, T7."U_EXO_UBI_OR" BinCode,
	--		'N' propuesto, 'N' EsLote, 'Y' SePuedeGestionar  ,COALESCE(TArt."NumInSale",1) NumPerMsr ,T2."PickQtty"+T2."RelQtty" CantidadPick,
	--		T7."U_EXO_UBI_DE" BinCodeDestino,tArt."CodeBars" CodeBars, T8."Attr1Val" UbiORSitu, T9."Attr1Val" UbiDESitu,IFNULL(T12."Substitute",'') Substitute,
	--		T7."U_EXO_LOT_ID",T7."U_EXO_TBULTO"
	--FROM "OWTR" T0 
	--	INNER JOIN "WTR1" T1 ON T0."DocEntry" = T1."DocEntry"
	--	INNER JOIN "OITM" TArt on TArt."ItemCode" = T1."ItemCode"
	--	INNER JOIN "PKL1" T2 ON T2."AbsEntry" = T0."U_EXO_NUMPICE" AND T2."PickEntry" = T0."U_EXO_LINPICE"
	--	INNER JOIN "WTQ1" T7 ON T2."BaseObject" = T7."ObjType" AND T7."DocEntry" = T2."OrderEntry" AND T7."LineNum" = T2."OrderLine"
	--	LEFT JOIN "OBIN" T8 ON T7."U_EXO_UBI_OR"=T8."BinCode"
	--	LEFT JOIN "OBIN" T9 ON T7."U_EXO_UBI_DE"=T9."BinCode"
	--		LEFT JOIN "OWTQ" T10 ON T7."DocEntry"=T10."DocEntry"
	--			LEFT JOIN "OSCN" T12 ON T10."CardCode"=T12."CardCode" and TArt."ItemCode"=T12."ItemCode"
	--WHERE  TART."ManBtchNum" = 'N' AND T0."U_EXO_NUMPICE" = :ClavePicking;

 
		
	--los que faltan, picking completo, que le iremos restando lo de arriba	
	DECLARE CURSOR cPicking_3 FOR
	SELECT T1."DocEntry" AbsEntry, T2."LineNum" PickEntry, T2."ItemCode" ItemCode, T2."Dscription" ItemName, 
			case when coalesce(T15."U_EXO_CANT",'-1')='-1' THEN T2."Quantity" ELSE T15."U_EXO_CANT" END   CantLinea,
			COALESCE(T2."unitMsr", '') Udm, 
			T2."WhsCode" Almacen, 
			T3."ManBtchNum" ConLote,
			COALESCE(T3."NumInSale",1) NumPerMsr,
			case when coalesce(T15."U_EXO_CANT",-1)=-1 THEN T2."Quantity" ELSE T15."U_EXO_CANT" END CantidadPick,
			T2."U_EXO_UBI_OR" BinCode,
			case when coalesce(T15."U_EXO_UBIDES",'')='' THEN T2."U_EXO_UBI_DE" ELSE T15."U_EXO_UBIDES" END BinCodeDestino,
			t3."CodeBars" CodeBars,T8."Attr1Val"  UbiORSitu, 
				case when coalesce(T99."Attr1Val",'')='' THEN T9."Attr1Val" ELSE T99."Attr1Val" end  UbiDESitu,
			IFNULL(T12."Substitute",'') Substitute,
			case when coalesce(T15."U_EXO_IDBULTO",'0')='0' THEN coalesce(T2."U_EXO_LOT_ID",'0') ELSE T15."U_EXO_IDBULTO" END U_EXO_LOT_ID,
			case when coalesce(T15."U_EXO_TBULTO",'')='' THEN coalesce(T2."U_EXO_TBULTO",'') ELSE T15."U_EXO_TBULTO" END U_EXO_TBULTO,
			T15."U_EXO_LOTE" Lote, T15."U_EXO_COMENT" Comentario
	FROM "OWTQ" T1
		INNER JOIN "WTQ1" T2 ON T1."DocEntry" = T2."DocEntry"
		INNER JOIN "OITM" T3 ON T3."ItemCode" = T2."ItemCode"
		LEFT JOIN "OBIN" T8 ON T2."U_EXO_UBI_OR"=T8."BinCode"
		LEFT JOIN "OBIN" T9 ON T2."U_EXO_UBI_DE"=T9."BinCode" 
		LEFT JOIN "OSCN" T12 ON T1."CardCode"=T12."CardCode" and T3."ItemCode"=T12."ItemCode"
		LEFT JOIN "@EXO_PACKINGL" T15 ON T1."U_EXO_PACKING"=T15."Code" AND T15."U_EXO_LINEA"=T2."LineNum"
			LEFT JOIN "OBIN" T99 ON T15."U_EXO_UBIDES"=T99."BinCode"
	WHERE T1."DocEntry" = :ClavePicking;
		
		
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
		EsLote NVARCHAR(1),    
		SePuedeGestionar NVARCHAR(1),
		NumPerMsr DECIMAL(16,6),
		CantidadPick DECIMAL(19,6),
		BinCodeDestino NVARCHAR(228),
		CodeBars NVARCHAR(228),
		BinCodeSitu NVARCHAR(30),
		BinCodeDestinoSitu NVARCHAR(30),
		Substitute NVARCHAR(100),
		IDBulto NVARCHAR(100),
		TipoBulto NVARCHAR(100),
		comentario NVARCHAR(100)
	);
	

	FOR c_row_1 AS c_1 DO
		INSERT INTO #tmp VALUES (c_row_1.AbsEntry, c_row_1.LinPicket,c_row_1.ItemCode,c_row_1.ItemName,c_row_1.PickQtty,c_row_1.Cantidad,
		c_row_1.Udm,c_row_1.BatchNum,c_row_1.BinCode,c_row_1.propuesto,c_row_1.EsLote,c_row_1.SePuedeGestionar,c_row_1.NumPerMsr,c_row_1.CantidadPick,c_row_1.BinCodeDestino ,c_row_1.CodeBars, c_row_1.UbiORSitu,c_row_1.UbiDESitu,c_row_1.Substitute,
		c_row_1.U_EXO_LOT_ID,c_row_1.U_EXO_TBULTO,c_row_1.Comentario);
	END FOR;

	--FOR c_row_2 AS c_2 DO
	--	INSERT INTO #tmp VALUES (c_row_2.AbsEntry, c_row_2.LinPicket,c_row_2.ItemCode,c_row_2.ItemName,c_row_2.PickQtty,c_row_2.Cantidad,
	--	c_row_2.Udm,c_row_2.BatchNum,c_row_2.BinCode,c_row_2.propuesto,c_row_2.EsLote,c_row_2.SePuedeGestionar,c_row_2.NumPerMsr,c_row_2.CantidadPick,c_row_2.BinCodeDestino ,c_row_2.CodeBars, c_row_2.UbiORSitu,c_row_2.UbiDESitu,c_row_2.Substitute,
	--	c_row_2.U_EXO_LOT_ID,c_row_2.U_EXO_TBULTO);
	--END FOR;

	FOR c_row_3 AS cPicking_3 DO

	--modificar para cada linea de la solicitud buscar linea / idbulto / articulo
 		SELECT coalesce(SUM(Cantidad),0) INTO vCantidadAsignada 
 		FROM #tmp  
 		WHERE AbsEntry = :ClavePicking and PickEntry = c_row_3.PickEntry AND  IDBulto=c_row_3.U_EXO_LOT_ID;
 		
 		IF COALESCE(vCantidadAsignada,0) < c_row_3.CantLinea THEN
			vCantidadFaltante :=  c_row_3.CantLinea - vCantidadAsignada;
			
			IF c_row_3.ConLote = 'Y' THEN
				vCantidadFaltante :=  c_row_3.CantLinea - vCantidadAsignada;
				vCantidadPropuesta:= vCantidadFaltante;
				
				--llamamos a procedimiento propongo lote, que nos devuelve lote,cantidad, ubicacion
				--CALL EXO_GP_PROPONGO_LOTE(c_row_3.Almacen,c_row_3.ItemCode,vCantidadFaltante,vLotePropuesto,vCantidadPropuesta,vUbicProp);
			ELSE
				vLotePropuesto :='';
  				vCantidadPropuesta := vCantidadFaltante;
  	 			vUbicProp :='';
				--llamamos a propongo ubicacion para uqe nos devuelva la ubicacion del articulo
				--SELECT EXO_GP_PROPONGO_UBICACION(c_row_3.Almacen, c_row_3.ItemCode,'', vCantidadPropuesta, 'V') INTO vUbicProp FROM DUMMY;
			
			END IF;
			
			--hacer el insert para la tabla
			INSERT INTO #tmp VALUES (:ClavePicking, c_row_3.PickEntry,c_row_3.ItemCode,c_row_3.ItemName,vCantidadFaltante,:vCantidadPropuesta,
			c_row_3.Udm,c_row_3.Lote,c_row_3.BinCode,'Y',c_row_3.ConLote,'N',c_row_3.NumPerMsr,c_row_3.CantidadPick,c_row_3.BinCodeDestino,c_row_3.CodeBars,c_row_3.UbiORSitu,c_row_3.UbiDESitu,c_row_3.Substitute,c_row_3.U_EXO_LOT_ID,c_row_3.U_EXO_TBULTO,c_row_3.Comentario);
			
 		END IF;
 		
	END FOR;

	TablaDatos = SELECT AbsEntry,PickEntry,ItemCode,ItemName,MAX(CantidadTotal) CantidadTotal,SUM(Cantidad) Cantidad,Udm,BatchNum,BinCode,Propuesto,EsLote,SePuedeGestionar ,NumPerMsr,CantidadPick,BinCodeDestino,CodeBars,BinCodeSitu,BinCodeDestinoSitu,Substitute,IDBulto,TipoBulto,Comentario
	FROM #tmp t0 left join OBIN t1 on t0.BinCode=t1."BinCode"
	GROUP BY AbsEntry,PickEntry,ItemCode,ItemName,Udm,BatchNum,BinCode,Propuesto,EsLote,SePuedeGestionar ,NumPerMsr,CantidadPick,BinCodeDestino,CodeBars,BinCodeSitu,BinCodeDestinoSitu,Substitute,IDBulto,TipoBulto,Comentario;

	
	
	


	DROP TABLE #tmp;

END;
