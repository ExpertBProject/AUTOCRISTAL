Imports SAPbouiCOM

Public Class EXO_PNEC
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Public _Width As Integer = 328
    Public _Height As Integer = 113
    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            'Apaño por un error que da EXO_Basic.dll al consultar infoEvento.FormTypeEx
            Try
                If infoEvento.FormTypeEx <> "" Then

                End If
            Catch ex As Exception
                Return False
            End Try

            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PNEC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                    If EventHandler_FORM_RESIZE_After(infoEvento) = False Then
                                        Return False
                                    End If
                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PNEC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PNEC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PNEC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                End If
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_FORM_RESIZE_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_FORM_RESIZE_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Freeze(True)
            If pVal.ActionSuccess = True Then
                'Posicionamos campos
                oForm.Items.Item("grdALM").Width = _Width
                oForm.Items.Item("grdALM").Height = _Height
                CType(oForm.Items.Item("grdALM").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
            End If



            EventHandler_FORM_RESIZE_After = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_VALIDATE_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_VALIDATE_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Freeze(True)
            If pVal.ItemUID = "grd_DOC" Then
                Dim dCant As Double = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Order").Cells.Item(pVal.Row).Value.ToString)
                Dim sArt As String = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("EUROCODE").Cells.Item(pVal.Row).Value.ToString
                Dim sProv As String = oForm.DataSources.UserDataSources.Item("UDPROV").Value
                Dim sProvD As String = oForm.DataSources.UserDataSources.Item("UDPROVD").Value
                Dim sCatalogo As String = ""
                If pVal.ColUID.ToUpper = "ORDER" And pVal.ItemChanged = True Then
                    If dCant = 0 Then
                        sProv = ""
                    Else
                        If sProv.Trim = "" Then
                            sProv = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Cod. Prov. 1").Cells.Item(pVal.Row).Value.ToString
                            sProvD = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Proveedor 1").Cells.Item(pVal.Row).Value.ToString
                        End If
                    End If

                    If sProv <> "" Then
                        'buscamos el catálogo
                        sSQL = "SELECT ""Substitute"" FROM OSCN WHERE ""CardCode""='" & sProv & "' and ""ItemCode""='" & sArt & "'"
                        sCatalogo = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    Else
                        sCatalogo = ""
                    End If
#Region "Fecha Prevista"
                    If sProv.Trim <> "" Then
                        sSQL = "SELECT ""U_EXO_TSUM"" FROM OCRD WHERE ""CardCode""='" & sProv & "' "
                        Dim dTR As Double = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)
                        Dim dFechaPrevista As Date = New Date(Now.Year, Now.Month, Now.Day)
                        dFechaPrevista = dFechaPrevista.AddDays(dTR)
                        If oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Fecha Prev.").Cells.Item(pVal.Row).Value Is Nothing Then
                            oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Fecha Prev.").Cells.Item(pVal.Row).Value = dFechaPrevista
                        End If
                    Else
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Fecha Prev.").Cells.Item(pVal.Row).Value = Nothing
                    End If
#End Region
                    If oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(pVal.Row).Value.ToString = "" Then
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nº Catálogo").Cells.Item(pVal.Row).Value = sCatalogo
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(pVal.Row).Value = sProv
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nombre").Cells.Item(pVal.Row).Value = sProvD
                        CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
                    End If

                    Dim sALM As String = oForm.DataSources.UserDataSources.Item("UDALM").Value
                    If dCant = 0 Then
                        sALM = ""
                    End If

                    oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Alm.Destino").Cells.Item(pVal.Row).Value = sALM
                ElseIf pVal.ColUID = "Prov.Pedido" Then
                    sProv = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(pVal.Row).Value.ToString
                    If sProv <> "" Then
                        'buscamos el catálogo
                        sSQL = "SELECT ""Substitute"" FROM OSCN WHERE ""CardCode""='" & sProv & "' and ""ItemCode""='" & sArt & "'"
                        sCatalogo = objGlobal.refDi.SQL.sqlStringB1(sSQL)
#Region "Fecha Prevista"
                        If sProv.Trim <> "" Then
                            sSQL = "SELECT ""U_EXO_TSUM"" FROM OCRD WHERE ""CardCode""='" & sProv & "' "
                            Dim dTR As Double = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)
                            Dim dFechaPrevista As Date = New Date(Now.Year, Now.Month, Now.Day)
                            dFechaPrevista = dFechaPrevista.AddDays(dTR)
                            If oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Fecha Prev.").Cells.Item(pVal.Row).Value Is Nothing Then
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Fecha Prev.").Cells.Item(pVal.Row).Value = dFechaPrevista
                            End If
                        Else
                            oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Fecha Prev.").Cells.Item(pVal.Row).Value = Nothing
                        End If
#End Region
                    Else
                        sCatalogo = ""
                    End If
                    oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nº Catálogo").Cells.Item(pVal.Row).Value = sCatalogo
                End If
            ElseIf pVal.ItemUID = "txtProv" And oForm.DataSources.UserDataSources.Item("UDPROV").Value.ToString.Trim = "" Then
                oForm.DataSources.UserDataSources.Item("UDPROVD").Value = ""
            End If
            EventHandler_VALIDATE_After = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing

        EventHandler_Choose_FromList_Before = False

        Try
            oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "txtProv" Then 'Proveedores
                oCFLEvento = CType(pVal, IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions
                oCond = oConds.Add
                oCond.Alias = "CardType"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "S"
                'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            ElseIf pVal.ItemUID = "grd_DOC" And pVal.ColUID.ToUpper = "PROV.PEDIDO" Then
                oCFLEvento = CType(pVal, IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions
                oCond = oConds.Add
                oCond.Alias = "CardType"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "S"
                'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            End If

            EventHandler_Choose_FromList_Before = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oDataTable As SAPbouiCOM.DataTable = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sMensaje As String = ""
        EventHandler_Choose_FromList_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing

                Return True
            End If

            oCFLEvento = CType(pVal, IChooseFromListEvent)

            oDataTable = oCFLEvento.SelectedObjects
            If Not oDataTable Is Nothing Then
                Select Case oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).ObjectType
                    Case "64" 'Almacenes
                        Try
                            If pVal.ItemUID = "grd_DOC" And pVal.ColUID.Trim = "Alm.Destino" Then
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Alm.Destino").Cells.Item(pVal.Row).Value = oDataTable.GetValue("WhsCode", 0).ToString
                            ElseIf pVal.ItemUID = "grd_DOC" And pVal.ColUID.Trim = "Alm.Origen" Then
                                Dim dCantTraslado As Double = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Traslado").Cells.Item(pVal.Row).Value.ToString)
                                If dCantTraslado > 0 Then
                                    oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Alm.Origen").Cells.Item(pVal.Row).Value = oDataTable.GetValue("WhsCode", 0).ToString
                                Else
                                    oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Alm.Origen").Cells.Item(pVal.Row).Value = ""
                                    sMensaje = "Antes de indicar un almacén de origen, debe indicar la cantidad a trasladar."
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                End If
                            End If

                        Catch ex As Exception

                        End Try
                    Case "2" 'Proveedor
                        Try
                            If pVal.ItemUID = "txtProv" Then
                                oForm.DataSources.UserDataSources.Item("UDPROVD").ValueEx = oDataTable.GetValue("CardName", 0).ToString
                                oForm.DataSources.UserDataSources.Item("UDPROV").ValueEx = oDataTable.GetValue("CardCode", 0).ToString
                            ElseIf pVal.ItemUID = "grd_DOC" And pVal.ColUID.Trim = "Prov.Pedido" Then
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nombre").Cells.Item(pVal.Row).Value = oDataTable.GetValue("CardName", 0).ToString
                                CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(pVal.Row).Value = oDataTable.GetValue("CardCode", 0).ToString
                            End If
                        Catch ex As Exception

                        End Try
                    Case "4" ' Articulos
                        Try
                            Select Case pVal.ItemUID
                                Case "txtARTD"
                                    oForm.DataSources.UserDataSources.Item("UDARTD").ValueEx = oDataTable.GetValue("ItemCode", 0).ToString
                                Case "txtARTH"
                                    oForm.DataSources.UserDataSources.Item("UDARTH").ValueEx = oDataTable.GetValue("ItemCode", 0).ToString
                            End Select
                        Catch ex As Exception

                        End Try
                    Case "52" ' Grupo de artículos
                        Try
                            Select Case pVal.ItemUID
                                Case "txtGRUPOD"
                                    oForm.DataSources.UserDataSources.Item("UDGRUD").ValueEx = oDataTable.GetValue("ItmsGrpCod", 0).ToString
                                Case "txtGRUPOH"
                                    oForm.DataSources.UserDataSources.Item("UDGRUH").ValueEx = oDataTable.GetValue("ItmsGrpCod", 0).ToString
                            End Select
                        Catch ex As Exception

                        End Try
                End Select
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.FormDatatable(oDataTable)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByVal pVal As ItemEvent) As Boolean
#Region "variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = "" : Dim sSQLGrid As String = ""
        Dim sArtD As String = "" : Dim sArtH As String = ""
        Dim sGrupoD As String = "" : Dim sGrupoH As String = ""
        Dim sClas As String = ""
        Dim sAlmacenes As String = ""
        Dim sProveedorPR As String = "" : Dim sNomProveedorPR As String = ""
        Dim dtArticulos As Data.DataTable = Nothing
        Dim dtProveedores As Data.DataTable = Nothing
        Dim dtTarifas As Data.DataTable = Nothing
        Dim sMensaje As String = ""
        Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim dFechaAnt As Date = dFecha.AddYears(-1)
        Dim dFechaSemestre As Date = dFecha.AddMonths(-6)

        Dim expression As String = ""
        Dim sortOrder As String = ""

        Dim sError As String = ""
        Dim oOPQT As SAPbobsCOM.Documents = Nothing : Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim sFecha As String = ""
        Dim sAlmacenDestino As String = ""
        Dim bEsPrimera As Boolean = True
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oOWTQ As SAPbobsCOM.StockTransfer = Nothing
        Dim sIndicator As String = "" : Dim sSucursal As String = "" : Dim sSerie As String = ""
#End Region
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            sArtD = oForm.DataSources.UserDataSources.Item("UDARTD").ValueEx.ToString
            sArtH = oForm.DataSources.UserDataSources.Item("UDARTH").ValueEx.ToString
            sGrupoD = oForm.DataSources.UserDataSources.Item("UDGRUD").ValueEx.ToString
            sGrupoH = oForm.DataSources.UserDataSources.Item("UDGRUH").ValueEx.ToString
            sClas = oForm.DataSources.UserDataSources.Item("UDCLAS").ValueEx.ToString
            sProveedorPR = oForm.DataSources.UserDataSources.Item("UDPROV").ValueEx.ToString



            Select Case pVal.ItemUID
                Case "btnCARGAR"
                    If ComprobarALM(oForm, "DTALM") = True Then
#Region "Comprobar si ha elegido un almacen"
                        sAlmacenes = ""
                        For i As Integer = 0 To oForm.DataSources.DataTables.Item("DTALM").Rows.Count - 1
                            If oForm.DataSources.DataTables.Item("DTALM").GetValue("Sel", i).ToString = "Y" Then
                                If sAlmacenes = "" Then
                                    sAlmacenes = "'" & oForm.DataSources.DataTables.Item("DTALM").GetValue("WhsCode", i).ToString & "' "
                                Else
                                    sAlmacenes &= ", '" & oForm.DataSources.DataTables.Item("DTALM").GetValue("WhsCode", i).ToString & "' "
                                End If

                            End If
                        Next
                        If sAlmacenes = "" Then
                            sMensaje = "No ha indicado un almacén. Es obligatorio indicar uno."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                            Return False
                        End If
#End Region
                        Dim iEncuentra As Integer = 0

                        sSQLGrid = ""
                        sSQL = "SELECT OITM.""ItemCode"",OITM.""ItemName"" FROM OITM INNER JOIN OITB ON OITB.""ItmsGrpCod""=OITM.""ItmsGrpCod"" WHERE OITB.""U_EXO_GESNEC""='Si' 
                             And OITM.""QryGroup2""<>'Y' "

                        If sArtD.Trim <> "" Then
                            sSQL &= " AND OITM.""ItemCode"">='" & sArtD.Trim & "' "
                        End If
                        If sArtH.Trim <> "" Then
                            sSQL &= " AND OITM.""ItemCode""<='" & sArtH.Trim & "' "
                        End If
                        If sGrupoD.Trim <> "" Then
                            sSQL &= " AND OITM.""ItmsGrpCod"">='" & sGrupoD.Trim & "' "
                        End If
                        If sGrupoD.Trim <> "" Then
                            sSQL &= " AND OITM.""ItmsGrpCod""<='" & sGrupoH.Trim & "' "
                        End If

                        If sProveedorPR.Trim <> "" Then
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Proveedor:" & sProveedorPR, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            sSQL &= " AND OITM.""CardCode""='" & sProveedorPR.Trim & "' "
                        End If
                        dtArticulos = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                        If dtArticulos.Rows.Count > 0 Then
                            sMensaje = "Cargando datos..."
                            oForm.Freeze(True)
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
#Region "Agrupado"
                            For a As Integer = 0 To dtArticulos.Rows.Count - 1
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Artículo: " & dtArticulos.Rows(a).Item("ItemCode").ToString & " - " & dtArticulos.Rows(a).Item("ItemName").ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                dtProveedores = Nothing
                                If iEncuentra <> 0 Then
                                    sSQLGrid &= " UNION ALL "
                                End If
                                sSQLGrid &= "(Select '" & dtArticulos.Rows(a).Item("ItemCode").ToString & "' ""EUROCODE"", '" & dtArticulos.Rows(a).Item("ItemName").ToString & "' ""Descripción"" 
                                            , 0.00 ""Order"", CAST('     ' AS VARCHAR(50)) ""Prov.Pedido"", CAST('     ' AS VARCHAR(150)) ""Nombre"" 
                                            , CAST('     ' AS VARCHAR(50)) ""Nº Catálogo"",  CAST('     ' AS DATE) ""Fecha Prev."" 
                                            , 0.00 ""Traslado"", CAST('     ' AS VARCHAR(50)) ""Alm.Origen"", CAST('     ' AS VARCHAR(50)) ""Alm.Destino"" 
                                            , ((Select IFNULL(SUM(INV1.""Quantity""),0) FROM INV1 INNER JOIN OINV On INV1.""DocEntry"" = OINV.""DocEntry"" 
                                                WHERE OINV.""CANCELED""='N' and OINV.""DocDate"" <='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "' 
                                                And OINV.""DocDate"">='" & dFechaAnt.Year.ToString("0000") & dFechaAnt.Month.ToString("00") & dFechaAnt.Day.ToString("00") & "' 
                                                And INV1.""WhsCode"" in (" & sAlmacenes & ") and INV1.""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "')- 
                                            (SELECT IFNULL(SUM(RIN1.""Quantity""),0) FROM RIN1 INNER JOIN ORIN ON RIN1.""DocEntry"" = ORIN.""DocEntry"" 
                                                WHERE ORIN.""CANCELED""='N' and ORIN.""DocDate"" <='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "' 
                                                And ORIN.""DocDate"">='" & dFechaAnt.Year.ToString("0000") & dFechaAnt.Month.ToString("00") & dFechaAnt.Day.ToString("00") & "' 
                                                And RIN1.""WhsCode"" in (" & sAlmacenes & ") and RIN1.""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "')) ""24Q-" & sAlmacenes & """ 
                                            ,(SELECT SUM(IFNULL(""OnHand"",0))  FROM OITW WHERE ""WhsCode"" in (" & sAlmacenes & ") and ""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "') ""Stock " & sAlmacenes & """ 
                                            , (SELECT SUM(""OpenQty"") FROM (
                                                                                (SELECT SUM(IFNULL(""OpenQty"",0)) ""OpenQty"" FROM POR1 where ""LineStatus""<>'C' and ""WhsCode"" in (" & sAlmacenes & ") 
                                                                                    and ""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "') 
                                            UNION ALL 
                                            (Select SUM(IFNULL(""OpenQty"",0)) ""OpenQty"" FROM WTQ1 where ""LineStatus""<>'C' and ""WhsCode"" in (" & sAlmacenes & ") and ""WhsCode""<>""FromWhsCod"" 
                                                 and ""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "') ) T) ""PTE " & sAlmacenes & """ 
                                            , (CASE WHEN (SELECT SUM(IFNULL(PCH1.""Quantity"",0)) FROM PCH1 INNER JOIN OPCH ON PCH1.""DocEntry"" = OPCH.""DocEntry""  
                                                WHERE OPCH.""DocDate""<='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "' 
                                                And OPCH.""DocDate"">='" & dFechaSemestre.Year.ToString("0000") & dFechaSemestre.Month.ToString("00") & dFechaSemestre.Day.ToString("00") & "' 
                                                And PCH1.""WhsCode"" in (" & sAlmacenes & ") and PCH1.""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "')- 
                                            (SELECT SUM(IFNULL(RPC1.""Quantity"", 0)) FROM RPC1 INNER JOIN ORPC On RPC1.""DocEntry"" = ORPC.""DocEntry""  
                                            WHERE ORPC.""DocDate""<='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "' 
                                            And ORPC.""DocDate"">='" & dFechaSemestre.Year.ToString("0000") & dFechaSemestre.Month.ToString("00") & dFechaSemestre.Day.ToString("00") & "' 
                                            And RPC1.""WhsCode"" in (" & sAlmacenes & ") and RPC1.""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "')>0  
                                            And 
                                            (SELECT SUM(IFNULL(INV1.""Quantity"", 0)) FROM INV1 INNER JOIN OINV On INV1.""DocEntry"" = OINV.""DocEntry"" 
                                                WHERE OINV.""CANCELED""='N' and OINV.""DocDate"" <='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "' 
                                                And OINV.""DocDate"">='" & dFechaAnt.Year.ToString("0000") & dFechaAnt.Month.ToString("00") & dFechaAnt.Day.ToString("00") & "' 
                                                And INV1.""WhsCode"" in (" & sAlmacenes & ") and INV1.""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "')- 
                                            (SELECT SUM(IFNULL(RIN1.""Quantity"",0)) FROM RIN1 INNER JOIN ORIN ON RIN1.""DocEntry"" = ORIN.""DocEntry"" 
                                            WHERE ORIN.""CANCELED""='N' and ORIN.""DocDate"" <='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "' 
                                                And ORIN.""DocDate"">='" & dFechaAnt.Year.ToString("0000") & dFechaAnt.Month.ToString("00") & dFechaAnt.Day.ToString("00") & "' 
                                                And RIN1.""WhsCode"" in (" & sAlmacenes & ") and RIN1.""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "')=0  
                                            then 'SI' else 'NO' END) ""N " & sAlmacenes & """ "
#Region "Proveedores"
                                Dim dPedir As Double = 0
                                'Tenemos que saber si el Artículo es de Stock
                                sSQL = "SELECT ""QryGroup2"" FROM OITM WHERE ""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "'"
                                Dim sEsStock As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                If sEsStock = "Y" Then
#Region "QryGroup1='Y'"
                                    sSQLGrid &= " , CAST('    ' AS VARCHAR(50)) ""Cod. Prov. " & (1).ToString & """ "
                                    sSQLGrid &= " , CAST('STOCK' AS VARCHAR(150)) ""Proveedor " & (1).ToString & """ "
#Region "Cálculo Pedir"
                                    dPedir = 0

                                    dPedir = Calculo_Pedir(objGlobal, oForm, dFecha, dFechaAnt, dtArticulos.Rows(a).Item("ItemCode").ToString, sAlmacenes, "")
                                    If dPedir < 0 Then
                                        dPedir = 0
                                    End If
                                    sSQLGrid &= ", '" & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, Math.Round(dPedir, 2, MidpointRounding.ToEven), EXO_GLOBALES.FuenteInformacion.Otros) & "' ""Pedir " & (1).ToString & "_" & sAlmacenes & """ "
#End Region
                                    '2º Proveedor es el que tenga por defecto
                                    sSQL = "SELECT TOP 1 OITM.""CardCode"", OCRD.""CardName"" FROM OITM INNER JOIN OCRD ON OITM.""CardCode""=OCRD.""CardCode"" "
                                    sSQL &= " WHERE OITM.""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "' "
                                    dtProveedores = Nothing : dtProveedores = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                                    If dtProveedores.Rows.Count > 0 Then
                                        For t = 1 To 1
                                            Try
                                                sSQLGrid &= " ,'" & dtProveedores.Rows(t).Item("CardCode").ToString & "' ""Cod. Prov. " & (t + 1).ToString & """ "
                                                sSQLGrid &= " ,'" & dtProveedores.Rows(t).Item("CardName").ToString & "' ""Proveedor " & (t + 1).ToString & """ "
#Region "Cálculo Pedir"
                                                dPedir = 0

                                                dPedir = Calculo_Pedir(objGlobal, oForm, dFecha, dFechaAnt, dtArticulos.Rows(a).Item("ItemCode").ToString, sAlmacenes, dtProveedores.Rows(t).Item("CardCode").ToString)
                                                If dPedir < 0 Then
                                                    dPedir = 0
                                                End If
                                                sSQLGrid &= ", '" & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, Math.Round(dPedir, 2, MidpointRounding.ToEven), EXO_GLOBALES.FuenteInformacion.Otros) & "' ""Pedir " & (t + 1).ToString & "_" & sAlmacenes & """ "
#End Region
                                            Catch ex As Exception
                                                sSQLGrid &= " , CAST(' ' AS VARCHAR(50)) ""Cod. Prov. " & (t + 1).ToString & """ "
                                                sSQLGrid &= " , CAST(' ' AS VARCHAR(150)) ""Proveedor " & (t + 1).ToString & """ "
#Region "Cálculo Pedir"
                                                dPedir = 0
                                                dPedir = Calculo_Pedir(objGlobal, oForm, dFecha, dFechaAnt, dtArticulos.Rows(a).Item("ItemCode").ToString, sAlmacenes, "")
                                                If dPedir < 0 Then
                                                    dPedir = 0
                                                End If
                                                sSQLGrid &= ", '" & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, Math.Round(dPedir, 2, MidpointRounding.ToEven), EXO_GLOBALES.FuenteInformacion.Otros) & "' ""Pedir " & (t + 1).ToString & "_" & sAlmacenes & """ "
#End Region
                                            End Try

                                        Next
                                    Else
                                        For t = 1 To 1
                                            sSQLGrid &= " , CAST(' ' AS VARCHAR(50)) ""Cod. Prov. " & (t + 1).ToString & """ "
                                            sSQLGrid &= " , CAST(' ' AS VARCHAR(150)) ""Proveedor " & (t + 1).ToString & """ "
#Region "Cálculo Pedir"
                                            dPedir = 0
                                            dPedir = Calculo_Pedir(objGlobal, oForm, dFecha, dFechaAnt, dtArticulos.Rows(a).Item("ItemCode").ToString, sAlmacenes, "")
                                            If dPedir < 0 Then
                                                dPedir = 0
                                            End If
                                            sSQLGrid &= ", '" & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, Math.Round(dPedir, 2, MidpointRounding.ToEven), EXO_GLOBALES.FuenteInformacion.Otros) & "' ""Pedir " & (t + 1).ToString & "_" & sAlmacenes & """ "
#End Region
                                        Next
                                    End If
#End Region

                                Else
#Region "Art QryGroup1=N"
                                    sSQL = "Select TOP 2 OCRD.""CardName"" ""NomProveedor"", ITM2.* FROM ITM2 INNER JOIN OCRD On OCRD.""CardCode""=ITM2.""VendorCode"" "
                                    sSQL &= " WHERE ITM2.""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "' "
                                    dtProveedores = Nothing : dtProveedores = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                                    If dtProveedores.Rows.Count > 0 Then
                                        For t = 0 To 1
                                            Try
                                                sSQLGrid &= " ,'" & dtProveedores.Rows(t).Item("VendorCode").ToString & "' ""Cod. Prov. " & (t + 1).ToString & """ "
                                                sSQLGrid &= " ,'" & dtProveedores.Rows(t).Item("NomProveedor").ToString & "' ""Proveedor " & (t + 1).ToString & """ "
#Region "Cálculo Pedir"
                                                dPedir = 0
                                                dPedir = Calculo_Pedir_Agrupado(objGlobal, oForm, dFecha, dFechaAnt, dtArticulos.Rows(a).Item("ItemCode").ToString, sAlmacenes, dtProveedores.Rows(t).Item("VendorCode").ToString)
                                                If dPedir < 0 Then
                                                    dPedir = 0
                                                End If
                                                sSQLGrid &= ", '" & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, Math.Round(dPedir, 2, MidpointRounding.ToEven), EXO_GLOBALES.FuenteInformacion.Otros) & "' ""Pedir " & (t + 1).ToString & "_" & sAlmacenes & """ "
#End Region
                                            Catch ex As Exception
                                                sSQLGrid &= " , CAST(' ' AS VARCHAR(50)) ""Cod. Prov. " & (t + 1).ToString & """ "
                                                sSQLGrid &= " , CAST(' ' AS VARCHAR(150)) ""Proveedor " & (t + 1).ToString & """ "
#Region "Cálculo Pedir"
                                                dPedir = 0

                                                dPedir = Calculo_Pedir_Agrupado(objGlobal, oForm, dFecha, dFechaAnt, dtArticulos.Rows(a).Item("ItemCode").ToString, sAlmacenes, "")
                                                If dPedir < 0 Then
                                                    dPedir = 0
                                                End If
                                                sSQLGrid &= ", '" & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, Math.Round(dPedir, 2, MidpointRounding.ToEven), EXO_GLOBALES.FuenteInformacion.Otros) & "' ""Pedir " & (t + 1).ToString & "_" & sAlmacenes & """ "
#End Region
                                            End Try

                                        Next
                                    Else
                                        For t = 0 To 1
                                            sSQLGrid &= " , CAST(' ' AS VARCHAR(50)) ""Cod. Prov. " & (t + 1).ToString & """ "
                                            sSQLGrid &= " , CAST(' ' AS VARCHAR(150)) ""Proveedor " & (t + 1).ToString & """ "
#Region "Cálculo Pedir"
                                            dPedir = 0
                                            dPedir = Calculo_Pedir_Agrupado(objGlobal, oForm, dFecha, dFechaAnt, dtArticulos.Rows(a).Item("ItemCode").ToString, sAlmacenes, "")
                                            If dPedir < 0 Then
                                                dPedir = 0
                                            End If
                                            sSQLGrid &= ", '" & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, Math.Round(dPedir, 2, MidpointRounding.ToEven), EXO_GLOBALES.FuenteInformacion.Otros) & "' ""Pedir " & (t + 1).ToString & "_" & sAlmacenes & """ "
#End Region
                                        Next
                                    End If
#End Region
                                End If
#End Region

#Region "Tarifas"
                                sSQL = "SELECT IFNULL(ITM1.""Price"",0) ""Precio"", OPLN.""ListNum"", OPLN.""ListName"" FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum"" 
                                        WHERE ITM1.""ItemCode""='" & dtArticulos.Rows(a).Item("ItemCode").ToString & "' 
                                        And OPLN.""U_EXO_TARCOM""='Si' "
                                dtTarifas = Nothing : dtTarifas = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                                For t = 0 To dtTarifas.Rows.Count - 1
                                    sSQLGrid &= ", " & dtTarifas.Rows(t).Item("Precio").ToString.Replace(",", ".") & " ""Tarifa " & dtTarifas.Rows(t).Item("ListName").ToString & """ "
                                Next
#End Region

                                sSQLGrid &= " FROM DUMMY) "
                                iEncuentra += 1
                            Next
#End Region
                        Else
                            sMensaje = "No existen artículos con Gestión de necesidades."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                            Return False
                        End If
                        oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQLGrid)
                        FormateaGridDOC(oForm)
                        sMensaje = "Fin de la carga de datos."
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                Case "btnGen"
                    oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    If objGlobal.SBOApp.MessageBox("¿Esta seguro de generar los documentos según su parametrización?", 1, "Sí", "No") = 1 Then
                        If oForm.DataSources.DataTables.Item("DT_DOC").Rows.Count > 0 Then
                            Dim dt As SAPbouiCOM.DataTable = Nothing : Dim dtDatos As New System.Data.DataTable
                            oForm.Freeze(True)
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Generando Documentos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            dt = Nothing : dt = oForm.DataSources.DataTables.Item("DT_DOC")

                            'Añadimos Columnas
                            For iCol As Integer = 0 To 9
                                dtDatos.Columns.Add(dt.Columns.Item(iCol).Name)
                            Next

                            'Añadimos los registros
                            For iRow As Integer = 0 To dt.Rows.Count - 1
                                Dim oRow As DataRow = dtDatos.NewRow
                                For iCol As Integer = 0 To 9
                                    oRow.Item(dt.Columns.Item(iCol).Name) = dt.Columns.Item(iCol).Cells.Item(iRow).Value
                                Next
                                dtDatos.Rows.Add(oRow)
                            Next
#Region "Solicitud de pedido"
#Region "Filtro y Orden"
                            Dim dtSolPedido As New System.Data.DataTable

                            expression = "Order>0 and Prov.Pedido<>'' and Alm.Destino<>'' "
                            sortOrder = "Prov.Pedido, Alm.Destino ASC"

                            Try
                                dtSolPedido = dtDatos.Select(expression, sortOrder).CopyToDataTable()
                            Catch ex As Exception

                            End Try

#End Region
                            Dim sProvPedido As String = ""
                            If dtSolPedido.Rows.Count > 0 Then
                                Dim bPlinea As Boolean = True
                                If objGlobal.compañia.InTransaction = True Then
                                    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                                objGlobal.compañia.StartTransaction()
                                bEsPrimera = True
                                For Each MiDataRow As DataRow In dtSolPedido.Rows
                                    If sProvPedido <> MiDataRow("Prov.Pedido").ToString Or sAlmacenDestino <> MiDataRow("Alm.Destino").ToString Then
                                        If bEsPrimera = False Then
                                            If oOPQT.Add() <> 0 Then
                                                sError = objGlobal.compañia.GetLastErrorCode.ToString & " / " & objGlobal.compañia.GetLastErrorDescription.Replace("'", "")
                                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Else
                                                objGlobal.compañia.GetNewObjectCode(sDocEntry)
                                                sSQL = "SELECT ""DocNum"" FROM OPQT WHERE ""DocEntry""=" & sDocEntry
                                                oRs.DoQuery(sSQL)

                                                If oRs.RecordCount > 0 Then
                                                    sDocNum = oRs.Fields.Item("DocNum").Value.ToString
                                                Else
                                                    sDocNum = ""
                                                End If
                                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se ha generado la Sol. de pedido Nº " & sDocNum & " para el proveedor " & sProvPedido, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            End If
                                        End If
                                        bEsPrimera = False : bPlinea = True
                                        sProvPedido = MiDataRow("Prov.Pedido").ToString : sAlmacenDestino = MiDataRow("Alm.Destino").ToString
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Generando Sol. de Pedido para el Proveedor " & sProvPedido & " y almacén destino " & sAlmacenDestino, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        oOPQT = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations), SAPbobsCOM.Documents)
#Region "Búsqueda de la serie"

                                        sIndicator = Now.Year.ToString("0000") & "-" & Now.Month.ToString("00")
                                        sSQL = "SELECT ""Indicator"" FROM OFPR WHERE ""Code""='" & sIndicator & "' "
                                        sIndicator = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                        sSQL = "SELECT ""U_EXO_SUCURSAL"" FROM ""OWHS"" WHERE ""WhsCode""='" & sAlmacenDestino & "' "
                                        sSucursal = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                        sSQL = "SELECT ""Series"" FROM ""NNM1"" WHERE ""Indicator""='" & sIndicator & "' and ""Remark""='" & sSucursal & "' and ""ObjectCode""='540000006' "
                                        sSerie = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                        If sSerie = "" Then
                                            If objGlobal.SBOApp.MessageBox("No encuentra la serie del almacén " & sAlmacenDestino & "¿Continuamos con la serie primaria?", 1, "Sí", "No") = 1 Then
                                                sSQL = "SELECT ""Series"" FROM ""NNM1"" WHERE ""SeriesName""='Primario' and ""ObjectCode""='540000006' "
                                                sSerie = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                                oOPQT.Series = CType(sSerie, Integer)
                                            Else
                                                sMensaje = "El usuario ha cancelado el proceso al no encontrar la serie correspondiente al almacén " & sAlmacenDestino
                                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                objGlobal.SBOApp.MessageBox(sMensaje)
                                                Return False
                                            End If
                                        Else
                                            oOPQT.Series = CType(sSerie, Integer)
                                        End If
#End Region
                                        oOPQT.CardCode = sProvPedido
                                        sFecha = MiDataRow("Fecha Prev.").ToString : dFecha = CDate(sFecha)
                                        oOPQT.RequriedDate = dFecha
                                        oOPQT.Comments = "Documento creado desde planificador de necesidades"
                                    End If
                                    If bPlinea = False Then
                                        oOPQT.Lines.Add()
                                    Else
                                        bPlinea = False
                                    End If
                                    If MiDataRow("Nº Catálogo").ToString.Trim <> "" Then
                                        oOPQT.Lines.SupplierCatNum = MiDataRow("Nº Catálogo").ToString
                                    Else
                                        oOPQT.Lines.ItemCode = MiDataRow("EUROCODE").ToString
                                    End If
                                    oOPQT.Lines.Quantity = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, MiDataRow("Order").ToString)
                                    oOPQT.Lines.RequiredQuantity = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, MiDataRow("Order").ToString)
                                    sFecha = MiDataRow("Fecha Prev.").ToString : dFecha = CDate(sFecha)
                                    oOPQT.Lines.RequiredDate = dFecha
                                    oOPQT.Lines.WarehouseCode = sAlmacenDestino
                                Next
                                If oOPQT.Add() <> 0 Then
                                    sError = objGlobal.compañia.GetLastErrorCode.ToString & " / " & objGlobal.compañia.GetLastErrorDescription.Replace("'", "")
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    objGlobal.compañia.GetNewObjectCode(sDocEntry)
                                    sSQL = "SELECT ""DocNum"" FROM OPQT WHERE ""DocEntry""=" & sDocEntry
                                    oRs.DoQuery(sSQL)

                                    If oRs.RecordCount > 0 Then
                                        sDocNum = oRs.Fields.Item("DocNum").Value.ToString
                                    Else
                                        sDocNum = ""
                                    End If
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se ha generado la Sol. de pedido Nº " & sDocNum & " para el proveedor " & sProvPedido, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                                If objGlobal.compañia.InTransaction = True Then
                                    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                End If
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - No existen datos para generar Solicitud de pedido.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If


#End Region
#Region "Solicitud de traslado"
#Region "Filtro y Orden"
                            Dim dtSolTraslado As New System.Data.DataTable

                            expression = "Traslado>0 and Alm.Origen<>'' and Alm.Destino<>'' "
                            sortOrder = "Alm.Origen, Alm.Destino ASC"

                            Try
                                dtSolTraslado = dtDatos.Select(expression, sortOrder).CopyToDataTable()
                            Catch ex As Exception

                            End Try
#End Region
                            Dim sAlmOrigen As String = "" : Dim sAlmDestino As String = ""
                            If dtSolTraslado.Rows.Count > 0 Then
                                Dim bPlinea As Boolean = True
                                If objGlobal.compañia.InTransaction = True Then
                                    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                                objGlobal.compañia.StartTransaction()
                                bEsPrimera = True
                                For Each MiDataRow As DataRow In dtSolTraslado.Rows
                                    If sAlmOrigen <> MiDataRow("Alm.origen").ToString Or sAlmDestino <> MiDataRow("Alm.Destino").ToString Then
                                        If bEsPrimera = False Then
                                            If oOWTQ.Add() <> 0 Then
                                                sError = objGlobal.compañia.GetLastErrorCode.ToString & " / " & objGlobal.compañia.GetLastErrorDescription.Replace("'", "")
                                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Else
                                                objGlobal.compañia.GetNewObjectCode(sDocEntry)
                                                sSQL = "SELECT ""DocNum"" FROM OWTQ WHERE ""DocEntry""=" & sDocEntry
                                                oRs.DoQuery(sSQL)

                                                If oRs.RecordCount > 0 Then
                                                    sDocNum = oRs.Fields.Item("DocNum").Value.ToString
                                                Else
                                                    sDocNum = ""
                                                End If
                                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se ha generado la Sol. de Traslado Nº " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            End If
                                        End If
                                        bEsPrimera = False : bPlinea = True
                                        sAlmOrigen = MiDataRow("Alm.Origen").ToString : sAlmDestino = MiDataRow("Alm.Destino").ToString
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Generando Sol. de Traslado del almacén de origen " & sAlmOrigen & " y almacén destino " & sAlmDestino, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        oOWTQ = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
#Region "Búsqueda de la serie"
                                        sIndicator = Now.Year.ToString("0000") & "-" & Now.Month.ToString("00")
                                        sSQL = "SELECT ""Indicator"" FROM OFPR WHERE ""Code""='" & sIndicator & "' "
                                        sIndicator = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                        sSQL = "SELECT ""U_EXO_SUCURSAL"" FROM ""OWHS"" WHERE ""WhsCode""='" & sAlmOrigen & "' "
                                        sSucursal = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                        sSQL = "SELECT ""Series"" FROM ""NNM1"" WHERE ""Indicator""='" & sIndicator & "' and ""Remark""='" & sSucursal & "' and ""ObjectCode""='1250000001' "
                                        sSerie = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                        If sSerie = "" Then
                                            If objGlobal.SBOApp.MessageBox("No encuentra la serie del almacén " & sAlmOrigen & "¿Continuamos con la serie primaria?", 1, "Sí", "No") = 1 Then
                                                sSQL = "SELECT ""Series"" FROM ""NNM1"" WHERE ""SeriesName""='Primario' and ""ObjectCode""='1250000001' "
                                                sSerie = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                                oOWTQ.Series = CType(sSerie, Integer)
                                            Else
                                                sMensaje = "El usuario ha cancelado el proceso al no encontrar la serie correspondiente al almacén " & sAlmOrigen
                                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                objGlobal.SBOApp.MessageBox(sMensaje)
                                                Return False
                                            End If
                                        Else
                                            oOWTQ.Series = CType(sSerie, Integer)
                                        End If
#End Region
                                        oOWTQ.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest
                                        oOWTQ.DocDate = New Date(Now.Year, Now.Month, Now.Day)
                                        oOWTQ.TaxDate = New Date(Now.Year, Now.Month, Now.Day)
                                        oOWTQ.DueDate = New Date(Now.Year, Now.Month, Now.Day)
                                        oOWTQ.FromWarehouse = sAlmOrigen
                                        oOWTQ.ToWarehouse = sAlmDestino
                                        oOWTQ.UserFields.Fields.Item("U_EXO_TIPO").Value = "ITC"
                                        oOWTQ.UserFields.Fields.Item("U_EXO_STATUSP").Value = "P"
                                        oOWTQ.Comments = "Documento creado desde planificador de necesidades"
                                    End If
                                    If bPlinea = False Then
                                        oOWTQ.Lines.Add()
                                    Else
                                        bPlinea = False
                                    End If
                                    oOWTQ.Lines.ItemCode = MiDataRow("EUROCODE").ToString
                                    oOWTQ.Lines.Quantity = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, MiDataRow("Traslado").ToString)
                                Next
                                If oOWTQ.Add() <> 0 Then
                                    sError = objGlobal.compañia.GetLastErrorCode.ToString & " / " & objGlobal.compañia.GetLastErrorDescription.Replace("'", "")
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    objGlobal.compañia.GetNewObjectCode(sDocEntry)
                                    sSQL = "SELECT ""DocNum"" FROM OWTQ WHERE ""DocEntry""=" & sDocEntry
                                    oRs.DoQuery(sSQL)

                                    If oRs.RecordCount > 0 Then
                                        sDocNum = oRs.Fields.Item("DocNum").Value.ToString
                                    Else
                                        sDocNum = ""
                                    End If
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se ha generado la Sol. de traslado Nº " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                                If objGlobal.compañia.InTransaction = True Then
                                    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                End If
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - No existen datos para generar Solicitud de traslado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If
#End Region
                            sMensaje = "Fin del proceso."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                        Else
                            sMensaje = "No existen registros para generar documentos."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                        End If

                    Else
                        sMensaje = "El usuario ha cancelado la generación de documentos."
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox(sMensaje)
                    End If

            End Select

            EventHandler_ItemPressed_After = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOPQT, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOWTQ, Object))
        End Try
    End Function
    Private Function Calculo_Pedir(ByRef oObjglobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form, ByVal dFecha As Date, ByVal dFechaAnt As Date,
                                                 ByVal sArticulo As String, ByVal sAlmacenes As String, ByVal sProveedor As String) As Double
        Calculo_Pedir = 0
        Dim sSQL As String = ""
        Try
            sSQL = "SELECT ((SELECT IFNULL(SUM(INV1.""Quantity""),0) FROM INV1 INNER JOIN OINV ON INV1.""DocEntry"" = OINV.""DocEntry"" "
            sSQL &= " WHERE OINV.""DocDate"" <='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "' "
            sSQL &= " And OINV.""DocDate"">='" & dFechaAnt.Year.ToString("0000") & dFechaAnt.Month.ToString("00") & dFechaAnt.Day.ToString("00") & "' "
            sSQL &= " And INV1.""WhsCode"" in (" & sAlmacenes & ") and INV1.""ItemCode""='" & sArticulo & "')- "
            sSQL &= " (SELECT IFNULL(SUM(RIN1.""Quantity""),0) FROM RIN1 INNER JOIN ORIN ON RIN1.""DocEntry"" = ORIN.""DocEntry"" "
            sSQL &= " WHERE ORIN.""DocDate"" <='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "' "
            sSQL &= " And ORIN.""DocDate"">='" & dFechaAnt.Year.ToString("0000") & dFechaAnt.Month.ToString("00") & dFechaAnt.Day.ToString("00") & "' "
            sSQL &= " And RIN1.""WhsCode"" in (" & sAlmacenes & ") and RIN1.""ItemCode""='" & sArticulo & "')) "
            sSQL &= " FROM DUMMY "
            Dim d24Q As Double = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)
            sSQL = "SELECT IFNULL(""OnHand""+""OnOrder"",0)  FROM OITW WHERE ""WhsCode"" in (" & sAlmacenes & ") and ""ItemCode""='" & sArticulo & "' "
            Dim dSTOCK As Double = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)
            sSQL = "SELECT IFNULL(SUM(""OpenQty""),0) FROM POR1 where ""LineStatus""<>'C' and ""WhsCode""in (" & sAlmacenes & ")  and ""ItemCode""='" & sArticulo & "'"
            Dim dPTE As Double = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)

            If d24Q > 3 And (dSTOCK + dPTE) < 2 Then
                Calculo_Pedir = 2 - dSTOCK - dPTE
            Else
                Dim dTR As Double = 0
                If sProveedor.Trim <> "" Then
                    sSQL = "SELECT ""U_EXO_TSUM"" FROM OCRD WHERE ""CardCode""='" & sProveedor & "' "
                    dTR = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)
                End If


                Dim dQ As Double = d24Q / 12
                Dim dSTOCK_m As Double = dQ * 2
                Dim dSTOCK_S As Double = d24Q / 12
                Dim dMGS As Double = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.UserDataSources.Item("UDDIAS").ValueEx.ToString)
                Dim dPP As Double = dQ * ((dMGS / 15) + dTR)
                Calculo_Pedir = dSTOCK_m + dSTOCK_S + dPP - dPTE - dSTOCK
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function Calculo_Pedir_Agrupado(ByRef oObjglobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form, ByVal dFecha As Date, ByVal dFechaAnt As Date,
                                                 ByVal sArticulo As String, ByVal sAlmacenes As String, ByVal sProveedor As String) As Double
        Calculo_Pedir_Agrupado = 0
        Dim sSQL As String = ""
        Try
            sSQL = "SELECT IFNULL(SUM(INV1.""OpenQty""),0) FROM INV1 INNER JOIN OINV ON INV1.""DocEntry"" = OINV.""DocEntry"" "
            sSQL &= " WHERE OINV.""DocStatus""<>'C' and OINV.""DocDate""<='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "' "
            sSQL &= " And OINV.""DocDate"">='" & dFechaAnt.Year.ToString("0000") & dFechaAnt.Month.ToString("00") & dFechaAnt.Day.ToString("00") & "' "
            sSQL &= " And INV1.""WhsCode"" in (" & sAlmacenes & ") and INV1.""ItemCode""='" & sArticulo & "' "
            Dim d24Q As Double = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)
            sSQL = "SELECT IFNULL(""OnHand""+""OnOrder"",0)  FROM OITW WHERE ""WhsCode"" in (" & sAlmacenes & ") and ""ItemCode""='" & sArticulo & "' "
            Dim dSTOCK As Double = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)
            sSQL = "SELECT IFNULL(SUM(""OpenQty""),0) FROM POR1 where ""LineStatus""<>'C' and ""WhsCode"" in (" & sAlmacenes & ") and ""ItemCode""='" & sArticulo & "'"
            Dim dPTE As Double = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)

            If d24Q > 3 And (dSTOCK + dPTE) < 2 Then
                Calculo_Pedir_Agrupado = 2 - dSTOCK - dPTE
            Else
                Dim dTR As Double = 0
                If sProveedor.Trim <> "" Then
                    sSQL = "SELECT ""U_EXO_TSUM"" FROM OCRD WHERE ""CardCode""='" & sProveedor & "' "
                    dTR = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)
                End If


                Dim dQ As Double = d24Q / 12
                Dim dSTOCK_m As Double = dQ * 2
                Dim dSTOCK_S As Double = dSTOCK
                Dim dMGS As Double = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.UserDataSources.Item("UDDIAS").ValueEx.ToString)
                Dim dPP As Double = dQ * ((dMGS / 15) + dTR)
                Calculo_Pedir_Agrupado = dSTOCK_m + dSTOCK_S + dPP - dPTE - dSTOCK
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ComprobarALM(ByRef oForm As SAPbouiCOM.Form, ByVal sFra As String) As Boolean
        Dim bLineasSel As Boolean = False
        Dim sMensaje As String = ""
        ComprobarALM = False

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sFra).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sFra).GetValue("Sel", i).ToString = "Y" Then
                    bLineasSel = True
                    Exit For
                End If
            Next

            If bLineasSel = False Then
                sMensaje = "Tiene que seleccionar al menos un almacén. Revise el filtro."
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox(sMensaje)
                Exit Function
            End If

            ComprobarALM = bLineasSel

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub FormateaGridDOC(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Dim iColumnas As Integer = 1
        Try
            oform.Freeze(True)
            iColumnas = CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Count
            'CType(oform.Items.Item("grdALM").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            'oColumnChk = CType(CType(oform.Items.Item("grdALM").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            'oColumnChk.Editable = True
            'oColumnChk.Width = 30

            For i = 0 To iColumnas - 1
                CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                Dim sTitulo As String = CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).TitleObject.Caption.ToString.ToUpper
                If sTitulo.Contains("EUROCODE") Then
                    oColumnTxt.LinkedObjectType = "4"
                ElseIf sTitulo.Contains("ORDER") Or sTitulo.Contains("TRASLADO") Then
                    oColumnTxt.RightJustified = True
                    oColumnTxt.Editable = True
                ElseIf sTitulo.Contains("PROV.PEDIDO") Then
                    oColumnTxt.Editable = True
                    oColumnTxt.ChooseFromListUID = "CFLPPED"
                    oColumnTxt.ChooseFromListAlias = "CardCode"
                    oColumnTxt.LinkedObjectType = "2"
                ElseIf sTitulo.Contains("ALM.ORIGEN") Then
                    oColumnTxt.Editable = True
                    oColumnTxt.ChooseFromListUID = "CFLALMO"
                    oColumnTxt.ChooseFromListAlias = "WhsCode"
                ElseIf sTitulo.Contains("ALM.DESTINO") Then
                    oColumnTxt.Editable = True
                    oColumnTxt.ChooseFromListUID = "CFLALMD"
                    oColumnTxt.ChooseFromListAlias = "WhsCode"
                ElseIf Left(sTitulo, 3) = "24Q" Then
                    oColumnTxt.RightJustified = True
                ElseIf sTitulo.Contains("STOCK") Or sTitulo.Contains("PTE") Or sTitulo.Contains("TARIFA") Or sTitulo.Contains("PEDIR") Then
                    oColumnTxt.RightJustified = True
                ElseIf Left(sTitulo, 2) = "N " Then
                    oColumnTxt.RightJustified = True
                ElseIf sTitulo.Contains("FECHA PREV.") Then
                    oColumnTxt.Editable = True
                End If
            Next



            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        SBOApp_MenuEvent = True
        Dim sSQL As String = ""
        Try
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnPNEC"
                        If CargarForm() = False Then
                            Exit Function
                        End If
                End Select
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally

        End Try
    End Function
    Public Function CargarForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)

        CargarForm = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_PNEC.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un Error interno") = True Then 'Falta de autorización
                    Exit Function
                Else
                    objGlobal.SBOApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                End If
            End Try
            oForm.DataSources.UserDataSources.Item("UDCLAS").ValueEx = "A"
            sSQL = "Select 'Y' as ""Sel"", ""WhsCode"", ""WhsName"" FROM ""OWHS"" order by ""WhsCode"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DTALM").ExecuteQuery(sSQL)
            FormateaGridALM(oForm)
            oForm.DataSources.UserDataSources.Item("UDDIAS").ValueEx = "7"
            oForm.Items.Item("btnGen").Enabled = False

            CargarForm = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub FormateaGridALM(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grdALM").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdALM").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 1 To 2
                CType(oform.Items.Item("grdALM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grdALM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
            Next



            CType(oform.Items.Item("grdALM").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
End Class
