Imports SAPbouiCOM

Public Class EXO_PNEC
    Private objGlobal As EXO_UIAPI.EXO_UIAPI

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
                oForm.Items.Item("grdALM").Width = INICIO._WidthALM
                oForm.Items.Item("grdALM").Height = INICIO._HeightALM
                CType(oForm.Items.Item("grdALM").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

                oForm.Items.Item("grdGRU").Width = INICIO._WidthGRU
                oForm.Items.Item("grdGRU").Height = INICIO._HeightGRU
                CType(oForm.Items.Item("grdGRU").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

                oForm.Items.Item("grdCLAS").Width = INICIO._WidthCLAS
                oForm.Items.Item("grdCLAS").Height = INICIO._HeightCLAS
                CType(oForm.Items.Item("grdCLAS").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

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
                            sProv = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(pVal.Row).Value.ToString
                            sProvD = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nombre").Cells.Item(pVal.Row).Value.ToString
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
                    Dim sALM As String = "" : Dim bVarios As Boolean = False
                    For i As Integer = 0 To oForm.DataSources.DataTables.Item("DTALM").Rows.Count - 1
                        If oForm.DataSources.DataTables.Item("DTALM").GetValue("Sel", i).ToString = "Y" Then
                            If sALM = "" Then
                                sALM = "'" & oForm.DataSources.DataTables.Item("DTALM").GetValue("Cod.", i).ToString & "' "
                                bVarios = False
                            Else
                                sALM &= ", '" & oForm.DataSources.DataTables.Item("DTALM").GetValue("Cod.", i).ToString & "' "
                                bVarios = True
                            End If

                        End If
                    Next

                    If dCant = 0 Then
                        sALM = ""
                    End If
                    If bVarios = True Then
                        sALM = ""
                    End If
                    oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Alm.Destino").Cells.Item(pVal.Row).Value = sALM.Replace("'", "")
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
        Dim sSQL As String = ""
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
                                'Buscamos el tiempo de suministro
                                ssql = "SELECT ""U_EXO_TSUM"" FROM OCRD WHERE ""CardCode""='" & oDataTable.GetValue("CardCode", 0).ToString & "' "
                                oForm.DataSources.UserDataSources.Item("UDTSM").ValueEx = objGlobal.refDi.SQL.sqlStringB1(sSQL)
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
        Dim sAlmacenes As String = "" : Dim sGrupos As String = "" : Dim sCLAS As String = ""
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
        Dim sTSUM As String = "" : Dim sMGS As String = ""
#End Region
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            sArtD = oForm.DataSources.UserDataSources.Item("UDARTD").ValueEx.ToString
            sArtH = oForm.DataSources.UserDataSources.Item("UDARTH").ValueEx.ToString

            sCLAS = ""
            For i As Integer = 0 To oForm.DataSources.DataTables.Item("DTCLAS").Rows.Count - 1
                If oForm.DataSources.DataTables.Item("DTCLAS").GetValue("Sel", i).ToString = "Y" Then
                    If sCLAS = "" Then
                        sCLAS = "'" & oForm.DataSources.DataTables.Item("DTCLAS").GetValue("Clas", i).ToString & "' "
                    Else
                        sCLAS &= ", '" & oForm.DataSources.DataTables.Item("DTCLAS").GetValue("Clas", i).ToString & "' "
                    End If
                End If
            Next

            sGrupos = ""
            For i As Integer = 0 To oForm.DataSources.DataTables.Item("DTGRU").Rows.Count - 1
                If oForm.DataSources.DataTables.Item("DTGRU").GetValue("Sel", i).ToString = "Y" Then
                    If sGrupos = "" Then
                        sGrupos = "'" & oForm.DataSources.DataTables.Item("DTGRU").GetValue("Cod.", i).ToString & "' "
                    Else
                        sGrupos &= ", '" & oForm.DataSources.DataTables.Item("DTGRU").GetValue("Cod.", i).ToString & "' "
                    End If
                End If
            Next

            sProveedorPR = oForm.DataSources.UserDataSources.Item("UDPROV").ValueEx.ToString
            sTSUM = oForm.DataSources.UserDataSources.Item("UDTSM").ValueEx.ToString
            sMGS = oForm.DataSources.UserDataSources.Item("UDDIAS").ValueEx.ToString

            Select Case pVal.ItemUID
                Case "grd_DOC"
                    If pVal.ColUID = "Sel." Then
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Filtrando datos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Dim dt As SAPbouiCOM.DataTable = Nothing
                        dt = Nothing : dt = oForm.DataSources.DataTables.Item("DT_DOC")

                        If dt.Columns.Item(0).Cells.Item(pVal.Row).Value.ToString = "Y" Then
                            Dim oRow As DataRow = INICIO._dtDatos.NewRow
                            For iCol As Integer = 0 To 12
                                oRow.Item(dt.Columns.Item(iCol).Name) = dt.Columns.Item(iCol).Cells.Item(pVal.Row).Value
                            Next
                            oRow.Item("ROW") = pVal.Row
                            INICIO._dtDatos.Rows.Add(oRow)
                        Else
                            INICIO._dtDatos.Rows.Remove(INICIO._dtDatos.Rows.Find(New Object() {pVal.Row}))
                        End If
                    End If
                Case "btnCARGAR"
                    If ComprobarALM(oForm, "DTALM") = True Then
#Region "Comprobar si ha elegido un almacen"
                        sAlmacenes = ""
                        For i As Integer = 0 To oForm.DataSources.DataTables.Item("DTALM").Rows.Count - 1
                            If oForm.DataSources.DataTables.Item("DTALM").GetValue("Sel", i).ToString = "Y" Then
                                If sAlmacenes = "" Then
                                    sAlmacenes = "'" & oForm.DataSources.DataTables.Item("DTALM").GetValue("Cod.", i).ToString & "' "
                                Else
                                    sAlmacenes &= ", '" & oForm.DataSources.DataTables.Item("DTALM").GetValue("Cod.", i).ToString & "' "
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

                        sMensaje = "Cargando datos..."
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        sSQLGrid = "Select 'N' ""Sel."", T0.""ItemCode"" ""EUROCODE"", T0.""ItemName"" ""Descripción"",  T0.""ItmsGrpCod"" ""Grupo"", 
IFNULL(((T1.""Ventas_24Q""/12)*2)+(T1.""Ventas_24Q""/12) +((T1.""Ventas_24Q""/12)*((" & sMGS & "/15)+" & sTSUM & ")-(PDTE.""PDTE"")-(STOCK.""Stock"")),0) ""Pedir"",
0.00 ""Order"", ifnull(T7.""Provee_III"",CAST('     ' AS VARCHAR(50))) ""Prov.Pedido"", ifnull(T7.""CardName"",CAST('     ' AS VARCHAR(150))) ""Nombre"", 
CAST('     ' AS VARCHAR(50)) ""Nº Catálogo"",  CAST('     ' AS DATE) ""Fecha Prev."", 0.00 ""Traslado"", CAST('     ' AS VARCHAR(50)) ""Alm.Origen"", 
CAST('     ' AS VARCHAR(50)) ""Alm.Destino"", IFNull(T0.""CardCode"", 'S_PROV._Principal') as ""Proveedor Principal"",
CASE WHEN TX.""24Q"" > 3 then 'A'
	 WHEN TX.""24Q"" <= 3 and TX.""24Q"">0 and TX.""8Q"" > 0 Then 'B' 
	 WHEN TX.""24Q"" <= 3 and  TX.""A_8Q"" > 0 Then 'D'
	 WHEN TX.""C_12Q"" > 0 and TX.""24Q"" = 0  Then 'E' 
	 WHEN TX.""24Q"" = 0  Then 'F' 
	 ELSE 'OJO' end  as ""Clasificación"",
CASE WHEN TX.""C_12Q"" = 0 and TX.""24Q"" = 0 then 'S' else 'N' End As ""Nuevo"", T1.""Ventas_24Q"",
T4.""24M_Q_AL0"", T4.""24M_Q_AL14"", T4.""24M_Q_AL16"", T4.""24M_Q_AL7"", T4.""24M_Q_AL8"",
IFNULL(T5.""Stock_AL0"",0) ""Stock_Al0"", IFNULL(T5.""Stock_AL14"",0) ""Stock_AL14"", IFNULL(T5.""Stock_AL16"",0) ""Stock_AL16"", 
IFNULL(T5.""Stock_AL7"",0) ""Stock_AL7"", IFNULL(T5.""Stock_AL8"",0) ""Stock_AL8"", IFNULL(T6.""Pdte_AL0"",0) ""Pdte_AL0"" , 
IFNULL(T6.""Pdte_AL14"",0) ""Pdte_AL14"", IFNULL(T6.""Pdte_AL16"",0) ""Pdte_AL16"", IFNULL(T6.""Pdte_AL7"",0) ""Pdte_AL7"" , 
IFNULL(T6.""Pdte_AL8"",0) ""Pdte_AL8"", IFNULL(T7.""Provee"",CAST('     ' AS VARCHAR(50))) ""Provee"", 
IFNULL(T7.""Provee_II"" ,CAST('     ' AS VARCHAR(50))) ""Provee_II"", IFNULL(T7.""Provee_III"",CAST('     ' AS VARCHAR(50))) ""Provee_III"",
IFNULL(T7.""Mejor_P"",CAST('     ' AS VARCHAR(50)))  ""Mejor_P"" , TAR1.""TC_AGC"",  TAR2.""TC_NORDGLASS"", TAR3.""TC_GLAVISTA"",
TAR4.""TC_LUCAS"",  TAR5.""TC_PILK_UNID"", TAR6.""TC_PILK_CAJA"",TAR7.""TC_PILK_PLANTA"",  TAR8.""TC_RIO"", TAR9.""TC_SISECAM"",TAR10.""TC_FUYAO""
From OITM T0 "

#Region "Tarifas"
                        sSQLGrid &= " LEFT JOIN (SELECT OPLN.""ListNum"" || ' - ' ||OPLN.""ListName""  ""Lista"", ITM1.""ItemCode"", ITM1.""Price"" ""TC_AGC""
                                                    FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum""
                                                    WHERE OPLN.""U_EXO_TARCOM""='Si' and OPLN.""ListName""='TC_AGC'
                                                    Order by ITM1.""ItemCode"", OPLN.""ListNum"") TAR1 ON TAR1.""ItemCode""=T0.""ItemCode"" "

                        sSQLGrid &= " LEFT JOIN (SELECT OPLN.""ListNum"" || ' - ' ||OPLN.""ListName""  ""Lista"", ITM1.""ItemCode"", ITM1.""Price"" ""TC_NORDGLASS""
                                                    FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum""
                                                    WHERE OPLN.""U_EXO_TARCOM""='Si' and OPLN.""ListName""='TC_NORDGLASS'
                                                    Order by ITM1.""ItemCode"", OPLN.""ListNum"") TAR2 ON TAR2.""ItemCode""=T0.""ItemCode"" "

                        sSQLGrid &= " LEFT JOIN (SELECT OPLN.""ListNum"" || ' - ' ||OPLN.""ListName""  ""Lista"", ITM1.""ItemCode"", ITM1.""Price"" ""TC_GLAVISTA""
                                                    FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum""
                                                    WHERE OPLN.""U_EXO_TARCOM""='Si' and OPLN.""ListName""='TC_GLAVISTA'
                                                    Order by ITM1.""ItemCode"", OPLN.""ListNum"") TAR3 ON TAR3.""ItemCode""=T0.""ItemCode"" "

                        sSQLGrid &= " LEFT JOIN (SELECT OPLN.""ListNum"" || ' - ' ||OPLN.""ListName""  ""Lista"", ITM1.""ItemCode"", ITM1.""Price"" ""TC_LUCAS""
                                                    FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum""
                                                    WHERE OPLN.""U_EXO_TARCOM""='Si' and OPLN.""ListName""='TC_LUCAS'
                                                    Order by ITM1.""ItemCode"", OPLN.""ListNum"") TAR4 ON TAR4.""ItemCode""=T0.""ItemCode"" "

                        sSQLGrid &= " LEFT JOIN (SELECT OPLN.""ListNum"" || ' - ' ||OPLN.""ListName""  ""Lista"", ITM1.""ItemCode"", ITM1.""Price"" ""TC_PILK_UNID""
                                                    FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum""
                                                    WHERE OPLN.""U_EXO_TARCOM""='Si' and OPLN.""ListName""='TC_PILK_UNID'
                                                    Order by ITM1.""ItemCode"", OPLN.""ListNum"") TAR5 ON TAR5.""ItemCode""=T0.""ItemCode"" "

                        sSQLGrid &= " LEFT JOIN (SELECT OPLN.""ListNum"" || ' - ' ||OPLN.""ListName""  ""Lista"", ITM1.""ItemCode"", ITM1.""Price"" ""TC_PILK_CAJA""
                                                    FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum""
                                                    WHERE OPLN.""U_EXO_TARCOM""='Si' and OPLN.""ListName""='TC_PILK_CAJA'
                                                    Order by ITM1.""ItemCode"", OPLN.""ListNum"") TAR6 ON TAR6.""ItemCode""=T0.""ItemCode"" "

                        sSQLGrid &= " LEFT JOIN (SELECT OPLN.""ListNum"" || ' - ' ||OPLN.""ListName""  ""Lista"", ITM1.""ItemCode"", ITM1.""Price"" ""TC_PILK_PLANTA""
                                                    FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum""
                                                    WHERE OPLN.""U_EXO_TARCOM""='Si' and OPLN.""ListName""='TC_PILK_PLANTA'
                                                    Order by ITM1.""ItemCode"", OPLN.""ListNum"") TAR7 ON TAR7.""ItemCode""=T0.""ItemCode"" "

                        sSQLGrid &= " LEFT JOIN (SELECT OPLN.""ListNum"" || ' - ' ||OPLN.""ListName""  ""Lista"", ITM1.""ItemCode"", ITM1.""Price"" ""TC_RIO""
                                                    FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum""
                                                    WHERE OPLN.""U_EXO_TARCOM""='Si' and OPLN.""ListName""='TC_RIO'
                                                    Order by ITM1.""ItemCode"", OPLN.""ListNum"") TAR8 ON TAR8.""ItemCode""=T0.""ItemCode"" "

                        sSQLGrid &= " LEFT JOIN (SELECT OPLN.""ListNum"" || ' - ' ||OPLN.""ListName""  ""Lista"", ITM1.""ItemCode"", ITM1.""Price"" ""TC_SISECAM""
                                                    FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum""
                                                    WHERE OPLN.""U_EXO_TARCOM""='Si' and OPLN.""ListName""='TC_SISECAM'
                                                    Order by ITM1.""ItemCode"", OPLN.""ListNum"") TAR9 ON TAR9.""ItemCode""=T0.""ItemCode"" "

                        sSQLGrid &= " LEFT JOIN (SELECT OPLN.""ListNum"" || ' - ' ||OPLN.""ListName""  ""Lista"", ITM1.""ItemCode"", ITM1.""Price"" ""TC_FUYAO""
                                                    FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum""
                                                    WHERE OPLN.""U_EXO_TARCOM""='Si' and OPLN.""ListName""='TC_FUYAO'
                                                    Order by ITM1.""ItemCode"", OPLN.""ListNum"") TAR10 ON TAR10.""ItemCode""=T0.""ItemCode"" "
                        'sSQL = "SELECT DISTINCT 0 ""Precio"", OPLN.""ListNum"", OPLN.""ListName"" FROM OPLN 
                        '            WHERE OPLN.""U_EXO_TARCOM""='Si' "
                        'dtTarifas = Nothing : dtTarifas = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                        'For t = 0 To dtTarifas.Rows.Count - 1
                        '    sSQLGrid &= ", " & dtTarifas.Rows(t).Item("ListNum").ToString & " ""Tarifa " & dtTarifas.Rows(t).Item("ListName").ToString & """, 0.00  ""Precio " & dtTarifas.Rows(t).Item("ListName").ToString & """ "
                        'Next
#End Region
                        sSQLGrid &= "
LEFT JOIN (Select X.""ItemCode"" as ""ItemCode"", Sum(X.""24Q"") as ""24Q"", Sum(X.""8Q"") as ""8Q"" ,SUM(X.""C_12Q"") as ""C_12Q"",SUM(X.""A_8Q"") as ""A_8Q"" 
			FROM (Select T0.""ItemCode"" , T0.""WhsCode"", Coalesce(T1.""Ventas_Ult_Año"",0) as ""24Q"", Coalesce(T2.""Ventas_8Q"",0) as ""8Q"",
			coalesce(T3.""Compras_Ult_SEM"",0) as ""C_12Q"" , Coalesce(T4.""Ventas_A_8Q"",0) as ""A_8Q""
			From OITW  T0
			Left Join ""EXO_MRP_Ventas24Q"" T1 On T0.""ItemCode"" = T1.""ItemCode"" And T0.""WhsCode""  = T1.""WhsCode""
            Left Join ""EXO_MRP_Ventas8Q""  T2 On T0.""ItemCode"" = T2.""ItemCode"" And T0.""WhsCode""  = T2.""WhsCode""
            Left Join ""EXO_MRP_ComprasSemestre"" T3 On T0.""ItemCode"" = T3.""ItemCode"" And T0.""WhsCode""  = T3.""WhsCode""
            Left Join ""EXO_MRP_VentasA_8Q"" T4 On T0.""ItemCode"" = T4.""ItemCode"" And T0.""WhsCode""  = T4.""WhsCode""
            Where T0.""WhsCode"" IN (" & sAlmacenes & ") 
		  ) as X Group by X.""ItemCode"" ) TX  ON TX.""ItemCode"" = T0.""ItemCode""
LEFT JOIN (select T0.""ItemCode"", Sum(COALESCE(T0.""OnOrder"" - TX.""CantidadSolTraInt"" , 0)) as ""PDTE""
            from OITW T0 
            left join ( Select      T1.""ItemCode"" ,  T1.""WhsCode"" ,      coalesce(Sum(T1.""OpenQty""),  0) as ""CantidadSolTraInt"" 
                        from OWTQ T0 
                        LEFT JOIN WTQ1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" 
                        Where T0.""DocStatus"" = 'O' and T1.""LineStatus"" = 'O' and T1.""FromWhsCod"" = T1.""WhsCode"" 
                        and T1.""WhsCode"" in (" & sAlmacenes & ") group by T1.""ItemCode"" ,   T1.""WhsCode"" 
                      )TX ON T0.""ItemCode"" = TX.""ItemCode"" and T0.""WhsCode"" = TX.""WhsCode"" 
            Where  T0.""OnOrder"" > 0  Group BY T0.""ItemCode""
         )PDTE ON PDTE.""ItemCode""= T0.""ItemCode""
LEFT JOIN (Select T1.""ItemCode"", Sum(T1.""OnHand"") as  ""Stock""
            From OITW T1 Where T1.""WhsCode"" in (" & sAlmacenes & ") Group by T1.""ItemCode"" 
          )STOCK ON STOCK.""ItemCode""= T0.""ItemCode""
Left Join(Select ""ItemCode"", Sum(""Ventas_Ult_Año"") as ""Ventas_24Q"" 
          FROM ""EXO_MRP_Ventas24Q""  Where ""WhsCode"" IN (" & sAlmacenes & " ) Group  by ""ItemCode"" 
          )  T1  On T1.""ItemCode"" = T0.""ItemCode""
Left Join ""EXO_MRP_StocksActuales"" T5 on T0.""ItemCode"" = T5.""ItemCode""
LEFT JOIN ""EXO_MRP_Ventas_MED_24Q"" T4 on T0.""ItemCode"" = T4.""ItemCode""
Left Join ""EXO_MRP_Pdte"" T6 on T0.""ItemCode"" = T6.""ItemCode""
Left Join(Select T0.""ItemCode"",Case WHen T0.""QryGroup1"" = 'Y' then 'STOCK' ELSE T0.""CardCode"" END as ""Provee"", 
	 		Case WHen T0.""QryGroup1"" = 'Y' then T0.""CardCode"" ELSE MAX(T1.""VendorCode"") END as ""Provee_II"",
	 		TY.""CardCode"" as ""Provee_III"" ,	 TY.""CardName"",TY.""CardCode"" || '_' || TY.""Price"" as ""Mejor_P"" 
			From OITM T0
			Left Join ITM2 T1 ON T1.""ItemCode"" = T0.""ItemCode"" And T1.""VendorCode"" <> T0.""CardCode""
            Left Join( Select T0.""ItemCode"", T0.""PriceList"" , T0.""Price"", T3.""CardCode"", T3.""CardName""
						From ITM1 T0 INNER Join (Select T0.""ItemCode"" , MIn(T0.""Price"") As ""Precio_Min""
										From ITM1 T0
										INNER Join OPLN T2 ON T2.""ListNum"" = T0.""PriceList"" And T2.""U_EXO_TARCOM"" = 'Si'
                                        Left Join OCRD T1 On T1.""ListNum"" = T0.""PriceList""
                                        Where Coalesce(T1.""U_EXO_TSUM"", 0) <= " & sTSUM & "
										And T0.""Price"" <> 0
										Group by T0.""ItemCode""
                                        Order By T0.""ItemCode""
                                    ) TX On TX.""ItemCode"" = T0.""ItemCode"" And T0.""Price"" = TX.""Precio_Min""
			 			Left Join OCRD T3 On T3.""ListNum"" = T0.""PriceList""
             			Left Join OPLN T4 on T0.""PriceList"" = T4.""ListNum""
             			Where t0.""Price"" > 0 And T4.""U_EXO_TARCOM"" = 'Si'
					) TY 	ON TY.""ItemCode"" = T0.""ItemCode""
			GROUP BY  T0.""ItemCode"" , T0.""QryGroup1"" , T0.""CardCode"", 	TY.""CardName"" ,  TY.""CardCode"" , TY.""Price""	
		   )  T7 ON T7.""ItemCode"" = T0.""ItemCode""
WHERE  1=1 and T0.""QryGroup2""='N' and T0.""validFor""='Y'"
                        If sProveedorPR <> "" Then
                            sSQLGrid &= " and IFNull(T0.""CardCode"", 'S_PROV._Principal') ='" & sProveedorPR & "' "
                        End If
                        If sArtD <> "" Then
                            sSQLGrid &= " and T0.""ItemCode"" >='" & sArtD & "' "
                        End If
                        If sArtH <> "" Then
                            sSQLGrid &= " and T0.""ItemCode"" <='" & sArtH & "' "
                        End If

                        If sGrupos <> "" Then
                            sSQLGrid &= " and T0.""ItmsGrpCod"" in (" & sGrupos & ") "
                        End If
                        If sCLAS <> "" Then
                            sSQLGrid &= " and (CASE WHEN TX.""24Q"" > 3 then 'A'
	 WHEN TX.""24Q"" <= 3 and TX.""24Q"">0 and TX.""8Q"" > 0 Then 'B' 
	 WHEN TX.""24Q"" <= 3 and  TX.""A_8Q"" > 0 Then 'D'
	 WHEN TX.""C_12Q"" > 0 and TX.""24Q"" = 0  Then 'E' 
	 WHEN TX.""24Q"" = 0  Then 'F' 
	 ELSE 'OJO' end) in (" & sCLAS & ") "
                        End If
                        sSQLGrid &= " order by T0.""ItemCode"" "
                        oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQLGrid)
                        FormateaGridDOC(oForm)
#Region "Rellenamos Tabla Temporal"
                        INICIO._dtDatos = New System.Data.DataTable
                        Dim dt As SAPbouiCOM.DataTable = Nothing
                        dt = Nothing : dt = oForm.DataSources.DataTables.Item("DT_DOC")
                        'Añadimos Columnas                       
                        For iCol As Integer = 0 To 12
                            INICIO._dtDatos.Columns.Add(dt.Columns.Item(iCol).Name)
                        Next
                        INICIO._dtDatos.Columns.Add("ROW")
                        Dim primaryKey(0) As System.Data.DataColumn
                        primaryKey(0) = INICIO._dtDatos.Columns.Item("ROW")
                        INICIO._dtDatos.PrimaryKey = CType(primaryKey, Data.DataColumn())
#End Region

                        sMensaje = "Fin de la carga de datos."
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                Case "btnGen"
                    oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    If objGlobal.SBOApp.MessageBox("¿Esta seguro de generar los documentos según su parametrización?", 1, "Sí", "No") = 1 Then
                        If oForm.DataSources.DataTables.Item("DT_DOC").Rows.Count > 0 Then
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Generando Documentos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            oForm.Freeze(True)
#Region "Solicitud de pedido"
#Region "Filtro y Orden"
                            Dim dtSolPedido As New System.Data.DataTable

                            expression = "Order>0 and Prov.Pedido<>'' and Alm.Destino<>'' "
                            'sortOrder = "Prov.Pedido, Alm.Destino ASC"

                            Try
                                dtSolPedido = INICIO._dtDatos.Select(expression).CopyToDataTable()
                                sortOrder = "Prov.Pedido, Alm.Destino ASC"
                                dtSolPedido.DefaultView.Sort = sortOrder
                                dtSolPedido = dtSolPedido.DefaultView.ToTable()

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


                            Try
                                dtSolTraslado = INICIO._dtDatos.Select(expression).CopyToDataTable()
                                sortOrder = "Alm.Origen, Alm.Destino ASC"
                                dtSolTraslado.DefaultView.Sort = sortOrder
                                dtSolTraslado = dtSolTraslado.DefaultView.ToTable()
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
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = False
            oColumnChk.Width = 30

            For i = 1 To iColumnas - 1
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
                ElseIf sTitulo.Contains("PRECIO") Then
                    oColumnTxt.Editable = True
                    oColumnTxt.RightJustified = True
                ElseIf sTitulo.Contains("VENTAS") Or sTitulo.Contains("24M_Q") Or sTitulo.Contains("STOCK") Or sTitulo.Contains("PDTE") Or sTitulo.Contains("TARIFA") Or sTitulo.Contains("PEDIR") Then
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
            sSQL = "SELECT 'N' as ""Sel"", 'A' ""Clas"" FROM DUMMY
                    UNION ALL
                    SELECT 'N' as ""Sel"", 'B' ""Clas"" FROM DUMMY 
                    UNION ALL
                    SELECT 'N' as ""Sel"", 'D' ""Clas"" FROM DUMMY
                    UNION ALL
                    SELECT 'N' as ""Sel"", 'E' ""Clas"" FROM DUMMY
                    UNION ALL
                    Select 'N' as ""Sel"", 'F' ""Clas"" FROM DUMMY"
            oForm.DataSources.DataTables.Item("DTCLAS").ExecuteQuery(sSQL)
            FormateaGridCLAS(oForm)

            sSQL = "SELECT 'N' as ""Sel"", ""ItmsGrpCod"" ""Cod."", ""ItmsGrpNam"" ""Grupo"" FROM OITB WHERE ""U_EXO_GESNEC""='Si'"
            oForm.DataSources.DataTables.Item("DTGRU").ExecuteQuery(sSQL)
            FormateaGridGRU(oForm)

            sSQL = "Select 'Y' as ""Sel"", ""WhsCode"" ""Cod."", ""WhsName"" ""Almacén"" FROM ""OWHS"" order by ""WhsCode"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DTALM").ExecuteQuery(sSQL)
            FormateaGridALM(oForm)
            oForm.DataSources.UserDataSources.Item("UDDIAS").ValueEx = "7"
            oForm.DataSources.UserDataSources.Item("UDTSM").ValueEx = "0"
            oForm.Items.Item("btnGen").Enabled = False
            'CType(oForm.Items.Item("txtARTD").Specific, SAPbouiCOM.EditText).Active = True
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
    Private Sub FormateaGridGRU(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grdGRU").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdGRU").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 1 To 2
                CType(oform.Items.Item("grdGRU").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grdGRU").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
            Next



            CType(oform.Items.Item("grdGRU").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGridCLAS(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grdCLAS").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdCLAS").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30


            CType(oform.Items.Item("grdCLAS").Specific, SAPbouiCOM.Grid).Columns.Item(1).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oColumnTxt = CType(CType(oform.Items.Item("grdCLAS").Specific, SAPbouiCOM.Grid).Columns.Item(1), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False




            CType(oform.Items.Item("grdCLAS").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
End Class
