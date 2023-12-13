Imports System.Globalization
Imports System.Runtime.CompilerServices
Imports SAPbouiCOM

Public Class EXO_PNEC
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Private Shared qtyWhs As Integer
    Private Shared qtyOrder As Double

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
                                    Return EventHandler_VALIDATE_Before(infoEvento)
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                    If EventHandler_MATRIX_LINK_PRESSED(infoEvento) = False Then
                                        Return False
                                    End If
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
                                    If EventHandler_GOT_FOCUS_After(infoEvento) = False Then
                                        Return False
                                    End If
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
    Private Function EventHandler_MATRIX_LINK_PRESSED(ByVal pVal As ItemEvent) As Boolean

        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim sTipo As String = ""
        EventHandler_MATRIX_LINK_PRESSED = False

        Try
            oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)


            Select Case pVal.ItemUID
                Case "grd_DOC"
                    Dim gridData = CType(oForm.Items.Item("grd_DOC").Specific, Grid)
                    If gridData.Columns.Item(0).TitleObject.Caption = "Codigo de Respuesta" Then
                        sTipo = CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).DataTable.GetValue("ObjType", pVal.Row).ToString
                        CType(gridData.Columns.Item(3), SAPbouiCOM.EditTextColumn).LinkedObjectType = sTipo
                        '    If pVal.Row = 0 Then
                        '        CType(gridData.Columns.Item(2), SAPbouiCOM.EditTextColumn).LinkedObjectType = "540000006"
                        '    Else
                        '        CType(gridData.Columns.Item(2), SAPbouiCOM.EditTextColumn).LinkedObjectType = "112"
                        '    End If
                    End If

            End Select

            EventHandler_MATRIX_LINK_PRESSED = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
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

                oForm.Items.Item("grdCOMP").Width = INICIO._WidthCOMP
                oForm.Items.Item("grdCOMP").Height = INICIO._HeightCOMP
                CType(oForm.Items.Item("grdCOMP").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

                oForm.Items.Item("grdVENT").Width = INICIO._WidthVENT
                oForm.Items.Item("grdVENT").Height = INICIO._HeightVENT
                CType(oForm.Items.Item("grdVENT").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

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
    Private Function EventHandler_VALIDATE_Before(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        EventHandler_VALIDATE_Before = True
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "grd_DOC" Then
                If pVal.ColUID.ToUpper = "ORDER" Or pVal.ColUID.ToUpper = "AL0" Or pVal.ColUID.ToUpper = "AL7" Or pVal.ColUID.ToUpper = "AL8" Or pVal.ColUID.ToUpper = "AL14" Or pVal.ColUID.ToUpper = "AL16" Then
                    Dim order = CType(oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Value, Double)
                    Dim uc = CType(oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("UC").Cells.Item(pVal.Row).Value, Double)

                    If (order <> 0) Then
                        If (order Mod uc <> 0) Then
                            objGlobal.SBOApp.SetStatusBarMessage("El numero ingresado no es multiplo de la UC")
                            Return False
                        End If
                    End If
                End If
            End If
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
                Dim gridOrder = CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid)
                Dim dCant As Double = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Order").Cells.Item(pVal.Row).Value.ToString)
                Dim sArt As String = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("EUROCODE").Cells.Item(pVal.Row).Value.ToString
                Dim sProv As String = oForm.DataSources.UserDataSources.Item("UDPROV").Value
                Dim sProvD As String = oForm.DataSources.UserDataSources.Item("UDPROVD").Value
                Dim sCatalogo As String = ""

                If pVal.ColUID = "Prov.Pedido" And pVal.ItemChanged = True Then
                    If oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value.ToString = "" Then
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("UC").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = 1
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nombre").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = ""
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nº Catálogo").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = ""
                    End If

                    If (pVal.ItemChanged) Then
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Sel.").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = "Y"
                        CalculateGridCheckValues(oForm, pVal.Row)
                    End If
                End If
                If pVal.ColUID.ToUpper = "ORDER" And pVal.ItemChanged = True Then

                    oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("OrigOrder").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Order").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value
                    If dCant = 0 Then
                        sProv = ""
                    Else
                        If sProv.Trim = "" Then
                            sProv = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value.ToString
                            sProvD = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nombre").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value.ToString
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
                    If oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value.ToString = "" Then
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nº Catálogo").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = sCatalogo
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = sProv
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nombre").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = sProvD
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

                    oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Alm.Destino").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = sALM.Replace("'", "")

                    If (pVal.ItemChanged) Then
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Sel.").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = "Y"
                        CalculateGridCheckValues(oForm, pVal.Row)
                    End If
                ElseIf pVal.ColUID = "AL0" Or pVal.ColUID = "AL7" Or pVal.ColUID = "AL8" Or pVal.ColUID = "AL14" Or pVal.ColUID = "AL16" Then
                    Dim alOrder = CType(oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("OrigOrder").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value, Integer)
                    Dim al0 = CType(oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("AL0").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value, Integer)
                    Dim al7 = CType(oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("AL7").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value, Integer)
                    Dim al8 = CType(oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("AL8").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value, Integer)
                    Dim al14 = CType(oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("AL14").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value, Integer)
                    Dim al16 = CType(oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("AL16").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value, Integer)

                    If (qtyWhs = 1) Then
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Order").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = If((alOrder - (al0 + al7 + al8 + al14 + al16)) < 0, 1, (alOrder - (al0 + al7 + al8 + al14 + al16)))
                    Else
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("OrigOrder").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = al0 + al7 + al8 + al14 + al16
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Order").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = al0 + al7 + al8 + al14 + al16
                    End If

                    If (pVal.ItemChanged) Then
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Sel.").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = "Y"
                        CalculateGridCheckValues(oForm, pVal.Row)
                    End If

                ElseIf pVal.ColUID = "Prov.Pedido" Then
                    sProv = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value.ToString
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
                            If oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Fecha Prev.").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value Is Nothing Then
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Fecha Prev.").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = dFechaPrevista
                            End If
                        Else
                            oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Fecha Prev.").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = Nothing
                        End If
#End Region
                    Else
                        sCatalogo = ""
                    End If
                    oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nº Catálogo").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = sCatalogo

                    If (pVal.ItemChanged) Then
                        oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Sel.").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value = "Y"
                        CalculateGridCheckValues(oForm, pVal.Row)
                    End If
                End If
                If pVal.ColUID = "Order" Or pVal.ColUID = "Prov.Pedido" Or pVal.ColUID = "Traslado" Or pVal.ColUID = "Alm.Origen" Or pVal.ColUID = "Alm.Destino" Then
                    Dim dCantOrder As Double = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Order").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value.ToString)
                    Dim dCantTraslado As Double = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Traslado").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value.ToString)
                    Dim sProveedor As String = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value.ToString
                    Dim AlmOrigen As String = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Alm.Origen").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value.ToString
                    Dim AlmDestino As String = oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Alm.Destino").Cells.Item(gridOrder.GetDataTableRowIndex(pVal.Row)).Value.ToString
                    Dim blueBackColor As Integer = RGB(52, 135, 255) 'Convert.ToInt32("c002", 16)

                    If dCantOrder > 0 Then
                        If sProveedor <> "" And AlmDestino <> "" Then
                            Filtra_Sel(objGlobal, oForm, pVal, "Y")
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 6, blueBackColor)
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 7, blueBackColor)
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 13, blueBackColor)
                        Else
                            Filtra_Sel(objGlobal, oForm, pVal, "N")
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 6, -1)
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 7, -1)
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 13, -1)
                        End If
                    ElseIf dCantTraslado > 0 Then
                        If AlmOrigen <> "" And AlmDestino <> "" Then
                            Filtra_Sel(objGlobal, oForm, pVal, "Y")
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 11, blueBackColor)
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 12, blueBackColor)
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 13, blueBackColor)
                        Else
                            Filtra_Sel(objGlobal, oForm, pVal, "N")
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 11, -1)
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 12, -1)
                            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 13, -1)
                        End If
                    Else
                        Filtra_Sel(objGlobal, oForm, pVal, "N")
                        CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 6, -1)
                        CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 7, -1)
                        CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 11, -1)
                        CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 12, -1)
                        CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).CommonSetting.SetCellBackColor(pVal.Row + 1, 13, -1)
                    End If
                End If
            ElseIf pVal.ItemUID = " ThentxtProv" And oForm.DataSources.UserDataSources.Item("UDPROV").Value.ToString.Trim = "" Then
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

        Dim sAlmacenes As String = ""
        Dim sMensaje As String = ""
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
            ElseIf pVal.ItemUID = "txtALM" Then
                'comprobamos que tenga almacenes seleccionados
#Region "Comprobar si ha elegido un almacen"
                sAlmacenes = ""
                For i As Integer = 0 To oForm.DataSources.DataTables.Item("DTALM").Rows.Count - 1
                    If oForm.DataSources.DataTables.Item("DTALM").GetValue("Sel", i).ToString = "Y" Then
                        If sAlmacenes = "" Then
                            sAlmacenes = oForm.DataSources.DataTables.Item("DTALM").GetValue("Cod.", i).ToString
                        Else
                            sAlmacenes &= ", " & oForm.DataSources.DataTables.Item("DTALM").GetValue("Cod.", i).ToString
                        End If

                    End If
                Next

                If sAlmacenes = "" Then
                    sMensaje = "No ha indicado un almacén. Es obligatorio indicar uno."
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objGlobal.SBOApp.MessageBox(sMensaje)
                    Return False
                Else
                    Dim sAAlmacenes() As String = sAlmacenes.Split(CType(",", Char))
                    Dim iAlmacenes As Integer = 0
                    oCFLEvento = CType(pVal, IChooseFromListEvent)

                    oConds = New SAPbouiCOM.Conditions
                    For Each Almacen In sAAlmacenes
                        iAlmacenes += 1
                        If iAlmacenes > 1 Then
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        End If
                        oCond = oConds.Add
                        oCond.Alias = "WhsCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = Almacen
                    Next
                    oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
                End If
#End Region
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

            Dim grid = CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid)
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
                            ElseIf pVal.ItemUID = "txtALM" Then
                                oForm.DataSources.UserDataSources.Item("UDALM").ValueEx = oDataTable.GetValue("WhsCode", 0).ToString
                            End If

                        Catch ex As Exception

                        End Try
                    Case "2" 'Proveedor
                        Try
                            If pVal.ItemUID = "txtProv" Then
                                oForm.DataSources.UserDataSources.Item("UDPROVD").ValueEx = oDataTable.GetValue("CardName", 0).ToString
                                oForm.DataSources.UserDataSources.Item("UDPROV").ValueEx = oDataTable.GetValue("CardCode", 0).ToString
                                'Buscamos el tiempo de suministro
                                sSQL = "Select ""U_EXO_TSUM"" FROM OCRD WHERE ""CardCode""='" & oDataTable.GetValue("CardCode", 0).ToString & "' "
                                oForm.DataSources.UserDataSources.Item("UDTSM").ValueEx = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                            ElseIf pVal.ItemUID = "grd_DOC" And pVal.ColUID.Trim = "Prov.Pedido" Then
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Nombre").Cells.Item(grid.GetDataTableRowIndex(pVal.Row)).Value = oDataTable.GetValue("CardName", 0).ToString.Substring(0, 5)
                                CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Prov.Pedido").Cells.Item(grid.GetDataTableRowIndex(pVal.Row)).Value = oDataTable.GetValue("CardCode", 0).ToString

                                'Debemos buscar la Unidad de compra
                                Dim sUnidadCompra As Double = 1
                                sSQL = "SELECT IFNULL(""U_EXO_UC"",1) ""UC"" FROM OSCN WHERE ""CardCode""='" & oDataTable.GetValue("CardCode", 0).ToString & "'"
                                sSQL &= " and ""ItemCode""='" & oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("EUROCODE").Cells.Item(pVal.Row).Value.ToString & "'"
                                sUnidadCompra = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)
                                If sUnidadCompra = 0 Then
                                    sUnidadCompra = 1
                                End If
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("UC").Cells.Item(grid.GetDataTableRowIndex(pVal.Row)).Value = sUnidadCompra.ToString

                                Dim alPed = CType(oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Pedir").Cells.Item(grid.GetDataTableRowIndex(pVal.Row)).Value, Double)

                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Order").Cells.Item(pVal.Row).Value = (Math.Round(alPed / sUnidadCompra) * sUnidadCompra).ToString
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("OrigOrder").Cells.Item(pVal.Row).Value = (Math.Round(alPed / sUnidadCompra) * sUnidadCompra).ToString
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("AL0").Cells.Item(pVal.Row).Value = 0
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("AL7").Cells.Item(pVal.Row).Value = 0
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("AL8").Cells.Item(pVal.Row).Value = 0
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("AL14").Cells.Item(pVal.Row).Value = 0
                                oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("AL16").Cells.Item(pVal.Row).Value = 0

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
    Private Function EventHandler_GOT_FOCUS_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            Select Case pVal.ItemUID
                Case "grd_DOC"
                    If INICIO._iRowGrid <> pVal.Row And pVal.Row <> -1 Then
                        INICIO._iRowGrid = pVal.Row
                        CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item("RowsHeader").Click(pVal.Row)
                    End If
            End Select
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)

        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByVal pVal As ItemEvent) As Boolean
#Region "variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = "" : Dim sSQLGrid As String = ""
        Dim sArtD As String = "" : Dim sArtH As String = ""
        Dim sAlmacenes As String = "" : Dim sGrupos As String = "" : Dim sCLAS As String = ""
        Dim sTarifas As String = "" : Dim sATarifas() As String = Nothing
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
            If sProveedorPR = "" Then
                oForm.DataSources.UserDataSources.Item("UDPROVD").ValueEx = ""
            End If
            sTSUM = oForm.DataSources.UserDataSources.Item("UDTSM").ValueEx.ToString
            sMGS = oForm.DataSources.UserDataSources.Item("UDDIAS").ValueEx.ToString

            Select Case pVal.ItemUID
                Case "grd_DOC"
                    If pVal.ColUID = "Sel." Then
                        CalculateGridCheckValues(oForm, pVal.Row)
                    End If
                    'If INICIO._iRowGrid <> pVal.Row And pVal.Row <> -1 Then
                    '    INICIO._iRowGrid = pVal.Row
                    '    CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item("RowsHeader").Click(pVal.Row)
                    'End If

                Case "btnCARGAR"
                    If ComprobarALM(oForm, "DTALM") = True Then
#Region "Comprobar si ha elegido un almacen"
                        qtyWhs = 0
                        sAlmacenes = ""
                        For i As Integer = 0 To oForm.DataSources.DataTables.Item("DTALM").Rows.Count - 1
                            If oForm.DataSources.DataTables.Item("DTALM").GetValue("Sel", i).ToString = "Y" Then
                                If sAlmacenes = "" Then
                                    sAlmacenes = "'" & oForm.DataSources.DataTables.Item("DTALM").GetValue("Cod.", i).ToString & "' "
                                Else
                                    sAlmacenes &= ", '" & oForm.DataSources.DataTables.Item("DTALM").GetValue("Cod.", i).ToString & "' "
                                End If
                                qtyWhs = qtyWhs + 1
                            End If
                        Next
                        If sAlmacenes = "" Then
                            sMensaje = "No ha indicado un almacén. Es obligatorio indicar uno."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                            Return False
                        End If
#End Region

#Region "Comprobar si ha elegido una tarifa"
                        'Debemos comprobar las tarifas que han seleccionado.
                        '1º de compras y luego ventas
                        sTarifas = "" : Dim iTarifa As Integer = 0
                        For i As Integer = 0 To oForm.DataSources.DataTables.Item("DTCOMP").Rows.Count - 1
                            If oForm.DataSources.DataTables.Item("DTCOMP").GetValue("Sel", i).ToString = "Y" Then
                                If sTarifas = "" Then
                                    sTarifas = oForm.DataSources.DataTables.Item("DTCOMP").GetValue("Nombre", i).ToString.Trim
                                Else
                                    sTarifas &= ", " & oForm.DataSources.DataTables.Item("DTCOMP").GetValue("Nombre", i).ToString.Trim
                                End If
                            End If
                        Next

                        For i As Integer = 0 To oForm.DataSources.DataTables.Item("DTVENT").Rows.Count - 1
                            If oForm.DataSources.DataTables.Item("DTVENT").GetValue("Sel", i).ToString = "Y" Then
                                If sTarifas = "" Then
                                    sTarifas = oForm.DataSources.DataTables.Item("DTVENT").GetValue("Nombre", i).ToString.Trim
                                Else
                                    sTarifas &= ", " & oForm.DataSources.DataTables.Item("DTVENT").GetValue("Nombre", i).ToString.Trim
                                End If
                            End If
                        Next

                        If sTarifas = "" Then
                            sMensaje = "No ha indicado una Tarifa a mostrar. Es obligatorio indicar una."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                            Return False
                        End If

                        sATarifas = sTarifas.Split(CType(",", Char))
#End Region

                        qtyOrder = 0
                        Dim almDef = oForm.DataSources.UserDataSources.Item("UDALM").Value
                        sMensaje = "Cargando datos..."
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)

                        sSQLGrid = "Select 
                                        'N' ""Sel."", 
                                        T0.""ItemCode"" ""EUROCODE"", 
                                        LEFT(T0.""ItemName"",35) ""Descripción"",  
                                        CASE WHEN TX.""24Q"" > 3 then 'A'
	                                         WHEN TX.""24Q"" <= 3 and TX.""24Q"">0 and TX.""8Q"" > 0 Then 'B' 
	                                         WHEN TX.""24Q"" <= 3 and  TX.""A_8Q"" > 0 Then 'D'
	                                         WHEN TX.""C_12Q"" > 0 and TX.""24Q"" = 0  Then 'E' 
	                                         WHEN TX.""24Q"" = 0  Then 'F' 
	                                         ELSE 'OJO' end  as ""T"",
                                        CASE WHEN TX.""C_12Q"" = 0 and TX.""24Q"" = 0 then 'S' else 'N' End As ""N"", 
                                        SUBSTRING(TG.""ItmsGrpNam"",0,4) " & If(qtyWhs = 1, " ""Familia""", " ""Grupo""") & ", 
                                        CAST(T1.""Ventas_24Q"" AS INTEGER) AS ""VA"","


                        Dim stockTT As String = "0"
                        If (sAlmacenes.Contains("AL0")) Then
                            stockTT &= "+0+CAST(IFNULL(T5.""Stock_AL0"",0) AS INTEGER)"
                        End If
                        If (sAlmacenes.Contains("AL14")) Then
                            stockTT &= "+0+CAST(IFNULL(T5.""Stock_AL14"",0) AS INTEGER)"
                        End If
                        If (sAlmacenes.Contains("AL16")) Then
                            stockTT &= "+0+CAST(IFNULL(T5.""Stock_AL16"",0) AS INTEGER)"
                        End If
                        If (sAlmacenes.Contains("AL7")) Then
                            stockTT &= "+0+CAST(IFNULL(T5.""Stock_AL7"",0) AS INTEGER)"
                        End If
                        If (sAlmacenes.Contains("AL8")) Then
                            stockTT &= "+0+CAST(IFNULL(T5.""Stock_AL8"",0) AS INTEGER)"
                        End If
                        stockTT &= "+0 ""ST"", "

                        sSQLGrid &= stockTT
                        sSQLGrid &= "   CAST(IFNULL(T5.""Stock_AL0"",0) AS INTEGER) ""S_A0"", 
                                        CAST(IFNULL(T5.""Stock_AL14"",0) AS INTEGER) ""S_A14"", 
                                        CAST(IFNULL(T5.""Stock_AL16"",0) AS INTEGER) ""S_A16"", 
                                        CAST(IFNULL(T5.""Stock_AL7"",0) AS INTEGER) ""S_A7"", 
                                        CAST(IFNULL(T5.""Stock_AL8"",0) AS INTEGER) ""S_A8"","

                        Dim stockpdte As String = "0"
                        If (sAlmacenes.Contains("AL0")) Then
                            stockpdte &= "+0+CAST(IFNULL(T6.""Pdte_AL0"",0) AS INTEGER)"
                        End If
                        If (sAlmacenes.Contains("AL14")) Then
                            stockpdte &= "+0+CAST(IFNULL(T6.""Pdte_AL14"",0) AS INTEGER)"
                        End If
                        If (sAlmacenes.Contains("AL16")) Then
                            stockpdte &= "+0+CAST(IFNULL(T6.""Pdte_AL16"",0) AS INTEGER)"
                        End If
                        If (sAlmacenes.Contains("AL7")) Then
                            stockpdte &= "+0+CAST(IFNULL(T6.""Pdte_AL7"",0) AS INTEGER)"
                        End If
                        If (sAlmacenes.Contains("AL8")) Then
                            stockpdte &= "+0+CAST(IFNULL(T6.""Pdte_AL8"",0) AS INTEGER)"
                        End If
                        stockpdte &= "+0 ""PT"", "

                        sSQLGrid &= stockpdte
                        sSQLGrid &= "   CAST(IFNULL(T6.""Pdte_AL0"",0) AS INTEGER) ""P_A0"" , 
                                        CAST(IFNULL(T6.""Pdte_AL14"",0) AS INTEGER) ""P_A14"", 
                                        CAST(IFNULL(T6.""Pdte_AL16"",0) AS INTEGER) ""P_A16"", 
                                        CAST(IFNULL(T6.""Pdte_AL7"",0) AS INTEGER) ""P_A7"" , 
                                        CAST(IFNULL(T6.""Pdte_AL8"",0) AS INTEGER) ""P_A8"", 
                                        T4.""24M_Q_AL0"" ""VM_A0"", 
                                        T4.""24M_Q_AL14"" ""VM_A14"", 
                                        T4.""24M_Q_AL16"" ""VM_A16"", 
                                        T4.""24M_Q_AL7"" ""VM_A7"", 
                                        T4.""24M_Q_AL8"" ""VM_A8"", 
                                        ( (CASE WHEN "

                        Dim pedirStr As String = "(0"
                        If (sAlmacenes.Contains("AL0")) Then
                            pedirStr &= "+0+ROUND(ROUND((CASE WHEN IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL0"",0)*12, 0)/24),0)*2) + IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL0"",0)*12, 0)/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL0"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL0"",0) - IFNULL(T6.""Pdte_AL0"",0)),0) < 0 THEN 0 ELSE IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL0"",0)*12, 0)/24),0)*2) + IFNULL((T1.""Ventas_24Q""/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL0"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL0"",0) - IFNULL(T6.""Pdte_AL0"",0)),0) END),0)/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1)"
                        End If
                        If (sAlmacenes.Contains("AL7")) Then
                            pedirStr &= "+0+ROUND(ROUND((CASE WHEN IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL7"",0)*12, 0)/24),0)*2) + IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL7"",0)*12, 0)/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL7"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL7"",0) - IFNULL(T6.""Pdte_AL7"",0)),0) < 0 THEN 0 ELSE IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL7"",0)*12, 0)/24),0)*2) + IFNULL((T1.""Ventas_24Q""/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL7"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL7"",0) - IFNULL(T6.""Pdte_AL7"",0)),0) END),0)/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1)"
                        End If
                        If (sAlmacenes.Contains("AL8")) Then
                            pedirStr &= "+0+ROUND(ROUND((CASE WHEN IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL8"",0)*12, 0)/24),0)*2) + IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL8"",0)*12, 0)/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL8"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL8"",0) - IFNULL(T6.""Pdte_AL8"",0)),0) < 0 THEN 0 ELSE IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL8"",0)*12, 0)/24),0)*2) + IFNULL((T1.""Ventas_24Q""/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL8"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL8"",0) - IFNULL(T6.""Pdte_AL8"",0)),0) END),0)/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1)"
                        End If
                        If (sAlmacenes.Contains("AL14")) Then
                            pedirStr &= "+0+ROUND(ROUND((CASE WHEN IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL14"",0)*12, 0)/24),0)*2) + IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL14"",0)*12, 0)/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL14"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL14"",0) - IFNULL(T6.""Pdte_AL14"",0)),0) < 0 THEN 0 ELSE IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL14"",0)*12, 0)/24),0)*2) + IFNULL((T1.""Ventas_24Q""/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL14"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL14"",0) - IFNULL(T6.""Pdte_AL14"",0)),0) END),0)/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1)"
                        End If
                        If (sAlmacenes.Contains("AL16")) Then
                            pedirStr &= "+0+ROUND(ROUND((CASE WHEN IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL16"",0)*12, 0)/24),0)*2) + IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL16"",0)*12, 0)/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL16"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL16"",0) - IFNULL(T6.""Pdte_AL16"",0)),0) < 0 THEN 0 ELSE IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL16"",0)*12, 0)/24),0)*2) + IFNULL((T1.""Ventas_24Q""/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL16"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL16"",0) - IFNULL(T6.""Pdte_AL16"",0)),0) END),0)/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1)"
                        End If

                        sSQLGrid &= pedirStr
                        sSQLGrid &= " +0) < 0 THEN 0 ELSE " & pedirStr & "+0)  END) ) ""Pedir"",
                                      CAST(IFNULL(OSCN.""U_EXO_UC"", 1) AS INTEGER) ""UC"","
                        sSQLGrid &= "CAST((CASE WHEN ROUND(ROUND((" & pedirStr & " )/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1))) < 0 THEN 0 ELSE ROUND(ROUND((" & pedirStr & ")/IFNULL(OSCN.""U_EXO_UC"", 1))*IFNULL(OSCN.""U_EXO_UC"",1))) END) AS INTEGER) ""OrigOrder"","
                        sSQLGrid &= "CAST((CASE WHEN ROUND(ROUND((" & pedirStr & " )/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1))) < 0 THEN 0 ELSE ROUND(ROUND((" & pedirStr & ")/IFNULL(OSCN.""U_EXO_UC"", 1))*IFNULL(OSCN.""U_EXO_UC"",1))) END) AS INTEGER) ""Order"","

                        Dim queryAl = String.Empty
                        If (sAlmacenes.Contains("AL0") And qtyWhs > 1) Then
                            queryAl &= "CAST( ROUND(ROUND((CASE WHEN IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL0"",0)*12, 0)/24),0)*2) + IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL0"",0)*12, 0)/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL0"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL0"",0) - IFNULL(T6.""Pdte_AL0"",0)),0) < 0 THEN 0 ELSE IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL0"",0)*12, 0)/24),0)*2) + IFNULL((T1.""Ventas_24Q""/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL0"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL0"",0) - IFNULL(T6.""Pdte_AL0"",0)),0) END),0)/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1) AS INTEGER) ""AL0"","
                        Else
                            queryAl &= "0 ""AL0"","
                        End If

                        If (sAlmacenes.Contains("AL7") And qtyWhs > 1) Then
                            queryAl &= "CAST(ROUND(ROUND((CASE WHEN IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL7"",0)*12, 0)/24),0)*2) + IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL7"",0)*12, 0)/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL7"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL7"",0) - IFNULL(T6.""Pdte_AL7"",0)),0) < 0 THEN 0 ELSE IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL7"",0)*12, 0)/24),0)*2) + IFNULL((T1.""Ventas_24Q""/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL7"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL7"",0) - IFNULL(T6.""Pdte_AL7"",0)),0) END),0)/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1) AS INTEGER) ""AL7"","
                        Else
                            queryAl &= "0 ""AL7"","
                        End If

                        If (sAlmacenes.Contains("AL8") And qtyWhs > 1) Then
                            queryAl &= "CAST(ROUND(ROUND((CASE WHEN IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL8"",0)*12, 0)/24),0)*2) + IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL8"",0)*12, 0)/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL8"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL8"",0) - IFNULL(T6.""Pdte_AL8"",0)),0) < 0 THEN 0 ELSE IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL8"",0)*12, 0)/24),0)*2) + IFNULL((T1.""Ventas_24Q""/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL8"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL8"",0) - IFNULL(T6.""Pdte_AL8"",0)),0) END),0)/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1) AS INTEGER) ""AL8"", "
                        Else
                            queryAl &= "0 ""AL8"","
                        End If

                        If (sAlmacenes.Contains("AL14") And qtyWhs > 1) Then
                            queryAl &= "CAST(ROUND(ROUND((CASE WHEN IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL14"",0)*12, 0)/24),0)*2) + IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL14"",0)*12, 0)/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL14"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL14"",0) - IFNULL(T6.""Pdte_AL14"",0)),0) < 0 THEN 0 ELSE IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL14"",0)*12, 0)/24),0)*2) + IFNULL((T1.""Ventas_24Q""/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL14"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL14"",0) - IFNULL(T6.""Pdte_AL14"",0)),0) END),0)/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1) AS INTEGER) ""AL14"","
                        Else
                            queryAl &= "0 ""AL14"","
                        End If

                        If (sAlmacenes.Contains("AL16") And qtyWhs > 1) Then
                            queryAl &= "CAST(ROUND(ROUND((CASE WHEN IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL16"",0)*12, 0)/24),0)*2) + IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL16"",0)*12, 0)/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL16"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL16"",0) - IFNULL(T6.""Pdte_AL16"",0)),0) < 0 THEN 0 ELSE IFNULL(((IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL16"",0)*12, 0)/24),0)*2) + IFNULL((T1.""Ventas_24Q""/24),0) + ((" & sMGS & "+IFNULL(OCRD.""U_EXO_TSUM""," & sTSUM & "))*IFNULL((IFNULL(IFNULL(T4.""24M_Q_AL16"",0)*12, 0)/24),0)) -  IFNULL(T5.""Stock_AL16"",0) - IFNULL(T6.""Pdte_AL16"",0)),0) END),0)/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1) AS INTEGER) ""AL16"","
                        Else
                            queryAl &= "0 ""AL16"","
                        End If

                        sSQLGrid &= queryAl
                        sSQLGrid &= "(CASE WHEN '" & sProveedorPR & "' = '' THEN IFNULL(T0.""CardCode"", CAST('     ' AS VARCHAR(50))) ELSE '" & sProveedorPR & "' END)  ""Prov.Pedido"",
                                        SUBSTRING(IFNULL(OCRD.""CardFName"",CAST('     ' AS VARCHAR(150))),0,5) ""Nombre"",
                                        IFNULL(OSCN.""Substitute"",CAST('     ' AS VARCHAR(50))) ""Nº Catálogo"",  
                                        CAST('     ' AS DATE) ""Fecha Prev."", 
                                        IFNull(T0.""CardCode"", 'S_PROV._Principal') as ""Proveedor Principal"",
                                        IFNULL(T7.""Provee"", CAST('     ' AS VARCHAR(50))) ""Provee"", 
                                        IFNULL(T7.""Provee_II"" ,CAST('     ' AS VARCHAR(50))) ""Provee_II"", 
                                        IFNULL(T7.""Provee_III"",CAST('     ' AS VARCHAR(50))) ""Provee_III"",
                                        IFNULL(T7.""Mejor_P"",CAST('     ' AS VARCHAR(50)))  ""Mejor_P"",
                                        0.00 ""Traslado"",
                                        CAST('     ' AS VARCHAR(50)) ""Alm.Origen"", 
                                        CAST('     ' AS VARCHAR(50)) ""Alm.Destino"""

#Region "Tarifas"

                        iTarifa = 0
                        For Each Tarifa In sATarifas
                            iTarifa += 1
                            sSQLGrid &= ", TAR" & iTarifa.ToString & ".""T-" & Tarifa.Trim & """ "
                        Next
#End Region

#Region "Origen Tablas"
                        sSQLGrid &= " FROM OITM T0 
LEFT JOIN OITB TG ON TG.""ItmsGrpCod"" = T0.""ItmsGrpCod"" 
LEFT JOIN OCRD ON T0.""CardCode""=OCRD.""CardCode"" "
#End Region

#Region "Tarifas"

                        iTarifa = 0
                        For Each Tarifa In sATarifas
                            iTarifa += 1
                            sSQLGrid &= " LEFT JOIN (SELECT OPLN.""ListNum"" || ' - ' ||OPLN.""ListName""  ""Lista"", ITM1.""ItemCode"", ITM1.""Price"" ""T-" & Tarifa.Trim & """
                                                    FROM ITM1 INNER JOIN OPLN ON ITM1.""PriceList""=OPLN.""ListNum""
                                                    WHERE OPLN.""ListName""='" & Tarifa.Trim & "'
                                                    Order by ITM1.""ItemCode"", OPLN.""ListNum"") TAR" & iTarifa.ToString & " ON TAR" & iTarifa.ToString & ".""ItemCode""=T0.""ItemCode"" "
                        Next
#End Region

#Region "Relacion Tablas y Condiciones"
                        sSQLGrid &= "
LEFT JOIN OSCN ON OSCN.""ItemCode""=T0.""ItemCode"" and OSCN.""CardCode""= T0.""CardCode""
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

                        If (oForm.DataSources.UserDataSources.Item("UD_Excl").Value = "Y") Then
                            sSQLGrid &= "and (CASE WHEN ROUND(ROUND((" & pedirStr & " )/IFNULL(OSCN.""U_EXO_UC"",1))*IFNULL(OSCN.""U_EXO_UC"",1))) < 0 THEN 0 ELSE ROUND(ROUND((" & pedirStr & ")/IFNULL(OSCN.""U_EXO_UC"", 1))*IFNULL(OSCN.""U_EXO_UC"",1))) END) > 0 "
                        End If

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
#End Region

                        'sQuery &= sSQLGrid & " ) X0 order by X0.""EUROCODE"""
                        sSQLGrid &= " order by T0.""ItemCode"" "
                        oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQLGrid)
                        FormateaGridDOC(oForm, sAlmacenes, If(qtyWhs = 1, False, True))

                        CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item("Sel.").Editable = True

#Region "Rellenamos Tabla Temporal"
                        INICIO._dtDatos = New System.Data.DataTable
                        Dim dt As SAPbouiCOM.DataTable = Nothing
                        dt = Nothing : dt = oForm.DataSources.DataTables.Item("DT_DOC")
                        'Añadimos Columnas                       
                        For iCol As Integer = 0 To dt.Columns.Count - 1
                            INICIO._dtDatos.Columns.Add(dt.Columns.Item(iCol).Name)
                        Next
                        INICIO._dtDatos.Columns.Add("ROW")
                        Dim primaryKey(0) As System.Data.DataColumn
                        primaryKey(0) = INICIO._dtDatos.Columns.Item("ROW")
                        INICIO._dtDatos.PrimaryKey = CType(primaryKey, Data.DataColumn())
#End Region

                        sMensaje = "Fin de la carga de datos."
                        oForm.Items.Item("btnGen").Enabled = True
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                Case "btnGen"
                    If objGlobal.SBOApp.MessageBox("¿Esta seguro de generar los documentos según su parametrización?", 1, "Sí", "No") = 1 Then
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Generando Documentos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        If (CreateDocuments(oForm) <> 0) Then
                            Return True
                        End If

                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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
    Private Sub CalculateGridCheckValues(ByRef oForm As Form, ByVal row As Integer)
        Dim gridOrder = CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid)
        Dim dt = oForm.DataSources.DataTables.Item("DT_DOC")
        Dim valueSel = dt.Columns.Item("Sel.").Cells.Item(gridOrder.GetDataTableRowIndex(row)).Value
        If (valueSel.Equals("Y")) Then
            qtyOrder += dt.Columns.Item("Order").Cells.Item(gridOrder.GetDataTableRowIndex(row)).Value.ToString().DoubleParseAdvanced
        Else
            qtyOrder -= dt.Columns.Item("Order").Cells.Item(gridOrder.GetDataTableRowIndex(row)).Value.ToString().DoubleParseAdvanced
        End If

        CType(CType(oForm.Items.Item("grd_DOC").Specific, Grid).Columns.Item("Order"), EditTextColumn).ColumnSetting.SumValue = qtyOrder.ToString

    End Sub
    Private Function CreateDocuments(ByRef oForm As Form) As Integer
        Try
            oForm.Freeze(True)
            Dim columns() As String = {"AL0", "AL7", "AL8", "AL14", "AL16"}
            Dim gridData = CType(oForm.Items.Item("grd_DOC").Specific, Grid)
            Dim qwewq = gridData.DataTable
            Dim tupleResult = New List(Of Tuple(Of Integer, String, String, String))
            Dim dtResp = oForm.DataSources.DataTables.Item("DT_Res")
            dtResp.Rows.Clear()

            Dim whsObject = XDocument.Parse(oForm.DataSources.DataTables.Item("DTALM").SerializeAsXML(BoDataTableXmlSelect.dxs_DataOnly))
            Dim whsSelected = whsObject.
                Descendants("Cell").
                Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode.NextNode, XElement).Value.ToString.Equals("Y")).
                Select(Function(attr) attr.Parent)

            If (whsSelected.Count = 0) Then
                objGlobal.SBOApp.SetStatusBarMessage("No hay linea de almacen seleccionada")
                Return -1
            End If

            Dim dataObject = XDocument.Parse(oForm.DataSources.DataTables.Item("DT_DOC").SerializeAsXML(BoDataTableXmlSelect.dxs_DataOnly))
            Dim rowsSelected = dataObject.
                Descendants("Cell").
                Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode.NextNode, XElement).Value.ToString.Equals("Y")).
                Select(Function(attr) attr.Parent).
                ToList()

            If (rowsSelected.Count = 0) Then
                objGlobal.SBOApp.SetStatusBarMessage("No hay lineas seleccionadas")
                Return -1
            End If

            Dim rowsGrouped = rowsSelected.
                              GroupBy(Function(v) New With {
                                                                Key CType(v.Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Prov.Pedido")).First().LastNode, XElement).Value}).
                              ToList()

            If (whsSelected.Count = 1) Then
                'Proceso Simple 
#Region "Creacion de solicitud de Compra"
                Dim dSumaCantidad As Double = 0
                For i As Integer = 0 To rowsGrouped.Count - 1
                    Dim prov = rowsGrouped(i).Key.Value
                    Dim solPedComp = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations), SAPbobsCOM.Documents)

                    Dim fechaPrevHead = CType(rowsGrouped(i).ToList()(0).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Fecha Prev.")).First().LastNode, XElement).Value
                    Dim docDatePedHead As Date = Date.Now
                    If (Not fechaPrevHead.Equals("00000000")) Then
                        docDatePedHead = DateTime.ParseExact(fechaPrevHead, "yyyyMMdd", CultureInfo.InvariantCulture)
                    End If
                    solPedComp.RequriedDate = docDatePedHead
                    solPedComp.CardCode = prov

                    For x As Integer = 0 To rowsGrouped(i).Count - 1
                        Dim fechaPrev = CType(rowsGrouped(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Fecha Prev.")).First().LastNode, XElement).Value
                        Dim itemCode = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("EUROCODE")).First().LastNode, XElement).Value
                        Dim order = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Order")).First().LastNode, XElement).Value
                        Dim al0 = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL0")).First().LastNode, XElement).Value
                        Dim al7 = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL7")).First().LastNode, XElement).Value
                        Dim al8 = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL8")).First().LastNode, XElement).Value
                        Dim al14 = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL14")).First().LastNode, XElement).Value
                        Dim al16 = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL16")).First().LastNode, XElement).Value

                        Dim docDatePed As DateTime = DateTime.Now
                        If (Not fechaPrev.Equals("00000000")) Then
                            DateTime.ParseExact(fechaPrev, "yyyyMMdd", CultureInfo.InvariantCulture)
                        End If

                        If order.DoubleParseAdvanced > 0 Then
                            solPedComp.Lines.ItemCode = itemCode
                            solPedComp.Lines.RequiredDate = docDatePed
                            solPedComp.Lines.ShipDate = docDatePed
                            solPedComp.Lines.Price = GetPrice(prov, itemCode, Single.Parse(order), docDatePed)
                            solPedComp.Lines.UnitPrice = solPedComp.Lines.Price
                            solPedComp.Lines.WarehouseCode = CType(CType(whsSelected.First().FirstNode.NextNode, XElement).LastNode, XElement).Value
                            solPedComp.Lines.Quantity = order.DoubleParseAdvanced
                            solPedComp.Lines.RequiredQuantity = solPedComp.Lines.Quantity
                            dSumaCantidad += solPedComp.Lines.Quantity
                            solPedComp.Lines.CostingCode = GetCostCenter(solPedComp.Lines.WarehouseCode)
                            If solPedComp.Lines.Quantity <> 0 Then
                                solPedComp.Lines.Add()
                            End If
                        End If
                    Next
                    Dim responsePedComp As Integer = -1
                    If dSumaCantidad > 0 Then
                        responsePedComp = solPedComp.Add()
                        Dim msgResp = String.Empty
                        If (responsePedComp = 0) Then
                            msgResp = $"Solicitud de pedido creado exitosamente {objGlobal.compañia.GetNewObjectKey()}"
                        Else
                            msgResp = $"Error creando la solicitud de pedido {objGlobal.compañia.GetLastErrorDescription}"
                        End If
                        tupleResult.Add(New Tuple(Of Integer, String, String, String)(responsePedComp, msgResp, "540000006", objGlobal.compañia.GetNewObjectKey()))
                    End If
                Next

#End Region
#Region "Creacion de Traslado"

                Dim trasDict = New Dictionary(Of String, List(Of MembersTransferRequest))
                For i As Integer = 0 To rowsSelected.Count - 1

                    Dim fechaPrev = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Fecha Prev.")).First().LastNode, XElement).Value
                    Dim itemCode = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("EUROCODE")).First().LastNode, XElement).Value
                    Dim order = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Order")).First().LastNode, XElement).Value
                    Dim al0 = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL0")).First().LastNode, XElement).Value
                    Dim al7 = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL7")).First().LastNode, XElement).Value
                    Dim al8 = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL8")).First().LastNode, XElement).Value
                    Dim al14 = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL14")).First().LastNode, XElement).Value
                    Dim al16 = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL16")).First().LastNode, XElement).Value
                    Dim quantity As Double = 0.0
                    Dim whsSel = CType(CType(whsSelected.First().FirstNode.NextNode, XElement).LastNode, XElement).Value

                    For j As Integer = 0 To columns.Count - 1
                        If (gridData.Columns.Item(columns(j)).Editable = True) Then
                            Dim isValueMoreThanZero As Boolean = False
                            Dim qtyToTake As Double = 0.0
                            Select Case columns(j)
                                Case "AL0"
                                    If (Double.Parse(al0) > 0) Then
                                        isValueMoreThanZero = True
                                        qtyToTake = al0.DoubleParseAdvanced
                                    End If
                                Case "AL7"
                                    If (Double.Parse(al7) > 0) Then
                                        isValueMoreThanZero = True
                                        qtyToTake = al7.DoubleParseAdvanced
                                    End If
                                Case "AL8"
                                    If (Double.Parse(al8) > 0) Then
                                        isValueMoreThanZero = True
                                        qtyToTake = al8.DoubleParseAdvanced
                                    End If
                                Case "AL14"
                                    If (Double.Parse(al14) > 0) Then
                                        isValueMoreThanZero = True
                                        qtyToTake = al14.DoubleParseAdvanced
                                    End If
                                Case "AL16"
                                    If (Double.Parse(al16) > 0) Then
                                        isValueMoreThanZero = True
                                        qtyToTake = al16.DoubleParseAdvanced
                                    End If
                            End Select

                            If (isValueMoreThanZero) Then
                                Dim line = New MembersTransferRequest
                                line.ItemCode = itemCode
                                line.FromWarehouseCode = columns(j)
                                line.WarehouseCode = whsSel
                                line.Quantity = qtyToTake

                                If (Not trasDict.ContainsKey(columns(j) & "-" & whsSel)) Then
                                    Dim lst = New List(Of MembersTransferRequest)
                                    lst.Add(line)
                                    trasDict.Add(columns(j) & "-" & whsSel, lst)
                                Else
                                    trasDict(columns(j) & "-" & whsSel).Add(line)
                                End If
                            End If

                        End If
                    Next
                Next

                For i As Integer = 0 To trasDict.Count - 1
                    Dim keyVal = trasDict.ElementAt(i).Key
                    Dim linesTransfer = trasDict(keyVal)
                    Dim solTras = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransferDraft), SAPbobsCOM.StockTransfer)
                    solTras.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest
                    solTras.UserFields.Fields.Item("U_EXO_TIPO").Value = "ITC"
                    solTras.FromWarehouse = linesTransfer(0).FromWarehouseCode
                    solTras.ToWarehouse = linesTransfer(0).WarehouseCode

                    For j As Integer = 0 To trasDict(keyVal).Count - 1
                        solTras.Lines.ItemCode = linesTransfer(j).ItemCode
                        solTras.Lines.FromWarehouseCode = linesTransfer(j).FromWarehouseCode
                        solTras.Lines.WarehouseCode = linesTransfer(j).WarehouseCode
                        solTras.Lines.Quantity = linesTransfer(j).Quantity
                        solTras.Lines.Add()
                    Next

                    Dim responseTrasl = solTras.Add()
                    Dim msgRespTras = String.Empty
                    Dim sDocEntry As String = ""
                    If (responseTrasl = 0) Then
                        sDocEntry = objGlobal.compañia.GetNewObjectKey()
                        msgRespTras = $"Solicitud de traslado de borrador creado exitosamente {sDocEntry}"
                    Else
                        sDocEntry = ""
                        msgRespTras = $"Error creando la solicitud de traslado {objGlobal.compañia.GetLastErrorDescription}"
                    End If

                    tupleResult.Add(New Tuple(Of Integer, String, String, String)(responseTrasl, msgRespTras, "112", sDocEntry))
                Next

#End Region

            ElseIf (whsSelected.Count > 1) Then
                Dim whsDef = oForm.DataSources.UserDataSources.Item("UDALM").Value.Trim
                If (String.IsNullOrEmpty(whsDef)) Then
                    'Proceso Multiple No Agrupado
#Region "Creacion de solicitud de Compra"
                    For i As Integer = 0 To rowsGrouped.Count - 1
                        Dim prov = rowsGrouped(i).Key.Value
                        Dim dicPed = New Dictionary(Of String, List(Of MembersPurchaseRequest))
                        For x As Integer = 0 To rowsGrouped(i).Count - 1

                            Dim fechaPrev = CType(rowsGrouped(i).ToList()(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Fecha Prev.")).First().LastNode, XElement).Value
                            Dim itemCode = CType(rowsGrouped(i).ToList()(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("EUROCODE")).First().LastNode, XElement).Value
                            Dim order = CType(rowsGrouped(i).ToList()(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Order")).First().LastNode, XElement).Value
                            Dim al0 = CType(rowsGrouped(i).ToList()(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL0")).First().LastNode, XElement).Value
                            Dim al7 = CType(rowsGrouped(i).ToList()(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL7")).First().LastNode, XElement).Value
                            Dim al8 = CType(rowsGrouped(i).ToList()(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL8")).First().LastNode, XElement).Value
                            Dim al14 = CType(rowsGrouped(i).ToList()(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL14")).First().LastNode, XElement).Value
                            Dim al16 = CType(rowsGrouped(i).ToList()(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL16")).First().LastNode, XElement).Value

                            Dim docDatePed As Date = Date.Now
                            If (Not fechaPrev.Equals("00000000")) Then
                                docDatePed = DateTime.ParseExact(fechaPrev, "yyyyMMdd", CultureInfo.InvariantCulture)
                            End If

                            For j As Integer = 0 To columns.Count - 1
                                If (gridData.Columns.Item(columns(j)).Editable = True) Then
                                    Dim isValueMoreThanZero As Boolean = False
                                    Dim qtyToTake As Double = 0.0
                                    Select Case columns(j)
                                        Case "AL0"
                                            If (Double.Parse(al0) > 0) Then
                                                isValueMoreThanZero = True
                                                qtyToTake = al0.DoubleParseAdvanced
                                            End If
                                        Case "AL7"
                                            If (Double.Parse(al7) > 0) Then
                                                isValueMoreThanZero = True
                                                qtyToTake = al7.DoubleParseAdvanced
                                            End If
                                        Case "AL8"
                                            If (Double.Parse(al8) > 0) Then
                                                isValueMoreThanZero = True
                                                qtyToTake = al8.DoubleParseAdvanced
                                            End If
                                        Case "AL14"
                                            If (Double.Parse(al14) > 0) Then
                                                isValueMoreThanZero = True
                                                qtyToTake = al14.DoubleParseAdvanced
                                            End If
                                        Case "AL16"
                                            If (Double.Parse(al16) > 0) Then
                                                isValueMoreThanZero = True
                                                qtyToTake = al16.DoubleParseAdvanced
                                            End If
                                    End Select

                                    If (isValueMoreThanZero) Then
                                        Dim line As MembersPurchaseRequest = New MembersPurchaseRequest
                                        line.ItemCode = itemCode
                                        line.WarehouseCode = columns(j)
                                        line.Quantity = qtyToTake
                                        line.RequiredDate = docDatePed
                                        line.ShipDate = docDatePed
                                        line.RequiredQuantity = line.Quantity
                                        line.Price = GetPrice(prov, itemCode, Single.Parse(order), docDatePed)
                                        line.UnitPrice = line.Price
                                        line.CostingCode = GetCostCenter(line.WarehouseCode)

                                        If (Not dicPed.ContainsKey(columns(j))) Then
                                            Dim list = New List(Of MembersPurchaseRequest)
                                            list.Add(line)

                                            dicPed.Add(columns(j), list)
                                        Else
                                            dicPed(columns(j)).Add(line)
                                        End If

                                    End If
                                End If
                            Next
                        Next
                        Dim dSumaCantidad As Double = 0
                        For x As Integer = 0 To dicPed.Count - 1
                            Dim key = dicPed.ElementAt(x).Key
                            Dim linesPed = dicPed(key)
                            Dim solPedComp = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations), SAPbobsCOM.Documents)
                            Dim maxDateLinq = From p In dicPed(key) Group p By p.RequiredDate Into g = Group Select MaxDate = g.Max(Function(p) p.RequiredDate)
                            solPedComp.CardCode = prov
                            solPedComp.RequriedDate = If(maxDateLinq.Count > 0, maxDateLinq.First(), Date.Now)

                            For line As Integer = 0 To linesPed.Count - 1
                                solPedComp.Lines.ItemCode = linesPed(line).ItemCode
                                solPedComp.Lines.WarehouseCode = linesPed(line).WarehouseCode
                                solPedComp.Lines.Price = linesPed(line).Price
                                solPedComp.Lines.UnitPrice = linesPed(line).UnitPrice
                                solPedComp.Lines.Quantity = linesPed(line).Quantity
                                solPedComp.Lines.RequiredQuantity = linesPed(line).RequiredQuantity
                                dSumaCantidad += solPedComp.Lines.Quantity
                                solPedComp.Lines.CostingCode = linesPed(line).CostingCode
                                solPedComp.Lines.RequiredDate = linesPed(line).RequiredDate
                                solPedComp.Lines.ShipDate = linesPed(line).ShipDate
                                If solPedComp.Lines.Quantity <> 0 Then
                                    solPedComp.Lines.Add()
                                End If
                            Next

                            If dSumaCantidad > 0 Then
                                Dim responsePedComp = solPedComp.Add()
                                Dim msgResp = String.Empty
                                Dim sDocEntry As String = objGlobal.compañia.GetNewObjectKey()
                                If (responsePedComp = 0) Then
                                    msgResp = $"Solicitud de pedido creado exitosamente {sDocEntry}"
                                Else
                                    msgResp = $"Error creando la solicitud de pedido {objGlobal.compañia.GetLastErrorDescription}"
                                    sDocEntry = ""
                                End If

                                tupleResult.Add(New Tuple(Of Integer, String, String, String)(responsePedComp, msgResp, "540000006", sDocEntry))
                            End If
                        Next
                    Next
#End Region
                Else
                    'Proceso Multiple Agrupado
                    Dim shouldCreateTransfer = objGlobal.SBOApp.MessageBox("Desea crear los traslados para los pedidos seleccionados?", 1, "Si", "No")
#Region "Creacion de solicitud de Compra"
                    Dim dSumaCantidad As Double = 0
                    For i As Integer = 0 To rowsGrouped.Count - 1
                        Dim prov = rowsGrouped(i).Key.Value
                        Dim solPedComp = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations), SAPbobsCOM.Documents)
                        Dim fechaPrevHead = CType(rowsGrouped(i).ToList()(0).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Fecha Prev.")).First().LastNode, XElement).Value
                        Dim docDatePedHead As Date = Date.Now
                        If (Not fechaPrevHead.Equals("00000000")) Then
                            docDatePedHead = DateTime.ParseExact(fechaPrevHead, "yyyyMMdd", CultureInfo.InvariantCulture)
                        End If
                        solPedComp.RequriedDate = docDatePedHead
                        solPedComp.CardCode = prov
                        Dim bprimeralinea As Boolean = True
                        For x As Integer = 0 To rowsGrouped(i).Count - 1
                            Dim fechaPrev = CType(rowsGrouped(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Fecha Prev.")).First().LastNode, XElement).Value
                            Dim itemCode = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("EUROCODE")).First().LastNode, XElement).Value
                            Dim order = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Order")).First().LastNode, XElement).Value
                            Dim al0 = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL0")).First().LastNode, XElement).Value
                            Dim al7 = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL7")).First().LastNode, XElement).Value
                            Dim al8 = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL8")).First().LastNode, XElement).Value
                            Dim al14 = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL14")).First().LastNode, XElement).Value
                            Dim al16 = CType(CType(rowsGrouped(i).ToList()(x), XElement).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL16")).First().LastNode, XElement).Value

                            Dim docDatePed As DateTime = DateTime.Now
                            If (Not fechaPrev.Equals("00000000")) Then
                                DateTime.ParseExact(fechaPrev, "yyyyMMdd", CultureInfo.InvariantCulture)
                            End If

                            If bprimeralinea Then
                                bprimeralinea = False
                            Else
                                If solPedComp.Lines.Quantity <> 0 Then
                                    solPedComp.Lines.Add()
                                End If
                            End If
                            solPedComp.Lines.ItemCode = itemCode
                            solPedComp.Lines.WarehouseCode = CType(CType(whsSelected.First().FirstNode.NextNode, XElement).LastNode, XElement).Value
                            solPedComp.Lines.Price = GetPrice(prov, itemCode, Single.Parse(order), docDatePed)
                            solPedComp.Lines.UnitPrice = solPedComp.Lines.Price
                            solPedComp.Lines.Quantity = order.DoubleParseAdvanced
                            solPedComp.Lines.RequiredQuantity = solPedComp.Lines.Quantity
                            dSumaCantidad += solPedComp.Lines.Quantity
                            solPedComp.Lines.CostingCode = GetCostCenter(solPedComp.Lines.WarehouseCode)
                            solPedComp.Lines.RequiredDate = docDatePed
                            solPedComp.Lines.ShipDate = docDatePed
                        Next
                        If dSumaCantidad > 0 Then
                            Dim responsePedComp = solPedComp.Add()
                            Dim msgResp = String.Empty
                            Dim sDocEntry As String = ""
                            If (responsePedComp = 0) Then
                                sDocEntry = objGlobal.compañia.GetNewObjectKey()
                                msgResp = $"Solicitud de pedido creado exitosamente {sDocEntry}"
                            Else
                                msgResp = $"Error creando la solicitud de pedido {objGlobal.compañia.GetLastErrorDescription}"
                                sDocEntry = ""
                            End If

                            tupleResult.Add(New Tuple(Of Integer, String, String, String)(responsePedComp, msgResp, "540000006", sDocEntry))
                        End If

                    Next
#End Region
                    If (shouldCreateTransfer = 1) Then
#Region "Creacion de Traslado"

                        Dim trasDict = New Dictionary(Of String, List(Of MembersTransferRequest))
                        For i As Integer = 0 To rowsSelected.Count - 1

                            'For x As Integer = 0 To rowsGrouped(x).Count - 1
                            Dim fechaPrev = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Fecha Prev.")).First().LastNode, XElement).Value
                                Dim itemCode = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("EUROCODE")).First().LastNode, XElement).Value
                                Dim order = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Order")).First().LastNode, XElement).Value
                                Dim al0 = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL0")).First().LastNode, XElement).Value
                                Dim al7 = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL7")).First().LastNode, XElement).Value
                                Dim al8 = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL8")).First().LastNode, XElement).Value
                                Dim al14 = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL14")).First().LastNode, XElement).Value
                                Dim al16 = CType(rowsSelected(i).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL16")).First().LastNode, XElement).Value

                                For j As Integer = 0 To columns.Count - 1
                                    If (gridData.Columns.Item(columns(j)).Editable = True) Then
                                        Dim qty As Double = 0
                                        Dim isValueMoreThanZero = False

                                        Select Case columns(j)
                                            Case "AL0"
                                                If (Not whsDef.Equals(columns(j))) Then
                                                    If (al0.DoubleParseAdvanced > 0) Then
                                                        isValueMoreThanZero = True
                                                        qty = al0.DoubleParseAdvanced
                                                    End If
                                                End If
                                            Case "AL7"
                                                If (Not whsDef.Equals(columns(j))) Then
                                                    If (al7.DoubleParseAdvanced > 0) Then
                                                        isValueMoreThanZero = True
                                                        qty = al7.DoubleParseAdvanced
                                                    End If
                                                End If
                                            Case "AL8"
                                                If (Not whsDef.Equals(columns(j))) Then
                                                    If (al8.DoubleParseAdvanced > 0) Then
                                                        isValueMoreThanZero = True
                                                        qty = al8.DoubleParseAdvanced
                                                    End If
                                                End If
                                            Case "AL14"
                                                If (Not whsDef.Equals(columns(j))) Then
                                                    If (al14.DoubleParseAdvanced > 0) Then
                                                        isValueMoreThanZero = True
                                                        qty = al14.DoubleParseAdvanced
                                                    End If
                                                End If
                                            Case "AL16"
                                                If (Not whsDef.Equals(columns(j))) Then
                                                    If (al16.DoubleParseAdvanced > 0) Then
                                                        isValueMoreThanZero = True
                                                        qty = al16.DoubleParseAdvanced
                                                    End If
                                                End If
                                        End Select

                                        If (isValueMoreThanZero) Then
                                            Dim line = New MembersTransferRequest
                                            line.ItemCode = itemCode
                                        line.FromWarehouseCode = whsDef
                                        line.WarehouseCode = columns(j)
                                        line.Quantity = qty

                                        If (Not trasDict.ContainsKey(columns(j) & "-" & whsDef)) Then
                                                Dim lst = New List(Of MembersTransferRequest)
                                                lst.Add(line)
                                                trasDict.Add(columns(j) & "-" & whsDef, lst)
                                            Else
                                                trasDict(columns(j) & "-" & whsDef).Add(line)
                                            End If
                                        End If
                                    End If
                                Next
                            'Next
                        Next

                        For i As Integer = 0 To trasDict.Count - 1
                            Dim keyVal = trasDict.ElementAt(i).Key
                            Dim linesTransfer = trasDict(keyVal)
                            Dim solTras = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransferDraft), SAPbobsCOM.StockTransfer)
                            solTras.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest
                            solTras.UserFields.Fields.Item("U_EXO_TIPO").Value = "ITC"
                            solTras.FromWarehouse = linesTransfer(0).FromWarehouseCode
                            solTras.ToWarehouse = linesTransfer(0).WarehouseCode

                            For j As Integer = 0 To trasDict(keyVal).Count - 1
                                solTras.Lines.ItemCode = linesTransfer(j).ItemCode
                                solTras.Lines.FromWarehouseCode = linesTransfer(j).FromWarehouseCode
                                solTras.Lines.WarehouseCode = linesTransfer(j).WarehouseCode
                                solTras.Lines.Quantity = linesTransfer(j).Quantity
                                solTras.Lines.Add()
                            Next

                            Dim responseTrasl = solTras.Add()
                            Dim msgRespTras = String.Empty
                            Dim sDocEntry As String = ""
                            If (responseTrasl = 0) Then
                                sDocEntry = objGlobal.compañia.GetNewObjectKey()
                                msgRespTras = $"Solicitud de traslado de borrador creado exitosamente {sDocEntry}"
                            Else
                                sDocEntry = ""
                                msgRespTras = $"Error creando la solicitud de traslado {objGlobal.compañia.GetLastErrorDescription}"
                            End If

                            tupleResult.Add(New Tuple(Of Integer, String, String, String)(responseTrasl, msgRespTras, "112", sDocEntry))
                        Next

#End Region
                    End If

                End If
            End If

            gridData.DataTable = dtResp
            For i As Integer = 0 To tupleResult.Count - 1
                dtResp.Rows.Add()
                dtResp.SetValue(0, i, tupleResult(i).Item1.ToString)
                dtResp.SetValue(1, i, tupleResult(i).Item2.ToString)
                dtResp.SetValue(2, i, tupleResult(i).Item3.ToString)
                dtResp.SetValue(3, i, tupleResult(i).Item4.ToString)
            Next

            For i As Integer = 0 To gridData.Columns.Count - 1
                gridData.Columns.Item(i).Editable = False
            Next

            gridData.Columns.Item(0).TitleObject.Caption = "Codigo de Respuesta"
            gridData.Columns.Item(1).TitleObject.Caption = "Mensaje"
            gridData.Columns.Item(2).TitleObject.Caption = "Tipo"
            CType(gridData.Columns.Item(2), SAPbouiCOM.EditTextColumn).Visible = False
            gridData.Columns.Item(3).TitleObject.Caption = "Documento"
            CType(gridData.Columns.Item(3), SAPbouiCOM.EditTextColumn).LinkedObjectType = "540000006"

            oForm.Items.Item("btnGen").Enabled = False
            Return 0
        Finally
            oForm.Freeze(False)
        End Try
    End Function
    Private Function GetPrice(ByVal cardCode As String, ByVal itemCode As String, ByVal amount As Single, ByVal refDate As Date) As Double
        Try
            Dim errResult As String = ""
            Dim vObj = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge), SAPbobsCOM.SBObob)
            Dim rs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            rs = vObj.GetItemPrice(cardCode, itemCode, amount, refDate)

            Return Double.Parse(rs.Fields.Item(0).Value.ToString)
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Private Function GetCostCenter(ByVal whsCode As String) As String
        Dim query As String = $"SELECT T1.""PrcCode"" FROM OWHS T0 Left Join OPRC t1 ON T1.""U_EXO_DELEGA"" = T0.""U_EXO_SUCURSAL"" WHERE T0.""WhsCode"" = '{whsCode}'"
        Dim rs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        rs.DoQuery(query)

        If (rs.RecordCount > 0) Then
            Return rs.Fields.Item("PrcCode").Value.ToString
        Else
            Return String.Empty
        End If
    End Function
    Private Sub Filtra_Sel(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form, ByVal pVal As ItemEvent, ByVal Valor As String)
        Try
            oForm.DataSources.DataTables.Item("DT_DOC").Columns.Item("Sel.").Cells.Item(pVal.Row).Value = Valor.Trim
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
                Try
                    INICIO._dtDatos.Rows.Remove(INICIO._dtDatos.Rows.Find(New Object() {pVal.Row}))
                Catch ex As Exception

                End Try

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
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
    Private Sub FormateaGridDOC(ByRef oform As SAPbouiCOM.Form, ByVal sAlmacenes As String, ByVal isMultiple As Boolean)
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
                    If (sTitulo.Contains("ORDER")) Then
                        If (Not isMultiple) Then
                            oColumnTxt.Editable = True
                        Else
                            oColumnTxt.Editable = False
                        End If
                    Else
                        oColumnTxt.Editable = True
                    End If
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
                ElseIf sTitulo.Contains("VA") Or sTitulo.Contains("VM_") Or sTitulo.Contains("ST") Or sTitulo.Contains("PDTE") Or sTitulo.Contains("T-") Or sTitulo.Contains("PEDIR") Or sTitulo.Contains("UC") Then
                    oColumnTxt.RightJustified = True
                ElseIf Left(sTitulo, 2) = "N " Then
                    oColumnTxt.RightJustified = True
                ElseIf sTitulo.Contains("FECHA PREV.") Then
                    oColumnTxt.Editable = True
                ElseIf sTitulo.Contains("AL0") Or sTitulo.Contains("AL7") Or sTitulo.Contains("AL8") Or sTitulo.Contains("AL14") Or sTitulo.Contains("AL16") Then
                    If (isMultiple) Then
                        If sAlmacenes.Contains(sTitulo) Then
                            oColumnTxt.Editable = True
                        End If
                    Else
                        If sAlmacenes.Contains(sTitulo) Then
                            oColumnTxt.Editable = False
                        Else
                            oColumnTxt.Editable = True
                        End If
                    End If
                End If

                oColumnTxt.TitleObject.Sortable = True
            Next

            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("OrigOrder").Visible = False
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("Traslado").Visible = False
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("Alm.Origen").Visible = False
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("Alm.Destino").Visible = False
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("Descripción").Width = 200
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item(If(Not isMultiple, "Familia", "Grupo")).Width = 40
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("VA").Width = 40
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("ST").Width = 40
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("S_A0").Width = 40
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("S_A0").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("S_A7").Width = 40
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("S_A7").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("S_A8").Width = 40
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("S_A8").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("S_A14").Width = 45
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("S_A14").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("S_A16").Width = 45
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("S_A16").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("PT").Width = 40
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("PT").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("P_A0").Width = 40
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("P_A0").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("P_A7").Width = 40
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("P_A7").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("P_A8").Width = 40
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("P_A8").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("P_A14").Width = 60
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("P_A14").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("P_A16").Width = 60
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("P_A16").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("VM_A0").Width = 50
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("VM_A7").Width = 50
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("VM_A8").Width = 50
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("VM_A14").Width = 60
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("VM_A16").Width = 60
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("AL0").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("AL7").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("AL8").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("AL14").RightJustified = True
            CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("AL16").RightJustified = True

            CType(CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("Order"), EditTextColumn).ColumnSetting.SumType = BoColumnSumType.bst_Manual
            CType(CType(oform.Items.Item("grd_DOC").Specific, Grid).Columns.Item("Order"), EditTextColumn).ColumnSetting.SumValue = "0"
            CType(oform.Items.Item("grd_DOC").Specific, Grid).AutoResizeColumns()
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

            sSQL = "SELECT 'N' as ""Sel"", ""ItmsGrpCod"" ""Cod."", ""ItmsGrpNam"" ""Familia"" FROM OITB WHERE ""U_EXO_GESNEC""='Si' ORDER BY ""ItmsGrpNam"""
            oForm.DataSources.DataTables.Item("DTGRU").ExecuteQuery(sSQL)
            FormateaGridGRU(oForm)

            sSQL = "Select 'N' as ""Sel"", ""WhsCode"" ""Cod."", ""WhsName"" ""Almacén"" FROM ""OWHS"" order by ""WhsName"""
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DTALM").ExecuteQuery(sSQL)
            FormateaGridALM(oForm)

#Region "Lista de precios de compras"
            sSQL = "SELECT 'N' as ""Sel"", ""ListNum"" ""LST"", ""ListName"" ""Nombre"" FROM OPLN WHERE ""U_EXO_TARCOM""='Si' ORDER BY ""ListName"""
            oForm.DataSources.DataTables.Item("DTCOMP").ExecuteQuery(sSQL)
            FormateaGridCOMP(oForm)
#End Region
#Region "Lista de precios de Ventas"
            sSQL = "SELECT 'N' as ""Sel"", ""ListNum"" ""LST"", ""ListName"" ""Nombre"" FROM OPLN WHERE ""U_EXO_TARCOM""='No' ORDER BY ""ListName"""
            oForm.DataSources.DataTables.Item("DTVENT").ExecuteQuery(sSQL)
            FormateaGridVENT(oForm)
#End Region

            oForm.DataSources.UserDataSources.Item("UDDIAS").ValueEx = "7"
            oForm.DataSources.UserDataSources.Item("UDTSM").ValueEx = "1"
            oForm.Items.Item("btnGen").Enabled = False
            'CType(oForm.Items.Item("txtARTD").Specific, SAPbouiCOM.EditText).Active = True
            CargarForm = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            oForm.State = BoFormStateEnum.fs_Maximized
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub FormateaGridVENT(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grdVENT").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdVENT").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 1 To 2
                CType(oform.Items.Item("grdVENT").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grdVENT").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
            Next



            CType(oform.Items.Item("grdVENT").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGridCOMP(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grdCOMP").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdCOMP").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 1 To 2
                CType(oform.Items.Item("grdCOMP").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grdCOMP").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
            Next



            CType(oform.Items.Item("grdCOMP").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
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

Module StringExtension
    <Extension()>
    Function DoubleParseAdvanced(ByVal strDouble As String) As Double
        Dim strDoubleNormalized As String
        If String.IsNullOrEmpty(strDouble) Then Return 0

        If strDouble.Contains(",") Then
            Dim strReplaced = strDouble.Replace(",", ".")
            Dim decimalSeparatorPos = strReplaced.LastIndexOf("."c)
            Dim strInteger = strReplaced.Substring(0, decimalSeparatorPos)
            Dim strFractional = strReplaced.Substring(decimalSeparatorPos)
            strInteger = strInteger.Replace(".", String.Empty)
            strDoubleNormalized = strInteger & strFractional
        Else
            strDoubleNormalized = strDouble
        End If

        Return Double.Parse(strDoubleNormalized, NumberStyles.Any, CultureInfo.InvariantCulture)
    End Function
End Module