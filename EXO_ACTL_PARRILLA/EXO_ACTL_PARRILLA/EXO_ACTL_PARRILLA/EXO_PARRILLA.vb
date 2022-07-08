Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_PARRILLA
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean

        Dim sSQL As String = ""
        Try
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnPARC"
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
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_PARRILLA.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                Else
                    objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
                    Exit Function
                End If
            End Try
            sSQL = "SELECT T2.""WhsCode"",T2.""WhsName"" "
            sSQL &= " From OUSR T0  "
            sSQL &= " INNER JOIN OHEM T1 ON T0.""USERID"" = T1.""userId"" "
            sSQL &= " INNER JOIN OWHS T2 ON T2.""U_EXO_SUCURSAL""=T1.""branch"" "
            sSQL &= " WHERE T0.""USERID""='" & objGlobal.compañia.UserSignature.ToString & "' order by T2.""WhsName"" "
            Try
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)
            Catch ex As Exception

            End Try

            sSQL = " SELECT CAST('-' as NVARCHAR(50)) ""TrnspCode"", CAST(' ' AS NVARCHAR(150))  ""TrnspName"" "
            sSQL &= " FROM DUMMY "
            sSQL &= " UNION ALL "
            sSQL &= " SELECT CAST(""TrnspCode"" as NVARCHAR(50)) ,""TrnspName"" "
            sSQL &= " From OSHP  "
            sSQL &= " ORDER By  ""TrnspName"" "
            Try
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbEXPE").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                CType(oForm.Items.Item("cbEXPE").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)
            Catch ex As Exception

            End Try

            sSQL = " SELECT CAST('-' as NVARCHAR(50)) ""territryID"", CAST(' ' AS NVARCHAR(150))  ""descript"" "
            sSQL &= " FROM DUMMY "
            sSQL &= " UNION ALL "
            sSQL &= "SELECT CAST(""territryID"" as NVARCHAR(50)),""descript"" "
            sSQL &= " From OTER  "
            sSQL &= " ORDER By  ""descript"" "
            Try
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbTERRI").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                CType(oForm.Items.Item("cbTERRI").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)
            Catch ex As Exception

            End Try

            CType(oForm.Items.Item("cbSAL").Specific, SAPbouiCOM.ComboBox).Select("TODOS", BoSearchKey.psk_ByValue)
            CType(oForm.Items.Item("cbENT").Specific, SAPbouiCOM.ComboBox).Select("TODOS", BoSearchKey.psk_ByValue)

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
    Public Function CargarFormRSTOCK(ByRef oFormParrilla As SAPbouiCOM.Form) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)
        Dim dtDatos As System.Data.DataTable = Nothing
        Dim dt As SAPbouiCOM.DataTable = Nothing
        CargarFormRSTOCK = False

        Try
            'Rellenar grid
            If oFormParrilla.DataSources.DataTables.Item("DTSPTE").Rows.Count > 0 Then
                oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
                oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_RSTOCK.srf")

                Try
                    oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
                Catch ex As Exception
                    If ex.Message.StartsWith("Form - already exists") = True Then
                        objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                        Exit Function
                    Else
                        objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
                        Exit Function
                    End If
                End Try

                dt = Nothing : dt = oFormParrilla.DataSources.DataTables.Item("DTSPTE")
                dtDatos = New System.Data.DataTable
                ComprobarDOCSEL(oFormParrilla, "DTSPTE", dtDatos, dt)
                sSQL = "SELECT ""ObjType"" ""TIPO"", ""DocEntry"" ""Nº INTERNO"", ""DocNum"" ""Documento"", ""LineNum"" ""Nº LINEA"", ""ItemCode"" ""ARTÍCULO"", ""ALMACEN"" ""ALMACÉN"", ""OpenQty"" ""CANTIDAD"" FROM ""EXO_ROTURA"" "
                If dtDatos.Rows.Count > 0 Then
                    sSQL &= " WHERE ""ALMACEN""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""DocEntry"" in ("
                    Dim bComa As Boolean = False
                    For Each MiDataRow As DataRow In dtDatos.Rows
                        If bComa = True Then
                            sSQL &= ", "
                        Else
                            bComa = True
                        End If
                        sSQL &= "'" & MiDataRow("Nº INTERNO").ToString & "' "
                    Next
                    sSQL &= ")"
                End If
                oForm.DataSources.DataTables.Item("DTSTOCK").ExecuteQuery(sSQL)
                FormateaGrid_RSTOCK(oForm)
                CargarFormRSTOCK = True
            Else
                objGlobal.SBOApp.StatusBar.SetText("No hay datos para mostrar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox("No hay datos para mostrar.")
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            If oForm IsNot Nothing Then
                oForm.Visible = True
            End If

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
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
                        Case "EXO_PARRILLA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

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
                        Case "EXO_PARRILLA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

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
                        Case "EXO_PARRILLA"
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
                        Case "EXO_PARRILLA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

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
    Private Function EventHandler_FORM_RESIZE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Items.Item("grdSPTE").Height = 140
            oForm.Items.Item("Item_5").Top = oForm.Items.Item("grdSPTE").Top - 15
            oForm.Items.Item("btLPicking").Top = oForm.Items.Item("grdSPTE").Top + 5
            oForm.Items.Item("btCPed").Top = oForm.Items.Item("grdSPTE").Top + 31
            oForm.Items.Item("btCALM").Top = oForm.Items.Item("grdSPTE").Top + 57
            oForm.Items.Item("btCCEXP").Top = oForm.Items.Item("grdSPTE").Top + 83
            oForm.Items.Item("btASS").Top = oForm.Items.Item("grdSPTE").Top + 109

            oForm.Items.Item("grdSLIB").Height = 140
            oForm.Items.Item("Item_6").Top = oForm.Items.Item("grdSLIB").Top - 15
            oForm.Items.Item("btGENALB").Top = oForm.Items.Item("grdSLIB").Top + 10
            oForm.Items.Item("btCCEXPL").Top = oForm.Items.Item("btGENALB").Top + 88

            oForm.Items.Item("grdSCOM").Height = 140
            oForm.Items.Item("Item_12").Top = oForm.Items.Item("grdSCOM").Top - 15
            oForm.Items.Item("btCCEXPC").Top = oForm.Items.Item("grdSCOM").Top + 5
            oForm.Items.Item("btImpD").Top = oForm.Items.Item("btCCEXPC").Top + 50
            oForm.Items.Item("btIMPE").Top = oForm.Items.Item("btImpD").Top + 50

            oForm.Items.Item("grdE").Height = 140
            oForm.Items.Item("Item_18").Top = oForm.Items.Item("grdE").Top - 15
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
        EventHandler_COMBO_SELECT_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                If pVal.ItemUID = "grdSCOM" And pVal.ColUID = "CLASE EXP." Then
                    Dim sExpe As String = CType(CType(oForm.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item("CLASE EXP."), SAPbouiCOM.ComboBoxColumn).GetSelectedValue(pVal.Row).Value.ToString
                    'Buscamos la agencia
                    sSQL = "SELECT ""U_EXO_AGE"" FROM OSHP WHERE ""TrnspCode""='" & sExpe & "' "
                    Dim sAGE As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    sAGE = IIf(sAGE = "", "-1", sAGE)
                    CType(oForm.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).DataTable.SetValue("AG. TRANSPORTE", pVal.Row, sAGE)
                End If
            End If

            EventHandler_COMBO_SELECT_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oItem, Object))
        End Try
    End Function
    Private Function EventHandler_MATRIX_LINK_PRESSED(ByVal pVal As ItemEvent) As Boolean

        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim sTipo As String = ""
        EventHandler_MATRIX_LINK_PRESSED = False

        Try
            oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                Return True
            End If
            Select Case pVal.ItemUID
                Case "grdSPTE", "grdSLIB"
                    oColumnTxt = CType(CType(oForm.Items.Item(pVal.ItemUID.ToString.Trim).Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.EditTextColumn)
                    sTipo = CType(oForm.Items.Item(pVal.ItemUID.ToString.Trim).Specific, SAPbouiCOM.Grid).DataTable.GetValue("T. SALIDA", pVal.Row).ToString

                    Select Case sTipo
                        Case "PEDVTA" 'Pedidos de ventas
                            oColumnTxt.LinkedObjectType = BoLinkedObject.lf_Order
                        Case "SDPROV" ' Sol de devolución de proveedor
                            oColumnTxt.LinkedObjectType = "234000032"
                        Case "SOLTRA" ' Sol de traslado
                            oColumnTxt.LinkedObjectType = BoLinkedObject.lf_StockTransfersRequest
                    End Select
                Case "grdSCOM"
                    oColumnTxt = CType(CType(oForm.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.EditTextColumn)
                    sTipo = CType(oForm.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).DataTable.GetValue("T. SALIDA", pVal.Row).ToString
                    Select Case sTipo
                        Case "ALBVTA" 'Albaranes de ventas
                            oColumnTxt.LinkedObjectType = BoLinkedObject.lf_DeliveryNotes
                        Case "SDPROV" 'Devolución de proveedor
                            oColumnTxt.LinkedObjectType = "21"
                        Case "SOLTRA" 'Sol de traslado
                            oColumnTxt.LinkedObjectType = BoLinkedObject.lf_StockTransfersRequest
                    End Select
                Case "grdE"
                    oColumnTxt = CType(CType(oForm.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.EditTextColumn)
                    sTipo = CType(oForm.Items.Item("grdE").Specific, SAPbouiCOM.Grid).DataTable.GetValue("T. ENTRADA", pVal.Row).ToString
                    Select Case sTipo
                        Case "PED" 'Pedidos de compra
                            oColumnTxt.LinkedObjectType = BoLinkedObject.lf_PurchaseOrder
                        Case "SDE" ' Solicitud de devolución de Clientes
                            oColumnTxt.LinkedObjectType = "234000031"
                        Case "STR" ' Solicitud de traslado Destino
                            oColumnTxt.LinkedObjectType = BoLinkedObject.lf_StockTransfersRequest
                    End Select
                    oColumnTxt = CType(CType(oForm.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(8), SAPbouiCOM.EditTextColumn)
            End Select



            EventHandler_MATRIX_LINK_PRESSED = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByVal pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oDataTable As DataTable = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        EventHandler_Choose_FromList_After = False

        Try
            oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                Return True
            End If

            oCFLEvento = CType(pVal, IChooseFromListEvent)

            oDataTable = oCFLEvento.SelectedObjects
            If Not oDataTable Is Nothing Then
                Select Case oCFLEvento.ChooseFromListUID
                    Case "CFLICD"
                        oDataTable = oCFLEvento.SelectedObjects

                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "txtICD" Then
                                Try
                                    CType(oForm.Items.Item("txtICD").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                                Catch ex As Exception

                                End Try
                            End If
                        End If
                    Case "CFLICH"
                        oDataTable = oCFLEvento.SelectedObjects

                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "txtICH" Then
                                Try
                                    CType(oForm.Items.Item("txtICH").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                                Catch ex As Exception

                                End Try
                            End If
                        End If
                End Select
            End If

            EventHandler_Choose_FromList_After = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.FormDatatable(oDataTable)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnFIL" ' Filtro
                    FiltrarPDTE(oForm)
                    FiltrarLIB(oForm)
                    FiltrarCOM(oForm)
                    FiltrarENT(oForm)
                Case "btLPicking" ' Liberar picking
                    If ComprobarDOC(oForm, "DTSPTE") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Liberando para picking... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If LiberarparaPicking(oForm, "DTSPTE", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarPDTE(oForm)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btCPed" ' Cerrar Documentos
                    If ComprobarDOC(oForm, "DTSPTE") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Cerrando documentos... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If CerrarDocumentos(oForm, "DTSPTE", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarPDTE(oForm)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btCCEXP" 'Cambio clase de expedición
                    If ComprobarDOC(oForm, "DTSPTE") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Cambiando clase de expedición... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If CambiarClaseExpedicion(oForm, "DTSPTE", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarPDTE(oForm)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btASS" 'Acceso a Art. sin Stocks
                    CargarFormRSTOCK(oForm)
                Case "btCALM" 'Cambio de almacén
                    If ComprobarDOCPED(oForm, "DTSPTE") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Cambiando almacén de documentos... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If CambiarAlmacen(oForm, "DTSPTE", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarPDTE(oForm)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btGENALB" 'Generar Albaranes
                    If ComprobarDOC(oForm, "DTSLIB") = True Then
                        objGlobal.SBOApp.StatusBar.SetText("Generando Documentos... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If Gen_DOC(oForm, "DTSLIB", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btCCEXPL" ' Cambio de clase de exp. liberadas
                    If ComprobarDOC(oForm, "DTSLIB") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Cambiando clase de expedición... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If CambiarClaseExpedicion(oForm, "DTSLIB", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btCCEXPC" ' Cambio de clase de exp. completadas
                    If ComprobarDOC(oForm, "DTSCOM") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Cambiando clase de expedición... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If CambiarClaseExpedicion(oForm, "DTSCOM", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btImpD" ' Impresión de documentos
                    If ComprobarDOC(oForm, "DTSCOM") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Imprimiendo documentos... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If Impresion_Doc(oForm, "DTSCOM", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btIMPE" 'Impresión de etiquetas
                    If ComprobarDOC(oForm, "DTSCOM") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Imprimiendo Etiquetas... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If Impresion_Doc(oForm, "DTSCOM", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Shared Function Impresion_ET(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Impresion_ET = False
#Region "VARIABLES"
        Dim oCmpSrv As SAPbobsCOM.CompanyService = oobjGlobal.compañia.GetCompanyService()
        Dim oReportLayoutService As SAPbobsCOM.ReportLayoutsService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), SAPbobsCOM.ReportLayoutsService)
        Dim oPrintParam As SAPbobsCOM.ReportLayoutPrintParams = CType(oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams), SAPbobsCOM.ReportLayoutPrintParams)
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim sLayout As String = "" : Dim sSQL As String = ""
#End Region

        Try
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" And oForm.DataSources.DataTables.Item(sData).GetValue("ESTADO", i).ToString = "EC" Then 'Sólo los registros que se han seleccionado
                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                    Select Case sTIPODOC
                        Case "PEDVTA"
#Region "Entregas"
                            sSQL = "SELECT ""DEFLT_REP"" FROM RTYP WHERE left(""CODE"",4)='DLN2' "
#End Region
                        Case "SOLTRA" ' Sol. de Traslado                           
#Region "Sol de traslado"
                            sSQL = "SELECT ""DEFLT_REP"" FROM RTYP WHERE left(""CODE"",4)='WTQ1' "
#End Region
                        Case "DPROV" ' Dev. de Proveedor
#Region "Dev de proveedor"
                            sSQL = "SELECT ""DEFLT_REP"" FROM RTYP WHERE left(""CODE"",4)='RPD2' "
#End Region
                    End Select
                    sLayout = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    oPrintParam.LayoutCode = sLayout 'codigo del formato importado en SAP
                    oPrintParam.DocEntry = sDocEntry 'parametro que se envia al crystal, DocEntry de la transaccion

                    oReportLayoutService.Print(oPrintParam)
                End If
            Next

            Impresion_ET = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oReportLayoutService = Nothing
            oCmpSrv = Nothing
        End Try
    End Function
    Public Shared Function Impresion_Doc(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Impresion_Doc = False
#Region "VARIABLES"
        Dim oCmpSrv As SAPbobsCOM.CompanyService = oobjGlobal.compañia.GetCompanyService()
        Dim oReportLayoutService As SAPbobsCOM.ReportLayoutsService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), SAPbobsCOM.ReportLayoutsService)
        Dim oPrintParam As SAPbobsCOM.ReportLayoutPrintParams = CType(oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams), SAPbobsCOM.ReportLayoutPrintParams)
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim sLayout As String = "" : Dim sSQL As String = ""
#End Region

        Try
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                    Select Case sTIPODOC
                        Case "PEDVTA"
#Region "Entregas"
                            sSQL = sSQL = "SELECT ""DEFLT_REP"" FROM RTYP WHERE left(""CODE"",4)='DLN2' "

#End Region
                        Case "SOLTRA" ' Sol. de Traslado                           
#Region "Sol de traslado"
                            sSQL = "SELECT ""DEFLT_REP"" FROM RTYP WHERE left(""CODE"",4)='WTQ1' "
#End Region
                        Case "DPROV" ' Dev. de Proveedor
#Region "Dev de proveedor"
                            sSQL = "SELECT ""DEFLT_REP"" FROM RTYP WHERE left(""CODE"",4)='RPD2' "
#End Region
                    End Select
                    sLayout = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    oPrintParam.LayoutCode = sLayout 'codigo del formato importado en SAP
                    oPrintParam.DocEntry = sDocEntry 'parametro que se envia al crystal, DocEntry de la transaccion

                    oReportLayoutService.Print(oPrintParam)
                End If
            Next

            Impresion_Doc = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oReportLayoutService = Nothing
            oCmpSrv = Nothing
        End Try
    End Function

    Public Shared Function Gen_DOC(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Gen_DOC = False
#Region "VARIABLES"
        Dim oRs As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sDocEntryFinal As String = "" : Dim sDocNumFinal As String = ""
        Dim oDocuments As SAPbobsCOM.Documents = Nothing : Dim oDocument_Lines As SAPbobsCOM.Document_Lines = Nothing
        Dim oDocFinal As SAPbobsCOM.Documents = Nothing : Dim oDocFinal_Lines As SAPbobsCOM.Document_Lines = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing : Dim oDocStockTransfer_Lines As SAPbobsCOM.StockTransfer_Lines = Nothing
#End Region

        Try
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                    Select Case sTIPODOC
                        Case "PEDVTA" ' Pedido 
                            oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders), SAPbobsCOM.Documents)

                            If oDocuments.GetByKey(sDocEntry) = True Then
                                oDocument_Lines = oDocuments.Lines
                                oDocFinal = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes), SAPbobsCOM.Documents)
                                oDocFinal_Lines = oDocFinal.Lines
                                oDocFinal.CardCode = oDocuments.CardCode
                                oDocFinal.DocDate = Now.Date
                                oDocFinal.TaxDate = Now.Date
                                For J = 0 To oDocument_Lines.Count - 1
                                    If (J > 0) Then
                                        oDocFinal_Lines.Add()
                                    End If
                                    oDocument_Lines.SetCurrentLine(J)
                                    oDocFinal_Lines.BaseObjectType = oDocuments.DocObjectCode
                                    oDocFinal_Lines.OrderEntry = oDocuments.DocEntry
                                    oDocFinal_Lines.OrderRowID = J
                                    oDocFinal_Lines.ReleasedQuantity = oDocument_Lines.RemainingOpenQuantity
                                Next
                                If oDocFinal.Add() <> 0 Then
                                    oobjGlobal.SBOApp.StatusBar.SetText("Error al generar la recepción del pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    oobjGlobal.compañia.GetNewObjectCode(sDocEntryFinal)
                                    sSQL = "SELECT ""DocNum"" FROM """ & oobjGlobal.compañia.CompanyDB & """.""OPDN"" WHERE ""DocEntry"" = " & sDocEntryFinal
                                    oRs.DoQuery(sSQL)
                                    If oRs.RecordCount > 0 Then
                                        sSQL = "UPDATE ORDR SET ""U_EXO_STATUSP""='C' WHERE ""DocEntry""=" & sDocEntry
                                        If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Actualizado Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                        sDocNumFinal = oRs.Fields.Item("DocNum").Value.ToString
                                        oobjGlobal.SBOApp.StatusBar.SetText("Entrega Nº: " & sDocNumFinal & " del Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    Else
                                        sDocNumFinal = "0"
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la entrega para el pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If

                                End If
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para para generar la recepción con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        Case "SOLTRA" ' Sol. de Traslado
                            oDocStockTransfer = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                            If oDocStockTransfer.GetByKey(sDocEntry) = True Then
                                oDocStockTransfer_Lines = oDocStockTransfer.Lines
                                oDocFinal = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer), SAPbobsCOM.StockTransfer)
                                oDocFinal_Lines = oDocFinal.Lines
                                For J = 0 To oDocStockTransfer_Lines.Count - 1
                                    If (J > 0) Then
                                        oDocFinal_Lines.Add()
                                    End If
                                    oDocStockTransfer_Lines.SetCurrentLine(J)
                                    oDocFinal_Lines.BaseObjectType = oDocStockTransfer.DocObjectCode
                                    oDocFinal_Lines.OrderEntry = oDocStockTransfer.DocEntry
                                    oDocFinal_Lines.OrderRowID = J
                                    oDocFinal_Lines.ReleasedQuantity = oDocStockTransfer_Lines.RemainingOpenQuantity
                                Next
                                If oDocFinal.Add() <> 0 Then
                                    oobjGlobal.SBOApp.StatusBar.SetText("Error al generar traslado de la Sol. de Traslado Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    oobjGlobal.compañia.GetNewObjectCode(sDocEntryFinal)
                                    sSQL = "SELECT ""DocNum"" FROM """ & oobjGlobal.compañia.CompanyDB & """.""OWTR"" WHERE ""DocEntry"" = " & sDocEntryFinal
                                    oRs.DoQuery(sSQL)
                                    If oRs.RecordCount > 0 Then
                                        sSQL = "UPDATE OWTQ SET ""U_EXO_STATUSP""='C' WHERE ""DocEntry""=" & sDocEntry
                                        If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Actualizado Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                        sDocNumFinal = oRs.Fields.Item("DocNum").Value.ToString
                                        oobjGlobal.SBOApp.StatusBar.SetText("Traslado Nº: " & sDocNumFinal & " de la Sol. de Traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    Else
                                        sDocNumFinal = "0"
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el traslado generado para la Sol. de Traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If

                                End If
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Traslado para generar el traslado con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        Case "SDPROV" ' Sol. de dev. 
                            sSQL = "UPDATE OPRR SET ""U_EXO_STATUSP""='C' WHERE ""DocEntry""=" & sDocEntry
                            If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                oobjGlobal.SBOApp.StatusBar.SetText("Picking Liberado Sol. de dev. de Proveedor con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("Error en Picking Liberado Sol. de dev. de Proveedor con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                    End Select
                End If
            Next

            Gen_DOC = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            oDocFinal = Nothing : oDocFinal_Lines = Nothing
            oDocuments = Nothing : oDocument_Lines = Nothing
            oDocStockTransfer = Nothing : oDocStockTransfer_Lines = Nothing
        End Try
    End Function
    Public Shared Function CambiarClaseExpedicion(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        CambiarClaseExpedicion = False
#Region "VARIABLES"
        Dim oRs As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sIC As String = "" : Dim sClaseExp As String = ""
        Dim sAgenenClase As String = "" : Dim sAgenciaListaNegra As String = ""
        Dim oDocuments As SAPbobsCOM.Documents = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing
#End Region

        Try
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                    sIC = oForm.DataSources.DataTables.Item(sData).GetValue("CÓDIGO", i).ToString
                    sClaseExp = oForm.DataSources.DataTables.Item(sData).GetValue("CLASE EXP.", i).ToString
#Region "Comprobamos que la clase de expedicion sea permitida y no este en la lista negra"
                    sSQL = " SELECT ""U_EXO_AGE"" FROM OSHP WHERE ""TrnspCode""='" & sClaseExp & "' "
                    sAgenenClase = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    sSQL = " SELECT ""U_EXO_COD"" FROM ""@EXO_LNEGRAL"" WHERE ""Code""='" & sIC & "' and ""U_EXO_COD""='" & sAgenenClase & "' "
                    sAgenciaListaNegra = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    Dim bActualiza As Boolean = True
                    If sAgenciaListaNegra <> "" Then
                        oobjGlobal.SBOApp.StatusBar.SetText("En el documento Nº: " & sDocNum & ", la clase de expedición tiene asignada la agencia """ & sAgenciaListaNegra & """ en la lista negra. No puede actualizarlo." & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        bActualiza = False
                    End If

                    If bActualiza = True Then
                        sSQL = " SELECT ""U_EXO_COD"" FROM ""@EXO_LNEGRAL"" WHERE ""Code""='" & sIC & "' and ""U_EXO_COD""='" & sAgenenClase & "' "
                        sAgenciaListaNegra = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                        If sAgenciaListaNegra <> "" Then
                            oobjGlobal.SBOApp.StatusBar.SetText("En el documento Nº: " & sDocNum & ", la clase de expedición tiene asignada la agencia """ & sAgenciaListaNegra & """ en la lista negra. No puede actualizarlo." & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Select Case sTIPODOC
                                Case "PEDVTA" ' Pedido de venta
                                    oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                                    If oDocuments.GetByKey(sDocEntry) = True Then
                                        oDocuments.TransportationCode = sClaseExp
                                        If oDocuments.Update() <> 0 Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar el pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Se ha actualizado la clase de expedición en el pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        End If
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para cambiar la clase de expedición con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                Case "SOLTRA" ' Sol. de Traslado                           
                                    oDocStockTransfer = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                                    If oDocStockTransfer.GetByKey(sDocEntry) = True Then
                                        oDocStockTransfer.UserFields.Fields.Item("U_EXO_CLASEE").Value = sClaseExp
                                        If oDocStockTransfer.Update() <> 0 Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar la Sol. de traslado Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Se ha actualizado la clase de expedición en la Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        End If
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Traslado  Nº: " & sDocNum & ". No s epuede cerrar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                Case "SDPROV" ' Sol. de dev. de Proveedor
                                    oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oGoodsReturnRequest), SAPbobsCOM.Documents)
                                    If oDocuments.GetByKey(sDocEntry) = True Then
                                        oDocuments.TransportationCode = sClaseExp
                                        If oDocuments.Update() <> 0 Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar la clase de expedición de la Sol. de Dev de proveedor Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Sol. de Dev de proveedor actualizada Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        End If
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Dev de proveedor para cerrarla con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                            End Select
                        End If
                    End If
#End Region
                End If
            Next

            CambiarClaseExpedicion = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

            oDocStockTransfer = Nothing
            oDocuments = Nothing
        End Try
    End Function
    Public Shared Function CerrarDocumentos(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        CerrarDocumentos = False
#Region "VARIABLES"
        Dim oRs As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim oDocuments As SAPbobsCOM.Documents = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing
#End Region

        Try
            If oobjGlobal.SBOApp.MessageBox("¿Está seguro que quiere cerrar los Documentos seleccionados?", 1, "Sí", "No") = 1 Then
                For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                    If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                        sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                        sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                        sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                        Select Case sTIPODOC
                            Case "PEDVTA" ' Pedido de venta
                                oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                                If oDocuments.GetByKey(sDocEntry) = True Then
                                    If oDocuments.Close() <> 0 Then
                                        oobjGlobal.SBOApp.StatusBar.SetText("Error al cerrar el pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("Pedido cerrado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para cerrarlo con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            Case "SOLTRA" ' Sol. de Traslado                           
                                oDocStockTransfer = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                                If oDocStockTransfer.GetByKey(sDocEntry) = True Then
                                    If oDocStockTransfer.Close() <> 0 Then
                                        oobjGlobal.SBOApp.StatusBar.SetText("Error al cerrar la Sol. de traslado Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("Sol. de traslado cerrada con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Traslado  Nº: " & sDocNum & ". No s epuede cerrar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            Case "SDPROV" ' Sol. de dev. de Proveedor
                                oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oGoodsReturnRequest), SAPbobsCOM.Documents)
                                If oDocuments.GetByKey(sDocEntry) = True Then
                                    If oDocuments.Close() <> 0 Then
                                        oobjGlobal.SBOApp.StatusBar.SetText("Error al cerrar la Sol. de Dev de proveedor Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("Sol. de Dev de proveedor cerrada Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Dev de proveedor para cerrarla con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                        End Select
                    End If
                Next
            End If
            CerrarDocumentos = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

            oDocStockTransfer = Nothing
            oDocuments = Nothing
        End Try
    End Function
    Public Shared Function LiberarparaPicking(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        LiberarparaPicking = False
#Region "VARIABLES"
        Dim oRs As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sDocEntryPicking As String = "" : Dim sDocNumPicking As String = ""
        Dim oPickLists As SAPbobsCOM.PickLists = Nothing : Dim oPickLists_Lines As SAPbobsCOM.PickLists_Lines = Nothing
        Dim oDocuments As SAPbobsCOM.Documents = Nothing : Dim oDocument_Lines As SAPbobsCOM.Document_Lines = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing : Dim oDocStockTransfer_Lines As SAPbobsCOM.StockTransfer_Lines = Nothing
#End Region

        Try
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                    Select Case sTIPODOC
                        Case "PEDVTA" ' Pedido de venta
                            oPickLists = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists), SAPbobsCOM.PickLists)
                            oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                            If oDocuments.GetByKey(sDocEntry) = True Then
                                oDocument_Lines = oDocuments.Lines
                                oPickLists = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists), SAPbobsCOM.PickLists)
                                oPickLists_Lines = oPickLists.Lines
                                For J = 0 To oDocument_Lines.Count - 1
                                    If (J > 0) Then
                                        oPickLists_Lines.Add()
                                    End If
                                    oDocument_Lines.SetCurrentLine(J)
                                    oPickLists_Lines.BaseObjectType = "17"
                                    oPickLists_Lines.OrderEntry = oDocuments.DocEntry
                                    oPickLists_Lines.OrderRowID = J
                                    oPickLists_Lines.ReleasedQuantity = oDocument_Lines.Quantity
                                Next
                                If oPickLists.Add() <> 0 Then
                                    oobjGlobal.SBOApp.StatusBar.SetText("Error al liberar Picking del pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    oobjGlobal.compañia.GetNewObjectCode(sDocEntryPicking)
                                    sSQL = "SELECT ""AbsEntry"" FROM """ & oobjGlobal.compañia.CompanyDB & """.""OPKL"" WHERE ""AbsEntry"" = " & sDocEntryPicking
                                    oRs.DoQuery(sSQL)
                                    If oRs.RecordCount > 0 Then
                                        sSQL = "UPDATE ORDR SET ""U_EXO_STATUSP""='L' WHERE ""DocEntry""=" & sDocEntry
                                        If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Actualizado Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                        sDocNumPicking = oRs.Fields.Item("AbsEntry").Value.ToString
                                        oobjGlobal.SBOApp.StatusBar.SetText("Picking Liberado Nº: " & sDocNumPicking & " del Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    Else
                                        sDocNumPicking = "0"
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el Picking generado para el pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If

                                End If
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para liberar Picking con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        Case "SOLTRA" ' Sol. de Traslado
                            oPickLists = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists), SAPbobsCOM.PickLists)
                            oDocStockTransfer = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                            If oDocStockTransfer.GetByKey(sDocEntry) = True Then
                                oDocStockTransfer_Lines = oDocStockTransfer.Lines
                                oPickLists = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists), SAPbobsCOM.PickLists)
                                oPickLists_Lines = oPickLists.Lines
                                For J = 0 To oDocStockTransfer_Lines.Count - 1
                                    If (J > 0) Then
                                        oPickLists_Lines.Add()
                                    End If
                                    oDocStockTransfer_Lines.SetCurrentLine(J)
                                    oPickLists_Lines.BaseObjectType = "1250000001"
                                    oPickLists_Lines.OrderEntry = oDocStockTransfer.DocEntry
                                    oPickLists_Lines.OrderRowID = J
                                    oPickLists_Lines.ReleasedQuantity = oDocStockTransfer_Lines.RemainingOpenQuantity
                                Next
                                If oPickLists.Add() <> 0 Then
                                    oobjGlobal.SBOApp.StatusBar.SetText("Error al liberar Picking de la Sol. de Traslado Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    oobjGlobal.compañia.GetNewObjectCode(sDocEntryPicking)
                                    sSQL = "SELECT ""AbsEntry"" FROM """ & oobjGlobal.compañia.CompanyDB & """.""OPKL"" WHERE ""AbsEntry"" = " & sDocEntryPicking
                                    oRs.DoQuery(sSQL)
                                    If oRs.RecordCount > 0 Then
                                        sSQL = "UPDATE OWTQ SET ""U_EXO_STATUSP""='L' WHERE ""DocEntry""=" & sDocEntry
                                        If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Actualizado Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                        sDocNumPicking = oRs.Fields.Item("AbsEntry").Value.ToString
                                        oobjGlobal.SBOApp.StatusBar.SetText("Picking Liberado Nº: " & sDocNumPicking & " de la Sol. de Traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    Else
                                        sDocNumPicking = "0"
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el Picking generado para la Sol. de Traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If

                                End If
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Traslado para liberar Picking con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        Case "SDPROV" ' Sol. de dev. de Proveedor
                            sSQL = "UPDATE OPRR SET ""U_EXO_STATUSP""='L' WHERE ""DocEntry""=" & sDocEntry
                            If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                oobjGlobal.SBOApp.StatusBar.SetText("Picking Liberado Sol. de dev. de Proveedor con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("Error en Picking Liberado Sol. de dev. de Proveedor con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                    End Select
                End If
            Next

            LiberarparaPicking = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            oPickLists = Nothing : oPickLists_Lines = Nothing
            oDocuments = Nothing : oDocument_Lines = Nothing
        End Try
    End Function
    Public Shared Function CambiarAlmacen(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        CambiarAlmacen = False
#Region "VARIABLES"
        Dim oRs As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim oDocuments As SAPbobsCOM.Documents = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing
        Dim bActualiza As Boolean = False
        Dim sDelPedido As String = "" : Dim sALMPedido As String = "" : Dim sALM As String = "" : Dim sDelALM As String = ""
        Dim sDocEntryTraslado As String = "" : Dim sDocNumTraslado As String = ""
#End Region

        Try
            If oobjGlobal.SBOApp.MessageBox("¿Está seguro de cambiar el almacén a los documentos seleccionados?", 1, "Sí", "No") = 1 Then
                For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                    If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                        sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                        sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                        sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                        bActualiza = False
                        Select Case sTIPODOC
                            Case "PEDVTA" ' Pedido de venta
                                oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                                If oDocuments.GetByKey(sDocEntry) = True Then
                                    sDelPedido = oDocuments.UserFields.Fields.Item("U_EXO_DELE").Value.ToString
                                    sSQL = "SELECT ""WhsCode"" FROM OWHS WHERE ""U_EXO_SUCURSAL""='" & sDelPedido & "' "
                                    sALMPedido = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                                    For lin = 0 To oDocuments.Lines.Count - 1
                                        oDocuments.Lines.SetCurrentLine(lin)
                                        sALM = oDocuments.Lines.WarehouseCode.ToString
                                        sSQL = "SELECT ""U_EXO_SUCURSAL"" FROm OWHS Where ""WhsCode""='" & sALM & "' "
                                        sDelALM = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                                        If sDelPedido <> sDelALM Then
                                            bActualiza = True
                                            Exit For
                                        End If
                                    Next
                                    If bActualiza = True Then
                                        If oobjGlobal.SBOApp.MessageBox("¿Quiere generar la solicitud de traslado para el pedido Nº" & sDocNum & "? ", 1, "Sí", "No") = 1 Then
#Region "Gen. Sol de Traslado"
                                            oDocStockTransfer = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                                            oDocStockTransfer.DocDate = Now.Date
                                            oDocStockTransfer.CardCode = oDocuments.CardCode
                                            oDocStockTransfer.FromWarehouse = sALM
                                            oDocStockTransfer.ToWarehouse = sALMPedido
                                            oDocStockTransfer.Comments = "Generado desde la Parrilla. Pedido Nº " & sDocNum
                                            If oDocStockTransfer.DocumentReferences.LineNumber.ToString > 1 Then
                                                oDocStockTransfer.DocumentReferences.Add()
                                                oDocStockTransfer.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_SalesOrder
                                                oDocStockTransfer.DocumentReferences.ReferencedDocEntry = sDocEntry
                                            Else
                                                oDocStockTransfer.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_SalesOrder
                                                oDocStockTransfer.DocumentReferences.ReferencedDocEntry = sDocEntry
                                            End If
                                            Dim bGrabalinea As Boolean = False
                                            For lin = 0 To oDocuments.Lines.Count - 1
                                                oDocuments.Lines.SetCurrentLine(lin)
                                                sALM = oDocuments.Lines.WarehouseCode.ToString
                                                sSQL = "SELECT ""U_EXO_SUCURSAL"" FROm OWHS Where ""WhsCode""='" & sALM & "' "
                                                sDelALM = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                                                If sDelPedido <> sDelALM Then

                                                    If bGrabalinea = True Then
                                                        oDocStockTransfer.Lines.Add()
                                                    Else
                                                        bGrabalinea = True
                                                    End If
                                                    oDocStockTransfer.Lines.ItemCode = oDocuments.Lines.ItemCode
                                                    oDocStockTransfer.Lines.Quantity = oDocuments.Lines.RemainingOpenQuantity
                                                    oDocStockTransfer.Lines.FromWarehouseCode = oDocuments.Lines.WarehouseCode.ToString
                                                    oDocuments.Lines.WarehouseCode = sALMPedido
                                                    oDocStockTransfer.Lines.WarehouseCode = sALMPedido
                                                End If
                                            Next
                                            If oDocStockTransfer.Add() <> 0 Then
                                                oobjGlobal.SBOApp.StatusBar.SetText("Error creando la Sol. de traslado. No se actualiza el pedido." & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                bActualiza = False
                                            Else
                                                sDocEntryTraslado = oobjGlobal.compañia.GetNewObjectKey()
                                                sSQL = "SELECT ""DocNum"" FROM OWTQ Where ""DocEntry""=" & sDocEntryTraslado
                                                sDocNumTraslado = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                                                oobjGlobal.SBOApp.StatusBar.SetText("Sol. de traslado creada con Nº: " & sDocNumTraslado, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                bActualiza = True
                                            End If
#End Region
                                        End If
                                        If bActualiza = True Then
                                            'hacemos referencia al documento creado
                                            If oDocuments.DocumentReferences.LineNumber.ToString > 1 Then
                                                oDocuments.DocumentReferences.Add()
                                                oDocuments.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_InventoryTransferRequest
                                                oDocuments.DocumentReferences.ReferencedDocEntry = sDocEntryTraslado
                                            Else
                                                oDocuments.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_InventoryTransferRequest
                                                oDocuments.DocumentReferences.ReferencedDocEntry = sDocEntryTraslado
                                            End If
                                            If oDocuments.Update() <> 0 Then
                                                oobjGlobal.SBOApp.StatusBar.SetText("Error modificar  el pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Else
                                                oobjGlobal.SBOApp.StatusBar.SetText("Pedido modificado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                            End If
                                        End If
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("El Pedido Nº: " & sDocNum & " no se modifica. La delegación del pedido es la misma que la del almacén.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If

                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para cambiar el almacén con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                        End Select
                    End If
                Next
            End If

            CambiarAlmacen = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

            oDocStockTransfer = Nothing
            oDocuments = Nothing
        End Try
    End Function
    Private Function ComprobarDOC(ByRef oForm As SAPbouiCOM.Form, ByVal sTABLA As String) As Boolean
        Dim bLineasSel As Boolean = False

        ComprobarDOC = False

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sTABLA).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sTABLA).GetValue("Sel", i).ToString = "Y" Then
                    bLineasSel = True
                    Exit For
                End If
            Next

            If bLineasSel = False Then
                objGlobal.SBOApp.MessageBox("Debe seleccionar al menos una línea.")
                Exit Function
            End If

            ComprobarDOC = bLineasSel

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ComprobarDOCPED(ByRef oForm As SAPbouiCOM.Form, ByVal sTABLA As String) As Boolean
        Dim bLineasSel As Boolean = False

        ComprobarDOCPED = False

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sTABLA).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sTABLA).GetValue("Sel", i).ToString = "Y" And oForm.DataSources.DataTables.Item(sTABLA).GetValue("T. SALIDA", i).ToString = "PEDVTA" Then
                    bLineasSel = True
                    Exit For
                End If
            Next

            If bLineasSel = False Then
                objGlobal.SBOApp.MessageBox("Debe seleccionar al menos una línea de pedido.")
                Exit Function
            End If

            ComprobarDOCPED = bLineasSel

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ComprobarDOCSEL(ByRef oForm As SAPbouiCOM.Form, ByVal sTABLA As String, ByRef dtDatos As System.Data.DataTable, ByRef dt As SAPbouiCOM.DataTable) As Boolean
        ComprobarDOCSEL = False

        Try
            For iCol As Integer = 0 To 12
                dtDatos.Columns.Add(dt.Columns.Item(iCol).Name)
            Next

            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sTABLA).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sTABLA).GetValue("Sel", i).ToString = "Y" Then
                    'Añadimos los registros
                    Dim oRow As DataRow = dtDatos.NewRow
                    For iCol As Integer = 0 To 12
                        oRow.Item(dt.Columns.Item(iCol).Name) = dt.Columns.Item(iCol).Cells.Item(i).Value
                    Next
                    dtDatos.Rows.Add(oRow)
                End If
            Next
            ComprobarDOCSEL = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub FiltrarPDTE(ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim sSalidas As String = ""
        Dim sICD As String = "" : Dim sICH As String = ""
        Dim sEXPE As String = "" : Dim sTerri As String = ""
        Dim sSQL As String = ""
#End Region
        Try
            sSalidas = oForm.DataSources.UserDataSources.Item("UDSAL").Value.ToString
            sICD = oForm.DataSources.UserDataSources.Item("UDICD").Value.ToString
            sICH = oForm.DataSources.UserDataSources.Item("UDICH").Value.ToString
            sEXPE = oForm.DataSources.UserDataSources.Item("UDEXPE").Value.ToString
            sTerri = oForm.DataSources.UserDataSources.Item("UDTERRI").Value.ToString
            oForm.Freeze(True)
            Select Case sSalidas
                Case "-"
                    sSQL = "SELECT CAST('' as nVARCHAR(50)) ""T. SALIDA"", CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", CAST('' as nVARCHAR(50)) ""Nº INTERNO"", CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " 'N' ""AUTORIZADO"", CAST('' as nVARCHAR(50)) ""CÓDIGO"",  CAST('' as nVARCHAR(150))	""EMPRESA"", CAST('' as nVARCHAR(50)) ""CLASE EXP."", 'N' ""ROT. STOCK"", "
                    sSQL &= " 'N' ""A"", CAST('' as nVARCHAR(50)) ""UBICACIÓN"", CAST('' as nVARCHAR(50)) ""ZONA TRANSPORTE"", 'N' ""Sel"" "
                    sSQL &= "FROM DUMMY "
                Case "TODOS"
#Region "Todos"
                    sSQL = "SELECT ""T. SALIDA"", ""DELEGACIÓN"", ""Nº INTERNO"", ""Nº DOCUMENTO"", ""AUTORIZADO"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", "
                    sSQL &= " ""ROT. STOCK"", ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel""  FROM ""EXO_PEDIDOS_VENTA"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and ""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""CLASE EXP.""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT ""T. SALIDA"", ""DELEGACIÓN"", ""Nº INTERNO"", ""Nº DOCUMENTO"", ""AUTORIZADO"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", ""ROT. STOCK"", "
                    sSQL &= " ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel"" FROM ""EXO_SOL_TRASLADO"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and ""FromWhsCod""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT ""T. SALIDA"", ""DELEGACIÓN"", ""Nº INTERNO"", ""Nº DOCUMENTO"", ""AUTORIZADO"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", ""ROT. STOCK"", "
                    sSQL &= " ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel"" FROM ""EXO_SOL_DEVOLUCION"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and ""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""CLASE EXP.""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "PEDVTA"
#Region "Pedidos de Ventas"
                    'sSQL = "SELECT DISTINCT CAST('PEDVTA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    'sSQL &= " T0.""Confirmed"" ""AUTORIZADO"", CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", "
                    'sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    'sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", 'N' ""Sel"" "
                    'sSQL &= "FROM ORDR T0 "
                    'sSQL &= " LEFT JOIN RDR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    'sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    'sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    'sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    'sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    'sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    'sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    'sSQL &= " WHERE TL.""LineStatus""='O' and T0.""Confirmed""='Y' and T0.""U_EXO_STATUSP""='P' "
                    sSQL = "SELECT ""T. SALIDA"", ""DELEGACIÓN"", ""Nº INTERNO"", ""Nº DOCUMENTO"", ""AUTORIZADO"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", "
                    sSQL &= " ""ROT. STOCK"", ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel""  FROM ""EXO_PEDIDOS_VENTA"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and ""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""CLASE EXP.""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "SOLTRA"
#Region "Sol de traslado"
                    'sSQL = "SELECT DISTINCT CAST('SOLTRA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    'sSQL &= " T0.""Confirmed"" ""AUTORIZADO"", CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50)) ""CLASE EXP."", "
                    'sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    'sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", 'N' ""Sel"" "
                    'sSQL &= "FROM OWTQ T0 "
                    'sSQL &= " LEFT JOIN WTQ1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    'sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    'sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    'sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    'sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    'sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    'sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    'sSQL &= " WHERE TL.""LineStatus""='O' and T0.""U_EXO_STATUSP""='P' "
                    sSQL = "SELECT ""T. SALIDA"", ""DELEGACIÓN"", ""Nº INTERNO"", ""Nº DOCUMENTO"", ""AUTORIZADO"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", ""ROT. STOCK"", "
                    sSQL &= " ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel"" FROM ""EXO_SOL_TRASLADO"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and ""FromWhsCod""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "SDPROV"
#Region "Sol de Devolución"
                    'sSQL = "SELECT DISTINCT CAST('SDPROV' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    'sSQL &= " T0.""Confirmed"" ""AUTORIZADO"", CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", "
                    'sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    'sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", 'N' ""Sel""  "
                    'sSQL &= "FROM OPRR T0 "
                    'sSQL &= " LEFT JOIN PRR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    'sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    'sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    'sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    'sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    'sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    'sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    'sSQL &= " WHERE TL.""LineStatus""='O' and T0.""Confirmed""='Y' and T0.""U_EXO_STATUSP""='P' "
                    sSQL = "SELECT ""T. SALIDA"", ""DELEGACIÓN"", ""Nº INTERNO"", ""Nº DOCUMENTO"", ""AUTORIZADO"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", ""ROT. STOCK"", "
                    sSQL &= " ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel"" FROM ""EXO_SOL_DEVOLUCION"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and ""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""CLASE EXP.""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
#End Region
            End Select
            oForm.DataSources.DataTables.Item("DTSPTE").ExecuteQuery(sSQL)
            FormateaGrid_SPTE(oForm)

                objGlobal.SBOApp.StatusBar.SetText("Datos de salidas Pdtes. Cargados con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub FiltrarLIB(ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim sSalidas As String = ""
        Dim sICD As String = "" : Dim sICH As String = ""
        Dim sEXPE As String = "" : Dim sTerri As String = ""
        Dim sSQL As String = ""
#End Region
        Try
            sSalidas = oForm.DataSources.UserDataSources.Item("UDSAL").Value.ToString
            sICD = oForm.DataSources.UserDataSources.Item("UDICD").Value.ToString
            sICH = oForm.DataSources.UserDataSources.Item("UDICH").Value.ToString
            sEXPE = oForm.DataSources.UserDataSources.Item("UDEXPE").Value.ToString
            sTerri = oForm.DataSources.UserDataSources.Item("UDTERRI").Value.ToString
            oForm.Freeze(True)
            Select Case sSalidas
                Case "-"
                    sSQL = "SELECT CAST('' as nVARCHAR(50)) ""T. SALIDA"", CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", CAST('' as nVARCHAR(50)) ""Nº INTERNO"", CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST('' as nVARCHAR(50)) ""CÓDIGO"",  CAST('' as nVARCHAR(150))	""EMPRESA"", CAST('' as nVARCHAR(50)) ""CLASE EXP."", 'N' ""ROT. STOCK"", "
                    sSQL &= " 'N' ""A"", CAST('' as nVARCHAR(50)) ""UBICACIÓN"", CAST('' as nVARCHAR(50)) ""ZONA TRANSPORTE"", 'N' ""Sel"" "
                    sSQL &= "FROM DUMMY "
                Case "TODOS"
#Region "Todos"
                    sSQL = "SELECT DISTINCT CAST('PEDVTA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"",  "
                    sSQL &= " IFNULL(PK.""Cant."",0) ""Cant."", IFNULL(PK.""Cant."" - PK.""Cant. P"",0) ""Cant. Pdte."", 'N' ""Sel"" "
                    sSQL &= "FROM ORDR T0 "
                    sSQL &= " LEFT JOIN RDR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN ""VEXO_PICKING"" PK ON PK.""BaseObject""= T0.""ObjType"" and PK.""OrderEntry""= TL.""DocEntry"" and PK.""OrderLine""= TL.""LineNum"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and T0.""Confirmed""='Y' and TL.""PickStatus""<>'N' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT DISTINCT CAST('SOLTRA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", "
                    sSQL &= " IFNULL(PK.""Cant."",0) ""Cant."", IFNULL(PK.""Cant."" - PK.""Cant. P"",0) ""Cant. Pdte."", 'N' ""Sel""  "
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN ""VEXO_PICKING"" PK ON PK.""BaseObject""= T0.""ObjType"" and PK.""OrderEntry""= TL.""DocEntry"" and PK.""OrderLine""= TL.""LineNum"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and T0.""U_EXO_STATUSP""='L' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT DISTINCT CAST('SDPROV' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", "
                    sSQL &= " IFNULL(TL.""Quantity"",0) ""Cant."", 0 ""Cant. Pdte."", 'N' ""Sel"" "
                    sSQL &= "FROM OPRR T0 "
                    sSQL &= " LEFT JOIN PRR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and T0.""Confirmed""='Y' and T0.""U_EXO_STATUSP""='L' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "PEDVTA"
#Region "Pedidos de Ventas"
                    sSQL = "SELECT DISTINCT CAST('PEDVTA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", "
                    sSQL &= " IFNULL(PK.""Cant."",0) ""Cant."", IFNULL(PK.""Cant."" - PK.""Cant. P"",0) ""Cant. Pdte."", 'N' ""Sel"" "
                    sSQL &= "FROM ORDR T0 "
                    sSQL &= " LEFT JOIN RDR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN ""VEXO_PICKING"" PK ON PK.""BaseObject""= T0.""ObjType"" and PK.""OrderEntry""= TL.""DocEntry"" and PK.""OrderLine""= TL.""LineNum"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and T0.""Confirmed""='Y' and TL.""PickStatus""<>'N' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "SOLTRA"
#Region "Sol de traslado"
                    sSQL = "SELECT DISTINCT CAST('SOLTRA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", "
                    sSQL &= " IFNULL(PK.""Cant."",0) ""Cant."", IFNULL(PK.""Cant."" - PK.""Cant. P"",0) ""Cant. Pdte."", 'N' ""Sel"" "
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN ""VEXO_PICKING"" PK ON PK.""BaseObject""= T0.""ObjType"" and PK.""OrderEntry""= TL.""DocEntry"" and PK.""OrderLine""= TL.""LineNum"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and T0.""U_EXO_STATUSP""='L' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "SDPROV"
#Region "Sol de Devolución"
                    sSQL = "SELECT DISTINCT CAST('SDPROV' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", "
                    sSQL &= " IFNULL(TL.""Quantity"",0) ""Cant."", 0 ""Cant. Pdte."",'N' ""Sel""  "
                    sSQL &= "FROM OPRR T0 "
                    sSQL &= " LEFT JOIN PRR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and T0.""Confirmed""='Y' and T0.""U_EXO_STATUSP""='L' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
            End Select
            oForm.DataSources.DataTables.Item("DTSLIB").ExecuteQuery(sSQL)
            FormateaGrid_SLIB(oForm)
                objGlobal.SBOApp.StatusBar.SetText("Datos de salida liberados Cargados con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub FiltrarCOM(ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim sSalidas As String = ""
        Dim sICD As String = "" : Dim sICH As String = ""
        Dim sEXPE As String = "" : Dim sTerri As String = ""
        Dim sSQL As String = ""
#End Region
        Try
            sSalidas = oForm.DataSources.UserDataSources.Item("UDSAL").Value.ToString
            sICD = oForm.DataSources.UserDataSources.Item("UDICD").Value.ToString
            sICH = oForm.DataSources.UserDataSources.Item("UDICH").Value.ToString
            sEXPE = oForm.DataSources.UserDataSources.Item("UDEXPE").Value.ToString
            sTerri = oForm.DataSources.UserDataSources.Item("UDTERRI").Value.ToString
            oForm.Freeze(True)
            Select Case sSalidas
                Case "-"
                    sSQL = "SELECT CAST('' as nVARCHAR(50)) ""T. SALIDA"", CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", CAST('' as nVARCHAR(50)) ""Nº INTERNO"", CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST('' as nVARCHAR(50)) ""CÓDIGO"",  CAST('' as nVARCHAR(150))	""EMPRESA"", CAST('' as nVARCHAR(50)) ""CLASE EXP."", CAST('' as nVARCHAR(50)) ""AG. TRANSPORTE"",  "
                    sSQL &= " 'PE' ""ESTADO"", 'N' ""Sel"" "
                    sSQL &= "FROM DUMMY "
                Case "TODOS"
#Region "Todos"
                    sSQL = "SELECT DISTINCT CAST('ALBVTA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"",  "
                    sSQL &= " CASE WHEN IFNULL(E.""U_EXO_DOCNUM"",'')='' THEN 'PE' ELSE 'EE' END ""ESTADO"", 'N' ""Sel"" "
                    sSQL &= "FROM ODLN T0 "
                    sSQL &= " LEFT JOIN DLN1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=T0.""TrnspCode"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON TL.""DocEntry""=E.""U_EXO_DOCENTRY"" and TL.""LineNum""=E.""U_EXO_LINNUM"" "
                    sSQL &= " WHERE T0.""U_EXO_STATUSP""='C' and (T0.""U_EXO_ESTPAC""='Pendiente' or T0.""U_EXO_ESTPAC""='En curso') "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= " SELECT DISTINCT CAST('SOLTRA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50)) ""CLASE EXP."", IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"", "
                    sSQL &= " CASE WHEN IFNULL(E.""U_EXO_DOCNUM"",'')='' THEN 'PE' ELSE 'EE' END ""ESTADO"", 'N' ""Sel"" "
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=T0.""U_EXO_CLASEE"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON TL.""DocEntry""=E.""U_EXO_DOCENTRY"" and TL.""LineNum""=E.""U_EXO_LINNUM"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " WHERE T0.""U_EXO_STATUSP""='C' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= " SELECT DISTINCT CAST('DPROV' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"",  "
                    sSQL &= " CASE WHEN IFNULL(E.""U_EXO_DOCNUM"",'')='' THEN 'PE' ELSE 'EE' END ""ESTADO"", 'N' ""Sel"" "
                    sSQL &= " FROM ORPD T0 "
                    sSQL &= " LEFT JOIN RPD1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=T0.""TrnspCode"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON TL.""DocEntry""=E.""U_EXO_DOCENTRY"" and TL.""LineNum""=E.""U_EXO_LINNUM"" "
                    sSQL &= " WHERE T0.""U_EXO_STATUSP""='C' and (T0.""U_EXO_ESTPAC""='Pendiente' or T0.""U_EXO_ESTPAC""='En curso') "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "PEDVTA"
#Region "Entregas de Ventas"
                    sSQL = " SELECT DISTINCT CAST('ALBVTA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"",  "
                    sSQL &= " CASE WHEN IFNULL(E.""U_EXO_DOCNUM"",'')='' THEN 'PE' ELSE 'EE' END ""ESTADO"", 'N' ""Sel"" "
                    sSQL &= " FROM ODLN T0 "
                    sSQL &= " LEFT JOIN DLN1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=T0.""TrnspCode"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON TL.""DocEntry""=E.""U_EXO_DOCENTRY"" and TL.""LineNum""=E.""U_EXO_LINNUM"" "
                    sSQL &= " WHERE T0.""U_EXO_STATUSP""='C' and (T0.""U_EXO_ESTPAC""='Pendiente' or T0.""U_EXO_ESTPAC""='En curso') "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "SOLTRA"
#Region "Sol de traslado"
                    sSQL = "SELECT DISTINCT CAST('SOLTRA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50)) ""CLASE EXP."", IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"", "
                    sSQL &= " CASE WHEN IFNULL(E.""U_EXO_DOCNUM"",'')='' THEN 'PE' ELSE 'EE' END ""ESTADO"", 'N' ""Sel"" "
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=T0.""U_EXO_CLASEE"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON TL.""DocEntry""=E.""U_EXO_DOCENTRY"" and TL.""LineNum""=E.""U_EXO_LINNUM"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " WHERE T0.""U_EXO_STATUSP""='C' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "SDPROV"
#Region "Devolución"
                    sSQL = "SELECT DISTINCT CAST('DPROV' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"",  "
                    sSQL &= " CASE WHEN IFNULL(E.""U_EXO_DOCNUM"",'')='' THEN 'PE' ELSE 'EE' END ""ESTADO"", 'N' ""Sel"" "
                    sSQL &= " FROM ORPD T0 "
                    sSQL &= " LEFT JOIN RPD1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=T0.""TrnspCode"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON TL.""DocEntry""=E.""U_EXO_DOCENTRY"" and TL.""LineNum""=E.""U_EXO_LINNUM"" "
                    sSQL &= " WHERE T0.""U_EXO_STATUSP""='C' and (T0.""U_EXO_ESTPAC""='Pendiente' or T0.""U_EXO_ESTPAC""='En curso') "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
            End Select
            oForm.DataSources.DataTables.Item("DTSCOM").ExecuteQuery(sSQL)
            FormateaGrid_SCOM(oForm)

                objGlobal.SBOApp.StatusBar.SetText("Datos de salida completadas Cargados con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub FiltrarENT(ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim sEntradas As String = ""
        Dim sICD As String = "" : Dim sICH As String = ""
        Dim sEXPE As String = "" : Dim sTerri As String = ""
        Dim sSQL As String = ""
#End Region
        Try
            sEntradas = oForm.DataSources.UserDataSources.Item("UDENT").Value.ToString
            sICD = oForm.DataSources.UserDataSources.Item("UDICD").Value.ToString
            sICH = oForm.DataSources.UserDataSources.Item("UDICH").Value.ToString
            sEXPE = oForm.DataSources.UserDataSources.Item("UDEXPE").Value.ToString
            sTerri = oForm.DataSources.UserDataSources.Item("UDTERRI").Value.ToString
            oForm.Freeze(True)
            Select Case sEntradas
                Case "-"
                    sSQL = "SELECT CAST('' as nVARCHAR(50)) ""T. ENTRADA"", CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", CAST('' as nVARCHAR(50)) ""Nº INTERNO"", CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST('' as nVARCHAR(50)) ""CÓDIGO"",  CAST('' as nVARCHAR(150))	""EMPRESA"", CAST('' as nVARCHAR(50)) ""ESTADO"", CAST('' as nVARCHAR(50)) ""DOC. ENTRADA"", "
                    sSQL &= " CAST('' as nVARCHAR(50)) ""ID DOC. ENTRADA"""
                    sSQL &= "FROM DUMMY "
                Case "TODOS"
#Region "Todos"
                    sSQL = "SELECT DISTINCT CAST('PED' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    sSQL &= " CAST((CASE WHEN T0.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"""
                    sSQL &= " FROM OPOR T0 "
                    sSQL &= " LEFT JOIN POR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN PDN1 T3 ON T0.""DocEntry""=T3.""BaseEntry"" and T0.""ObjType""=T3.""BaseType"" "
                    sSQL &= " Left JOIN OPDN T4 ON T3.""DocEntry""=T4.""DocEntry"" "
                    sSQL &= " WHERE ((TL.""LineStatus""='O' and T0.""DocDueDate""<='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "') "
                    sSQL &= " OR ( T4.""DocDate""='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "')) "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT DISTINCT CAST('STR' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    sSQL &= " CAST((CASE WHEN T0.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"""
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL On TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 On T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 On T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R On R.""DocEntry""=T0.""DocEntry"" And R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S On S.""DocEntry""=T0.""DocEntry"" And S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A On A.""CardCode""=T0.""CardCode"" And A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT On T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN WTR1 T3 ON T0.""DocEntry""=T3.""BaseEntry"" and T0.""ObjType""=T3.""BaseType"" "
                    sSQL &= " Left JOIN OWTR T4 ON T3.""DocEntry""=T4.""DocEntry"" "
                    sSQL &= " WHERE ((TL.""LineStatus""='O' and T0.""DocDueDate""<='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "') "
                    sSQL &= " OR ( T4.""DocDate""='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "')) "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT DISTINCT CAST('SDE' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    sSQL &= " CAST((CASE WHEN T0.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"""
                    sSQL &= "FROM ORRR T0 "
                    sSQL &= " LEFT JOIN RRR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN RDN1 T3 ON T0.""DocEntry""=T3.""BaseEntry"" and T0.""ObjType""=T3.""BaseType"" "
                    sSQL &= " Left JOIN ORDN T4 ON T3.""DocEntry""=T4.""DocEntry"" "
                    sSQL &= " WHERE ((TL.""LineStatus""='O' and T0.""DocDueDate""<='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "') "
                    sSQL &= " OR ( T4.""DocDate""='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "')) "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "PED"
#Region "Pedidos de compra"
                    sSQL = "SELECT DISTINCT CAST('PED' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    sSQL &= " CAST((CASE WHEN T0.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"""
                    sSQL &= " FROM OPOR T0 "
                    sSQL &= " LEFT JOIN POR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN PDN1 T3 ON T0.""DocEntry""=T3.""BaseEntry"" and T0.""ObjType""=T3.""BaseType"" "
                    sSQL &= " Left JOIN OPDN T4 ON T3.""DocEntry""=T4.""DocEntry"" "
                    sSQL &= " WHERE ((TL.""LineStatus""='O' and T0.""DocDueDate""<='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "') "
                    sSQL &= " OR ( T4.""DocDate""='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "')) "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "STR"
#Region "Sol de traslado en destino"
                    sSQL = "SELECT DISTINCT CAST('STR' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    sSQL &= " CAST((CASE WHEN T0.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"""
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL On TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 On T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 On T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R On R.""DocEntry""=T0.""DocEntry"" And R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S On S.""DocEntry""=T0.""DocEntry"" And S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A On A.""CardCode""=T0.""CardCode"" And A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT On T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN WTR1 T3 ON T0.""DocEntry""=T3.""BaseEntry"" and T0.""ObjType""=T3.""BaseType"" "
                    sSQL &= " Left JOIN OWTR T4 ON T3.""DocEntry""=T4.""DocEntry"" "
                    sSQL &= " WHERE ((TL.""LineStatus""='O' and T0.""DocDueDate""<='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "') "
                    sSQL &= " OR ( T4.""DocDate""='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "')) "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
                Case "SDE"
#Region "Sol de Devolución de cliente"
                    sSQL = "SELECT DISTINCT CAST('SDE' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    sSQL &= " CAST((CASE WHEN T0.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"""
                    sSQL &= "FROM ORRR T0 "
                    sSQL &= " LEFT JOIN RRR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN RDN1 T3 ON T0.""DocEntry""=T3.""BaseEntry"" and T0.""ObjType""=T3.""BaseType"" "
                    sSQL &= " Left JOIN ORDN T4 ON T3.""DocEntry""=T4.""DocEntry"" "
                    sSQL &= " WHERE ((TL.""LineStatus""='O' and T0.""DocDueDate""<='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "') "
                    sSQL &= " OR ( T4.""DocDate""='" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "')) "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
#End Region
            End Select
            oForm.DataSources.DataTables.Item("DTE").ExecuteQuery(sSQL)
            FormateaGrid_E(oForm)

                objGlobal.SBOApp.StatusBar.SetText("Datos de Entrada Cargados con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_E(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)

            For i = 0 To 8
                Select Case i
                    Case 0
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("PED", "Pedido de compra")
                        oColumnCb.ValidValues.Add("STR", "Sol. de traslado Destino")
                        oColumnCb.ValidValues.Add("SDE", "Solicitud de devolución de Clientes")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 2
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "22"
                        oColumnTxt.Editable = False
                    Case 8
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "20"
                        oColumnTxt.Editable = False

                    Case Else
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next
            CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_SCOM(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(9).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(9), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 0 To 8
                Select Case i
                    Case 0
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("ALBVTA", "Entrega de clientes")
                        oColumnCb.ValidValues.Add("SOLTRA", "Sol. de traslado Origen")
                        oColumnCb.ValidValues.Add("DPROV", "Dev. de Proveedor")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 2
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "17"
                        oColumnTxt.Editable = False
                    Case 4
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "2"
                        oColumnTxt.Editable = False
                    Case 6
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)

                        sSQL = " SELECT CAST(""TrnspCode"" as NVARCHAR(50)) ,""TrnspName"" "
                        sSQL &= " From OSHP  "
                        sSQL &= " ORDER By  ""TrnspName"" "
                        Try
                            objGlobal.funcionesUI.cargaCombo(oColumnCb.ValidValues, sSQL)
                            oColumnCb.ValidValues.Add("-1", " ")
                            oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        Catch ex As Exception

                        End Try
                        oColumnCb.Editable = True
                    Case 7
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)

                        sSQL = " SELECT ""CardCode"" ,""CardFName"" "
                        sSQL &= " From OCRD  "
                        Try
                            objGlobal.funcionesUI.cargaCombo(oColumnCb.ValidValues, sSQL)
                            oColumnCb.ValidValues.Add("-1", " ")
                            oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        Catch ex As Exception

                        End Try
                        oColumnCb.Editable = False
                    Case 8
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("PE", "Pendiente de expedición")
                        oColumnCb.ValidValues.Add("EE", "En Expedición")
                        oColumnCb.ValidValues.Add("EC", "Expedición cerrada")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case Else
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next
            CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_SLIB(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(13).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(13), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 0 To 12
                Select Case i
                    Case 0
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("PEDVTA", "Pedido de clientes")
                        oColumnCb.ValidValues.Add("SOLTRA", "Sol. de traslado Origen")
                        oColumnCb.ValidValues.Add("SDPROV", "Sol. de dev. de Proveedor")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 2
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "17"
                        oColumnTxt.Editable = False
                    Case 4
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "2"
                        oColumnTxt.Editable = False
                    Case 7
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)

                        oColumnCb.ValidValues.Add("Y", "Sí")
                        oColumnCb.ValidValues.Add("N", "No")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                        oColumnCb.Visible = False
                    Case 8
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)

                        oColumnCb.ValidValues.Add("Y", "Sí")
                        oColumnCb.ValidValues.Add("N", "No")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 9, 10
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Visible = False
                    Case 6
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)

                        sSQL = " SELECT CAST(""TrnspCode"" as NVARCHAR(50)) ,""TrnspName"" "
                        sSQL &= " From OSHP  "
                        sSQL &= " ORDER By  ""TrnspName"" "
                        Try
                            objGlobal.funcionesUI.cargaCombo(oColumnCb.ValidValues, sSQL)
                            oColumnCb.ValidValues.Add("-1", " ")
                            oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        Catch ex As Exception

                        End Try
                        oColumnCb.Editable = True
                    Case 11, 12
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.RightJustified = True
                    Case Else
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next
            CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_SPTE(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(12).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(12), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 0 To 11
                Select Case i
                    Case 0
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("PEDVTA", "Pedido de clientes")
                        oColumnCb.ValidValues.Add("SOLTRA", "Sol. de traslado Origen")
                        oColumnCb.ValidValues.Add("SDPROV", "Sol. de dev. de Proveedor")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 2
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "17"
                        oColumnTxt.Editable = False
                    Case 5
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "2"
                        oColumnTxt.Editable = False
                    Case 4, 8, 9
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)

                        oColumnCb.ValidValues.Add("Y", "Sí")
                        oColumnCb.ValidValues.Add("N", "No")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 5
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "2"
                        oColumnTxt.Editable = False
                    Case 7
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)

                        sSQL = " SELECT CAST(""TrnspCode"" as NVARCHAR(50)) ,""TrnspName"" "
                        sSQL &= " From OSHP  "
                        sSQL &= " ORDER By  ""TrnspName"" "
                        Try
                            objGlobal.funcionesUI.cargaCombo(oColumnCb.ValidValues, sSQL)
                            oColumnCb.ValidValues.Add("-1", " ")
                            oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        Catch ex As Exception

                        End Try
                        oColumnCb.Editable = True
                    Case Else
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next
            CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_RSTOCK(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)

            For i = 0 To 6
                Select Case i
                    Case 0
                        CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("17", "Pedido de clientes")
                        oColumnCb.ValidValues.Add("1250000001", "Sol. de traslado Origen")
                        oColumnCb.ValidValues.Add("234000032", "Sol. de dev. de Proveedor")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 1
                        CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "17"
                        oColumnTxt.Editable = False
                    Case 4
                        CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "4"
                        oColumnTxt.Editable = False
                    Case 6
                        CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.RightJustified = True
                        oColumnTxt.Editable = False
                    Case Else
                        CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next
            CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
End Class
