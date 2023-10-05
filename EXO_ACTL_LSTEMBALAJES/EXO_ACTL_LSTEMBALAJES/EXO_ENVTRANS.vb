Imports System.IO
Imports System.Xml
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports OfficeOpenXml

Public Class EXO_ENVTRANS
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        SBOApp_MenuEvent = False
        Dim sSQL As String = ""
        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.MenuUID
                    Case "1286" 'Cerrar
                        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.ActiveForm
                        If oForm IsNot Nothing Then
                            If oForm.TypeEx = "UDO_FT_EXO_ENVTRANS" Then
                                If Cerrar_ENVIO(oForm) = False Then
                                    Return False
                                Else
                                    Return True
                                End If
                            End If
                        End If
                End Select
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnGETR"
                        If CargarUDO() = False Then
                            Exit Function
                        End If
                    Case "1282"
                        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.ActiveForm
                        If oForm IsNot Nothing Then
                            If oForm.TypeEx = "UDO_FT_EXO_ENVTRANS" Then
                                EXO_GLOBALES.Modo_Anadir(oForm, objGlobal)
                                Cargar_Combos(oForm)

                                If objGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
                                    objGlobal.SBOApp.ActivateMenuItem("1304")
                                End If
                            End If
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
    Public Function Cerrar_ENVIO(ByRef oform As SAPbouiCOM.Form) As Boolean
#Region "Variables"
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sAlmacen As String = "" : Dim sStatus As String = ""
        Dim sDocEntryFinal As String = "" : Dim sDocNumFinal As String = ""
        Dim dtDatos As System.Data.DataTable = Nothing

        Dim Omercancias As SAPbobsCOM.Documents = Nothing
        Dim sDocEntryCerrar As String = "" : Dim sDocNumCerrar As String = "" : Dim sStatusCerrar As String = ""
        Dim iLinea As Integer = 0

        Dim oGeneralService As SAPbobsCOM.GeneralService = Nothing
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams = Nothing
        Dim oCompService As SAPbobsCOM.CompanyService = objGlobal.compañia.GetCompanyService()
#End Region
        Cerrar_ENVIO = False
        Try
            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            objGlobal.compañia.StartTransaction()

            sDocEntry = oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("DocEntry", 0)
            If sDocEntry = "" Then
                sDocEntry = CType(oform.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value.ToString()
            End If
            sDocNum = oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("DocNum", 0)
            If sDocNum = "" Then
                sDocNum = CType(oform.Items.Item("1_U_E").Specific, SAPbouiCOM.EditText).Value.ToString()
            End If
            sAlmacen = oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("U_EXO_ALMACEN", 0)
            If sAlmacen = "" Then
                sAlmacen = CType(oform.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            End If

            objGlobal.SBOApp.StatusBar.SetText("Cerrando Documento Nº " & sDocNum & "...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sStatus = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Status"" FROM ""@EXO_ENVTRANS"" Where ""DocEntry""=" & sDocEntry)
            objGlobal.SBOApp.StatusBar.SetText("Status: " & sStatus, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If sStatus.Trim = "O" Then
                sSQL = "SELECT TC.""DocEntry"", TC.""DocNum"", COUNT(""U_EXO_IDBULTO"" || ' - ' ||  ""U_EXO_TBULTO"") ""Cantidad"", MAX(""U_EXO_IDBULTO"") ""ID BULTO"", MAX(""U_EXO_TBULTO"") ""BULTO""
                    FROM ""@EXO_LSTEMBL"" TL
                    INNER JOIN ""@EXO_LSTEMB"" TC ON TC.""DocEntry""=TL.""DocEntry""
                    WHERE TC.""Status""='O' and TL.""DocEntry"" IN (SELECT T0.""DocEntry""
           				                        FROM ""@EXO_LSTEMB""  T0 
            			                        Left Join  OCRD T1 ON T0.""U_EXO_IC"" = T1.""CardCode"" 
            			                        where T0.""U_EXO_IDENVIO"" =" & sDocEntry & ")
                    GROUP BY TC.""DocEntry"", TC.""DocNum"", ""U_EXO_IDBULTO"" || ' - ' ||  ""U_EXO_TBULTO""
                    ORDER BY TC.""DocEntry"", MAX(""U_EXO_TBULTO""),MAX(""U_EXO_IDBULTO"")"
                dtDatos = New System.Data.DataTable
                dtDatos = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                If dtDatos.Rows.Count > 0 Then
                    Omercancias = CType(objGlobal.compañia.GetBusinessObject(BoObjectTypes.oInventoryGenExit), SAPbobsCOM.Documents)
                    Omercancias.DocDate = Date.Now
                    Omercancias.TaxDate = Date.Now
                    For Each MiDataRow As DataRow In dtDatos.Rows
                        If sDocEntryCerrar <> MiDataRow("DocEntry").ToString Then
                            sDocEntryCerrar = MiDataRow("DocEntry").ToString
                            sSQL = "SELECT ""DocNum"" FROM ""@EXO_LSTEMB"" WHERE ""DocEntry""=" & sDocEntryCerrar
                            sDocNumCerrar = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                            sSQL = "SELECT ""Status"" FROM ""@EXO_LSTEMB"" WHERE ""DocEntry""=" & sDocEntryCerrar
                            sStatusCerrar = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                            If sStatusCerrar = "O" Then
                                objGlobal.SBOApp.StatusBar.SetText("Cerrando Lista de embalaje Nº: " & sDocNumCerrar, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
                                'Cerramos el UDO
                                'Get a handle to the SM_MOR UDO
                                oGeneralService = oCompService.GetGeneralService("EXO_LSTEMB")
                                'Close UDO record
                                oGeneralParams = CType(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams), SAPbobsCOM.GeneralDataParams)
                                oGeneralParams.SetProperty("DocEntry", sDocEntryCerrar)
                                oGeneralService.Close(oGeneralParams)
                                objGlobal.SBOApp.StatusBar.SetText("Se ha cerrado la Lista de Embalaje Nº: " & sDocNumCerrar, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)
                            End If
#Region "Lineas"
                            sSQL = "SELECT L.* FROM ""@EXO_PAQL"" L INNER JOIN ""@EXO_PAQ"" C ON C.""Code""=L.""Code"" WHERE C.""Name""='" & MiDataRow("BULTO").ToString & "' ORDER BY L.""LineId"" "
                            oRs.DoQuery(sSQL)
                            For i = 0 To oRs.RecordCount - 1
                                If iLinea <> 0 Then
                                    Omercancias.Lines.Add()
                                End If
                                Omercancias.Lines.ItemCode = oRs.Fields.Item("U_EXO_ITEMCODE").Value.ToString
                                Omercancias.Lines.Quantity = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, MiDataRow("Cantidad").ToString) * EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oRs.Fields.Item("U_EXO_CANT").Value.ToString)
                                Omercancias.Lines.WarehouseCode = sAlmacen
                                ' Omercancias.Lines.BatchNumbers.BatchNumber = ""
                                'Omercancias.Lines.BatchNumbers.Quantity = Omercancias.Lines.Quantity
                                ' Omercancias.Lines.BatchNumbers.Add()
                                Omercancias.Lines.UserFields.Fields.Item("U_EXO_ENVTRDE").Value = sDocEntryCerrar
                                Omercancias.Lines.UserFields.Fields.Item("U_EXO_ENVTRDN").Value = sDocNumCerrar
                                iLinea += 1

                                oRs.MoveNext()
                            Next
#End Region
                        End If

                    Next
                    Omercancias.Comments = "Generado automáticamente al cerrar Envío - Transporte Nº " & sDocNum
                    If Omercancias.Add() <> 0 Then
                        objGlobal.SBOApp.StatusBar.SetText("Error al generar Salida de Mercancía. " & objGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        If objGlobal.compañia.InTransaction = True Then
                            objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        Return False
                    Else
                        objGlobal.compañia.GetNewObjectCode(sDocEntryFinal)

                        If objGlobal.compañia.InTransaction = True Then
                            objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If
                        sSQL = "Select ""DocNum"" FROM """ & objGlobal.compañia.CompanyDB & """.""OIGE"" WHERE ""DocEntry"" = " & sDocEntryFinal
                        oRs.DoQuery(sSQL)
                        If oRs.RecordCount > 0 Then
                            sDocNumFinal = oRs.Fields.Item("DocNum").Value.ToString
                            'Actualizamos el UDO
                            sSQL = "UPDATE ""@EXO_ENVTRANS"" SET ""U_EXO_CONEMB""='" & sDocEntryFinal & "' WHERE ""DocEntry""=" & sDocEntry
                            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                objGlobal.SBOApp.StatusBar.SetText("Se ha generado la Salida de mercancía Nº: " & sDocNumFinal, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                Return True
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("No se ha podido actualizar el envío - Transporte Nº: " & sDocNum & " con el Nº de documento generado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Return False
                            End If

                        Else
                            sDocNumFinal = "0"
                            If objGlobal.compañia.InTransaction = True Then
                                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                            objGlobal.SBOApp.StatusBar.SetText("No se encuentra la Salida de mercancía con Nº Interno: " & sDocEntryFinal, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If

                    End If
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado Lista de embalajes para cerrar. Revise los datos. ", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
                    objGlobal.SBOApp.MessageBox("No se ha encontrado Lista de embalajes para cerrar. Revise los datos. ")
                    If objGlobal.compañia.InTransaction = True Then
                        objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    Return False
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("Este documento ya está cerrado.", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox("Este documento ya está cerrado.")
                If objGlobal.compañia.InTransaction = True Then
                    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If
                Return False
            End If


        Catch ex As Exception
            Throw ex
        Finally
            If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then 'Para que el combo enseñe la descripción
                If objGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
                    objGlobal.SBOApp.ActivateMenuItem("1304")
                End If
            End If

            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(Omercancias, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompService, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oGeneralParams, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oGeneralService, Object))
        End Try
    End Function
    Public Function CargarUDO() As Boolean
        CargarUDO = False

        Try
            objGlobal.funcionesUI.cargaFormUdoBD("EXO_ENVTRANS")

            CargarUDO = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally

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
                        Case "UDO_FT_EXO_ENVTRANS"
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

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_ENVTRANS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                    'If EventHandler_MATRIX_LINK_PRESSED(infoEvento) = False Then
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_ENVTRANS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(objGlobal, infoEvento) = False Then
                                        Return False
                                    End If
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
                        Case "UDO_FT_EXO_ENVTRANS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

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
    Private Function EventHandler_GOT_FOCUS_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_GOT_FOCUS_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "22_U_Cb"
                    If oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_FIND_MODE Then
                        If pVal.ItemChanged = True Then
                            Cargar_Combos(oForm)
                        End If
                    End If
            End Select

            EventHandler_GOT_FOCUS_After = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_Before(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "1_U_FD"
                    If oForm.Mode = BoFormMode.fm_OK_MODE Then
                        Rellena_Grid(oForm)
                    ElseIf oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_UPDATE_MODE Then
                        objGlobal.SBOApp.StatusBar.SetText("Grabe primero para poder ver las Expediciones.", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox("Grabe primero para poder ver las expediciones.")
                    End If
            End Select

            EventHandler_ItemPressed_Before = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Sub Rellena_Grid(ByRef oform As SAPbouiCOM.Form)
#Region "Variables"
        Dim sSQL As String = ""
        Dim sDocEntry As String = ""

#End Region
        Try
            sDocEntry = oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("DocEntry", 0)
            If sDocEntry = "" Then
                sDocEntry = CType(oform.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value.ToString()
            End If
            objGlobal.SBOApp.StatusBar.SetText("Documento Nº " & sDocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = "SELECT ""IdExpedicion"", ""U_EXO_DESTINO"" ""Destino"", ""Direccion"", ""Resumenbulto"" ""Resumen Bulto"", ""Peso"", ""Volumen"" "
            sSQL &= " FROM ""EXO_DetalleBultosEnvioTransporte""  "
            sSQL &= " where ""IdEnvioTTE"" =" & sDocEntry
            oform.DataSources.DataTables.Item("DTEX").ExecuteQuery(sSQL)
            FormateaGrid(oform)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub FormateaGrid(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)

            For i = 0 To 5
                Select Case i
                    Case 0
                        CType(oform.Items.Item("grdEX").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdEX").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                        oColumnTxt.LinkedObjectType = "EXO_LSTEMB"
                    Case Else
                        CType(oform.Items.Item("grdEX").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdEX").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next
            CType(oform.Items.Item("grdEX").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Function EventHandler_ItemPressed_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "1"
                    If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                        Cargar_Combos(oForm)
                    End If
                Case "btnMan"
                    objGlobal.SBOApp.StatusBar.SetText("Imprimiendo... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    If Impresion(oForm, objGlobal) = False Then
                        Exit Function
                    End If
                    objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Case "btnEFich"
                    objGlobal.SBOApp.StatusBar.SetText("Exportando en el directorio... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    If Exportacion(oForm, objGlobal) = False Then
                        Exit Function
                    End If
                    objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
    Public Shared Function Exportacion(ByRef oForm As SAPbouiCOM.Form, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Exportacion = False
#Region "VARIABLES"
        Dim sSQL As String = ""
        Dim sDocEntry As String = "" : Dim sProveedor As String = "" : Dim sProvName As String = ""
        Dim oRsC As SAPbobsCOM.Recordset = Nothing
        Dim nFila As Integer = 0 : Dim colIndex As Integer = 0
        Dim sFormato As String = "" : Dim sPath As String = ""
        Dim sFile As String = "" : Dim EMail As String = ""
        Dim pck As ExcelPackage = Nothing
#End Region

        Try
            oRsC = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sDocEntry = oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("DocEntry", 0)
            If sDocEntry = "" Then
                sDocEntry = CType(oForm.Items.Item("1_U_E").Specific, SAPbouiCOM.EditText).Value.ToString()
            End If

            sProveedor = oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("U_EXO_AGTCODE", 0)
            If sProveedor = "" Then
                sProveedor = CType(oForm.Items.Item("23_U_E").Specific, SAPbouiCOM.EditText).Value.ToString()
            End If

            sProvName = oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("U_EXO_AGTNAME", 0)
            If sProvName = "" Then
                sProvName = CType(oForm.Items.Item("24_U_E").Specific, SAPbouiCOM.EditText).Value.ToString()
            End If

            sSQL = "SELECT ""U_EXO_FORMATO"" FROM OCRD WHERE ""CardCode""='" & sProveedor & "' "
            sFormato = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
            If sFormato = "" Or sFormato = "0" Then
                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el formato del proveedor. Vaya al proveedor y elija un formato.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                oobjGlobal.SBOApp.MessageBox("No se encuentra el formato del proveedor. Vaya al proveedor y elija un formato.")
                Return False
            End If

            sSQL = "SELECT ""U_EXO_DIR"" FROM OCRD WHERE ""CardCode""='" & sProveedor & "' "
            sPath = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
            If sPath = "" Then
                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el directorio para exportar. Vaya al proveedor y elija un direcotrio de trabajo.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                oobjGlobal.SBOApp.MessageBox("No se encuentra el directorio para exportar. Vaya al proveedor y elija un direcotrio de trabajo.")
                Return False
            Else
                If Not System.IO.Directory.Exists(sPath) Then
                    System.IO.Directory.CreateDirectory(sPath)
                End If
            End If

            sSQL = "SELECT ""QString"" FROM OUQR WHERE ""IntrnalKey""=" & sFormato
            sSQL = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
            sSQL = sSQL.Replace("[%0]", sDocEntry)
            oRsC.DoQuery(sSQL)
            If oRsC.RecordCount >= 1 Then
                sFile = sProveedor & "_" & sDocEntry & "_" & Now.Year.ToString & Right("00" & Now.Month.ToString, 2) & Right("00" & Now.Day.ToString, 2) & ".xlsx"
                sFile = sPath & "\" & sFile

                ' miramos si existe el fichero y lo borramos
                If File.Exists(sFile) Then
                    File.Delete(sFile)
                End If

                'datos de cabecera
                Dim newFile As FileInfo = New FileInfo(sFile)
                pck = New ExcelPackage(newFile)

                Dim wSheet As ExcelWorksheet = pck.Workbook.Worksheets.Add("AG. Transporte")

                nFila = 1
                colIndex = 0
                For ColCount = 0 To oRsC.Fields.Count - 1

                    colIndex = colIndex + 1
                    wSheet.Cells(nFila, colIndex).Value = oRsC.Fields.Item(ColCount).Description
                    wSheet.Cells(nFila, colIndex).Style.Font.Bold = True
                    'If colIndex = 6 Then
                    '    wSheet.Cells(nFila, colIndex).Value = "Atrasos"
                    '    wSheet.Cells(nFila, colIndex).Style.Font.Color.SetColor(System.Drawing.Color.Red)
                    'ElseIf colIndex > 6 And colIndex <= iDFirme Then
                    '    wSheet.Cells(nFila, colIndex).Style.Font.Color.SetColor(System.Drawing.Color.Green)
                    'End If
                Next

                While Not oRsC.EoF
                    'filas
                    nFila = nFila + 1
                    colIndex = 0
                    For ColCount = 0 To oRsC.Fields.Count - 1
                        colIndex = colIndex + 1
                        'If colIndex >= 4 Then
                        '    wSheet.Cells(nFila, colIndex).Style.Numberformat.Format = "0"

                        'End If
                        wSheet.Cells(nFila, colIndex).Value = oRsC.Fields.Item(ColCount).Value.ToString
                        'wSheet.Cells(nFila, colIndex).Style.Font.Bold = True
                        'If colIndex = 6 Then
                        '    wSheet.Cells(nFila, colIndex).Style.Font.Color.SetColor(System.Drawing.Color.Red)
                        'ElseIf colIndex > 6 And colIndex <= iDFirme Then
                        '    wSheet.Cells(nFila, colIndex).Style.Font.Color.SetColor(System.Drawing.Color.Green)
                        'End If
                    Next

                    If oRsC.EoF = False Then
                        oRsC.MoveNext()
                    End If
                End While
                pck.Save()

#Region "Enviar por mail"
                If oobjGlobal.SBOApp.MessageBox("¿Quiere enviar el documento generado?", 1, "Sí", "No") = 1 Then
                    Dim Ficherohtml As String = oobjGlobal.pathHistorico
                    Ficherohtml = oobjGlobal.pathHistorico & "\mail.htm"
                    Dim srCAB As StreamReader = New StreamReader(Ficherohtml)
                    Dim cuerpo As String = srCAB.ReadToEnd()
                    sSQL = "SELECT ""E_Mail"" FROM OCRD WHERE ""CardCode""='" & sProveedor & "' "
                    EMail = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    'Dim cuerpo As String = oobjGlobal.leerEmbebido(Me.GetType(), "mail.htm")
                    If EMail <> "" Then
                        oobjGlobal.SBOApp.StatusBar.SetText("Enviando Correo ...", BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        EXO_GLOBALES.Enviarmail(oobjGlobal, cuerpo, EMail, sProveedor, sProvName, sFile)
                    Else
                        oobjGlobal.SBOApp.StatusBar.SetText("El proveedor " & sProveedor & " - " & sProvName & " no tiene indicado el mail. ", BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If

                End If
#End Region
            End If
            Exportacion = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsC, Object))
        End Try
    End Function
    Public Shared Function Impresion(ByRef oForm As SAPbouiCOM.Form, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Impresion = False
#Region "VARIABLES"
        Dim oCmpSrv As SAPbobsCOM.CompanyService = oobjGlobal.compañia.GetCompanyService()
        Dim oReportLayoutService As SAPbobsCOM.ReportLayoutsService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), SAPbobsCOM.ReportLayoutsService)
        Dim oPrintParam As SAPbobsCOM.ReportLayoutPrintParams = CType(oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams), SAPbobsCOM.ReportLayoutPrintParams)
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim sLayout As String = "" : Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sDocEntryEnvTransporte As String = ""
#End Region

        Try
            oRs = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sDocEntryEnvTransporte = oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("DocEntry", 0).ToString.Trim

            sLayout = oobjGlobal.funcionesUI.refDi.OGEN.valorVariable("EXO_MANIFIESTO")
            If sLayout = "" Then
                oobjGlobal.SBOApp.StatusBar.SetText("Parámetro [EXO_MANIFIESTO] no tiene valor. Revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                sSQL = "SELECT DISTINCT ""DocEntry"" FROM ""@EXO_ENVTRANS"" WHERE ""DocEntry""=" & sDocEntryEnvTransporte
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    Dim sDirExportar As String = oobjGlobal.path & "\05.Rpt\ENVTRANS\"
                    Dim sRutaFicheros As String = oobjGlobal.path & "\05.Rpt\ENVTRANS\MANIFIESTO\"
                    If IO.Directory.Exists(sDirExportar) = False Then
                        IO.Directory.CreateDirectory(sDirExportar)
                    End If
                    If IO.Directory.Exists(sRutaFicheros) = False Then
                        IO.Directory.CreateDirectory(sRutaFicheros)
                    End If
                    Dim sCrystal As String = "MANIFIESTO.rpt"
                    EXO_GLOBALES.GetCrystalReportFile(oobjGlobal, sDirExportar & sCrystal, sLayout)
                    oobjGlobal.SBOApp.StatusBar.SetText("Layout " & sDirExportar & sCrystal, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)

                    For p = 0 To oRs.RecordCount - 1
                        Dim sDocEntryEnv As String = oRs.Fields.Item("DocEntry").Value.ToString.Trim

                        Dim sTipoImp As String = "IMP"
                        'Imprimimos la etiqueta
                        GenerarImpCrystal(oobjGlobal, sDirExportar, sCrystal, sDocEntryEnv, sTipoImp, sRutaFicheros, "")

                        oRs.MoveNext()
                    Next
                Else
                    oobjGlobal.SBOApp.StatusBar.SetText("No tiene Lista de embalajes. No puede imprimir la etiqueta.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                End If
            End If

            Impresion = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oReportLayoutService = Nothing
            oCmpSrv = Nothing
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Shared Sub GenerarImpCrystal(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal rutaCrystal As String, ByVal sCrystal As String,
                                       ByVal sDocEntry As String, ByVal sTIPOIMP As String, ByVal sDir As String, ByRef sReport As String)

        Dim oCRReport As ReportDocument = Nothing
        Dim oFileDestino As DiskFileDestinationOptions = Nothing
        Dim sServer As String = ""
        Dim sDriver As String = ""
        Dim sBBDD As String = ""
        Dim sUser As String = ""
        Dim sPwd As String = ""
        Dim sConnection As String = ""
        Dim oLogonProps As NameValuePairs2 = Nothing

        Dim conrepor As DataSourceConnections = Nothing
        Dim sImpresora As String = "" : Dim nCopias As Integer = 1
        Dim sSQL As String = ""
        Try

            oCRReport = New ReportDocument()

            oCRReport.Load(rutaCrystal & sCrystal)

            oCRReport.DataSourceConnections.Clear()

            oObjGlobal.SBOApp.StatusBar.SetText("DocEntry: " & sDocEntry, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)

            'Establecemos las conexiones a la BBDD
            sServer = oObjGlobal.funcionesUI.refDi.OGEN.valorVariable("SERVIDOR_HANA") ' objGlobal.compañia.Server
            'sServer = objGlobal.refDi.SQL.dameCadenaConexion.ToString
            sBBDD = oObjGlobal.compañia.CompanyDB
            sUser = oObjGlobal.refDi.SQL.usuarioSQL
            sPwd = oObjGlobal.refDi.SQL.claveSQL

            sDriver = "HDBODBC"
            sConnection = "DRIVER={" & sDriver & "};UID=" & sUser & ";PWD=" & sPwd & ";SERVERNODE=" & sServer & ";DATABASE=" & sBBDD & ";"
            'sConnection = "DRIVER={" & sDriver & "};" & sServer & ";DATABASE=" & sBBDD & ";"
            oObjGlobal.SBOApp.StatusBar.SetText("Conectando: " & sConnection, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
            oLogonProps = oCRReport.DataSourceConnections(0).LogonProperties
            oLogonProps.Set("Provider", sDriver)
            oLogonProps.Set("Connection String", sConnection)


            'Establecemos los parámetros para el report.
            oCRReport.SetParameterValue("Id_envío", sDocEntry)
            oCRReport.SetParameterValue("Schema@", sBBDD)


            oCRReport.DataSourceConnections(0).SetLogonProperties(oLogonProps)
            oCRReport.DataSourceConnections(0).SetConnection(sServer, sBBDD, False)
            oObjGlobal.SBOApp.StatusBar.SetText("Connection String: " & sConnection, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)

            For Each oSubReport As ReportDocument In oCRReport.Subreports
                For Each oConnection As IConnectionInfo In oSubReport.DataSourceConnections
                    oConnection.SetConnection(sServer, sBBDD, False)
                    oConnection.SetLogon(sUser, sPwd)
                Next
            Next
            oObjGlobal.SBOApp.StatusBar.SetText("Actualizado conect Subreport...", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)

            Select Case sTIPOIMP
                Case "PDF"
#Region "Exportar a PDF"
                    'Preparamos para la exportación
                    If IO.Directory.Exists(sDir) = False Then
                        IO.Directory.CreateDirectory(sDir)
                    End If
                    sReport = sDir & "Et_Bultos_" & sDocEntry & ".pdf"
                    'Compruebo si existe y lo borro
                    If IO.File.Exists(sReport) Then
                        IO.File.Delete(sReport)
                    End If
                    oObjGlobal.SBOApp.StatusBar.SetText("Generando pdf para envio impresión...Espere por favor", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)

                    oCRReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat

                    oFileDestino = New CrystalDecisions.Shared.DiskFileDestinationOptions
                    oFileDestino.DiskFileName = sReport

                    'Le pasamos al reporte el parámetro destino del reporte (ruta)
                    oCRReport.ExportOptions.DestinationOptions = oFileDestino

                    'Le indicamos que el reporte no es para mostrarse en pantalla, sino, que es para guardar en disco
                    oCRReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile

                    'Finalmente exportamos el reporte a PDF
                    oCRReport.Export()
                    '            oCRReport.ExportToDisk(ExportFormatType.PortableDocFormat, sReport)
#End Region
                Case "IMP"
#Region "Imprimir a impresora"
                    'Buscamos la impresora por defecto
                    'Dim instance As New Printing.PrinterSettings
                    'sImpresora = instance.PrinterName
                    sImpresora = oObjGlobal.refDi.SQL.sqlStringB1("SELECT ""U_EXO_IMPDOC"" FROM OUSR WHERE ""USERID""='" & oObjGlobal.compañia.UserSignature.ToString & "' ")
                    'oObjGlobal.SBOApp.StatusBar.SetText("Impresora: " & sImpresora, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)
                    oObjGlobal.SBOApp.StatusBar.SetText("Buscando Impresora " & sImpresora & "...Espere por favor", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                    If EXO_GLOBALES.IsPrinterOnline(sImpresora) = True Then
                        oObjGlobal.SBOApp.StatusBar.SetText("Imprimiendo en " & sImpresora & "...Espere por favor", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                        oCRReport.PrintOptions.NoPrinter = False
                        oCRReport.PrintOptions.PrinterName = sImpresora
                        oCRReport.PrintToPrinter(nCopias, False, 0, 9999)
                    Else
                        oObjGlobal.SBOApp.StatusBar.SetText("La impresora seleccionada en el usuario no se encuentra o está offline. Por favor verifique la parametrización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    End If
#End Region
            End Select

            'Cerramos
            oCRReport.Close()
            oCRReport.Dispose()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oObjGlobal.SBOApp.StatusBar.SetText("Fin del proceso de impresión.", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success)
            oCRReport = Nothing
            oFileDestino = Nothing
        End Try
    End Sub
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
        Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim sFecha As String = ""
        Dim sAlmacen As String = ""
#End Region

        EventHandler_COMBO_SELECT_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True And oForm.Mode = BoFormMode.fm_ADD_MODE Then
                If pVal.ItemUID = "4_U_Cb" Then
                    If CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value IsNot Nothing Then
                        Dim sSerie As String = CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                        Dim iNum As Integer
                        iNum = oForm.BusinessObject.GetNextSerialNumber(CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, oForm.BusinessObject.Type.ToString)
                        oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("DocNum", 0, iNum.ToString)
                        objGlobal.SBOApp.StatusBar.SetText("Serie:" & sSerie & " - " & iNum.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                End If
            End If
            If oForm.Visible = True Then
                If pVal.ItemUID = "22_U_Cb" Then ' Almacen
                    sFecha = CType(oForm.Items.Item("21_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                    If sFecha = "" Then
                        sFecha = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                        oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_DOCDATE", 0, sFecha)
                    End If
                    sAlmacen = CType(oForm.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    'Expedición
                    sSQL = "Select ""TrnspCode"",""TrnspName"" FROM OSHP WHERE ""Active""='Y' and ""TrnspCode"" in ("
                    sSQL &= " SELECT distinct  ""TrnspCode"" FROM ("
                    sSQL &= " Select T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM ORDR T0 "
                    sSQL &= " Inner JOIN RDR1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    sSQL &= " UNION ALL "
                    sSQL &= " Select T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM ODLN T0 "
                    sSQL &= " Inner JOIN DLN1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    sSQL &= " UNION ALL "
                    sSQL &= "Select  T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM OPRR T0 "
                    sSQL &= " Inner JOIN PRR1 t1 on  T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    sSQL &= " UNION ALL "
                    sSQL &= " Select T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM  ORPD T0 "
                    sSQL &= " Inner JOIN RPD1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    sSQL &= " UNION ALL "
                    sSQL &= "Select  T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM OWTQ T0 "
                    sSQL &= " Inner JOIN WTQ1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' )"
                    sSQL &= " ) ORDER BY ""TrnspName"""
                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                ElseIf pVal.ItemUID = "20_U_Cb" Then 'Clase de expedición
                    If CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value IsNot Nothing Then
                        Dim sExpedicion As String = CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                        sSQL = "Select IFNULL(""U_EXO_AGE"",'') FROM OSHP WHERE ""TrnspCode""='" & sExpedicion & "' "
                    End If
                    Dim sAgeCod As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_AGTCODE", 0, sAgeCod)
                    sSQL = "SELECT ""CardName"" FROM OCRD WHERE ""CardCode""='" & sAgeCod & "' "
                    Dim sAgeNom As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_AGTNAME", 0, sAgeNom)
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
    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True And oForm.TypeEx = "UDO_FT_EXO_ENVTRANS" Then
                oForm.Freeze(True)
                'No dejamos que modifique la cabecera
                oItem = oForm.Items.Item("22_U_Cb")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem = oForm.Items.Item("4_U_Cb")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem = oForm.Items.Item("20_U_Cb")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem = oForm.Items.Item("21_U_E")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem = oForm.Items.Item("23_U_E")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                Cargar_Combos(oForm)
                Dim sAgencia As String = CType(oForm.Items.Item("23_U_E").Specific, SAPbouiCOM.EditText).Value
                Cargar_Combo_Matricula_Conductor_Plataforma(oForm, sAgencia)

                If objGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
                    objGlobal.SBOApp.ActivateMenuItem("1304")
                End If
#Region "Botón Manifiesto Transporte"
                oItem = oForm.Items.Add("btnMan", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                oItem.Left = oForm.Items.Item("2").Left + 80
                oItem.Width = oForm.Items.Item("2").Width * 2
                oItem.Top = oForm.Items.Item("2").Top
                oItem.Height = oForm.Items.Item("2").Height
                oItem.Enabled = False
                Dim oBtnAct As SAPbouiCOM.Button
                oBtnAct = CType(oItem.Specific, Button)
                oBtnAct.Caption = "Imp. Manifiesto Transporte"
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
#End Region

#Region "Botón Exportar a fichero"
                oItem = oForm.Items.Add("btnEFich", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                oItem.Left = oForm.Items.Item("btnMan").Left + oForm.Items.Item("btnMan").Width + 5
                oItem.Width = oForm.Items.Item("btnMan").Width
                oItem.Top = oForm.Items.Item("btnMan").Top
                oItem.Height = oForm.Items.Item("btnMan").Height
                oItem.Enabled = False
                oBtnAct = CType(oItem.Specific, Button)
                oBtnAct.Caption = "Exportar Fichero"
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
#End Region
            End If

            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oItem, Object))
        End Try
    End Function
    Private Sub Cargar_Combos(ByRef oform As SAPbouiCOM.Form)
#Region "Variables"
        Dim sClaseExp As String = ""
        Dim sSucursal As String = ""
        Dim sSerieDef As String = ""
        Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim sFecha As String = ""
        Dim sSQL As String = ""
        Dim sAlmacendef As String = ""
        Dim sExpedicion As String = ""
#End Region
        Try
            sFecha = CType(oform.Items.Item("21_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
            If sFecha = "" Then
                sFecha = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_DOCDATE", 0, sFecha)
            End If

            'Almacen
            sSQL = "SELECT ""Branch"" FROM OUSR WHERE ""USERID""=" & objGlobal.compañia.UserSignature.ToString
            sSucursal = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            sSQL = " SELECT ""WhsCode"",""WhsName"" FROM OWHS"
            sSQL &= " WHERE ""Inactive""='N' and ""U_EXO_SUCURSAL""='" & sSucursal & "' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            If oform.Mode = BoFormMode.fm_ADD_MODE Then
                'Poner almacen por defecto
                Try
                    sSQL = " SELECT TOP 1 ""WhsCode"" FROM OWHS"
                    sSQL &= " WHERE ""Inactive""='N' and ""U_EXO_SUCURSAL""='" & sSucursal & "' "
                    sAlmacendef = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    CType(oform.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(sAlmacendef, BoSearchKey.psk_ByValue)
                Catch ex As Exception

                End Try
            Else
                sAlmacendef = oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").GetValue("U_EXO_ALMACEN", 0)
            End If
            oform.Items.Item("22_U_Cb").DisplayDesc = True

            'Expedición
            sSQL = "SELECT ""TrnspCode"",""TrnspName"" FROM OSHP WHERE ""Active""='Y' "
            sSQL &= " ORDER BY ""TrnspName"""
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)


            'Series 
            sSQL = "SELECT ""Series"",""SeriesName"" FROM NNM1 WHERE ""ObjectCode""='EXO_ENVTRANS' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oform.Items.Item("4_U_Cb").DisplayDesc = True



            If oform.Mode = BoFormMode.fm_ADD_MODE Then
                'Expedición por defecto
                If CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues.Count > 0 Then
                    CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)
                    sExpedicion = CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                Else
                    sExpedicion = ""
                End If

                'Poner serie por defecto y el num. de documento
                sSQL = " SELECT ""DfltSeries"" FROM ONNM WHERE ""ObjectCode""='EXO_ENVTRANS' "
                sSerieDef = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                CType(oform.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(sSerieDef, BoSearchKey.psk_ByValue)

                EXO_GLOBALES.Poner_DocNum(oform, sSerieDef, objGlobal)

                'Como en la expedición tenemos la agencia, pues tenemos que rellenarlo automático
                sSQL = "SELECT IFNULL(""U_EXO_AGE"",'') FROM OSHP WHERE ""TrnspCode""='" & sExpedicion & "' "
                Dim sAgeCod As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_AGTCODE", 0, sAgeCod)
                sSQL = "SELECT ""CardName"" FROM OCRD WHERE ""CardCode""='" & sAgeCod & "' "
                Dim sAgeNom As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_AGTNAME", 0, sAgeNom)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Cargar_Combo_Matricula_Conductor_Plataforma(ByRef oform As SAPbouiCOM.Form, ByVal sAgencia As String)
#Region "Variables"
        Dim sSQL As String = ""

#End Region
        Try

            'Matricula
            sSQL = "SELECT ""U_EXO_VEHICULO"",""U_EXO_DES"" FROM ""@EXO_VEHIAGL"" WHERE ""Code""='" & sAgencia & "' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("25_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oform.Items.Item("25_U_Cb").DisplayDesc = True

            'Conductor 
            sSQL = "SELECT ""U_EXO_COND"",(""U_EXO_NOMBRE"" || ' ' || ""U_EXO_APE"") ""Nombre"" FROM ""@EXO_CONAGL"" WHERE ""Code""='" & sAgencia & "' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("26_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oform.Items.Item("26_U_Cb").DisplayDesc = True

            ''Plataforma
            'sSQL = "SELECT ""U_EXO_PLATA"",""U_EXO_PLATAD"" FROM ""@EXO_PLATAAGL"" WHERE ""Code""='" & sAgencia & "' "
            'objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_5").ValidValues, sSQL)
            'CType(oform.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_5").DisplayDesc = True

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function EventHandler_Choose_FromList_Before(ByVal pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing

        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""

        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        EventHandler_Choose_FromList_Before = False

        Try
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            If pVal.ItemUID = "23_U_E" Then 'Agencia de transporte
                oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
                oCFLEvento = CType(pVal, IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions
                oCond = oConds.Add
                oCond.Alias = "QryGroup1"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"
                'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            ElseIf pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_1_5" Then
                oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
                Dim sAgencia As String = CType(oForm.Items.Item("23_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                sSQL = "SELECT ""U_EXO_PLATA"" FROM ""@EXO_PLATAAGL"" WHERE ""Code""='" & sAgencia & "'"
                oRs.DoQuery(sSQL)
                oConds = New SAPbouiCOM.Conditions
                oCond = oConds.Add
                oCond.Alias = "Code"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "0"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                For i = 0 To oRs.RecordCount - 1
                    oCond = oConds.Add
                    oCond.Alias = "Code"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = oRs.Fields.Item("U_EXO_PLATA").Value.ToString
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oRs.MoveNext()
                Next
                If oConds.Count > 0 Then oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                oForm.ChooseFromLists.Item("CFLPT").SetConditions(oConds)
            End If

            EventHandler_Choose_FromList_Before = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByVal pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oDataTable As DataTable = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing

        Dim sNombre As String = ""
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
                oDataTable = oCFLEvento.SelectedObjects
                Select Case oCFLEvento.ChooseFromListUID
                    Case "CFLAT"
                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "23_U_E" Then
                                CType(oForm.Items.Item("24_U_E").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardName", 0).ToString
                                Cargar_Combo_Matricula_Conductor_Plataforma(oForm, oDataTable.GetValue("CardCode", 0).ToString)
                            End If
                        End If
                    Case "CFLAB"
                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_1_1" Then
                                sNombre = oDataTable.GetValue("Name", 0).ToString
                                oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANSAB").SetValue("U_EXO_AGNAME", pVal.Row - 1, sNombre)
                            End If
                        End If
                    Case "CFLPT"
                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_1_5" Then
                                sNombre = oDataTable.GetValue("Name", 0).ToString
                                oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANSAB").SetValue("U_EXO_PLATAD", pVal.Row - 1, sNombre)
                            End If
                        End If
                End Select
            End If

            EventHandler_Choose_FromList_After = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.FormDatatable(oDataTable)

            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument
        Dim sFecha As String = "" : Dim sAlmacen As String = "" : Dim sSucursal As String = ""
        Dim sSQL As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_ENVTRANS"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Mode = BoFormMode.fm_OK_MODE Then
                                    'Almacen
                                    'sSQL = "SELECT ""Branch"" FROM OUSR WHERE ""USERID""=" & objGlobal.compañia.UserSignature.ToString
                                    'sSucursal = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                    sSQL = " SELECT ""WhsCode"",""WhsName"" FROM OWHS"
                                    sSQL &= " WHERE ""Inactive""='N' "
                                    'ssql &= " And ""U_EXO_SUCURSAL""='" & sSucursal & "' "
                                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                                    ''Expedición
                                    'sSQL = "Select ""TrnspCode"",""TrnspName"" FROM OSHP "
                                    'objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

                                    Dim sAgencia As String = CType(oForm.Items.Item("23_U_E").Specific, SAPbouiCOM.EditText).Value
                                    Cargar_Combo_Matricula_Conductor_Plataforma(oForm, sAgencia)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If Comprobar_existe(oForm) = False Then
                                    Return False
                                Else
                                    Return True
                                End If
                        End Select

                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_ENVTRANS"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                End Select
            End If

            Return True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function Comprobar_existe(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Comprobar_existe = False
        Dim sClaseExp As String = "" : Dim sAlmacen As String = ""
        Dim sFecha As String = "" : Dim sDocNum As String = "" : Dim sSerie As String = ""
        Dim sSQL As String = ""
        Try
            sSerie = CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            sFecha = CType(oForm.Items.Item("21_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
            sClaseExp = CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            sAlmacen = CType(oForm.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            sSQL = " SELECT ""DocNum"" FROM ""@EXO_ENVTRANS"" Where ""U_EXO_CEXP""='" & sClaseExp & "' and ""U_EXO_DOCDATE""='" & sFecha & "' and ""U_EXO_ALMACEN""='" & sAlmacen & "' and ""Series""=" & sSerie
            sDocNum = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            If sDocNum = "" Then
                Return True
            Else
                objGlobal.SBOApp.StatusBar.SetText("Ya existe el documento Nº " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If
            Comprobar_existe = True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
