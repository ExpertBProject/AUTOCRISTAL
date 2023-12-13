Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class EXO_LSTEMB
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
                            If oForm.TypeEx = "UDO_FT_EXO_LSTEMB" Then
                                If Gestion_Cerrar() = False Then
                                    Return False
                                End If
                            End If
                        End If
                End Select
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnLEmb"
                        If CargarUDO() = False Then
                            Exit Function
                        End If
                    Case "1282"
                        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.ActiveForm
                        If oForm IsNot Nothing Then
                            If oForm.TypeEx = "UDO_FT_EXO_LSTEMB" Then
                                EXO_GLOBALES.Modo_Anadir(oForm, objGlobal)
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
    Public Function Gestion_Cerrar() As Boolean
        Gestion_Cerrar = False

        Try
            objGlobal.SBOApp.StatusBar.SetText("No se puede cerrar desde este menú. Se debe cerrar desde la ventana ""Envío - Transporte"" ", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
            objGlobal.SBOApp.MessageBox("No se puede cerrar desde este menú. Se debe cerrar desde la ventana ""Envío - Transporte"" ")

            Return False
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally

        End Try
    End Function
    Public Function CargarUDO() As Boolean
        CargarUDO = False

        Try
            objGlobal.funcionesUI.cargaFormUdoBD("EXO_LSTEMB")

            CargarUDO = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally

        End Try
    End Function
    Public Function SBOApp_ProgressBarEvent(infoEvento As ProgressBarEvent) As Boolean
        Try
            If infoEvento.EventType = SAPbouiCOM.BoProgressBarEventTypes.pbet_ProgressBarStopped And infoEvento.BeforeAction Then
                'Fail to handle document numbering:
            End If
            Return True
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
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
                        Case "UDO_FT_EXO_LSTEMB"
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
                        Case "UDO_FT_EXO_LSTEMB"
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
                        Case "UDO_FT_EXO_LSTEMB"
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

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_LSTEMB"
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
    Private Function EventHandler_ItemPressed_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnET" 'Impresión de etiquetas
                    'Calculando datos
                    objGlobal.SBOApp.StatusBar.SetText("Imprimiendo Etiquetas... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    If Impresion_ET(oForm, objGlobal) = False Then
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
    Public Shared Function Impresion_ET(ByRef oForm As SAPbouiCOM.Form, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Impresion_ET = False
#Region "VARIABLES"
        Dim oCmpSrv As SAPbobsCOM.CompanyService = oobjGlobal.compañia.GetCompanyService()
        Dim oReportLayoutService As SAPbobsCOM.ReportLayoutsService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), SAPbobsCOM.ReportLayoutsService)
        Dim oPrintParam As SAPbobsCOM.ReportLayoutPrintParams = CType(oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams), SAPbobsCOM.ReportLayoutPrintParams)
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim sLayout As String = "" : Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sDocEntryLstEmbalaje As String = ""
#End Region

        Try
            oRs = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sDocEntryLstEmbalaje = oForm.DataSources.DBDataSources.Item("@EXO_LSTEMB").GetValue("DocEntry", 0).ToString.Trim

            sLayout = oobjGlobal.funcionesUI.refDi.OGEN.valorVariable("EXO_ETPARRILLA")
            If sLayout = "" Then
                oobjGlobal.SBOApp.StatusBar.SetText("Parámetro [EXO_ETPARRILLA] no tiene valor. Revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                sSQL = "SELECT DISTINCT ""DocEntry"",""U_EXO_IDBULTO"" FROM ""@EXO_LSTEMBL"" WHERE ""DocEntry""=" & sDocEntryLstEmbalaje
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    Dim sDirExportar As String = oobjGlobal.path & "\05.Rpt\PARRILLADOC\"
                    Dim sRutaFicheros As String = oobjGlobal.path & "\05.Rpt\PARRILLADOC\ET_CREADAS\"
                    If IO.Directory.Exists(sDirExportar) = False Then
                        IO.Directory.CreateDirectory(sDirExportar)
                    End If
                    If IO.Directory.Exists(sRutaFicheros) = False Then
                        IO.Directory.CreateDirectory(sRutaFicheros)
                    End If
                    Dim sCrystal As String = "ETIQUETASBULTOS.rpt"
                    EXO_GLOBALES.GetCrystalReportFile(oobjGlobal, sDirExportar & sCrystal, sLayout)
                    oobjGlobal.SBOApp.StatusBar.SetText("Layout " & sDirExportar & sCrystal, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)

                    For p = 0 To oRs.RecordCount - 1
                        Dim sDocEntryLstEmb As String = oRs.Fields.Item("DocEntry").Value.ToString.Trim
                        Dim sIDBulto As String = oRs.Fields.Item("U_EXO_IDBULTO").Value.ToString.Trim

                        Dim sTipoImp As String = "IMP"
                        'Imprimimos la etiqueta
                        GenerarImpCrystalET(oobjGlobal, sDirExportar, sCrystal, sDocEntryLstEmb, sIDBulto, sTipoImp, sRutaFicheros, "")

                        oRs.MoveNext()
                    Next
                Else
                    oobjGlobal.SBOApp.StatusBar.SetText("No tiene Lista de embalajes. No puede imprimir la etiqueta.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                End If
            End If

            Impresion_ET = True
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
    Public Shared Sub GenerarImpCrystalET(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal rutaCrystal As String, ByVal sCrystal As String,
                                       ByVal sDocEntry As String, ByVal sIDBULTO As String, ByVal sTIPOIMP As String, ByVal sDir As String, ByRef sReport As String)

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

            oObjGlobal.SBOApp.StatusBar.SetText("DocEntry: " & sDocEntry & " - IDBULTO: " & sIDBULTO, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)

            'Establecemos las conexiones a la BBDD
            sServer = oObjGlobal.funcionesUI.refDi.OGEN.valorVariable("SERVIDOR_HANA") ' objGlobal.compañia.Server
            'sServer = objGlobal.refDi.SQL.dameCadenaConexion.ToString
            sBBDD = oObjGlobal.compañia.CompanyDB
            sUser = oObjGlobal.refDi.SQL.usuarioSQL
            sPwd = oObjGlobal.refDi.SQL.claveSQL

            'sDriver = "HDBODBC"
            'sConnection = "DRIVER={" & sDriver & "};UID=" & sUser & ";PWD=" & sPwd & ";SERVERNODE=" & sServer & ";DATABASE=" & sBBDD & ";"
            sDriver = "B1CRHPROXY"
            sConnection = "DRIVER={B1CRHPROXY};UID=B1SLDUSER;PWD=Aut0Cr1Sb01;SERVERNODE=10.10.1.13:30015;DATABASE=SBO_AUTOCRISTAL_PRUEBAS;"
            oObjGlobal.SBOApp.StatusBar.SetText("Conectando: " & sConnection, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
            oLogonProps = oCRReport.DataSourceConnections(0).LogonProperties
            oLogonProps.Set("Provider", sDriver)
            oLogonProps.Set("Connection String", sConnection)


            'Establecemos los parámetros para el report.
            oCRReport.SetParameterValue("DocEntry", sDocEntry)
            oCRReport.SetParameterValue("ID_Bulto", sIDBULTO)
            'oCRReport.SetParameterValue("Schema@", sBBDD)


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
                    sImpresora = oObjGlobal.refDi.SQL.sqlStringB1("SELECT ""U_EXO_IMPBUL"" FROM OUSR WHERE ""USERID""='" & oObjGlobal.compañia.UserSignature.ToString & "' ")
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
                        objGlobal.SBOApp.StatusBar.SetText("Grabe primero para poder ver el resumen.", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox("Grabe primero para poder ver el resumen.")
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
            sDocEntry = oform.DataSources.DBDataSources.Item("@EXO_LSTEMB").GetValue("DocEntry", 0)
            If sDocEntry = "" Then
                sDocEntry = CType(oform.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value.ToString()
            End If
            objGlobal.SBOApp.StatusBar.SetText("Documento Nº " & sDocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = "SELECT DISTINCT T0.""U_EXO_IDBULTO"" ""ID BULTO"", T0.""U_EXO_TBULTO"" ""BULTO"", IFNULL( T0.""Volumen"",0.00) ""VOLUMEN"", ifnull(T0.""Peso"",0.00) ""PESO"" "
            sSQL &= "FROM ""EXO_DetalleBultosExpediciones""  T0 "
            sSQL &= " WHERE T0.""IdExpedición"" =" & sDocEntry
            oform.DataSources.DataTables.Item("DTRES").ExecuteQuery(sSQL)
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

            For i = 0 To 3
                Select Case i
                    Case 2, 3
                        CType(oform.Items.Item("grdRES").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdRES").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                        oColumnTxt.RightJustified = True
                    Case Else
                        CType(oform.Items.Item("grdRES").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdRES").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next
            CType(oform.Items.Item("grdRES").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Function EventHandler_MATRIX_LINK_PRESSED(ByVal pVal As ItemEvent) As Boolean

        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oLinkedButton As SAPbouiCOM.LinkedButton = Nothing
        Dim oColumn As SAPbouiCOM.Column = Nothing
        Dim sTipo As String = ""
        EventHandler_MATRIX_LINK_PRESSED = False

        Try
            oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                Return True
            End If


            oMatrix = CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
            oColumn = oMatrix.Columns.Item("C_0_5")
            oLinkedButton = CType(oColumn.ExtendedObject, SAPbouiCOM.LinkedButton)

            sTipo = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            Select Case sTipo
                Case "PEDVTA" 'Pedidos de ventas
                    oLinkedButton.LinkedObject = BoLinkedObject.lf_Order
                Case "SDPROV" ' Sol de devolución de proveedor
                    oLinkedButton.LinkedObject = CType("234000032", BoLinkedObject)
                Case "SOLTRA" ' Sol de traslado
                    oLinkedButton.LinkedObject = BoLinkedObject.lf_StockTransfersRequest
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
        Dim sCardCode As String = ""
        Dim sBulto As String = "" : Dim sTBulto As String = ""
        Dim sArticulo As String = ""
        Dim sSQL As String = ""
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
                    Case "CFL_IC"
                        oDataTable = oCFLEvento.SelectedObjects

                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "22_U_E" Then
                                Try
                                    sCardCode = oDataTable.GetValue("CardCode", 0).ToString
                                    CType(oForm.Items.Item("23_U_E").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardName", 0).ToString
                                    'Cargamos Combo de dirección
                                    sSQL = "Select ""Address"" FROM CRD1 WHERE ""CardCode""='" & sCardCode & "' and ""AdresType""='S' "
                                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("24_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                                Catch ex As Exception

                                End Try
                            End If
                        End If
                    Case "CFLOPKG"
                        sTBulto = oDataTable.GetValue("PkgType", 0).ToString
                        sBulto = oDataTable.GetValue("U_EXO_DESBUL", 0).ToString
                        oForm.DataSources.DBDataSources.Item("@EXO_LSTEMBL").SetValue("U_EXO_TBULTO", pVal.Row - 1, sTBulto)
                        oForm.DataSources.DBDataSources.Item("@EXO_LSTEMBL").SetValue("U_EXO_BULTO", pVal.Row - 1, sBulto)
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sBulto
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sTBulto
                    Case "CFLART"
                        sArticulo = oDataTable.GetValue("ItemName", 0).ToString
                        oForm.DataSources.DBDataSources.Item("@EXO_LSTEMBL").SetValue("U_EXO_ITEMNAME", pVal.Row - 1, sArticulo)
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_9").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sArticulo
                        sArticulo = oDataTable.GetValue("ItemCode", 0).ToString
                        oForm.DataSources.DBDataSources.Item("@EXO_LSTEMBL").SetValue("U_EXO_ITEMCODE", pVal.Row - 1, sArticulo)
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_8").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sArticulo
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
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sFecha As String = "" : Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day) : Dim sClaseExp As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
        Dim sAlmacen As String = ""
        EventHandler_COMBO_SELECT_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True And oForm.Mode = BoFormMode.fm_ADD_MODE Then
                If pVal.ItemUID = "4_U_Cb" Then
                    If CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value IsNot Nothing Then
                        Dim sSerie As String = CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                        EXO_GLOBALES.Poner_DocNum(oForm, sSerie, objGlobal)
                    End If
                ElseIf pVal.ItemUID = "20_U_Cb" Then
                    'Actualizamos el combo de cargar	Id – Envío – Tte
                    sFecha = CType(oForm.Items.Item("26_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                    If sFecha = "" Then
                        sFecha = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                        oForm.DataSources.DBDataSources.Item("@EXO_LSTEMB").SetValue("U_EXO_DOCDATE", 0, sFecha)
                    End If
                    If CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sClaseExp = CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    Else
                        sClaseExp = ""
                    End If
                    sSQL = " SELECT '-'  ""DocEntry"", ' ' ""DocNum"" FROM DUMMY"
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT CAST(""DocEntry"" as nVARCHAR),CAST(""DocNum"" as nVARCHAR) FROM ""@EXO_ENVTRANS"" WHERE ""Status""='O' "
                    sSQL &= " and ""U_EXO_DOCDATE""='" & sFecha & "' and ""U_EXO_CEXP""='" & sClaseExp & "' "
                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("1_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                    oForm.Items.Item("1_U_Cb").DisplayDesc = True
                ElseIf pVal.ItemUID = "22_U_Cb" Then
                    'sFecha = CType(oForm.Items.Item("26_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                    'If sFecha = "" Then
                    '    sFecha = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                    '    oForm.DataSources.DBDataSources.Item("@EXO_LSTEMB").SetValue("U_EXO_DOCDATE", 0, sFecha)
                    'End If
                    'sAlmacen = CType(oForm.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    ''Expedición
                    'sSQL = "SELECT ""TrnspCode"",""TrnspName"" FROM OSHP WHERE ""Active""='Y' and ""TrnspCode"" in ("
                    'sSQL &= " SELECT distinct  ""TrnspCode"" FROM ("
                    'sSQL &= " Select T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM ORDR T0 "
                    'sSQL &= " Inner JOIN RDR1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    'sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    'sSQL &= " UNION ALL "
                    'sSQL &= " Select T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM ODLN T0 "
                    'sSQL &= " Inner JOIN DLN1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    'sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    'sSQL &= " UNION ALL "
                    'sSQL &= "Select  T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM OPRR T0 "
                    'sSQL &= " Inner JOIN PRR1 t1 on  T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    'sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    'sSQL &= " UNION ALL "
                    'sSQL &= " Select T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM  ORPD T0 "
                    'sSQL &= " Inner JOIN RPD1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    'sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' "
                    'sSQL &= " UNION ALL "
                    'sSQL &= "Select  T0.""DocNum"", T0.""DocDueDate"", T0.""TrnspCode"", T0.""DocStatus"" FROM OWTQ T0 "
                    'sSQL &= " Inner JOIN WTQ1 t1 on T1.""DocEntry"" = T0.""DocEntry"" and T1.""WhsCode"" = '" & sAlmacen & "' "
                    'sSQL &= " Where T0.""DocDueDate"" = '" & sFecha & "' )"
                    'sSQL &= " ) ORDER BY ""TrnspName"""
                    'objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
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
            If oForm.Visible = True And oForm.TypeEx = "UDO_FT_EXO_LSTEMB" Then
                oForm.Freeze(True)
#Region "Botón Imprimir Etiqueta"
                oItem = oForm.Items.Add("btnET", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                oItem.Left = oForm.Items.Item("2").Left + 80
                oItem.Width = oForm.Items.Item("2").Width + 25
                oItem.Top = oForm.Items.Item("2").Top
                oItem.Height = oForm.Items.Item("2").Height
                oItem.Enabled = False
                Dim oBtnAct As SAPbouiCOM.Button
                oBtnAct = CType(oItem.Specific, Button)
                oBtnAct.Caption = "Imprimir Bulto"
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
#End Region
                Cargar_Combos(oForm)
                If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                    EXO_GLOBALES.Modo_Anadir(oForm, objGlobal)
                End If

                If objGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
                    If oForm.Mode <> BoFormMode.fm_OK_MODE Then
                        oForm.Mode = BoFormMode.fm_OK_MODE
                    End If
                    objGlobal.SBOApp.ActivateMenuItem("1304")
                End If
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
        Dim sSQL As String = ""
        Dim sFecha As String = "" : Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim sClaseExp As String = "" : Dim sSucursal As String = ""
        Dim sAlmacendef As String = ""
#End Region
        Try
            sFecha = CType(oform.Items.Item("26_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
            If sFecha = "" Then
                sFecha = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                oform.DataSources.DBDataSources.Item("@EXO_LSTEMB").SetValue("U_EXO_DOCDATE", 0, sFecha)
            End If

            sSQL = "SELECT ""Branch"" FROM OUSR WHERE ""USERID""=" & objGlobal.compañia.UserSignature.ToString
            sSucursal = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            sSQL = " SELECT ""WhsCode"",""WhsName"" FROM OWHS"
            sSQL &= " WHERE ""Inactive""='N' and ""U_EXO_SUCURSAL""='" & sSucursal & "' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oform.Items.Item("22_U_Cb").DisplayDesc = True

            'Series 
            sSQL = "SELECT ""Series"",""SeriesName"" FROM NNM1 WHERE ""ObjectCode""='EXO_LSTEMB' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oform.Items.Item("4_U_Cb").DisplayDesc = True

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
                sAlmacendef = oform.DataSources.DBDataSources.Item("@EXO_LSTEMB").GetValue("U_EXO_ALMACEN", 0)
            End If

            'Expedición
            sSQL = "SELECT ""TrnspCode"",""TrnspName"" FROM OSHP WHERE ""Active""='Y' "
            sSQL &= " ORDER BY ""TrnspName"""
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            If CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues.Count > 0 Then
                CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)
                sClaseExp = CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            Else
                sClaseExp = ""
            End If


            'cargar	Id – Envío – Tte
            sSQL = " Select '-'  ""DocEntry"", ' ' ""DocNum"" FROM DUMMY"
            sSQL &= " UNION ALL "
            sSQL &= "SELECT CAST(""DocEntry"" as nVARCHAR),CAST(""DocNum"" as nVARCHAR) FROM ""@EXO_ENVTRANS"" WHERE ""Status""='O' "
            'sSQL &= " and ""U_EXO_DOCDATE""='" & sFecha & "' and ""U_EXO_CEXP""='" & sClaseExp & "' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("1_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oform.Items.Item("1_U_Cb").DisplayDesc = True


        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument

        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_LSTEMB"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Mode = BoFormMode.fm_OK_MODE Then
                                    Carga_combos_DATA(oForm, objGlobal)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        End Select

                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_LSTEMB"
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
    Private Sub Carga_combos_DATA(ByRef oForm As SAPbouiCOM.Form, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI)
        Dim sSQL As String = ""
        Dim sFecha As String = ""
        Dim sClaseExp As String = ""
        Dim sCardCode As String = ""
        Dim sAlmacen As String = ""
        Try

            'cargar	Id – Envío – Tte
            sSQL = " Select '-'  ""DocEntry"", ' ' ""DocNum"" FROM DUMMY"
            sSQL &= " UNION ALL "
            sSQL &= "SELECT CAST(""DocEntry"" as nVARCHAR),CAST(""DocNum"" as nVARCHAR) FROM ""@EXO_ENVTRANS"" WHERE ""Status""='O' "
            'sSQL &= " and ""U_EXO_DOCDATE""='" & sFecha & "' and ""U_EXO_CEXP""='" & sClaseExp & "' "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("1_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oForm.Items.Item("1_U_Cb").DisplayDesc = True

            'Cargamos Combo de dirección
            sCardCode = CType(oForm.Items.Item("22_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
            sSQL = "SELECT ""Address"" FROM CRD1 WHERE ""CardCode""='" & sCardCode & "' and ""AdresType""='S' "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("24_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

            'If objGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
            '    objGlobal.SBOApp.ActivateMenuItem("1304")
            'End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
