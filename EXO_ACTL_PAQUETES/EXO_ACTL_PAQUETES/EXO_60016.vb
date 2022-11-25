Imports SAPbouiCOM
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class EXO_60016
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
                        Case "60016"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

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
                        Case "60016"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "60016"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                                    If EventHandler_GOT_FOCUS_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "60016"
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
    Private Function EventHandler_Form_Load(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item

        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Visible = False

            'Buscar XML de update
            objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

#Region "Botones"
            oItem = oForm.Items.Add("btnPAQ", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 5
            oItem.Width = oForm.Items.Item("2").Width
            oItem.Top = oForm.Items.Item("2").Top
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Enabled = False
            Dim oBtnAct As SAPbouiCOM.Button
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Paquetes"
            oItem.TextStyle = 1
            oItem.LinkTo = "2"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            oItem = oForm.Items.Add("btnImpPAQ", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("btnPAQ").Left + oForm.Items.Item("btnPAQ").Width + 5
            oItem.Width = oForm.Items.Item("2").Width
            oItem.Top = oForm.Items.Item("2").Top
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Enabled = False
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Imprimir"
            oItem.TextStyle = 1
            oItem.LinkTo = "2"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oItem.Visible = False
#End Region

            oForm.Visible = True

            EventHandler_Form_Load = True

        Catch ex As Exception
            oForm.Visible = True
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_GOT_FOCUS_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_GOT_FOCUS_After = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "3" Then
                INICIO._sLineaSel = pVal.Row.ToString
            End If

            EventHandler_GOT_FOCUS_After = True

        Catch ex As Exception
            oForm.Visible = True
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnPAQ"
                    If pVal.ActionSuccess = True Then
                        If CargarUDOPAQ(oForm) = False Then
                            Exit Function
                        End If
                    End If
                Case "btnImpPAQ"
                    If pVal.ActionSuccess = True Then
                        If ImprimirET(oForm) = False Then
                            Exit Function
                        End If
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Function ImprimirET(ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Variables"

        Dim sOPKG As String = ""

        Dim iLineaSel As Integer = 0 : Dim bLineaSel As Boolean = False
        Dim sExiste As String = "" : Dim sMensaje As String = ""
        Dim sSQL As String = ""
        Dim sMatrix As String = ""

        Dim rutaCrystal As String = "" : Dim sRutaFicheros As String = "" : Dim sReport As String = "" : Dim sTipoImp As String = ""
        Dim sCrystal As String = "Et_Paquete.rpt"
#End Region

        ImprimirET = False

        Try
            sMatrix = "3"

            If INICIO._sLineaSel.Trim <> "" Then
                bLineaSel = True
                iLineaSel = CType(INICIO._sLineaSel, Integer)
            End If
            If oForm.Mode <> BoFormMode.fm_OK_MODE Then
                sMensaje = "Por favor, guarde primero los datos."
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox(sMensaje)
                Return True
            End If
            If bLineaSel = True Then

                sOPKG = oForm.DataSources.DBDataSources.Item("OPKG").GetValue("PkgType", iLineaSel - 1).ToUpper

                If sOPKG <> "" Then
                    rutaCrystal = objGlobal.path & "\05.Rpt\ETIQUETAS\"
                    sTipoImp = "IMP"
                    'Imprimimos la etiqueta
                    GenerarImpCrystal(rutaCrystal, sCrystal, sOPKG, sRutaFicheros, sReport, sTipoImp)
                Else
                    sMensaje = "Por favor, seleccione una línea válida."
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objGlobal.SBOApp.MessageBox(sMensaje)
                End If
            Else
                sMensaje = "Tiene que seleccionar una línea."
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox(sMensaje)
            End If



            ImprimirET = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally

        End Try
    End Function
    Public Sub GenerarImpCrystal(ByVal rutaCrystal As String, ByVal sCrystal As String, ByVal sCode As String, ByVal sFileName As String, ByRef sReport As String, ByVal sTipoImp As String)

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

            oCRReport.Load(rutaCrystal & "\" & sCrystal)

            oCRReport.DataSourceConnections.Clear()

            'Establecemos las conexiones a la BBDD
            sServer = "10.10.1.13:30015" ' objGlobal.compañia.Server
            'sServer = objGlobal.refDi.SQL.dameCadenaConexion.ToString
            sBBDD = objGlobal.compañia.CompanyDB
            sUser = objGlobal.refDi.SQL.usuarioSQL
            sPwd = objGlobal.refDi.SQL.claveSQL

            sDriver = "HDBODBC"
            sConnection = "DRIVER={" & sDriver & "};UID=" & sUser & ";PWD=" & sPwd & ";SERVERNODE=" & sServer & ";DATABASE=" & sBBDD & ";"
            'sConnection = "DRIVER={" & sDriver & "};" & sServer & ";DATABASE=" & sBBDD & ";"
            objGlobal.SBOApp.StatusBar.SetText("Conectando: " & sConnection, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
            oLogonProps = oCRReport.DataSourceConnections(0).LogonProperties
            oLogonProps.Set("Provider", sDriver)
            oLogonProps.Set("Connection String", sConnection)

            oCRReport.DataSourceConnections(0).SetLogonProperties(oLogonProps)
            oCRReport.DataSourceConnections(0).SetConnection(sServer, sBBDD, False)

            For Each oSubReport As ReportDocument In oCRReport.Subreports
                For Each oConnection As IConnectionInfo In oSubReport.DataSourceConnections
                    oConnection.SetConnection(sServer, sBBDD, False)
                    oConnection.SetLogon(sUser, sPwd)
                Next
            Next
            'Establecemos los parámetros para el report.
            oCRReport.SetParameterValue("Code", sCode)
            Select Case sTipoImp
                Case "PDF"
#Region "Exportar a PDF"
                    'Preparamos para la exportación
                    sReport = sFileName & "PEDIDO_" & sCode & ".pdf"
                    'Compruebo si existe y lo borro
                    If IO.File.Exists(sReport) Then
                        IO.File.Delete(sReport)
                    End If
                    objGlobal.SBOApp.StatusBar.SetText("Generando pdf para envio impresión...Espere por favor", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)

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
                    'Buscamos la impresora
                    sSQL = "SELECT ""Fax"" FROM OUSR WHERE ""USERID""='" & objGlobal.compañia.UserSignature.ToString & "' "
                    sImpresora = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    If EXO_GLOBALES.IsPrinterOnline(sImpresora) = True Then
                        objGlobal.SBOApp.StatusBar.SetText("Imprimiendo en " & sImpresora & "...Espere por favor", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                        oCRReport.PrintOptions.PrinterName = sImpresora
                        oCRReport.PrintToPrinter(nCopias, False, 0, 0)
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("La impresora seleccionada en el usuario no se encuentra o está offline. Por favor verifique la parametrización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
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
            oCRReport = Nothing
            oFileDestino = Nothing
        End Try
    End Sub
    Public Function CargarUDOPAQ(ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Variables"
        Dim sTipo As String = ""
        Dim sOPKG As String = ""
        Dim sDes As String = ""
        Dim iLineaSel As Integer = 0 : Dim bLineaSel As Boolean = False
        Dim sExiste As String = "" : Dim sMensaje As String = ""
        Dim sSQL As String = ""
        Dim sMatrix As String = ""
#End Region

        CargarUDOPAQ = False

        Try
            sMatrix = "3"

            If INICIO._sLineaSel.Trim <> "" Then
                bLineaSel = True
                iLineaSel = CType(INICIO._sLineaSel, Integer)
            End If
            If oForm.Mode <> BoFormMode.fm_OK_MODE Then
                sMensaje = "Por favor, guarde primero los datos."
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox(sMensaje)
                Return True
            End If
            If bLineaSel = True Then

                sOPKG = oForm.DataSources.DBDataSources.Item("OPKG").GetValue("PkgCode", iLineaSel - 1).ToUpper
                sTipo = oForm.DataSources.DBDataSources.Item("OPKG").GetValue("PkgType", iLineaSel - 1).ToUpper
                sDes = oForm.DataSources.DBDataSources.Item("OPKG").GetValue("U_EXO_DESBUL", iLineaSel - 1).ToUpper 'CType(CType(oForm.Items.Item(sMatrix).Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_DESBUL").Cells.Item(iLineaSel - 1).Specific, SAPbouiCOM.EditText).Value.ToString


                If sOPKG <> "" Then
                    'Si no existe, creamos el artículo
                    sSQL = "SELECT ""Code"" FROM ""@EXO_PAQ"" WHERE ""Code""='" & sOPKG & "' "
                    sExiste = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    If sExiste = "" Then
                        'Presentamos UDO Y escribimos los datos de la cabecera
                        INICIO._sOPKG = sOPKG
                        INICIO._sTipo = sTipo
                        INICIO._sDes = sDes
                        INICIO._sLineaSel = iLineaSel.ToString
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_PAQ")
                    Else
                        INICIO._sOPKG = ""
                        INICIO._sTipo = ""
                        INICIO._sDes = ""
                        INICIO._sLineaSel = iLineaSel.ToString
                        objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_PAQ", sOPKG)
                    End If
                Else
                    sMensaje = "Por favor, seleccione una línea válida."
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objGlobal.SBOApp.MessageBox(sMensaje)
                End If
            Else
                sMensaje = "Tiene que seleccionar una línea."
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox(sMensaje)
            End If



            CargarUDOPAQ = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally

        End Try
    End Function
End Class
