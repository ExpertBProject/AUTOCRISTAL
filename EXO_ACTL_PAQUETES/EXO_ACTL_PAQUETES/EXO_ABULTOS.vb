Imports SAPbouiCOM
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Public Class EXO_ABULTOS
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
                        Case "UDO_FT_EXO_ABULTOS"
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
                        Case "UDO_FT_EXO_ABULTOS"
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
                        Case "UDO_FT_EXO_ABULTOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(objGlobal, infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_ABULTOS"
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
    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item = Nothing
        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                oItem = oForm.Items.Item("btnImp")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                Carga_Des(oForm, objGlobal)
                If oForm.Mode = BoFormMode.fm_OK_MODE Then
                    objGlobal.SBOApp.ActivateMenuItem("1304")
                End If

            End If

            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oItem, Object))
        End Try
    End Function
    Private Sub Carga_Des(ByRef oForm As SAPbouiCOM.Form, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI)
        Dim sSQL As String = ""
        Try
            sSQL = "SELECT ""Code"" ""Código"", ""Name"" ""Delegación"" FROM OUBR ORDER BY ""Name"" "
            oObjGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("14_U_CB").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim rutaCrystal As String = "" : Dim sRutaFicheros As String = "" : Dim sReport As String = "" : Dim sTipoImp As String = ""
        Dim sCrystal As String = "Et_ABultos.rpt"
        Dim sCode As String = ""
#End Region


        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "btnImp" Then
                rutaCrystal = objGlobal.path & "\05.Rpt\ETIQUETAS\"
                sCode = oForm.DataSources.DBDataSources.Item("@EXO_ABULTOS").GetValue("Code", 0).ToUpper
                sTipoImp = "IMP"
                'Imprimimos la etiqueta
                GenerarImpCrystal(rutaCrystal, sCrystal, sCode, sRutaFicheros, sReport, sTipoImp)

            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean

        Dim sSQL As String = ""
        Try
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnABUL"
                        If EXO_GLOBALES.CargarUDO(objGlobal, "EXO_ABULTOS", "") = False Then
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
            sServer = objGlobal.funcionesUI.refDi.OGEN.valorVariable("SERVIDOR_HANA") ' objGlobal.compañia.Server
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
                        oCRReport.PrintToPrinter(nCopias, False, 0, 999)
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
            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso de impresión.", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_None)
            oCRReport = Nothing
            oFileDestino = Nothing
        End Try
    End Sub
End Class
