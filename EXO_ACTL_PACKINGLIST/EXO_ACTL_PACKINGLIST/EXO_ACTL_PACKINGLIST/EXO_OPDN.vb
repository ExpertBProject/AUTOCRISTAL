Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OPDN
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "Variables globales"
    Dim _sDocEntryODRF As String = ""
#End Region
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        If actualizar Then
            cargaCampos()
        End If
    End Sub
    Private Sub cargaCampos()
        Dim result As Boolean = False
        Dim sSQL As String = ""
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_PACKING.xml")
            result = objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_PACKING", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_WTQ1.xml")
            result = objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDFs_WTQ1", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OPOR.xml")
            result = objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDFs_OPOR", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "UBIDESTINO.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_UbicacionDestinoEntradaCompra_2", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        End If
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function

    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "143"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "143"
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
                        Case "143"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "143"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Dim oItem As SAPbouiCOM.Item
        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            oForm.Visible = False
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oItem = oForm.Items.Add("btnTR", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 5
            oItem.Width = oForm.Items.Item("2").Width + 100
            oItem.Top = oForm.Items.Item("2").Top
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Enabled = False
            oItem.LinkTo = "2"
            Dim oBtnAct As SAPbouiCOM.Button
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Gen. Lista de Embalaje"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oForm.Visible = True


            _sDocEntryODRF = oForm.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString.Trim
            EventHandler_Form_Load = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            If oForm IsNot Nothing Then oForm.Visible = True

            Throw exCOM
        Catch ex As Exception
            If oForm IsNot Nothing Then oForm.Visible = True

            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim iTamMatrix As Integer = 0
        Dim bSelLinea As Boolean = False
        Dim sMensaje As String = ""
        Dim sEstado As String = ""
        Dim sCancelado As String = ""
        Dim sDocNum As String = ""
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            'Comprobamos que exista el directorio y sino, lo creamos
            sCancelado = oForm.DataSources.DBDataSources.Item("OPDN").GetValue("CANCELED", 0).ToString.Trim
            sEstado = oForm.DataSources.DBDataSources.Item("OPDN").GetValue("DocStatus", 0).ToString.Trim
            Select Case pVal.ItemUID
                Case "btnTR"
                    If sEstado = "O" And sCancelado = "N" Then
                        EXO_GLOBALES.Gen_Lista_Embalaje(objGlobal.compañia, objGlobal, oForm)
                    Else
                        If sCancelado = "Y" Then
                            sMensaje = "La Entrada de Mercancía está cancelada, no se puede generar la Lista de embalaje."
                        Else
                            If sEstado = "C" Then
                                sMensaje = "La Entrada de Mercancía está cerrada, no se puede generar la Lista de embalaje."
                            End If
                        End If
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox(sMensaje)
                    End If
                Case "1"
                    'If CType(oForm.Items.Item("81").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                    '    sEstado = CType(oForm.Items.Item("81").Specific, SAPbouiCOM.ComboBox).Selected.Description.ToString
                    '    If sEstado = "Borrador" Then
                    '        'Copiamos las líneas al definitivo
                    '        sDocNum = oForm.DataSources.DBDataSources.Item("OPDN").GetValue("DocNum", 0).ToString.Trim
                    '    End If
                    'End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))

        End Try
    End Function
    Public Overrides Function SBOApp_FormDataEvent(infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "143"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        End Select

                End Select

            Else
                Select Case infoEvento.FormTypeEx
                    Case "143"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then
                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                    If oForm.Visible = True Then
                                        Dim sDocEntryOPDN As String = oForm.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString.Trim
                                        If sDocEntryOPDN <> "" Then
                                            'Copiamos de @EXO_PACKINGL del borrador al definitivo.
                                            Dim sSQL As String = ""
                                            sSQL = "INSERT INTO ""@EXO_PACKINGL"" (""Code"", ""LineId"", ""Object"", ""LogInst"", ""U_EXO_USUARIO"", ""U_EXO_CAT"", ""U_EXO_CODE"", ""U_EXO_CANT"",  "
                                            sSQL &= " ""U_EXO_LOTE"",  ""U_EXO_FFAB"", ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"",""U_EXO_LINEA"")  "
                                            sSQL &= " SELECT '" & sDocEntryOPDN & "', ""LineId"", 'OPDN', '0',  '" & objGlobal.compañia.UserName.ToString & "', ""U_EXO_CAT"", ""U_EXO_CODE"", ""U_EXO_CANT"", ""U_EXO_LOTE"",   "
                                            sSQL &= " ""U_EXO_FFAB"", ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"", ""U_EXO_LINEA"" "
                                            sSQL &= " FROM ""@EXO_PACKINGL""  where ""Object""='ODRF' and ""Code""='" & _sDocEntryODRF & "' "
                                            objGlobal.refDi.SQL.sqlUpdB1(sSQL)
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                End Select
            End If

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

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
End Class
