Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OCRD
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

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
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    'If EventHandler_ItemPressed_After(infoEvento) = False Then
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"
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
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "134"
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

            oItem = oForm.Items.Add("btnAgencia", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO)
            oItem.Left = oForm.Items.Item("2").Left + 210
            oItem.Width = oForm.Items.Item("2").Width + 75
            oItem.Top = oForm.Items.Item("2").Top
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Enabled = False
            Dim oBtnAg As SAPbouiCOM.ButtonCombo
            oBtnAg = CType(oItem.Specific, ButtonCombo)
            oBtnAg.Caption = "Agencias"
            oBtnAg.ExpandType = BoExpandType.et_DescriptionOnly
            oBtnAg.ValidValues.Add("Servicios - Delegación", "Servicios - Delegación")
            oBtnAg.ValidValues.Add("Bultos Ag. vs Bultos", "Bultos Ag. vs Bultos")
            oBtnAg.ValidValues.Add("Conductores", "Conductores")
            oBtnAg.ValidValues.Add("Vehículos", "Vehículos")
            oBtnAg.ValidValues.Add("Plataformas", "Plataformas")
            oBtnAg.Item.AffectsFormMode = False
            oBtnAg.Item.LinkTo = "540002072"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            'oItem = oForm.Items.Add("btnSRVDEL", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem.Left = oForm.Items.Item("2").Left + 175
            'oItem.Width = oForm.Items.Item("2").Width + 70
            'oItem.Top = oForm.Items.Item("2").Top
            'oItem.Height = oForm.Items.Item("2").Height
            'oItem.Enabled = False
            'Dim oBtnAct As SAPbouiCOM.Button
            'oBtnAct = CType(oItem.Specific, Button)
            'oBtnAct.Caption = "Servicios - Delegación"
            'oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'oItem = oForm.Items.Add("btnBLTAG", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem.Left = oForm.Items.Item("2").Left + 315
            'oItem.Width = oForm.Items.Item("2").Width + 70
            'oItem.Top = oForm.Items.Item("2").Top
            'oItem.Height = oForm.Items.Item("2").Height
            'oItem.Enabled = False
            'Dim obtnBLTAG As SAPbouiCOM.Button
            'obtnBLTAG = CType(oItem.Specific, Button)
            'obtnBLTAG.Caption = "Bultos Ag. vs Bultos"
            'oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oForm.Visible = True

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
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sPropiedad As String = "" : Dim sIC As String = "" : Dim sNombre As String = ""
        Dim sMensaje As String = ""
        Dim sExiste As String = ""
        EventHandler_COMBO_SELECT_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            'Comprobamos que exista el directorio y sino, lo creamos
            sIC = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString.Trim
            sPropiedad = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("QryGroup1", 0).ToString.Trim
            sNombre = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardName", 0).ToString.Trim
            Select Case pVal.ItemUID
                Case "btnAgencia"
                    If sPropiedad = "Y" Then
                        If oForm.Mode = BoFormMode.fm_OK_MODE Then
                            Acciones_Agencias(objGlobal, oForm, sIC, sNombre)
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, guarde primero los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox("Por favor, guarde primero los datos")
                        End If
                    Else
                        sMensaje = "IC: " & sNombre & " no es Agencia. No puede acceder a las acciones específicas."
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox(sMensaje)
                    End If
                    CType(oForm.Items.Item("btnAgencia").Specific, SAPbouiCOM.ButtonCombo).Caption = "Agencias"
            End Select

            EventHandler_COMBO_SELECT_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))

        End Try
    End Function
    'Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
    '    Dim oForm As SAPbouiCOM.Form = Nothing
    '    Dim sPropiedad As String = "" : Dim sIC As String = "" : Dim sNombre As String = ""
    '    Dim sMensaje As String = ""
    '    Dim sExiste As String = ""
    '    EventHandler_ItemPressed_After = False

    '    Try
    '        oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
    '        'Comprobamos que exista el directorio y sino, lo creamos
    '        sIC = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString.Trim
    '        sPropiedad = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("QryGroup1", 0).ToString.Trim
    '        sNombre = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardName", 0).ToString.Trim
    '        Select Case pVal.ItemUID
    '            Case "btnSRVDEL"
    '                If sPropiedad = "Y" Then
    '                    If oForm.Mode = BoFormMode.fm_OK_MODE Then
    '                        'Si no existe, creamos el IC
    '                        sExiste = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Code"" FROM ""@EXO_SERVICIOS"" WHERE ""Code""='" & sIC & "' ")

    '                        If sExiste = "" Then
    '                            EXO_TSERVICIOS._sIC = sIC
    '                            EXO_TSERVICIOS._sDes = sNombre
    '                            'Presentamos UDO Y escribimos los datos de la cabecera
    '                            objGlobal.funcionesUI.cargaFormUdoBD("EXO_SERVICIOS")
    '                        Else
    '                            EXO_TSERVICIOS._sIC = ""
    '                            EXO_TSERVICIOS._sDes = ""
    '                            'Presentamos la pantalla los los datos                              
    '                            objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_SERVICIOS", sIC)
    '                        End If
    '                    Else
    '                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, guarde primero los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '                        objGlobal.SBOApp.MessageBox("Por favor, guarde primero los datos")
    '                    End If
    '                Else
    '                    sMensaje = "IC: " & sNombre & " no es Agencia. No puede acceder a los Servicios - Delegación."
    '                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '                    objGlobal.SBOApp.MessageBox(sMensaje)
    '                End If
    '            Case "btnBLTAG"
    '                If sPropiedad = "Y" Then
    '                    If oForm.Mode = BoFormMode.fm_OK_MODE Then
    '                        'Si no existe, creamos el IC
    '                        sExiste = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Code"" FROM ""@EXO_BULTOAG"" WHERE ""Code""='" & sIC & "' ")

    '                        If sExiste = "" Then
    '                            EXO_TSERVICIOS._sIC = sIC
    '                            EXO_TSERVICIOS._sDes = sNombre
    '                            'Presentamos UDO Y escribimos los datos de la cabecera
    '                            objGlobal.funcionesUI.cargaFormUdoBD("EXO_BULTOAG")
    '                        Else
    '                            EXO_TSERVICIOS._sIC = ""
    '                            EXO_TSERVICIOS._sDes = ""
    '                            'Presentamos la pantalla los los datos                              
    '                            objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_BULTOAG", sIC)
    '                        End If
    '                    Else
    '                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, guarde primero los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '                        objGlobal.SBOApp.MessageBox("Por favor, guarde primero los datos")
    '                    End If
    '                Else
    '                    sMensaje = "IC: " & sNombre & " no es Agencia. No puede acceder a Bultos Ag. vs Bultos."
    '                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '                    objGlobal.SBOApp.MessageBox(sMensaje)
    '                End If
    '        End Select

    '        EventHandler_ItemPressed_After = True

    '    Catch exCOM As System.Runtime.InteropServices.COMException
    '        Throw exCOM
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        oForm.Freeze(False)
    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))

    '    End Try
    'End Function
    Private Sub Acciones_Agencias(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form, ByVal sIC As String, ByVal sNombre As String)

        Dim ocombo As SAPbouiCOM.ButtonCombo = Nothing
        Dim Accion As String = ""
        Dim sExiste As String = ""
        Try
            ocombo = CType(oForm.Items.Item("btnAgencia").Specific, SAPbouiCOM.ButtonCombo)
            If ocombo.Selected Is Nothing Then
                Exit Sub
            End If
            Accion = ocombo.Selected.Value.Trim
            Select Case Accion
                Case "Servicios - Delegación"
                    'Si no existe, creamos el IC
                    sExiste = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Code"" FROM ""@EXO_SERVICIOS"" WHERE ""Code""='" & sIC & "' ")

                    If sExiste = "" Then
                        EXO_TSERVICIOS._sIC = sIC
                        EXO_TSERVICIOS._sDes = sNombre
                        'Presentamos UDO Y escribimos los datos de la cabecera
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_SERVICIOS")
                    Else
                        EXO_TSERVICIOS._sIC = ""
                        EXO_TSERVICIOS._sDes = ""
                        'Presentamos la pantalla los los datos                              
                        objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_SERVICIOS", sIC)
                    End If
                Case "Bultos Ag. vs Bultos"
                    'Si no existe, creamos el IC
                    sExiste = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Code"" FROM ""@EXO_BULTOAG"" WHERE ""Code""='" & sIC & "' ")

                    If sExiste = "" Then
                        EXO_BULTOAG._sIC = sIC
                        EXO_BULTOAG._sDes = sNombre
                        'Presentamos UDO Y escribimos los datos de la cabecera
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_BULTOAG")
                    Else
                        EXO_BULTOAG._sIC = ""
                        EXO_BULTOAG._sDes = ""
                        'Presentamos la pantalla los los datos                              
                        objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_BULTOAG", sIC)
                    End If
                Case "Conductores"
                    'Si no existe, creamos el IC
                    sExiste = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Code"" FROM ""@EXO_CONAG"" WHERE ""Code""='" & sIC & "' ")

                    If sExiste = "" Then
                        EXO_CONAG._sIC = sIC
                        EXO_CONAG._sDes = sNombre
                        'Presentamos UDO Y escribimos los datos de la cabecera
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_CONAG")
                    Else
                        EXO_CONAG._sIC = ""
                        EXO_CONAG._sDes = ""
                        'Presentamos la pantalla los los datos                              
                        objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_CONAG", sIC)
                    End If
                Case "Vehículos"
                    'Si no existe, creamos el IC
                    sExiste = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Code"" FROM ""@EXO_VEHIAG"" WHERE ""Code""='" & sIC & "' ")

                    If sExiste = "" Then
                        EXO_VEHIAG._sIC = sIC
                        EXO_VEHIAG._sDes = sNombre
                        'Presentamos UDO Y escribimos los datos de la cabecera
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_VEHIAG")
                    Else
                        EXO_VEHIAG._sIC = ""
                        EXO_VEHIAG._sDes = ""
                        'Presentamos la pantalla los los datos                              
                        objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_VEHIAG", sIC)
                    End If
                Case "Plataformas"
                    'Si no existe, creamos el IC
                    sExiste = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Code"" FROM ""@EXO_PLATAAG"" WHERE ""Code""='" & sIC & "' ")

                    If sExiste = "" Then
                        EXO_PLATAAG._sIC = sIC
                        EXO_PLATAAG._sDes = sNombre
                        'Presentamos UDO Y escribimos los datos de la cabecera
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_PLATAAG")
                    Else
                        EXO_PLATAAG._sIC = ""
                        EXO_PLATAAG._sDes = ""
                        'Presentamos la pantalla los los datos                              
                        objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_PLATAAG", sIC)
                    End If
            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
