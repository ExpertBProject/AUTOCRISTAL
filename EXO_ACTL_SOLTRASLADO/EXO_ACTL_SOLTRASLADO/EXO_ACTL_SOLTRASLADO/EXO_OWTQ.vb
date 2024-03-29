﻿Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OWTQ
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        If actualizar Then
            cargaCampos()
        End If
    End Sub
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OWTQ.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDFs_OWTQ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_WTQ1.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDFs_WTQ1", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                        Case "1250000940"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_Validate_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "1250000940"
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
                        Case "1250000940"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                    If EventHandler_LOST_FOCUS_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "1250000940"
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

            oForm.Freeze(True)
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            oItem = oForm.Items.Add("cbTipoSol", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.LinkTo = "1470000101"
            oItem.Top = oForm.Items.Item("1470000101").Top + oForm.Items.Item("1470000101").Height + 2
            oItem.Left = oForm.Items.Item("1470000101").Left
            oItem.Height = oForm.Items.Item("1470000101").Height
            oItem.Width = oForm.Items.Item("1470000101").Width
            oItem.DisplayDesc = True
            oItem.Enabled = True
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OWTQ", "U_EXO_TIPO")
            oItem = oForm.Items.Add("lblTipoSol", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("cbTipoSol").Top
            oItem.Left = oForm.Items.Item("1470000099").Left
            oItem.Height = oForm.Items.Item("1470000099").Height
            oItem.Width = oForm.Items.Item("1470000099").Width
            oItem.LinkTo = "cbTipoSol"
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Tipo Solicitud"

            oItem = oForm.Items.Add("cbStatusP", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.LinkTo = "234000002"
            oItem.Top = oForm.Items.Item("234000002").Top + oForm.Items.Item("234000002").Height + 2
            oItem.Left = oForm.Items.Item("234000002").Left
            oItem.Height = oForm.Items.Item("234000002").Height
            oItem.Width = oForm.Items.Item("234000002").Width
            oItem.DisplayDesc = True
            oItem.Enabled = True
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OWTQ", "U_EXO_STATUSP")
            oItem = oForm.Items.Add("lblStatusP", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("cbStatusP").Top
            oItem.Left = oForm.Items.Item("234000001").Left
            oItem.Height = oForm.Items.Item("234000001").Height
            oItem.Width = oForm.Items.Item("234000001").Width
            oItem.LinkTo = "cbStatusP"
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Status Picking"

            'oItem = oForm.Items.Add("btnUS", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem.Left = oForm.Items.Item("2").Left + 175
            'oItem.Width = oForm.Items.Item("2").Width + 70
            'oItem.Top = oForm.Items.Item("2").Top
            'oItem.Height = oForm.Items.Item("2").Height
            'oItem.Enabled = False
            'Dim oBtnAct As SAPbouiCOM.Button
            'oBtnAct = CType(oItem.Specific, Button)
            'oBtnAct.Caption = "Usuarios asignados"
            'oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oForm.Freeze(False)

            EventHandler_Form_Load = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            If oForm IsNot Nothing Then oForm.Visible = True

            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            If oForm IsNot Nothing Then oForm.Visible = True

            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_Validate_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sALMD As String = "" : Dim sALMH As String = ""
        EventHandler_Validate_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            'If pVal.ItemUID = "18" Or pVal.ItemUID = "1470000101" Then
            sALMD = CType(oForm.Items.Item("18").Specific, SAPbouiCOM.EditText).Value.ToString
            sALMH = CType(oForm.Items.Item("1470000101").Specific, SAPbouiCOM.EditText).Value.ToString

            If sALMD = sALMH Then
                CType(oForm.Items.Item("cbTipoSol").Specific, SAPbouiCOM.ComboBox).Select("INT", BoSearchKey.psk_ByValue)
            Else
                CType(oForm.Items.Item("cbTipoSol").Specific, SAPbouiCOM.ComboBox).Select("ITC", BoSearchKey.psk_ByValue)
            End If
            'End If


            EventHandler_Validate_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_LOST_FOCUS_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sALMD As String = "" : Dim sALMH As String = ""
        EventHandler_LOST_FOCUS_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "18" Or pVal.ItemUID = "1470000101" Then
                sALMD = CType(oForm.Items.Item("18").Specific, SAPbouiCOM.EditText).Value.ToString
                sALMH = CType(oForm.Items.Item("1470000101").Specific, SAPbouiCOM.EditText).Value.ToString

                If sALMD = sALMH Then
                    CType(oForm.Items.Item("cbTipoSol").Specific, SAPbouiCOM.ComboBox).Select("INT", BoSearchKey.psk_ByValue)
                Else
                    CType(oForm.Items.Item("cbTipoSol").Specific, SAPbouiCOM.ComboBox).Select("ITC", BoSearchKey.psk_ByValue)
                End If
            End If


            EventHandler_LOST_FOCUS_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sMensaje As String = ""
        Dim sDocEntry As String = "" : Dim sSerie As String = "" : Dim sDocnum As String = "" : Dim sAlmacen As String = "" : Dim sAlmacenD As String = ""
        Dim sExiste As String = ""
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            'Comprobamos que exista el directorio y sino, lo creamos
            sDocEntry = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("DocEntry", 0).ToString.Trim
            sSerie = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("Series", 0).ToString.Trim
            sDocnum = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("DocNum", 0).ToString.Trim
            sAlmacen = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("ToWhsCode", 0).ToString.Trim
            sAlmacenD = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("Filler", 0).ToString.Trim
            Select Case pVal.ItemUID
                Case "btnUS"
                    If oForm.Mode = BoFormMode.fm_OK_MODE Then
                        If sAlmacen = sAlmacenD Then
                            'Si no existe, creamos el IC
                            sExiste = objGlobal.refDi.SQL.sqlStringB1("SELECT ""DocEntry"" FROM ""@EXO_USSOL"" WHERE ""Code""='" & sDocEntry & "' ")

                            If sExiste = "" Then
                                EXO_USSOL._sDocEntry = sDocEntry
                                EXO_USSOL._sSerie = sSerie
                                EXO_USSOL._sDocNum = sDocnum
                                EXO_USSOL._sAlmacen = sAlmacen
                                'Presentamos UDO Y escribimos los datos de la cabecera
                                objGlobal.funcionesUI.cargaFormUdoBD("EXO_USSOL")
                            Else
                                EXO_USSOL._sDocEntry = ""
                                EXO_USSOL._sSerie = ""
                                EXO_USSOL._sDocNum = ""
                                EXO_USSOL._sAlmacen = sAlmacen
                                'Presentamos la pantalla los los datos                              
                                objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_USSOL", sDocEntry)
                            End If
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Sólo está activo para la asignación de usuarios de Traslados internos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox("Sólo está activo para la asignación de usuarios de Traslados internos.")
                        End If

                    Else
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, guarde primero los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox("Por favor, guarde primero los datos")
                    End If
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
End Class
