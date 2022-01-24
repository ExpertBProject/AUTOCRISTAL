Imports SAPbouiCOM

Public Class EXO_EUROCODES
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "Variables"
    Public Shared _sArticulo As String = ""
    Public Shared _sDes As String = ""
#End Region
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaCampos()
        End If
    End Sub
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_EUROCODES.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_EUROCODES", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_EUROCODES"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_EUROCODES"

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
                        Case "UDO_FT_EXO_EUROCODES"
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
                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_EUROCODES"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

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
    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                oItem = oForm.Items.Item("0_U_E")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                oItem = oForm.Items.Item("1_U_E")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


                If _sArticulo <> "" Then
                    oForm.Mode = BoFormMode.fm_ADD_MODE
                    oForm.DataSources.DBDataSources.Item("@EXO_EUROCODES").SetValue("Code", 0, _sArticulo)
                    oForm.DataSources.DBDataSources.Item("@EXO_EUROCODES").SetValue("Name", 0, _sDes)
                End If

            End If

            EventHandler_Form_Visible = True

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
    Private Function EventHandler_Choose_FromList_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sItemCode As String
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim sGroupCode As String = ""
        Dim bEsADR As Boolean = False
        Dim bDebeSerADR As Boolean = False
        Dim h As Integer = 1

        EventHandler_Choose_FromList_Before = False

        Try
            oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "14_U_E" Then 'Marca
                oCFLEvento = CType(pVal, IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions
                oCond = oConds.Add
                oCond.Alias = "U_EXO_COD"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                Dim sCode As String = Left(CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value.ToString.Trim, 2)
                oCond.CondVal = sCode
                'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            ElseIf pVal.ItemUID = "15_U_E" Then 'Modelo
                oCFLEvento = CType(pVal, IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions
                oCond = oConds.Add
                oCond.Alias = "U_EXO_COD"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                Dim sCode As String = Left(CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value.ToString.Trim, 4)
                oCond.CondVal = sCode
                'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            ElseIf pVal.ItemUID = "17_U_E" Then 'Luna
                oCFLEvento = CType(pVal, IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions
                oCond = oConds.Add
                oCond.Alias = "Code"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                Dim sCode As String = Mid(CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value.ToString.Trim, 5, 1)
                oCond.CondVal = sCode
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
                    Case "EXO_MARCAS"
                        Try
                            oForm.DataSources.UserDataSources.Item("UDMARCA").ValueEx = oDataTable.GetValue("Name", 0).ToString
                        Catch ex As Exception

                        End Try
                    Case "EXO_MODELOS"
                        Try
                            oForm.DataSources.UserDataSources.Item("UDMOD").ValueEx = oDataTable.GetValue("Name", 0).ToString
                        Catch ex As Exception

                        End Try
                    Case "EXO_LUNAS"
                        Try
                            oForm.DataSources.UserDataSources.Item("ULUN").ValueEx = oDataTable.GetValue("Name", 0).ToString
                        Catch ex As Exception

                        End Try
                    Case "EXO_TINTES"
                        Try
                            oForm.DataSources.UserDataSources.Item("UDTIN").ValueEx = oDataTable.GetValue("Name", 0).ToString
                        Catch ex As Exception

                        End Try
                    Case "EXO_VISERAS"
                        Try
                            oForm.DataSources.UserDataSources.Item("UDVIS").ValueEx = oDataTable.GetValue("Name", 0).ToString
                        Catch ex As Exception

                        End Try
                    Case "EXO_CUERPOS"
                        Try
                            oForm.DataSources.UserDataSources.Item("UDCUE").ValueEx = oDataTable.GetValue("Name", 0).ToString
                        Catch ex As Exception

                        End Try
                    Case "EXO_CARROCERIAS"
                        Try
                            oForm.DataSources.UserDataSources.Item("UDCAR").ValueEx = oDataTable.GetValue("Name", 0).ToString
                        Catch ex As Exception

                        End Try
                    Case "EXO_LUNASP"
                        Try
                            oForm.DataSources.UserDataSources.Item("UDPOS").ValueEx = oDataTable.GetValue("Name", 0).ToString
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
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_EUROCODES"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                Borra_Des(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                Borra_Des(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                Borra_Des(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_EUROCODES"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                'If oForm.Visible = True Then
                                Carga_Des(oForm)
                                'End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

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
    Private Sub Borra_Des(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Try
            oForm.DataSources.UserDataSources.Item("UDMARCA").ValueEx = ""
            oForm.DataSources.UserDataSources.Item("UDMOD").ValueEx = ""
            oForm.DataSources.UserDataSources.Item("ULUN").ValueEx = ""
            oForm.DataSources.UserDataSources.Item("UDTIN").ValueEx = ""
            oForm.DataSources.UserDataSources.Item("UDVIS").ValueEx = ""
            oForm.DataSources.UserDataSources.Item("UDCUE").ValueEx = ""
            oForm.DataSources.UserDataSources.Item("UDCAR").ValueEx = ""
            oForm.DataSources.UserDataSources.Item("UDPOS").ValueEx = ""
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Carga_Des(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Try
            oForm.DataSources.UserDataSources.Item("UDMARCA").ValueEx = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Name"" FROM ""@EXO_MARCAS"" WHERE ""Code""='" & CType(oForm.Items.Item("14_U_E").Specific, SAPbouiCOM.EditText).Value.ToString & "' ")
            oForm.DataSources.UserDataSources.Item("UDMOD").ValueEx = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Name"" FROM ""@EXO_MODELOS"" WHERE ""Code""='" & CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).Value.ToString & "' ")
            oForm.DataSources.UserDataSources.Item("ULUN").ValueEx = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Name"" FROM ""@EXO_LUNAS"" WHERE ""Code""='" & CType(oForm.Items.Item("17_U_E").Specific, SAPbouiCOM.EditText).Value.ToString & "' ")
            oForm.DataSources.UserDataSources.Item("UDTIN").ValueEx = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Name"" FROM ""@EXO_TINTES"" WHERE ""Code""='" & CType(oForm.Items.Item("18_U_E").Specific, SAPbouiCOM.EditText).Value.ToString & "' ")
            oForm.DataSources.UserDataSources.Item("UDVIS").ValueEx = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Name"" FROM ""@EXO_VISERAS"" WHERE ""Code""='" & CType(oForm.Items.Item("19_U_E").Specific, SAPbouiCOM.EditText).Value.ToString & "' ")
            oForm.DataSources.UserDataSources.Item("UDCUE").ValueEx = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Name"" FROM ""@EXO_CUERPOS"" WHERE ""Code""='" & CType(oForm.Items.Item("20_U_E").Specific, SAPbouiCOM.EditText).Value.ToString & "' ")
            oForm.DataSources.UserDataSources.Item("UDCAR").ValueEx = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Name"" FROM ""@EXO_CARROCERIAS"" WHERE ""Code""='" & CType(oForm.Items.Item("21_U_E").Specific, SAPbouiCOM.EditText).Value.ToString & "' ")
            oForm.DataSources.UserDataSources.Item("UDPOS").ValueEx = objGlobal.refDi.SQL.sqlStringB1("SELECT ""Name"" FROM ""@EXO_LUNASP"" WHERE ""Code""='" & CType(oForm.Items.Item("22_U_E").Specific, SAPbouiCOM.EditText).Value.ToString & "' ")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
