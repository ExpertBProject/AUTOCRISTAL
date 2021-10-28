Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_USSOL
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "Variables Globales"
    Public Shared _sDocEntry As String = ""
    Public Shared _sSerie As String = ""
    Public Shared _sDocNum As String = ""
    Public Shared _sAlmacen As String = ""
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

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_USSOL.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_USSOL", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                        Case "UDO_FT_EXO_USSOL"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
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
                        Case "UDO_FT_EXO_USSOL"
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
                        Case "UDO_FT_EXO_USSOL"
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
                        Case "UDO_FT_EXO_USSOL"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
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
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                oItem = oForm.Items.Item("14_U_E")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                oItem = oForm.Items.Item("cbSeries")
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                'Introducimos los valores en los combos
                CargarCombos(objGlobal, oForm)

                If _sDocEntry <> "" Then
                    oForm.Mode = BoFormMode.fm_ADD_MODE
                    oForm.DataSources.DBDataSources.Item("@EXO_USSOL").SetValue("Code", 0, _sDocEntry)
                    oForm.DataSources.DBDataSources.Item("@EXO_USSOL").SetValue("U_EXO_SERIES", 0, _sSerie)
                    oForm.DataSources.DBDataSources.Item("@EXO_USSOL").SetValue("U_EXO_DOCNUM", 0, _sDocNum)
                End If

            End If


            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oItem, Object))
        End Try
    End Function
    Private Sub CargarCombos(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Try
            sSQL = "SELECT ""Series"",""SeriesName"" FROM NNM1 WHERE ""ObjectCode""='1250000001' "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbSeries").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "1" Then
                If pVal.ActionSuccess = True Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then 'Después de añadir
                        objGlobal.SBOApp.ActivateMenuItem("1291")
                    End If
                End If
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
                    Case "171"
                        Try
                            Dim sNombre As String = oDataTable.GetValue("firstName", 0).ToString
                            Dim sNombre2 As String = oDataTable.GetValue("middleName", 0).ToString
                            Dim sApell As String = oDataTable.GetValue("lastName", 0).ToString

                            oForm.DataSources.DBDataSources.Item("@EXO_USSOLL").SetValue("U_EXO_NOMBRE", pVal.Row - 1, sNombre)
                            CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sNombre
                            oForm.DataSources.DBDataSources.Item("@EXO_USSOLL").SetValue("U_EXO_NOM2", pVal.Row - 1, sNombre2)
                            CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sNombre2
                            oForm.DataSources.DBDataSources.Item("@EXO_USSOLL").SetValue("U_EXO_APELL", pVal.Row - 1, sApell)
                            CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sApell
                        Catch ex As Exception
                            oForm.DataSources.DBDataSources.Item("@EXO_USSOLL").SetValue("U_EXO_NOMBRE", pVal.Row - 1, oDataTable.GetValue("firstName", 0).ToString)
                            oForm.DataSources.DBDataSources.Item("@EXO_USSOLL").SetValue("U_EXO_NOM2", pVal.Row - 1, oDataTable.GetValue("middleName", 0).ToString)
                            oForm.DataSources.DBDataSources.Item("@EXO_USSOLL").SetValue("U_EXO_APELL", pVal.Row - 1, oDataTable.GetValue("lastName", 0).ToString)
                        End Try

                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End Select
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
    Private Function EventHandler_Choose_FromList_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim sSucursal As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        EventHandler_Choose_FromList_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_1" Then 'Usuarios
                sSQL = "SELECT  ""OUBR"".""Code"" From ""OWHS"" LEFT JOIN ""OUBR"" ON ""OWHS"".""U_EXO_SUCURSAL""=""OUBR"".""Name"" WHERE ""WhsCode""='" & _sAlmacen & "' "
                sSucursal = objGlobal.refDi.SQL.sqlStringB1(sSQL)


                oCFLEvento = CType(pVal, IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions

                sSQL = "SELECT ""HEM6"".""empID"" ""US"" FROM ""HEM6"" INNER JOIN ""OHEM"" ON ""HEM6"".""empID""=""OHEM"".""empID"" "
                sSQL &= " INNER JOIN ""OHTY"" On ""HEM6"".""roleID""=""OHTY"".""typeID"" "
                sSQL &= " WHERE ""OHTY"".""name""='Almacén' and ""OHEM"".""branch""='" & sSucursal & "' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    For i = 0 To oRs.RecordCount - 1
                        oCond = oConds.Add
                        oCond.Alias = "empID"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = oRs.Fields.Item("US").Value.ToString
                        If i <> oRs.RecordCount - 1 Then
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        End If
                        oRs.MoveNext()
                    Next
                Else
                    oCond = oConds.Add
                    oCond.Alias = "branch"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = sSucursal
                End If

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            End If

            EventHandler_Choose_FromList_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.FormConditions(oConds)
            EXO_CleanCOM.CLiberaCOM.FormCondition(oCond)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
End Class
