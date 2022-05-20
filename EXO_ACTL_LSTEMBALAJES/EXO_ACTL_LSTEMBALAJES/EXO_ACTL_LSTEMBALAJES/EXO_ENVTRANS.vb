Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
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
                Case "1"
                    If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                        Cargar_Combos(oForm)
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
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
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
                    End If
                End If
            End If
            If oForm.Visible = True Then
                If pVal.ItemUID = "20_U_Cb" Then 'Clase de expedición
                    If CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value IsNot Nothing Then
                        Dim sExpedicion As String = CType(oForm.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                        sSQL = "SELECT IFNULL(""U_EXO_AGE"",'') FROM OSHP WHERE ""TrnspCode""='" & sExpedicion & "' "
                        Dim sAgeCod As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                        oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_AGTCODE", 0, sAgeCod)
                        sSQL = "SELECT ""CardName"" FROM OCRD WHERE ""CardCode""='" & sAgeCod & "' "
                        Dim sAgeNom As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                        oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_AGTNAME", 0, sAgeNom)
                    End If
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

                Cargar_Combos(oForm)
                Dim sAgencia As String = CType(oForm.Items.Item("23_U_E").Specific, SAPbouiCOM.EditText).Value
                Cargar_Combo_Matricula_Conductor_Plataforma(oForm, sAgencia)

                If objGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
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
#End Region
        Try
            'Almacen
            sSQL = "SELECT ""Branch"" FROM OUSR WHERE ""USERID""=" & objGlobal.compañia.UserSignature.ToString
            sSucursal = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            sSQL = " SELECT ""WhsCode"",""WhsName"" FROM OWHS"
            sSQL &= " WHERE ""Inactive""='N' and ""U_EXO_SUCURSAL""='" & sSucursal & "' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oform.Items.Item("22_U_Cb").DisplayDesc = True

            'Expedición
            sSQL = "SELECT ""TrnspCode"",""TrnspName"" FROM OSHP WHERE ""Active""='Y' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)


            'Series 
            sSQL = "SELECT ""Series"",""SeriesName"" FROM NNM1 WHERE ""ObjectCode""='EXO_ENVTRANS' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oform.Items.Item("4_U_Cb").DisplayDesc = True

            sFecha = CType(oform.Items.Item("21_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
            sClaseExp = CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString



            If oform.Mode = BoFormMode.fm_ADD_MODE Then
                'Poner fecha
                sFecha = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_DOCDATE", 0, sFecha)

                'Poner serie por defecto y el num. de documento
                sSQL = " SELECT ""DfltSeries"" FROM ONNM WHERE ""ObjectCode""='EXO_ENVTRANS' "
                sSerieDef = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                CType(oform.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(sSerieDef, BoSearchKey.psk_ByValue)
                Dim iNum As Integer
                iNum = oform.BusinessObject.GetNextSerialNumber(sSerieDef, oform.BusinessObject.Type.ToString)
                oform.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("DocNum", 0, iNum.ToString)

                'Poner almacen por defecto
                Try
                    sSQL = " SELECT TOP 1 ""WhsCode"" FROM OWHS"
                    sSQL &= " WHERE ""Inactive""='N' and ""U_EXO_SUCURSAL""='" & sSucursal & "' "
                    sAlmacendef = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    CType(oform.Items.Item("22_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(sAlmacendef, BoSearchKey.psk_ByValue)
                Catch ex As Exception

                End Try
                'Como en la expedición tenemos la agencia, pues tenemos que rellenarlo automático
                Dim sExpedicion As String = CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
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

            'Plataforma
            sSQL = "SELECT ""U_EXO_PLATA"",""U_EXO_PLATAD"" FROM ""@EXO_PLATAAGL"" WHERE ""Code""='" & sAgencia & "' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_5").ValidValues, sSQL)
            CType(oform.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_5").DisplayDesc = True

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
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing

        EventHandler_Choose_FromList_Before = False

        Try
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
            End If

            EventHandler_Choose_FromList_Before = True

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
        Dim sSQL As String = ""
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
                Select Case oCFLEvento.ChooseFromListUID
                    Case "CFLAT"
                        oDataTable = oCFLEvento.SelectedObjects

                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "23_U_E" Then
                                CType(oForm.Items.Item("24_U_E").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardName", 0).ToString
                                Cargar_Combo_Matricula_Conductor_Plataforma(oForm, oDataTable.GetValue("CardCode", 0).ToString)
                            End If
                        End If
                    Case "CFLIC"
                        oDataTable = oCFLEvento.SelectedObjects

                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_1" Then
                                sNombre = oDataTable.GetValue("CardName", 0).ToString
                                oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANSEX").SetValue("U_EXO_EMPRESA", pVal.Row - 1, sNombre)
                            End If
                        End If
                    Case "CFLAB"
                        oDataTable = oCFLEvento.SelectedObjects

                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "1_U_G" And pVal.ColUID = "C_1_1" Then
                                sNombre = oDataTable.GetValue("Name", 0).ToString
                                oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANSAB").SetValue("U_EXO_AGNAME", pVal.Row - 1, sNombre)
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
    Public Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument

        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_ENVTRANS"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Mode = BoFormMode.fm_OK_MODE Then
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
