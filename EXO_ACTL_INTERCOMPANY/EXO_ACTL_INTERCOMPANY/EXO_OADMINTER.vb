Imports SAPbouiCOM
Public Class EXO_OADMINTER
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaCampos()
            CargarScripts()
        End If
        cargamenu()
    End Sub
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_OADMINTERC.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_OADMINTERC", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


        End If


    End Sub
    Private Sub CargarScripts()
        Dim sScript As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then
            Try
                sScript = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_TABLE_REPLICATE.sql")
                objGlobal.refDi.SQL.executeNonQuery(sScript)
            Catch exCOM As System.Runtime.InteropServices.COMException
                Throw exCOM
            Catch ex As Exception
                Throw ex
            End Try

            Try
                sScript = objGlobal.funciones.leerEmbebido(Me.GetType(), "_98_EXO_CHK_OITM.sql")
                objGlobal.refDi.SQL.executeNonQuery(sScript)
            Catch exCOM As System.Runtime.InteropServices.COMException
                Throw exCOM
            Catch ex As Exception
                Throw ex
            End Try

            Try
                sScript = objGlobal.funciones.leerEmbebido(Me.GetType(), "_98_EXO_CHK_OITB.sql")
                objGlobal.refDi.SQL.executeNonQuery(sScript)
            Catch exCOM As System.Runtime.InteropServices.COMException
                Throw exCOM
            Catch ex As Exception
                Throw ex
            End Try

            Try
                sScript = objGlobal.funciones.leerEmbebido(Me.GetType(), "_98_EXO_CHK_OITG.sql")
                objGlobal.refDi.SQL.executeNonQuery(sScript)
            Catch exCOM As System.Runtime.InteropServices.COMException
                Throw exCOM
            Catch ex As Exception
                Throw ex
            End Try

            Try
                sScript = objGlobal.funciones.leerEmbebido(Me.GetType(), "_98_EXO_CHK_OMRC.sql")
                objGlobal.refDi.SQL.executeNonQuery(sScript)
            Catch exCOM As System.Runtime.InteropServices.COMException
                Throw exCOM
            Catch ex As Exception
                Throw ex
            End Try

            Try
                sScript = objGlobal.funciones.leerEmbebido(Me.GetType(), "_98_EXO_CHK_OSHP.sql")
                objGlobal.refDi.SQL.executeNonQuery(sScript)
            Catch exCOM As System.Runtime.InteropServices.COMException
                Throw exCOM
            Catch ex As Exception
                Throw ex
            End Try


        End If
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults

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
                        Case "UDO_FT_EXO_OADMINTERC"
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
                        Case "UDO_FT_EXO_OADMINTERC"

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
                        Case "UDO_FT_EXO_OADMINTERC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(objGlobal, infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_OADMINTERC"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

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
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_OADMINTERC"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If Valida_Campos_Lineas(oForm) = False Then
                                    Return False
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If Valida_Campos_Lineas(oForm) = False Then
                                    Return False
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_OADMINTERC"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

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
    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oMat As SAPbouiCOM.Matrix = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sCode As String = ""

        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                oMat = CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
                sSQL = "SELECT ""dbName"",""cmpName""" _
                & " FROM     ""SBOCOMMON"".""SRGC"""
                objGlobal.funcionesUI.cargaCombo(oMat.Columns.Item("C_0_1").ValidValues, sSQL)
                oMat.Columns.Item("C_0_1").ExpandType = BoExpandType.et_ValueDescription
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)


                oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                oRs.DoQuery("SELECT ""Code"" FROM ""@EXO_OADMINTERC""  ")

                If oRs.RecordCount > 0 Then
                    sCode = oRs.Fields.Item("Code").Value.ToString()

                End If
                If sCode <> "" Then
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    'oForm.Items.Item("0_U_E").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True)
                    CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.ComboBox).Select(sCode, BoSearchKey.psk_ByValue)
                    oForm.Items.Item("1").Click(BoCellClickType.ct_Regular)
                Else
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                End If
            End If
            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oMat, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCombo, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "EXO-MnConfE"
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_OADMINTERC")
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function Valida_Campos_Lineas(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Valida_Campos_Lineas = False
        Dim sEmpresa As String = ""
        Dim intCont As Integer = 0

        Try
            If oForm.Visible = True Then
                For i = 1 To CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).RowCount
                    intCont = 0
                    sEmpresa = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(i).Specific, SAPbouiCOM.ComboBox).Value
                    'for de todo, tengo que encontrar una
                    For j = 1 To CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).RowCount
                        If sEmpresa = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(j).Specific, SAPbouiCOM.ComboBox).Value Then
                            intCont = intCont + 1
                        End If
                    Next
                    If intCont > 1 Then
                        'esta dos veces metida la empresa, aviso de que no puede guardar así hasta que no quite una
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - La empresa " & sEmpresa & " no puede estar definida más de una vez", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objGlobal.SBOApp.MessageBox(" La empresa " & sEmpresa & " no puede estar definida más de una vez.")
                        Exit Function
                    End If

                Next
            End If

            Valida_Campos_Lineas = True
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
End Class
