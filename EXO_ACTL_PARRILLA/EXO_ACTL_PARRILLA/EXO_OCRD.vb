Imports SAPbouiCOM
Public Class EXO_OCRD
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
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

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
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "134"
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

    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Dim oItem As SAPbouiCOM.Item
        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            oForm.Freeze(True)
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

#Region "Delegación"
            'Campo de la delegación
            oItem = oForm.Items.Add("cbDele", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.LinkTo = "102"
            oItem.Top = oForm.Items.Item("228").Top
            oItem.Left = oForm.Items.Item("102").Left
            oItem.Height = oForm.Items.Item("102").Height
            oItem.Width = oForm.Items.Item("102").Width
            oItem.Enabled = True
            oItem.DisplayDesc = True
            oItem.FromPane = 1 : oItem.ToPane = 1
            CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OCRD", "U_EXO_DELE")
            oItem = oForm.Items.Add("lblDele", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("cbDele").Top
            oItem.Left = oForm.Items.Item("101").Left
            oItem.Height = oForm.Items.Item("101").Height
            oItem.Width = oForm.Items.Item("101").Width
            oItem.FromPane = 1 : oItem.ToPane = 1
            oItem.LinkTo = "cbDele"
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Delegación"
#End Region
#Region "Agencia de Transporte"
            'Agencia de Transporte
            oItem = oForm.Items.Add("cbATrans", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.LinkTo = "cbDele"
            oItem.Top = oForm.Items.Item("430").Top
            oItem.Left = oForm.Items.Item("102").Left
            oItem.Height = oForm.Items.Item("102").Height
            oItem.Width = oForm.Items.Item("102").Width
            oItem.Enabled = True
            oItem.DisplayDesc = True
            oItem.FromPane = 1 : oItem.ToPane = 1
            CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OCRD", "U_EXO_AGENCIA")
            oItem = oForm.Items.Add("lblATrans", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("cbATrans").Top
            oItem.Left = oForm.Items.Item("101").Left
            oItem.Height = oForm.Items.Item("101").Height
            oItem.Width = oForm.Items.Item("101").Width
            oItem.FromPane = 1 : oItem.ToPane = 1
            oItem.LinkTo = "cbATrans"
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Agencia de Transporte"
#End Region
#Region "Clave de Acceso"
            'Movemos el campo Clave de acceso
            oItem = oForm.Items.Item("185")
            oItem.Top = oForm.Items.Item("cbATrans").Top + oForm.Items.Item("cbATrans").Height
            oItem = oForm.Items.Item("184")
            oItem.Top = oForm.Items.Item("185").Top
#End Region
            'Introducimos los valores en los combos
            CargarCombos(objGlobal, oForm)

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
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub CargarCombos(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Try
            sSQL = "SELECT ""Code"",""Name"" FROM OUBR Order by ""Name"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbDele").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            Dim sClienteAct As String = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString.Trim

            sSQL = "SELECT * FROM ( "
            sSQL &= " SELECT ""CardCode"",""CardName"" FROM OCRD "
            sSQL &= " WHERE ""CardType""='S' and ""QryGroup1""='Y' "
            sSQL &= " And ""CardCode"" not in(SELECT ""U_EXO_COD"" FROM ""@EXO_LNEGRAL"" Where ""Code""='" & sClienteAct & "') "
            sSQL &= " UNION ALL "
            sSQL &= " SELECT '-', '' FROM DUMMY "
            sSQL &= " )T Order by ""CardName"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbATrans").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

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
                    Case "134"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "134"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Visible = True Then
                                    CargarCombos(objGlobal, oForm)
                                End If
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

End Class
