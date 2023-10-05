
Imports SAPbobsCOM
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
                        Case "134"
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
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

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

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

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

            'campos y botón
            oItem = oForm.Items.Add("cbFormato", BoFormItemTypes.it_COMBO_BOX)
            oItem.Top = oForm.Items.Item("1470002110").Top + oForm.Items.Item("1470002110").Height + 5
            oItem.Left = oForm.Items.Item("1470002110").Left
            oItem.Height = oForm.Items.Item("1470002110").Height
            oItem.Width = (oForm.Items.Item("1470002110").Width + 80)
            oItem.LinkTo = "1470002110"
            oItem.FromPane = 1
            oItem.ToPane = 1
            CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OCRD", "U_EXO_FORMATO")
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oItem = oForm.Items.Add("lblFormato", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("cbFormato").Top
            oItem.Left = oForm.Items.Item("1470002109").Left
            oItem.Height = oForm.Items.Item("1470002109").Height
            oItem.Width = oForm.Items.Item("1470002109").Width
            oItem.LinkTo = "cbFormato"
            oItem.FromPane = 1
            oItem.ToPane = 1
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Formato "

            oItem = oForm.Items.Add("txtDIR", BoFormItemTypes.it_EDIT)
            oItem.Top = oForm.Items.Item("cbFormato").Top
            oItem.Left = oForm.Items.Item("362").Left
            oItem.Height = oForm.Items.Item("362").Height
            oItem.Width = (oForm.Items.Item("362").Width * 2) - 50
            oItem.LinkTo = "362"
            oItem.FromPane = 1
            oItem.ToPane = 1
            CType(oItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "OCRD", "U_EXO_DIR")
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem = oForm.Items.Add("lblDIR", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("txtDIR").Top
            oItem.Left = oForm.Items.Item("358").Left
            oItem.Height = oForm.Items.Item("358").Height
            oItem.Width = oForm.Items.Item("358").Width
            oItem.LinkTo = "txtDIR"
            oItem.FromPane = 1
            oItem.ToPane = 1
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Directorio Exportar "

            oItem = oForm.Items.Add("tbnDIR", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("txtDIR").Left + oForm.Items.Item("txtDIR").Width + 2
            oItem.Width = oForm.Items.Item("111").Width
            oItem.Top = oForm.Items.Item("txtDIR").Top
            oItem.Height = oForm.Items.Item("111").Height
            Dim oBtnAg As SAPbouiCOM.Button
            oBtnAg = CType(oItem.Specific, Button)
            oBtnAg.Caption = "..."
            oBtnAg.Item.AffectsFormMode = False
            oBtnAg.Item.LinkTo = "txtDIR"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            CargaCombos(oForm)


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
    Private Sub CargaCombos(ByRef oform As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Try
            sSQL = "SELECT 0 ""IntrnalKey"", CAST(' ' as NVARCHAR(50)) ""QName""   FROM DUMMY 
                    UNION ALL
                    SELECT ""IntrnalKey"", ""QName"" FROM OUQR Q
                        INNER JOIN  OQCN C ON Q.""QCategory""=C.""CategoryId""
                        WHERE ""CatName""='LOGISTICA' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("cbFormato").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oform.Items.Item("cbFormato").DisplayDesc = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sDirOrigen As String = ""
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            Select Case pVal.ItemUID
                Case "tbnDIR"
                    'Tenemos que controlar que es cliente o web
                    If objGlobal.SBOApp.ClientType = SAPbouiCOM.BoClientType.ct_Browser Then
                        sDirOrigen = objGlobal.SBOApp.GetFileFromBrowser() 'Modificar
                        sDirOrigen = IO.Path.GetDirectoryName(sDirOrigen)
                    Else
                        'Controlar el tipo de fichero que vamos a abrir según campo de formato
                        sDirOrigen = objGlobal.funciones.OpenDialogFolder("Elige Directorio de trabajo")
                    End If

                    CType(oForm.Items.Item("txtDIR").Specific, SAPbouiCOM.EditText).Value = sDirOrigen
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
