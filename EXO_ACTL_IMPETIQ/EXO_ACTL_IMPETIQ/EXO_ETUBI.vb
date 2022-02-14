Imports SAPbouiCOM
Public Class EXO_ETUBI
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim sMensaje As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.ActiveForm
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnETUB"
                        If CargarForm() = False Then
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
    Public Function CargarForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)

        CargarForm = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_ETUBI.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un Error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try
            CargaCombo_Ubicaciones(oForm)
            CType(oForm.Items.Item("2").Specific, SAPbouiCOM.Button).Caption = "Imprimir"
            CargarForm = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub CargaCombo_Ubicaciones(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim sAlmacen As String = ""
        Try
            sSQL = ""
            sAlmacen = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            If sAlmacen <> "" Then
                sSQL = "SELECT ""BinCode"" ""Lote"", ""BinCode"" FROM OBIN WHERE ""WhsCode""='" & sAlmacen & "' "
            Else
                sSQL = "SELECT ""BinCode"" ""Lote"", ""BinCode"" FROM OBIN "
            End If
            oForm.Items.Item("cbDesde").DisplayDesc = False
            oForm.Items.Item("cbHasta").DisplayDesc = False
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbDesde").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbHasta").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)


        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
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
                        Case "EXO_ETUBI"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_ETUBI"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        Return False
                                    End If
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
                        Case "EXO_ETUBI"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_ETUBI"
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
    Private Function EventHandler_COMBO_SELECT_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_COMBO_SELECT_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            CType(oForm.Items.Item("2").Specific, SAPbouiCOM.Button).Caption = "Imprimir"

            EventHandler_COMBO_SELECT_After = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_Before(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim sDesde As String = "" : Dim sHasta As String = ""
        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "2"
                    If CType(oForm.Items.Item("cbDesde").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sDesde = CType(oForm.Items.Item("cbDesde").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    Else
                        sDesde = ""
                    End If

                    If CType(oForm.Items.Item("cbHasta").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sHasta = CType(oForm.Items.Item("cbHasta").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    Else
                        sHasta = ""
                    End If
                    Menu_Imprimir_Etiquetas_UBI(sDesde, sHasta)
            End Select

            EventHandler_ItemPressed_Before = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function

    Private Function Menu_Imprimir_Etiquetas_UBI(ByVal sDesde As String, ByVal sHasta As String) As Boolean
#Region "Variables"
        Dim rutaCrystal As String = "" : Dim sRutaFicheros As String = objGlobal.pathHistorico : Dim sReport As String = "" : Dim sTipoImp As String = ""
        Dim sCrystal As String = "Et_Ubicaciones.rpt"
#End Region
        Menu_Imprimir_Etiquetas_UBI = False

        Try
            rutaCrystal = objGlobal.path & "\05.Rpt\ETIQUETAS\"

            sTipoImp = "IMP"
            'Imprimimos la etiqueta
            EXO_GLOBALES.GenerarImpCrystal_Rangos(objGlobal, rutaCrystal, sCrystal, sDesde, sHasta, sRutaFicheros, sReport, sTipoImp, objGlobal.compañia.UserSignature.ToString)




            Menu_Imprimir_Etiquetas_UBI = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
End Class
