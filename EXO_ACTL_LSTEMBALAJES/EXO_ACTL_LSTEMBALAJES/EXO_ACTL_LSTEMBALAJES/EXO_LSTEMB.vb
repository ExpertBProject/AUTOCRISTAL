Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_LSTEMB
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean

        Dim sSQL As String = ""
        Try
            If infoEvento.BeforeAction = True Then
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnLEmb"
                        If CargarUDO() = False Then
                            Exit Function
                        End If
                    Case "1282"
                        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.ActiveForm
                        Modo_Anadir(oForm)
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
            objGlobal.funcionesUI.cargaFormUdoBD("EXO_LSTEMB")

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
                        Case "UDO_FT_EXO_LSTEMB"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    'If EventHandler_ItemPressed_After(infoEvento) = False Then
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_LSTEMB"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

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
                        Case "UDO_FT_EXO_LSTEMB"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(objGlobal, infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_LSTEMB"
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
    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True And oForm.Mode = BoFormMode.fm_ADD_MODE Then
                Modo_Anadir(oForm)
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
        Dim sSQL As String = ""
        Dim sSerieDef As String = ""
#End Region
        Try
            'Expedición
            sSQL = "SELECT ""TrnspCode"",""TrnspName"" FROM OSHP WHERE ""Active""='Y' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            CType(oform.Items.Item("20_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)
            'Series 
            sSQL = "SELECT ""Series"",""SeriesName"" FROM NNM1 WHERE ""ObjectCode""='EXO_LSTEMB' "
            objGlobal.funcionesUI.cargaCombo(CType(oform.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            oform.Items.Item("4_U_Cb").DisplayDesc = True
            sSQL = " SELECT ""DfltSeries"" FROM ONNM WHERE ""ObjectCode""='EXO_LSTEMB' "
            sSerieDef = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            CType(oform.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(sSerieDef, BoSearchKey.psk_ByValue)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Modo_Anadir(ByRef oForm As SAPbouiCOM.Form)
        Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim sFecha As String = ""
        Try

            Cargar_Combos(oForm)
            'Poner fecha
            sFecha = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
            oForm.DataSources.DBDataSources.Item("@EXO_LSTEMB").SetValue("U_EXO_DOCDATE", 0, sFecha)
            ' CType(oForm.Items.Item("26_U_E").Specific, SAPbouiCOM.EditText).Value = sFecha

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
