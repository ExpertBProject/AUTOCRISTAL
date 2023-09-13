Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_RSTOCK
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
                        Case "EXO_RSTOCK"
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
                        Case "EXO_RSTOCK"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                    If EventHandler_MATRIX_LINK_PRESSED(infoEvento) = False Then
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_RSTOCK"
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
                        Case "EXO_RSTOCK"
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
    Private Function EventHandler_MATRIX_LINK_PRESSED(ByVal pVal As ItemEvent) As Boolean

        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim sTipo As String = ""
        EventHandler_MATRIX_LINK_PRESSED = False

        Try
            oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                Return True
            End If

            oColumnTxt = CType(CType(oForm.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(1), SAPbouiCOM.EditTextColumn)
            Dim grid = CType(oForm.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid)
            sTipo = grid.DataTable.GetValue("TIPO", pVal.Row).ToString
            Dim nroInt = grid.DataTable.GetValue("Nº INTERNO", grid.GetDataTableRowIndex(pVal.Row)).ToString
            Select Case pVal.ColUID
                Case "Documento"
                    Select Case sTipo
                        Case "17" 'Albaranes de ventas
                            objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_Order, "", nroInt)
                        Case "1250000001" 'Sol de traslado
                            objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_StockTransfersRequest, "", nroInt)
                        Case "234000032" 'Devolución de proveedor
                            objGlobal.SBOApp.OpenForm("234000032", "", nroInt)
                    End Select
                Case "ARTÍCULO"
                    CType(CType(oForm.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item("ARTÍCULO"), SAPbouiCOM.EditTextColumn).LinkedObjectType = "4"

            End Select

            Return True

            EventHandler_MATRIX_LINK_PRESSED = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
End Class