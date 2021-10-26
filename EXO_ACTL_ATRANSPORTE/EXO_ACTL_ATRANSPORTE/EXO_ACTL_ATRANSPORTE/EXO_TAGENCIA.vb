Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_TAGENCIA
    Inherits EXO_UIAPI.EXO_DLLBase

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        If actualizar Then
            cargaCampos()
        End If
        cargamenu()
    End Sub
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_TAGENCIA.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_TAGENCIA", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
        Try
            If objGlobal.SBOApp.Menus.Exists("EXO-MnLPAT") = True Then
                Path = objGlobal.refDi.OGEN.pathGeneral & "\02.Menus"  'objGlobal.compañia.conexionSAP.path & "\02.Menus"
                If Path <> "" Then
                    If IO.File.Exists(Path & "\MnLPAT.png") = True Then
                        objGlobal.SBOApp.Menus.Item("EXO-MnLPAT").Image = Path & "\MnLPAT.png"
                    End If
                End If
            End If
        Catch ex As Exception
            objGlobal.SBOApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

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
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnDTAG"
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_TAGENCIA")
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
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_TAGENCIA"
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
                        Case "UDO_FT_EXO_TAGENCIA"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    'If EventHandler_Validate_Before(infoEvento) = False Then
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_TAGENCIA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_TAGENCIA"

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
    Private Function EventHandler_Choose_FromList_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_Choose_FromList_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_3" Then 'Tarifa

                Dim sTTarifa As String = ""
                If CType(oForm.Items.Item("13_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                    sTTarifa = CType(oForm.Items.Item("13_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value
                Else
                    sTTarifa = ""
                End If

                oCFLEvento = CType(pVal, IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions

                oCond = oConds.Add
                oCond.Alias = "U_EXO_TTARIFA"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = sTTarifa

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
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCond, Object))
        End Try
    End Function
    'Private Function EventHandler_Validate_Before(ByRef pVal As ItemEvent) As Boolean
    '    Dim oForm As SAPbouiCOM.Form = Nothing
    '    Dim dFechaD As Date = New Date(Now.Year, Now.Month, Now.Day)
    '    Dim dFechaH As Date = New Date(Now.Year, Now.Month, Now.Day)
    '    EventHandler_Validate_Before = False

    '    Try
    '        oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

    '        If pVal.ItemUID = "0_U_G" Then
    '            If pVal.ColUID = "C_0_1" Or pVal.ColUID = "C_0_2" Then
    '                Dim sFechaD As String = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value
    '                If sFechaD.Trim <> "" Then
    '                    dFechaD = New Date(CType(Left(sFechaD, 4).ToString, Integer), CType(Mid(sFechaD, 5, 2).ToString, Integer), CType(Right(sFechaD, 2).ToString, Integer))
    '                End If
    '                Dim sFechaH As String = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value
    '                If sFechaH.Trim <> "" Then
    '                    dFechaH = New Date(CType(Left(sFechaH, 4).ToString, Integer), CType(Mid(sFechaH, 5, 2).ToString, Integer), CType(Right(sFechaH, 2).ToString, Integer))
    '                End If

    '            End If
    '            End If

    '        EventHandler_Validate_Before = True

    '    Catch exCOM As System.Runtime.InteropServices.COMException
    '        Throw exCOM
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        EXO_CleanCOM.CLiberaCOM.Form(oForm)
    '    End Try
    'End Function
End Class
