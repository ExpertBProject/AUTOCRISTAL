Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_ORDR
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaCampos()
        End If
    End Sub

    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_ORDR.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDFs_ORDR", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                        Case "139"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "139"
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
                        Case "139"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "139"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

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
    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Dim oItem As SAPbouiCOM.Item
        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            oForm.Visible = False
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

#Region "Delegación"
            'Campo de la delegación
            oItem = oForm.Items.Add("cbDele", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.LinkTo = "222"
            oItem.Top = oForm.Items.Item("103").Top
            oItem.Left = oForm.Items.Item("20").Left
            oItem.Height = oForm.Items.Item("20").Height
            oItem.Width = oForm.Items.Item("20").Width
            oItem.Enabled = True
            oItem.DisplayDesc = True
            oItem.FromPane = 0 : oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "ORDR", "U_EXO_DELE")
            oItem = oForm.Items.Add("lblDele", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("cbDele").Top
            oItem.Left = oForm.Items.Item("230").Left
            oItem.Height = oForm.Items.Item("230").Height
            oItem.Width = oForm.Items.Item("230").Width
            oItem.FromPane = 0 : oItem.ToPane = 0
            oItem.LinkTo = "cbDele"
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Delegación"
#End Region
#Region "Agencia de Transporte"
            'Agencia de Transporte
            oItem = oForm.Items.Add("cbATrans", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.LinkTo = "cbDele"
            oItem.Top = oForm.Items.Item("27").Top
            oItem.Left = oForm.Items.Item("cbDele").Left
            oItem.Height = oForm.Items.Item("cbDele").Height
            oItem.Width = oForm.Items.Item("cbDele").Width
            oItem.Enabled = True
            oItem.DisplayDesc = True
            oItem.FromPane = 0 : oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OCRD", "U_EXO_AGENCIA")
            oItem = oForm.Items.Add("lblATrans", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("cbATrans").Top
            oItem.Left = oForm.Items.Item("lblDele").Left
            oItem.Height = oForm.Items.Item("lblDele").Height
            oItem.Width = oForm.Items.Item("lblDele").Width
            oItem.FromPane = 0 : oItem.ToPane = 0
            oItem.LinkTo = "cbATrans"
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Agencia de Transporte"
#End Region

            'Introducimos los valores en los combos
            CargarCombos(objGlobal, oForm)

            oForm.Visible = True

            EventHandler_Form_Load = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            If oForm IsNot Nothing Then oForm.Visible = True

            Throw exCOM
        Catch ex As Exception
            If oForm IsNot Nothing Then oForm.Visible = True

            Throw ex
        Finally
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
    Private Function EventHandler_VALIDATE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sDelegacion As String = "" : Dim sAgencia As String = "" : Dim sIC As String = ""
        Dim sSQL As String = ""
        Dim sTable_Origen As String = ""
        EventHandler_VALIDATE_After = False
        Try
            If oForm.Mode = BoFormMode.fm_UPDATE_MODE Or oForm.Mode = BoFormMode.fm_ADD_MODE Then
                If pVal.ItemUID = "4" Then
                    sTable_Origen = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).DataBind.TableName
                    sIC = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("CardCode", 0).ToString.Trim
                    CargarCombos(objGlobal, oForm)
                    sSQL = "SELECT ""U_EXO_DELE"" FROM ""OCRD"" WHERE ""CardCode""='" & sIC & "'"
                    sDelegacion = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    sSQL = "SELECT ""U_EXO_AGENCIA"" FROM ""OCRD"" WHERE ""CardCode""='" & sIC & "'"
                    sAgencia = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    CType(oForm.Items.Item("cbDele").Specific, SAPbouiCOM.ComboBox).Select(sDelegacion, BoSearchKey.psk_ByValue)
                    CType(oForm.Items.Item("cbATrans").Specific, SAPbouiCOM.ComboBox).Select(sAgencia, BoSearchKey.psk_ByValue)
                End If
            End If
            EventHandler_VALIDATE_After = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "139"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "139"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Visible = True Then
                                    CargarCombos(objGlobal, oForm)
                                End If
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
End Class
