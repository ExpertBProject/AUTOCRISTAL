Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_DTARIFAS
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

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_DTARIFAS.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_DTARIFAS", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
        Dim sSQL As String = "" : Dim sCode As String = ""
        Try
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnDTar"
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_DTARIFAS")
                    Case "1282"
                        oForm = objGlobal.SBOApp.Forms.ActiveForm()
                        If oForm.Visible = True Then
#Region "Buscamos el número siguiente para poner al code"
                            Try
                                If oForm.TypeEx = "UDO_FT_EXO_DTARIFAS" And oForm.Visible = True Then
                                    sSQL = "SELECT ifnull(MAX(CAST(""Code"" as INT)),'0')+1 ""CODIGO"" FROM ""@EXO_DTARIFAS"" "
                                    sCode = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                    CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value = sCode
                                End If

                            Catch ex As Exception
                                Throw ex
                            End Try
#End Region
                        End If
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
                        Case "UDO_FT_EXO_DTARIFAS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_after(infoEvento) = False Then
                                        Return False
                                    End If
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
                        Case "UDO_FT_EXO_DTARIFAS"

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
                        Case "UDO_FT_EXO_DTARIFAS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_FORM_VISIBLE_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_DTARIFAS"

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
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
#Region "Varibales"
        Dim sArchivo As String = objGlobal.refDi.OGEN.pathGeneral & "\08.Historico\DOC_CARGADOS\" & objGlobal.compañia.CompanyDB & "\PAGTRANSPORTE\"
        Dim sTipoArchivo As String = "" : Dim sNomFICH As String = ""
        Dim sArchivoOrigen As String = ""
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sCodigo As String = "" : Dim sTTarifa As String = ""
        Dim sMensaje As String = ""
#End Region

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            'Comprobamos que exista el directorio y sino, lo creamos
            sCodigo = CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
            If CType(oForm.Items.Item("13_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                sTTarifa = CType(oForm.Items.Item("13_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value
            Else
                sTTarifa = ""
            End If
            Select Case pVal.ItemUID
                Case "btn_Fich"
                    If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere agregar el fichero en la lista activa?", 1, "Sí", "No") = 1 Then
                        'Tiene que tener código para cargar
                        If sCodigo <> "" And sTTarifa <> "" Then
                            'Capturamos el fichero
#Region "Coger y leer fichero"
                            sTipoArchivo = "Ficheros CSV|*.csv|Texto|*.txt"
                            'Tenemos que controlar que es cliente o web
                            If objGlobal.SBOApp.ClientType = SAPbouiCOM.BoClientType.ct_Browser Then
                                sArchivoOrigen = objGlobal.SBOApp.GetFileFromBrowser() 'Modificar
                            Else
                                'Controlar el tipo de fichero que vamos a abrir según campo de formato
                                sArchivoOrigen = objGlobal.funciones.OpenDialogFiles("Abrir archivo como", sTipoArchivo)
                            End If

                            If Len(sArchivoOrigen) = 0 Then
                                CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = ""
                                objGlobal.SBOApp.MessageBox("Debe indicar un archivo a importar.")
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Debe indicar un archivo a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                oForm.Items.Item("btn_Carga").Enabled = False
                                Exit Function
                            Else
                                CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = sArchivoOrigen
                                sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                                sArchivo = sArchivo & sNomFICH
                                'Hacemos copia de seguridad para tratarlo
                                Copia_Seguridad(sArchivoOrigen, sArchivo)
#Region "Tarifa de Agencia de transporte"
                                'Ahora abrimos el fichero para tratarlo
                                TratarFichero(sArchivo, oForm)
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Fin de la Actualización...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
                            End If
#End Region
                        Else
                            sMensaje = "Introduzca un código y un tipo de tarifa primero. Revise los datos."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                        End If
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se ha cancelado la importación de precios de la Agencia de transporte activa.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Function CargarFormImp(ByVal sTTarifa As String, ByRef oFormDTarifa As SAPbouiCOM.Form) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing

        CargarFormImp = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_PATIMP.srf")
            'oFP.XmlData = oFP.XmlData.Replace("modality=""0""", "modality=""1""")
            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
                EXO_PATIMP._oformTAgencia = oFormDTarifa
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try

            CargarFormImp = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_FORM_VISIBLE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim sCode As String = ""
        EventHandler_FORM_VISIBLE_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                oForm.Freeze(True)

                objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                CargarCombos(objGlobal, oForm)
                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_5").RightJustified = True
                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_6").RightJustified = True
                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_7").RightJustified = True

                If oForm.Mode = BoFormMode.fm_ADD_MODE Then
#Region "Buscamos el número siguiente para poner al code"
                    Try
                        sSQL = "SELECT ifnull(MAX(CAST(""Code"" as INT)),'0')+1 ""CODIGO"" FROM ""@EXO_DTARIFAS"" "
                        sCode = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                        CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value = sCode
                    Catch ex As Exception
                        Throw ex
                    End Try
#End Region
                End If
            End If


            EventHandler_FORM_VISIBLE_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Sub CargarCombos(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Try
            sSQL = "SELECT ""Code"" ""Código"", ""Name"" ""País"" FROM OCRY ORDER BY ""Name"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ValidValues, sSQL)

            sSQL = "SELECT ""Code"" ""Código"", ""Name"" ""Localidad"" FROM OCST WHERE ""Country""='ES' ORDER BY ""Name"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").ValidValues, sSQL)

            sSQL = "SELECT * FROM("
            sSQL &= " SELECT '-' ""Código"", CAST(' ' as varchar(100)) ""Tipo"" FROM DUMMY  "
            sSQL &= " UNION ALL "
            sSQL &= "SELECT CAST(""PkgCode"" as varchar(100)) ""Código"", ""PkgType"" ""Tipo"" FROM OPKG  "
            sSQL &= " ) T ORDER BY ""Tipo"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").ValidValues, sSQL)

            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Editable = True
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_5").Editable = False
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_6").Editable = False

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function EventHandler_COMBO_SELECT_after(ByRef pval As ItemEvent) As Boolean
        EventHandler_COMBO_SELECT_after = False
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sTBulto As String = ""

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pval.FormUID)
            Select Case pval.ItemUID
                Case "13_U_Cb"
                    If CType(oForm.Items.Item("13_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sTBulto = CType(oForm.Items.Item("13_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value
                        Select Case sTBulto
                            Case "B" : Deshabilitar_TBULTO(oForm)
                            Case "V" : Habilitar_TBULTO(oForm)

                        End Select
                    End If

            End Select

            EventHandler_COMBO_SELECT_after = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Sub Deshabilitar_TBULTO(ByRef oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            For i = 1 To CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).RowCount
                If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Cells.Item(i).Specific, SAPbouiCOM.ComboBox).Selected Is Nothing Then
                    CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Cells.Item(i).Specific, SAPbouiCOM.ComboBox).Select("-", BoSearchKey.psk_ByValue)
                End If
                CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_5").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value = "0.00"
                CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_6").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value = "0.00"
            Next
            CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Active = True
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Editable = True
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_5").Editable = False
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_6").Editable = False
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub Habilitar_TBULTO(ByRef oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            For i = 1 To CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).RowCount
                CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Cells.Item(i).Specific, SAPbouiCOM.ComboBox).Select("-", BoSearchKey.psk_ByValue)
            Next
            CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Active = True
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Editable = False
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_5").Editable = True
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_6").Editable = True
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Public Overrides Function SBOApp_FormDataEvent(infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument
        Dim sDocEntry As String = ""

        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_DTARIFAS"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_DTARIFAS"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If EventHandler_FORM_DATA_LOAD(oForm) = False Then
                                    Return False
                                End If
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
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            oXml = Nothing
        End Try
    End Function
    Private Function EventHandler_FORM_DATA_LOAD(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sTBulto As String = ""
        EventHandler_FORM_DATA_LOAD = False
        Try
            If oForm.Visible = True Then
                If CType(oForm.Items.Item("13_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                    sTBulto = CType(oForm.Items.Item("13_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value
                    Select Case sTBulto
                        Case "B" : Deshabilitar_TBULTO(oForm)
                        Case "V" : Habilitar_TBULTO(oForm)

                    End Select
                End If
                CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = ""
                CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Active = True
            End If

            EventHandler_FORM_DATA_LOAD = True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub Copia_Seguridad(ByVal sArchivoOrigen As String, ByVal sArchivo As String)
        'Comprobamos el directorio de copia que exista
        Dim sPath As String = ""
        sPath = IO.Path.GetDirectoryName(sArchivo)
        If IO.Directory.Exists(sPath) = False Then
            IO.Directory.CreateDirectory(sPath)
        End If
        If IO.File.Exists(sArchivo) = True Then
            IO.File.Delete(sArchivo)
        End If
        'Subimos el archivo
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Comienza la Copia de seguridad del fichero - " & sArchivoOrigen & " -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        If objGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
            Dim fs As FileStream = New FileStream(sArchivoOrigen, FileMode.Open, FileAccess.Read)
            Dim b(CInt(fs.Length() - 1)) As Byte
            fs.Read(b, 0, b.Length)
            fs.Close()
            Dim fs2 As New System.IO.FileStream(sArchivo, IO.FileMode.Create, IO.FileAccess.Write)
            fs2.Write(b, 0, b.Length)
            fs2.Close()
        Else
            My.Computer.FileSystem.CopyFile(sArchivoOrigen, sArchivo)
        End If
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Copia de Seguridad realizada Correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
    Private Sub TratarFichero(ByVal sArchivo As String, ByRef oForm As SAPbouiCOM.Form)
        Dim myStream As StreamReader = Nothing
        Dim Reader As XmlTextReader = New XmlTextReader(myStream)
        Dim sSQL As String = ""
        Dim sExiste As String = "" ' Para comprobar si existen los datos
        Dim sDelimitador As String = "2"
        Try
            objGlobal.SBOApp.StatusBar.SetText("Cargando datos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            EXO_GLOBALES.TratarFichero_TXT(sArchivo, sDelimitador, oForm, objGlobal.compañia, objGlobal.SBOApp, objGlobal)


            objGlobal.SBOApp.StatusBar.SetText("Se terminado de leer el fichero. Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.SBOApp.MessageBox("Se terminado de leer el fichero. Fin del proceso")
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub
End Class
