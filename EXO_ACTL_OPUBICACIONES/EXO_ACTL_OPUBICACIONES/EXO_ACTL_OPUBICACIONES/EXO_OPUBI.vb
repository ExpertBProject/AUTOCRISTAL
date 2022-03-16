Imports SAPbouiCOM
Public Class EXO_OPUBI
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

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UTs_EXO_TMPOPUBI.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UTs_EXO_TMPOPUBI", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "EXO-MnOUB"
                        If CargarFormOPUBI() = False Then
                            Exit Function
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
    Public Function CargarFormOPUBI() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing

        CargarFormOPUBI = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_OPUBI.srf")
            oFP.XmlData = oFP.XmlData.Replace("modality=""0""", "modality=""1""")
            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)

            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try

            CargarFormOPUBI = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_OPUBI"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

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
                        Case "EXO_OPUBI"
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
                        Case "EXO_OPUBI"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        Return False
                                    End If
                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_OPUBI"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
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
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oDataTable As SAPbouiCOM.DataTable = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_Choose_FromList_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing

                Return True
            End If

            oCFLEvento = CType(pVal, IChooseFromListEvent)

            oDataTable = oCFLEvento.SelectedObjects
            If Not oDataTable Is Nothing Then
                Select Case oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).ObjectType
                    Case "64"
                        Try
                            Dim sCodALM As String = oDataTable.GetValue("WhsCode", 0).ToString
                            Dim sNombre As String = oDataTable.GetValue("WhsName", 0).ToString

                            oForm.DataSources.UserDataSources.Item("UDALM").Value = sCodALM
                            oForm.DataSources.UserDataSources.Item("UDADES").Value = sNombre
                        Catch ex As Exception
                            CType(oForm.Items.Item("txtALM").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("WhsCode", 0).ToString
                            CType(oForm.Items.Item("txtDES").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("WhsName", 0).ToString
                        End Try

                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End Select
            End If

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.FormDatatable(oDataTable)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnCARGAR"  'Cargamos datos del grid
                    Cargar_Grid(oForm)
                Case "btnGen" 'gen Solicitud de traslado
                    Gen_Sol_Traslado(oForm)
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

    Private Sub Cargar_Grid(ByRef oform As SAPbouiCOM.Form)
#Region "Variables"
        Dim sSQL As String = ""
        Dim iDoc As Integer = 0 'Contador de Code de documentos
#End Region
        Try
            oform.Freeze(True)

            Limpiar_Grid(oform)
#Region "Cargar datos tabla temporal"
            'Tengo que buscar en la tabla el último numero de documento
            iDoc = objGlobal.refDi.SQL.sqlNumericaB1("SELECT isnull(MAX(cast(""Code"" as int)),0) FROM ""@EXO_TMPOPUBI"" ")

            sSQL = " ""U_EXO_USUARIO""='" & objGlobal.compañia.UserName.ToString & "'  "
            objGlobal.refDi.SQL.sqlStringB1(sSQL)
#End Region
#Region "Cargar Datos Grid"
            objGlobal.SBOApp.StatusBar.SetText("Cargando en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' ""Sel."", ""U_EXO_ITEMCODE"" ""Cod. Artículo"", ""U_EXO_ITEMNAME"" ""Descripción"",""U_EXO_CANT"" ""Cantidad"", ""U_EXO_UBACT"" ""Ubicación actual"", "
            sSQL &= " ""U_EXO_ZONAACT"" as ""Zona almacén de Rotación actual"", ""U_EXO_CLAACT"" as ""Clasificación actual de Rotación"" , ""U_EXO_TRASLADO"" as ""Traslado"", "
            sSQL &= " ""U_EXO_UBIDES"" ""Ubicación destino"" "
            sSQL &= " From ""@EXO_TMPOPUBI"" "
            sSQL &= " WHERE ""U_EXO_USUARIO""='" & objGlobal.compañia.UserName.ToString & "' "
            sSQL &= " ORDER BY ""Code"", ""LineId"" "
            'Cargamos grid
            oform.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oform)
#End Region

            oform.Freeze(False)
            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso de carga.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.SBOApp.MessageBox("Fin del proceso de carga.")
        Catch exCOM As System.Runtime.InteropServices.COMException
            oform.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oform.Freeze(False)
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub Limpiar_Grid(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Try
            oForm.Freeze(True)
            'Limpiamos grid
            'Borrar tablas temporales por usuario activo
            sSQL = "DELETE FROM ""@EXO_TMPOPUBI"" where ""U_EXO_USUARIO""='" & objGlobal.compañia.UserName.ToString & "'  "
            objGlobal.refDi.SQL.sqlUpdB1(sSQL)

            'Ahora cargamos el Grid con los datos guardados
            objGlobal.SBOApp.StatusBar.SetText("Cargando en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' ""Sel."", ""U_EXO_ITEMCODE"" ""Cod. Artículo"", ""U_EXO_ITEMNAME"" ""Descripción"",""U_EXO_CANT"" ""Cantidad"", ""U_EXO_UBACT"" ""Ubicación actual"", "
            sSQL &= " ""U_EXO_ZONAACT"" as ""Zona almacén de Rotación actual"", ""U_EXO_CLAACT"" as ""Clasificación actual de Rotación"" , ""U_EXO_TRASLADO"" as ""Traslado"", "
            sSQL &= " ""U_EXO_UBIDES"" ""Ubicación destino"" "
            sSQL &= " From ""@EXO_TMPOPUBI"" "
            sSQL &= " WHERE ""U_EXO_USUARIO""='" & objGlobal.compañia.UserName.ToString & "' "
            sSQL &= " ORDER BY ""Code"", ""LineId"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            'For i = 1 To 5
            '    CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            '    oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
            '    oColumnTxt.Editable = False
            '    If i = 2 Then
            '        oColumnTxt.LinkedObjectType = "112"
            '    End If
            'Next
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Private Sub Gen_Sol_Traslado(ByRef oform As SAPbouiCOM.Form)

        Try
            oform.Freeze(True)
            If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere generar la Sol. de traslado con los registros seleccionados?", 1, "Sí", "No") = 1 Then
                If ComprobarDOC(oform, "DT_DOC") = True Then
                    oform.Items.Item("btnGen").Enabled = False
                    'Generamos facturas
                    objGlobal.SBOApp.StatusBar.SetText("Creando Documento... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oform.Freeze(True)
                    If CrearDocumento(oform, "DT_DOC", "FACTURA") = False Then
                        Exit Sub
                    End If
                    oform.Freeze(False)
                    objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objGlobal.SBOApp.MessageBox("Fin del Proceso." & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                    oform.Items.Item("btnGen").Enabled = True
                End If
            End If


            oform.Freeze(False)
            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso de carga.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.SBOApp.MessageBox("Fin del proceso de carga.")
        Catch exCOM As System.Runtime.InteropServices.COMException
            oform.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oform.Freeze(False)
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Function ComprobarDOC(ByRef oForm As SAPbouiCOM.Form, ByVal sFra As String) As Boolean
        Dim bLineasSel As Boolean = False

        ComprobarDOC = False

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sFra).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sFra).GetValue("Sel", i).ToString = "Y" Then
                    bLineasSel = True
                    Exit For
                End If
            Next

            If bLineasSel = False Then
                objGlobal.SBOApp.MessageBox("Debe seleccionar al menos una línea.")
                Exit Function
            End If

            ComprobarDOC = bLineasSel

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function CrearDocumento(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByVal sTDoc As String) As Boolean
        CrearDocumento = False
#Region "Variables"
        Dim oDoc As SAPbobsCOM.Documents = Nothing
        Dim sExiste As String = "" ' Para comprobar si existen los datos
        Dim sErrorDes As String = ""
        Dim sDocAdd As String = ""
        Dim sMensaje As String = ""
#End Region

        Try
            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            objGlobal.compañia.StartTransaction()
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    oForm.DataSources.DataTables.Item(sData).SetValue("Estado", i, "WARNING")
                    sMensaje = "Falta definir cómo crear el documento"
                    oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", i, sMensaje)
                    'grabar el documento
                    'If oDoc.Add() <> 0 Then 'Si ocurre un error en la grabación entra
                    '    sErrorDes = objGlobal.compañia.GetLastErrorCode & " / " & objGlobal.compañia.GetLastErrorDescription
                    '    objGlobal.SBOApp.StatusBar.SetText(sErrorDes, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    oForm.DataSources.DataTables.Item(sData).SetValue("Estado", i, "ERROR")
                    '    oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", i, sErrorDes)
                    '    oForm.DataSources.DataTables.Item(sData).SetValue("DocEntry", i, "")
                    'Else
                    '    sDocAdd = objGlobal.compañia.GetNewObjectKey() 'Recoge el último documento creado
                    '    oForm.DataSources.DataTables.Item(sData).SetValue("DocEntry", i, sDocAdd)
                    '    'Buscamos el documento para crear un mensaje
                    '    sDocAdd = EXO_GLOBALES.GetValueDB(oCompany, """" & sTabla & """", """DocNum""", """DocEntry""=" & sDocAdd)
                    '    If sModo = "F" Then
                    '        sModo = ""
                    '    Else
                    '        sModo = " borrador "
                    '    End If
                    '    oForm.DataSources.DataTables.Item(sData).SetValue("Estado", i, "OK")
                    '    oForm.DataSources.DataTables.Item(sData).SetValue("Nº Documento", i, sDocAdd)
                    '    Select Case sTipoFac
                    '        Case "13" 'Factura de ventas
                    '            sMensaje = "(EXO) - Ha sido creada la factura " & sModo & " de ventas Nº" & sDocAdd
                    '        Case "14" 'Abono de ventas
                    '            sMensaje = "(EXO) - Ha sido creado el abono " & sModo & " de ventas Nº" & sDocAdd
                    '        Case "18" 'Factura de compras
                    '            sMensaje = "(EXO) - Ha sido creada la factura " & sModo & " de compras Nº" & sDocAdd
                    '        Case "19" 'Abono de compras
                    '            sMensaje = "(EXO) - Ha sido creado el abono " & sModo & " de compras Nº" & sDocAdd
                    '        Case "22" 'Pedido de compras
                    '            sMensaje = "(EXO) - Ha sido creado el pedido " & sModo & " de compras Nº" & sDocAdd
                    '    End Select
                    '    oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", i, sMensaje)
                    '    objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'End If
                End If
            Next

            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If

            CrearDocumento = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDoc, Object))
        End Try
    End Function
End Class
