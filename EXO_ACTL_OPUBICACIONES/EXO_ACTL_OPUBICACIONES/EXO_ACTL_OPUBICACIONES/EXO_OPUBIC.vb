Imports SAPbouiCOM
Public Class EXO_OPUBIC
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

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UTs_EXO_TMPOPUBIC.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UTs_EXO_TMPOPUBIC", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                    Case "EXO-MnOUBC"
                        If CargarFormOPUBIC() = False Then
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
    Public Function CargarFormOPUBIC() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing

        CargarFormOPUBIC = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_OPUBIC.srf")
            'oFP.XmlData = oFP.XmlData.Replace("modality=""0""", "modality=""1""")
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
            CType(oForm.Items.Item("btnGen").Specific, SAPbouiCOM.Button).Item.Enabled = False
            CargarFormOPUBIC = True

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
                        Case "EXO_OPUBIC"
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
                        Case "EXO_OPUBIC"
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
                        Case "EXO_OPUBIC"
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
                        Case "EXO_OPUBIC"
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
        Dim iDocLin As Integer = 0 'Contador de Lineas de documentos
        Dim sAlmacen As String = ""
        Dim sFechaD As String = "" : Dim dFechaD As Date = Now.Date.AddMonths(-3)
        Dim sFechaH As String = ""
        Dim dtDatos As System.Data.DataTable = Nothing
        Dim sUbOrigen As String = "" : Dim sUBOrigenTotal As String = "'' "
#End Region
        Try
            objGlobal.SBOApp.StatusBar.SetText("Buscando datos ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oform.Freeze(True)
            sAlmacen = oform.DataSources.UserDataSources.Item("UDALM").Value
            sFechaD = dFechaD.Year.ToString("0000") & dFechaD.Month.ToString("00") & dFechaD.Day.ToString("00")
            sFechaH = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00")
            Limpiar_Grid(oform)
#Region "Cargar datos tabla temporal"
            'Tengo que buscar en la tabla el último numero de documento

            sSQL = "SELECT I.""ItemCode"",IFNULL(I.""ItemName"",'') ""ItemName"" ,IFNULL(S.""OnHand"",0) ""OnHand"", W.""U_EXO_CMIN"" ""CMIN"", W.""U_EXO_CNEC"" ""CNEC"",
                    IFNULL(B.""BinCode"",'') ""UBSTANDARD"", ' ' ""UBO"",  ' ' ""DESTINO""
                    FROM ""OITM"" I
                    INNER JOIN ""OITW"" W ON I.""ItemCode""=W.""ItemCode"" and W.""WhsCode""='" & sAlmacen & "'
                    LEFT JOIN ""OBIN"" B ON W.""DftBinAbs""=B.""AbsEntry"" and W.""WhsCode""=B.""WhsCode""
                    LEFT JOIN (SELECT ""ItemCode"",""WhsCode"",	""BinAbs"", SUM(""OnHandQty"") ""OnHand""
                                FROM OBBQ GROUP BY ""ItemCode"",""WhsCode"",""BinAbs"")S ON S.""ItemCode""=W.""ItemCode"" and S.""WhsCode""=B.""WhsCode"" and S.""BinAbs""=B.""AbsEntry""
                    WHERE IFNULL(B.""BinCode"",'')<>'' and IFNULL(S.""OnHand"",0)<W.""U_EXO_CMIN""
                    ORDER BY I.""ItemName"""
            dtDatos = New System.Data.DataTable
            dtDatos = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
            If dtDatos.Rows.Count > 0 Then
                iDoc = objGlobal.refDi.SQL.sqlNumericaB1("SELECT ifnull(MAX(cast(""Code"" as int)),0)+1 FROM ""@EXO_TMPOPUBIC"" ")
                For Each MiDataRow As DataRow In dtDatos.Rows
                    iDocLin = objGlobal.refDi.SQL.sqlNumericaB1("SELECT ifnull(MAX(cast(""LineId"" as int)),0)+1 FROM ""@EXO_TMPOPUBIC"" WHERE ""Code""='" & iDoc.ToString & "'")
                    sUBOrigenTotal &= ", '" & MiDataRow("UBSTANDARD").ToString & "' "
                    Dim dCantidad As Double = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, MiDataRow("CMIN").ToString) - EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, MiDataRow("OnHand").ToString)
                    sSQL = "SELECT  ""BinCode"" FROM OBIN B 
                                                LEFT JOIN (SELECT ""ItemCode"",""WhsCode"",	""BinAbs"", SUM(""OnHandQty"") ""OnHand""
                                                            FROM OBBQ GROUP BY ""ItemCode"",""WhsCode"",""BinAbs"")S ON S.""ItemCode""='" & MiDataRow("ItemCode").ToString & "'
                                                                and S.""WhsCode""=B.""WhsCode"" and S.""BinAbs""=B.""AbsEntry"" 
                                                WHERE B.""WhsCode""='" & sAlmacen & "' and B.""Attr2Val"" ='Picking'
                                                And IFNULL(S.""OnHand"",0)>= " & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, MiDataRow("CMIN").ToString, EXO_GLOBALES.FuenteInformacion.Otros) &
                                              " And ""BinCode"" Not In (" & sUBOrigenTotal & ")"
                    sUbOrigen = objGlobal.refDi.SQL.sqlStringB1(sSQL)

                    sUBOrigenTotal &= ", '" & sUbOrigen & "' "

                    sSQL = "INSERT INTO ""@EXO_TMPOPUBIC"" values('" & iDoc.ToString & "'," & iDocLin.ToString & ",'EXO_TMPOPUBIC',0,'" & MiDataRow("ItemCode").ToString & "',
                            '" & MiDataRow("ItemName").ToString.Replace("'", "") & " '," & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, MiDataRow("OnHand").ToString, EXO_GLOBALES.FuenteInformacion.Otros) &
                            ", " & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, MiDataRow("CMIN").ToString, EXO_GLOBALES.FuenteInformacion.Otros) &
                            ", " & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, MiDataRow("CNEC").ToString, EXO_GLOBALES.FuenteInformacion.Otros) & ",'" & MiDataRow("UBSTANDARD").ToString &
                            "','" & sUbOrigen & "'," & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dCantidad, EXO_GLOBALES.FuenteInformacion.Otros) &
                            ", '" & MiDataRow("UBSTANDARD").ToString & "','" & objGlobal.compañia.UserName.ToString & "')"
                    objGlobal.refDi.SQL.sqlStringB1(sSQL)
                Next
            End If

#End Region
#Region "Cargar Datos Grid"
            objGlobal.SBOApp.StatusBar.SetText("Cargando en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' ""Sel"",  ""U_EXO_ITEMCODE"" ""Cod. Artículo"", ""U_EXO_ITEMNAME"" ""Descripción"",""U_EXO_CANT"" ""Cantidad"", ""U_EXO_CMIN"" ""Cant. Min."", "
            sSQL &= " ""U_EXO_CNEC"" ""Cant. Nec. "", ""U_EXO_UBI"" ""Ub. Std."", ""U_EXO_UBIO"" ""Ub. Origen"", ""U_EXO_TRASLADO"" ""Traslado"", ""U_EXO_UBID"" ""Ub. Destino"" "
            sSQL &= " From ""@EXO_TMPOPUBIC"" "
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
            CType(oform.Items.Item("btnGen").Specific, SAPbouiCOM.Button).Item.Enabled = True
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub Limpiar_Grid(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Try
            oForm.Freeze(True)
            'Limpiamos grid
            'Borrar tablas temporales por usuario activo
            sSQL = "DELETE FROM ""@EXO_TMPOPUBIC"" where ""U_EXO_USUARIO""='" & objGlobal.compañia.UserName.ToString & "'  "
            objGlobal.refDi.SQL.sqlUpdB1(sSQL)
            oForm.DataSources.UserDataSources.Item("UDMEN").Value = ""
            oForm.DataSources.UserDataSources.Item("UDDE").Value = ""

            'Ahora cargamos el Grid con los datos guardados
            sSQL = "SELECT 'Y' ""Sel"",  ""U_EXO_ITEMCODE"" ""Cod. Artículo"", ""U_EXO_ITEMNAME"" ""Descripción"",""U_EXO_CANT"" ""Cantidad"", ""U_EXO_CMIN"" ""Cant. Min."", "
            sSQL &= " ""U_EXO_CNEC"" ""Cant. Nec. "", ""U_EXO_UBI"" ""Ub. Std."", ""U_EXO_UBIO"" ""Ub. Origen"", ""U_EXO_TRASLADO"" ""Traslado"", ""U_EXO_UBID"" ""Ub. Destino"" "
            sSQL &= " From ""@EXO_TMPOPUBIC"" "
            sSQL &= " WHERE ""U_EXO_USUARIO""='" & objGlobal.compañia.UserName.ToString & "' "
            sSQL &= " ORDER BY ""Code"", ""LineId"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            'FormateaGrid(oForm)
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
        Dim sAlmacen As String = ""
        Try
            sAlmacen = oform.DataSources.UserDataSources.Item("UDALM").Value
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True

            For i = 1 To 9
                Select Case i
                    Case 1
                        CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                    oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                    oColumnTxt.Editable = False
                    oColumnTxt.LinkedObjectType = 4
                    Case 3, 4, 5, 8, 9
                        CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                        oColumnTxt.RightJustified = True
                    Case 7
                        CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = True
                    Case Else
                        CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next
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
                    If CrearDocumento(oform, "DT_DOC") = False Then
                        Exit Sub
                    End If
                    oform.Freeze(False)
                    objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objGlobal.SBOApp.MessageBox("Fin del Proceso." & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                    oform.Items.Item("btnGen").Enabled = True
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("El usuario ha cancelado el proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objGlobal.SBOApp.MessageBox("El usuario ha cancelado el proceso.")
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            oform.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oform.Freeze(False)
            Throw ex
        Finally
            CType(oform.Items.Item("btnGen").Specific, SAPbouiCOM.Button).Item.Enabled = False
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
    Private Function CrearDocumento(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String) As Boolean
        CrearDocumento = False
#Region "Variables"
        Dim sExiste As String = "" ' Para comprobar si existen los datos
        Dim sErrorDes As String = ""
        Dim sDocAdd As String = ""
        Dim sMensaje As String = ""
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing
        Dim sAlmacen As String = ""
        Dim sSQL As String = ""
        Dim sUbStandard As String = ""
#End Region

        Try
            sAlmacen = oForm.DataSources.UserDataSources.Item("UDALM").Value
            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            objGlobal.compañia.StartTransaction()
            oDocStockTransfer = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
            oDocStockTransfer.DocDate = Now.Date
            oDocStockTransfer.DueDate = Now.Date
            oDocStockTransfer.TaxDate = Now.Date
            oDocStockTransfer.FromWarehouse = sAlmacen
            oDocStockTransfer.ToWarehouse = sAlmacen
            oDocStockTransfer.UserFields.Fields.Item("U_EXO_STATUSP").Value = "P"
            oDocStockTransfer.UserFields.Fields.Item("U_EXO_TIPO").Value = "INU"
            oDocStockTransfer.Comments = "Generado automáticamente a través de Optimización de ubicaciones por rotación."
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    If (i > 0) Then
                        oDocStockTransfer.Lines.Add()
                    End If
                    oDocStockTransfer.Lines.ItemCode = oForm.DataSources.DataTables.Item(sData).GetValue("Cod. Artículo", i).ToString
                    oDocStockTransfer.Lines.ItemDescription = oForm.DataSources.DataTables.Item(sData).GetValue("Descripción", i).ToString
                    oDocStockTransfer.Lines.UserFields.Fields.Item("U_EXO_UBI_OR").Value = oForm.DataSources.DataTables.Item(sData).GetValue("Ubicación actual", i).ToString
                    oDocStockTransfer.Lines.UserFields.Fields.Item("U_EXO_UBI_DE").Value = oForm.DataSources.DataTables.Item(sData).GetValue("Ubicación destino", i).ToString
                    oDocStockTransfer.Lines.Quantity = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oForm.DataSources.DataTables.Item(sData).GetValue("Traslado", i).ToString)
                End If
            Next
            ' grabar el documento
            If oDocStockTransfer.Add() <> 0 Then 'Si ocurre un error en la grabación entra
                sErrorDes = objGlobal.compañia.GetLastErrorCode & " / " & objGlobal.compañia.GetLastErrorDescription
                objGlobal.SBOApp.StatusBar.SetText(sErrorDes, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oForm.DataSources.UserDataSources.Item("UDMEN").Value = "ERROR - " & sErrorDes
                oForm.DataSources.UserDataSources.Item("UDDE").Value = ""
            Else
                sDocAdd = objGlobal.compañia.GetNewObjectKey() 'Recoge el último documento creado
                oForm.DataSources.UserDataSources.Item("UDDE").Value = sDocAdd
                'Buscamos el documento para crear un mensaje
                sDocAdd = objGlobal.refDi.SQL.sqlStringB1("SELECT ""DocNum"" FROM OWTQ WHERE ""DocEntry""=" & sDocAdd)
                oForm.DataSources.UserDataSources.Item("UDMEN").Value = "OK - Nº Documento creado " & sDocAdd

                objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                If objGlobal.compañia.InTransaction = True Then
                    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If
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

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDocStockTransfer, Object))
        End Try
    End Function
End Class
