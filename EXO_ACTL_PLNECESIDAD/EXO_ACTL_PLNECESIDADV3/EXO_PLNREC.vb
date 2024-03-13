Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_PLNREC
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
                        Case "EXO_PLNREC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(objGlobal, infoEvento) = False Then
                                        GC.Collect()
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
                        Case "EXO_PLNREC"

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
                        Case "EXO_PLNREC"
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
                        Case "EXO_PLNREC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

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
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oDataTable As SAPbouiCOM.DataTable = Nothing
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)

            If pVal.ItemUID = "txtFiltro" Then
                oDataTable = oCFLEvento.SelectedObjects
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.UserDataSources.Item("UDFILTRO").Value = oDataTable.GetValue("Code", 0).ToString
                    Catch ex As Exception
                        CType(oForm.Items.Item("txtFiltro").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("Code", 0).ToString
                    End Try
                End If
            End If
            EventHandler_Choose_FromList_After = True

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

    Private Function EventHandler_ItemPressed_After(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
#End Region
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnA"
                    'Grabamos el filtro
                    'Comprobamos que haya escrito el código y el nombre para poder guardar.
                    If oForm.DataSources.UserDataSources.Item("UDFILTRO").Value <> "" Then
                        'Comprobaremos que exista
                        sSQL = "SELECT ""Code"" FROM ""@EXO_PLNHCO"" WHERE ""Code""='" & oForm.DataSources.UserDataSources.Item("UDFILTRO").Value & "'"
                        Dim sExiste As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                        If sExiste <> "" Then
                            objGlobal.SBOApp.StatusBar.SetText("Recuperando Filtro guardando...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            RecuperarFiltro(oForm)
                            oForm.Close()
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("El código que ha indicado no existe en los filtros guardados. Por favor, introduzca uno existente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If
                        'Si no existe mensaje

                    Else
                        objGlobal.SBOApp.StatusBar.SetText("Para poder recuperar un filtro debe escribir uno guardado. Operación cancelada.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso de recuperación.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub RecuperarFiltro(ByRef oForm As SAPbouiCOM.Form)
#Region "variables"
        Dim oformOrigen As SAPbouiCOM.Form = Nothing
        Dim dtCab As System.Data.DataTable = Nothing
        Dim sCode As String = ""
        Dim sSQL As String = ""
#End Region
        Try
            'Buscamos formulario Orieb para coger los datos.
            oformOrigen = objGlobal.SBOApp.Forms.Item(oForm.DataSources.UserDataSources.Item("UDFRM").Value)
            sCode = oForm.DataSources.UserDataSources.Item("UDFILTRO").Value.ToString
            sSQL = "SELECT ""U_EXO_EXPEDIR"", ""U_EXO_DART"", ""U_EXO_HART"", ""U_EXO_PROV"", ""U_EXO_MGS"", ""U_EXO_TSUMI"", ""U_EXO_WHS"", ""U_EXO_WHS2"" 
                    FROM ""@EXO_PLNHCO"" WHERE ""Code""='" & sCode & "' "
            dtCab = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
            If dtCab.Rows.Count > 0 Then
                oformOrigen.Freeze(True)
#Region "Cabecera"
                For i = 0 To dtCab.Rows.Count - 1
                    oformOrigen.DataSources.UserDataSources.Item("UD_Excl").Value = dtCab.Rows(i)("U_EXO_EXPEDIR").ToString()
                    oformOrigen.DataSources.UserDataSources.Item("UDARTD").Value = dtCab.Rows(i)("U_EXO_DART").ToString()
                    oformOrigen.DataSources.UserDataSources.Item("UDARTH").Value = dtCab.Rows(i)("U_EXO_HART").ToString()
                    oformOrigen.DataSources.UserDataSources.Item("UDPROV").Value = dtCab.Rows(i)("U_EXO_PROV").ToString()
                    oformOrigen.DataSources.UserDataSources.Item("UDPROVD").Value = objGlobal.refDi.SQL.sqlStringB1("SELECT ""CardName"" FROM OCRD WHERE ""CardCode""='" & dtCab.Rows(i)("U_EXO_PROV").ToString() & "'")
                    oformOrigen.DataSources.UserDataSources.Item("UDDIAS").Value = dtCab.Rows(i)("U_EXO_MGS").ToString()
                    oformOrigen.DataSources.UserDataSources.Item("UDTSM").Value = dtCab.Rows(i)("U_EXO_TSUMI").ToString()
                    oformOrigen.DataSources.UserDataSources.Item("UDALM").Value = dtCab.Rows(i)("U_EXO_WHS").ToString()
                    oformOrigen.DataSources.UserDataSources.Item("UDALM2").Value = dtCab.Rows(i)("U_EXO_WHS2").ToString()
                Next
#End Region
#Region "Articulos"
                sSQL = "SELECT Count(""LineId"") FROM ""@EXO_PLNHCO1"" WHERE ""Code""='" & sCode & "' "
                Dim dCuantos As Double = objGlobal.refDi.SQL.sqlNumericaB1(sSQL)
                If dCuantos > 0 Then
                    sSQL = "SELECT ""U_EXO_ITEMCODE"" ""Código"", ""U_EXO_ITEMNAME"" ""Descripción"" FROM ""@EXO_PLNHCO1"" WHERE ""Code""='" & sCode & "' Order BY ""LineId"" "
                Else
                    sSQL = "SELECT '' ""Código"", ' ' ""Descripción"" FROM DUMMY"
                End If

                oformOrigen.DataSources.DataTables.Item("DTART").ExecuteQuery(sSQL)
                EXO_PNEC.FormateaGridART(oformOrigen)
#End Region


#Region "Almacenes"
                sSQL = "SELECT ""U_Sel"" ""Sel"", ""U_EXO_WHSCODE"" ""Cod."",""U_EXO_WHSNAME"" ""Almacén"" FROM ""@EXO_PLNHCO2"" WHERE ""Code""='" & sCode & "' Order BY ""LineId"" "
                oformOrigen.DataSources.DataTables.Item("DTALM").ExecuteQuery(sSQL)
                EXO_PNEC.FormateaGridALM(oformOrigen)
#End Region
#Region "Almacenes2"
                sSQL = "SELECT ""U_EXO_SEL"" ""Sel"", ""U_EXO_WHSCODE"" ""Cod."",""U_EXO_WHSNAME"" ""Almacén"" FROM ""@EXO_PLNHCO7"" WHERE ""Code""='" & sCode & "' Order BY ""LineId"" "
                oformOrigen.DataSources.DataTables.Item("DTALM2").ExecuteQuery(sSQL)
                EXO_PNEC.FormateaGridALM2(oformOrigen)
#End Region
#Region "Clasificación"
                sSQL = "SELECT ""U_EXO_SEL"" ""Sel"", ""U_EXO_CLAS"" ""Clas"" FROM ""@EXO_PLNHCO3"" WHERE ""Code""='" & sCode & "' Order BY ""LineId"" "
                oformOrigen.DataSources.DataTables.Item("DTCLAS").ExecuteQuery(sSQL)
                EXO_PNEC.FormateaGridCLAS(oformOrigen)
#End Region
#Region "Grupos Art."
                sSQL = "SELECT ""U_EXO_SEL"" ""Sel"", ""U_EXO_GRUPO"" ""Cod."", ""U_EXO_DES"" ""Familia"" FROM ""@EXO_PLNHCO4"" WHERE ""Code""='" & sCode & "' Order BY ""LineId"" "
                oformOrigen.DataSources.DataTables.Item("DTGRU").ExecuteQuery(sSQL)
                EXO_PNEC.FormateaGridGRU(oformOrigen)
#End Region
#Region "Lst. de Compras"
                sSQL = "SELECT ""U_EXO_SEL"" ""Sel"", ""U_EXO_LST"" ""LST"", ""U_EXO_NOMBRE"" ""Nombre"" FROM ""@EXO_PLNHCO5"" WHERE ""Code""='" & sCode & "' Order BY ""LineId"" "
                oformOrigen.DataSources.DataTables.Item("DTCOMP").ExecuteQuery(sSQL)
                EXO_PNEC.FormateaGridCOMP(oformOrigen)
#End Region
#Region "Lst. de Ventas"
                sSQL = "SELECT ""U_EXO_SEL"" ""Sel"", ""U_EXO_LST"" ""LST"", ""U_EXO_NOMBRE"" ""Nombre"" FROM ""@EXO_PLNHCO6"" WHERE ""Code""='" & sCode & "' Order BY ""LineId"" "
                oformOrigen.DataSources.DataTables.Item("DTVENT").ExecuteQuery(sSQL)
                EXO_PNEC.FormateaGridVENT(oformOrigen)
#End Region
#Region "Lineas"
                EXO_PNEC.Cargar_Datos(objGlobal, oformOrigen)

                objGlobal.SBOApp.StatusBar.SetText("Recuperando Datos de las líneas...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Dim dtLineas As System.Data.DataTable = New Data.DataTable
                sSQL = "SELECT ""U_EXO_EUROCODE"", ""U_EXO_ORDER"", 
                        ""U_EXO_AL0"",""U_EXO_AL7"", ""U_EXO_AL8"", ""U_EXO_AL14"", ""U_EXO_AL16"", ""U_EXO_ORDE2"", ""U_EXO_PROV"", ""U_EXO_FECHA""
                        FROM ""@EXO_PLNHCO8"" WHERE ""Code""='" & sCode & "' Order BY ""LineId"" "

                dtLineas = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                If dtLineas.Rows.Count > 0 Then
                    Dim sXmlGrid As String = "" : Dim oXmlGrid As New Xml.XmlDocument : Dim oXmlNodesGrid As XmlNodeList = Nothing
                    sXmlGrid = oformOrigen.DataSources.DataTables.Item("DT_DOC").SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All)
                    oXmlGrid.LoadXml(sXmlGrid)
                    Dim oXmlNodeGrid As XmlNode = Nothing
                    For Each Row As DataRow In dtLineas.Rows
                        oXmlNodeGrid = oXmlGrid.SelectSingleNode("/DataTable/Rows/Row/Cells/Cell[./ColumnUid='EUROCODE' and ./Value='" & Row("U_EXO_EUROCODE").ToString() & "']/../..")
                        If Not IsNothing(oXmlNodeGrid) Then
                            oXmlNodeGrid.SelectSingleNode("./Cells/Cell[./ColumnUid='Sel.']/Value").InnerText = "Y"
                            oXmlNodeGrid.SelectSingleNode("./Cells/Cell[./ColumnUid='Order']/Value").InnerText = CInt(Row("U_EXO_ORDER").ToString().Replace(".", "")).ToString
                            oXmlNodeGrid.SelectSingleNode("./Cells/Cell[./ColumnUid='AL0']/Value").InnerText = CInt(Row("U_EXO_AL0").ToString().Replace(".", "")).ToString
                            oXmlNodeGrid.SelectSingleNode("./Cells/Cell[./ColumnUid='AL7']/Value").InnerText = CInt(Row("U_EXO_AL7").ToString().Replace(".", "")).ToString
                            oXmlNodeGrid.SelectSingleNode("./Cells/Cell[./ColumnUid='AL8']/Value").InnerText = CInt(Row("U_EXO_AL8").ToString().Replace(".", "")).ToString
                            oXmlNodeGrid.SelectSingleNode("./Cells/Cell[./ColumnUid='AL14']/Value").InnerText = CInt(Row("U_EXO_AL14").ToString().Replace(".", "")).ToString
                            oXmlNodeGrid.SelectSingleNode("./Cells/Cell[./ColumnUid='AL16']/Value").InnerText = CInt(Row("U_EXO_AL16").ToString().Replace(".", "")).ToString
                            If CInt(Row("U_EXO_ORDER").ToString()) > 0 Then
                                oXmlNodeGrid.SelectSingleNode("./Cells/Cell[./ColumnUid='Order2']/Value").InnerText = CInt(Row("U_EXO_ORDE2").ToString().Replace(".", "")).ToString
                            End If
                            If Row("U_EXO_PROV").ToString() <> "" Then
                                oXmlNodeGrid.SelectSingleNode("./Cells/Cell[./ColumnUid='Prov.Pedido']/Value").InnerText = Row("U_EXO_PROV").ToString()
                                sSQL = "SELECT SUBSTRING(IFNULL(OCRD.""CardFName"",CAST('     ' AS VARCHAR(150))),0,5) FROM OCRD WHERE ""CardCode""='" & Row("U_EXO_PROV").ToString() & "' "
                                oXmlNodeGrid.SelectSingleNode("./Cells/Cell[./ColumnUid='Nombre']/Value").InnerText = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                            End If
                            If Row("U_EXO_FECHA").ToString() <> "" Then
                                Dim fechaPrev As String = Row("U_EXO_FECHA").ToString()
                                Dim dFecha As Date = CDate(fechaPrev)
                                oXmlNodeGrid.SelectSingleNode("./Cells/Cell[./ColumnUid='Fecha Prev.']/Value").InnerText = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                            End If
                        End If
                    Next
                    oformOrigen.DataSources.DataTables.Item("DT_DOC").LoadSerializedXML(BoDataTableXmlSelect.dxs_DataOnly, oXmlGrid.InnerXml)
                End If
#End Region
                objGlobal.SBOApp.StatusBar.SetText("Se ha recuperado el filtro " & sCode & " correctamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido recuperar el filtro " & sCode & ". Revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            oformOrigen.Freeze(False)
        Catch ex As Exception
            oformOrigen.Freeze(False)
            Throw ex
        Finally
            oformOrigen.Freeze(False)
        End Try
    End Sub
End Class
