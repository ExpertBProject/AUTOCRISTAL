Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_PLNSAVE
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
                        Case "EXO_PLNSAVE"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(objGlobal, infoEvento) = False Then
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
                        Case "EXO_PLNSAVE"

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
                        Case "EXO_PLNSAVE"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PLNSAVE"
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
    Private Function EventHandler_ItemPressed_After(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
#End Region
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnA"
                    'Grabamos el filtro
                    'Comprobamos que haya escrito el código y el nombre para poder guardar.
                    If oForm.DataSources.UserDataSources.Item("UDFILTRO").Value <> "" And oForm.DataSources.UserDataSources.Item("UDNOM").Value <> "" Then
                        objGlobal.SBOApp.StatusBar.SetText("Guardando Filtro...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        GuardarFiltro(oForm)
                        oForm.Close()
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("No se ha podido guardar si no se indica un Código y un nombre. Operación cancelada.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso de guardado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub GuardarFiltro(ByRef oForm As SAPbouiCOM.Form)
#Region "variables"
        Dim oGeneralService As SAPbobsCOM.GeneralService = Nothing
        Dim oGeneralData As SAPbobsCOM.GeneralData = Nothing
        Dim oCompService As SAPbobsCOM.CompanyService = objGlobal.compañia.GetCompanyService()
        Dim oformOrigen As SAPbouiCOM.Form = Nothing


        Dim oChildren As SAPbobsCOM.GeneralDataCollection = Nothing
        Dim oChild As SAPbobsCOM.GeneralData = Nothing
#End Region
        Try
            oGeneralService = oCompService.GetGeneralService("EXO_PLNHCO")
            oGeneralData = CType(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData), SAPbobsCOM.GeneralData)
            oGeneralData.SetProperty("Code", oForm.DataSources.UserDataSources.Item("UDFILTRO").Value)
            oGeneralData.SetProperty("Name", oForm.DataSources.UserDataSources.Item("UDNOM").Value)

            'Buscamos formulario Orieb para coger los datos.
            oformOrigen = objGlobal.SBOApp.Forms.Item(oForm.DataSources.UserDataSources.Item("UDFRM").Value)
#Region "Datos de cabecera"
            Try
                oGeneralData.SetProperty("U_EXO_EXPEDIR", oformOrigen.DataSources.UserDataSources.Item("UD_Excl").Value)
            Catch ex As Exception
                oGeneralData.SetProperty("U_EXO_EXPEDIR", "N")
            End Try
            oGeneralData.SetProperty("U_EXO_DART", oformOrigen.DataSources.UserDataSources.Item("UDARTD").Value)
            oGeneralData.SetProperty("U_EXO_HART", oformOrigen.DataSources.UserDataSources.Item("UDARTH").Value)
            oGeneralData.SetProperty("U_EXO_MGS", oformOrigen.DataSources.UserDataSources.Item("UDDIAS").Value)
            oGeneralData.SetProperty("U_EXO_TSUMI", oformOrigen.DataSources.UserDataSources.Item("UDTSM").Value)
            oGeneralData.SetProperty("U_EXO_PROV", oformOrigen.DataSources.UserDataSources.Item("UDPROV").Value)
            oGeneralData.SetProperty("U_EXO_WHS", oformOrigen.DataSources.UserDataSources.Item("UDALM").Value)
            oGeneralData.SetProperty("U_EXO_WHS2", oformOrigen.DataSources.UserDataSources.Item("UDALM2").Value)
#End Region
#Region "Artículos"
            If oformOrigen.DataSources.DataTables.Item("DTART").Rows.Count > 0 Then
                oChildren = Nothing : oChild = Nothing
                oChildren = oGeneralData.Child("EXO_PLNHCO1")
                For j As Integer = 0 To oformOrigen.DataSources.DataTables.Item("DTART").Rows.Count - 1
                    If oformOrigen.DataSources.DataTables.Item("DTART").GetValue("Código", j).ToString <> "" Then
                        oChild = oChildren.Add()
                        oChild.SetProperty("U_EXO_ITEMCODE", oformOrigen.DataSources.DataTables.Item("DTART").GetValue("Código", j).ToString)
                        oChild.SetProperty("U_EXO_ITEMNAME", oformOrigen.DataSources.DataTables.Item("DTART").GetValue("Descripción", j).ToString)
                    End If
                Next
            End If
#End Region
#Region "Almacenes"
            If oformOrigen.DataSources.DataTables.Item("DTALM").Rows.Count > 0 Then
                oChildren = Nothing : oChild = Nothing
                oChildren = oGeneralData.Child("EXO_PLNHCO2")
                For j As Integer = 0 To oformOrigen.DataSources.DataTables.Item("DTALM").Rows.Count - 1
                    If oformOrigen.DataSources.DataTables.Item("DTALM").GetValue("Cod.", j).ToString <> "" Then
                        oChild = oChildren.Add()
                        oChild.SetProperty("U_Sel", oformOrigen.DataSources.DataTables.Item("DTALM").GetValue("Sel", j).ToString)
                        oChild.SetProperty("U_EXO_WHSCODE", oformOrigen.DataSources.DataTables.Item("DTALM").GetValue("Cod.", j).ToString)
                        oChild.SetProperty("U_EXO_WHSNAME", oformOrigen.DataSources.DataTables.Item("DTALM").GetValue("Almacén", j).ToString)
                    End If
                Next
            End If
#End Region
#Region "Almacenes 2"
            If oformOrigen.DataSources.DataTables.Item("DTALM2").Rows.Count > 0 Then
                oChildren = Nothing : oChild = Nothing
                oChildren = oGeneralData.Child("EXO_PLNHCO7")
                For j As Integer = 0 To oformOrigen.DataSources.DataTables.Item("DTALM2").Rows.Count - 1
                    If oformOrigen.DataSources.DataTables.Item("DTALM2").GetValue("Cod.", j).ToString <> "" Then
                        oChild = oChildren.Add()
                        oChild.SetProperty("U_EXO_SEL", oformOrigen.DataSources.DataTables.Item("DTALM2").GetValue("Sel", j).ToString)
                        oChild.SetProperty("U_EXO_WHSCODE", oformOrigen.DataSources.DataTables.Item("DTALM2").GetValue("Cod.", j).ToString)
                        oChild.SetProperty("U_EXO_WHSNAME", oformOrigen.DataSources.DataTables.Item("DTALM2").GetValue("Almacén", j).ToString)
                    End If
                Next
            End If
#End Region
#Region "Clasificación"
            If oformOrigen.DataSources.DataTables.Item("DTCLAS").Rows.Count > 0 Then
                oChildren = Nothing : oChild = Nothing
                oChildren = oGeneralData.Child("EXO_PLNHCO3")
                For j As Integer = 0 To oformOrigen.DataSources.DataTables.Item("DTCLAS").Rows.Count - 1
                    If oformOrigen.DataSources.DataTables.Item("DTCLAS").GetValue("Clas", j).ToString <> "" Then
                        oChild = oChildren.Add()
                        oChild.SetProperty("U_EXO_SEL", oformOrigen.DataSources.DataTables.Item("DTCLAS").GetValue("Sel", j).ToString)
                        oChild.SetProperty("U_EXO_CLAS", oformOrigen.DataSources.DataTables.Item("DTCLAS").GetValue("Clas", j).ToString)
                    End If
                Next
            End If
#End Region
#Region "Grupos Art."
            If oformOrigen.DataSources.DataTables.Item("DTGRU").Rows.Count > 0 Then
                oChildren = Nothing : oChild = Nothing
                oChildren = oGeneralData.Child("EXO_PLNHCO4")
                For j As Integer = 0 To oformOrigen.DataSources.DataTables.Item("DTGRU").Rows.Count - 1
                    If oformOrigen.DataSources.DataTables.Item("DTGRU").GetValue("Cod.", j).ToString <> "" Then
                        oChild = oChildren.Add()
                        oChild.SetProperty("U_EXO_SEL", oformOrigen.DataSources.DataTables.Item("DTGRU").GetValue("Sel", j).ToString)
                        oChild.SetProperty("U_EXO_GRUPO", oformOrigen.DataSources.DataTables.Item("DTGRU").GetValue("Cod.", j).ToString)
                        oChild.SetProperty("U_EXO_DES", oformOrigen.DataSources.DataTables.Item("DTGRU").GetValue("Familia", j).ToString)
                    End If
                Next
            End If
#End Region
#Region "Lst. de Compras"
            If oformOrigen.DataSources.DataTables.Item("DTCOMP").Rows.Count > 0 Then
                oChildren = Nothing : oChild = Nothing
                oChildren = oGeneralData.Child("EXO_PLNHCO5")
                For j As Integer = 0 To oformOrigen.DataSources.DataTables.Item("DTCOMP").Rows.Count - 1
                    If oformOrigen.DataSources.DataTables.Item("DTCOMP").GetValue("LST", j).ToString <> "" Then
                        oChild = oChildren.Add()
                        oChild.SetProperty("U_EXO_SEL", oformOrigen.DataSources.DataTables.Item("DTCOMP").GetValue("Sel", j).ToString)
                        oChild.SetProperty("U_EXO_LST", oformOrigen.DataSources.DataTables.Item("DTCOMP").GetValue("LST", j).ToString)
                        oChild.SetProperty("U_EXO_NOMBRE", oformOrigen.DataSources.DataTables.Item("DTCOMP").GetValue("Nombre", j).ToString)
                    End If
                Next
            End If
#End Region
#Region "Lst. de Ventas"
            If oformOrigen.DataSources.DataTables.Item("DTVENT").Rows.Count > 0 Then
                oChildren = Nothing : oChild = Nothing
                oChildren = oGeneralData.Child("EXO_PLNHCO6")
                For j As Integer = 0 To oformOrigen.DataSources.DataTables.Item("DTVENT").Rows.Count - 1
                    If oformOrigen.DataSources.DataTables.Item("DTVENT").GetValue("LST", j).ToString <> "" Then
                        oChild = oChildren.Add()
                        oChild.SetProperty("U_EXO_SEL", oformOrigen.DataSources.DataTables.Item("DTVENT").GetValue("Sel", j).ToString)
                        oChild.SetProperty("U_EXO_LST", oformOrigen.DataSources.DataTables.Item("DTVENT").GetValue("LST", j).ToString)
                        oChild.SetProperty("U_EXO_NOMBRE", oformOrigen.DataSources.DataTables.Item("DTVENT").GetValue("Nombre", j).ToString)
                    End If
                Next
            End If
#End Region
#Region "Lineas"
            Dim sXmlGrid As String = "" : Dim oXmlGrid As New Xml.XmlDocument : Dim oXmlNodesGrid As XmlNodeList = Nothing
            sXmlGrid = oformOrigen.DataSources.DataTables.Item("DT_DOC").SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All)
            oXmlGrid.LoadXml(sXmlGrid)

            oXmlNodesGrid = oXmlGrid.SelectNodes("/DataTable/Rows/Row/Cells/Cell[./ColumnUid='Sel.' and ./Value='Y']/../..")

            If oXmlNodesGrid IsNot Nothing AndAlso oXmlNodesGrid.Count > 0 Then
                oChildren = Nothing : oChild = Nothing
                oChildren = oGeneralData.Child("EXO_PLNHCO8")
                For i As Integer = 0 To oXmlNodesGrid.Count - 1
                    oChild = oChildren.Add()
                    Dim EUROCODE As String = oXmlNodesGrid.Item(i).SelectSingleNode("./Cells/Cell[./ColumnUid='EUROCODE']/Value").InnerText
                    oChild.SetProperty("U_EXO_EUROCODE", EUROCODE)
                    Dim order As String = oXmlNodesGrid.Item(i).SelectSingleNode("./Cells/Cell[./ColumnUid='Order']/Value").InnerText
                    oChild.SetProperty("U_EXO_ORDER", order)
                    Dim al0 = oXmlNodesGrid.Item(i).SelectSingleNode("./Cells/Cell[./ColumnUid='AL0']/Value").InnerText
                    oChild.SetProperty("U_EXO_AL0", al0)
                    Dim al7 = oXmlNodesGrid.Item(i).SelectSingleNode("./Cells/Cell[./ColumnUid='AL7']/Value").InnerText
                    oChild.SetProperty("U_EXO_AL7", al7)
                    Dim al8 = oXmlNodesGrid.Item(i).SelectSingleNode("./Cells/Cell[./ColumnUid='AL8']/Value").InnerText
                    oChild.SetProperty("U_EXO_AL8", al8)
                    Dim al14 = oXmlNodesGrid.Item(i).SelectSingleNode("./Cells/Cell[./ColumnUid='AL14']/Value").InnerText
                    oChild.SetProperty("U_EXO_AL14", al14)
                    Dim al16 = oXmlNodesGrid.Item(i).SelectSingleNode("./Cells/Cell[./ColumnUid='AL16']/Value").InnerText
                    oChild.SetProperty("U_EXO_AL16", al16)
                    Try
                        Dim order2 As String = oXmlNodesGrid.Item(i).SelectSingleNode("./Cells/Cell[./ColumnUid='Order2']/Value").InnerText
                        oChild.SetProperty("U_EXO_ORDE2", order2)
                    Catch ex As Exception

                    End Try
                    Dim Proveedor As String = oXmlNodesGrid.Item(i).SelectSingleNode("./Cells/Cell[./ColumnUid='Prov.Pedido']/Value").InnerText
                    oChild.SetProperty("U_EXO_PROV", Proveedor)
                    Dim fechaPrev As String = oXmlNodesGrid.Item(i).SelectSingleNode("./Cells/Cell[./ColumnUid='Fecha Prev.']/Value").InnerText
                    If fechaPrev <> "" And Not IsNothing(fechaPrev) And fechaPrev <> "00000000" Then
                        Dim dFecha As Date = New Date(CInt(Left(fechaPrev, 4)), CInt(Mid(fechaPrev, 5, 2)), CInt(Right(fechaPrev, 2)))
                        oChild.SetProperty("U_EXO_FECHA", dFecha)
                    End If
                Next
            End If



            'Dim dataObject = XDocument.Parse(oformOrigen.DataSources.DataTables.Item("DT_DOC").SerializeAsXML(BoDataTableXmlSelect.dxs_DataOnly))
            'Dim rowsSelected = dataObject.
            '    Descendants("Cell").
            '    Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode.NextNode, XElement).Value.ToString.Equals("Y")).
            '    Select(Function(attr) attr.Parent).
            '    ToList()

            'If (rowsSelected.Count > 0) Then
            '    oChildren = Nothing : oChild = Nothing
            '    oChildren = oGeneralData.Child("EXO_PLNHCO8")
            '    For x As Integer = 0 To rowsSelected.Count - 1
            '        oChild = oChildren.Add()
            '        Dim EUROCODE As String = CType(rowsSelected(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("EUROCODE")).First().LastNode, XElement).Value
            '        oChild.SetProperty("U_EXO_EUROCODE", EUROCODE)
            '        Dim order As String = CType(rowsSelected(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Order")).First().LastNode, XElement).Value
            '        oChild.SetProperty("U_EXO_ORDER", order)
            '        Dim al0 = CType(rowsSelected(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL0")).First().LastNode, XElement).Value
            '        oChild.SetProperty("U_EXO_AL0", al0)
            '        Dim al7 = CType(rowsSelected(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL7")).First().LastNode, XElement).Value
            '        oChild.SetProperty("U_EXO_AL7", al7)
            '        Dim al8 = CType(rowsSelected(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL8")).First().LastNode, XElement).Value
            '        oChild.SetProperty("U_EXO_AL8", al8)
            '        Dim al14 = CType(rowsSelected(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL14")).First().LastNode, XElement).Value
            '        oChild.SetProperty("U_EXO_AL14", al14)
            '        Dim al16 = CType(rowsSelected(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("AL16")).First().LastNode, XElement).Value
            '        oChild.SetProperty("U_EXO_AL16", al16)
            '        Try
            '            Dim order2 As String = CType(rowsSelected(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Order2")).First().LastNode, XElement).Value
            '            oChild.SetProperty("U_EXO_ORDE2", order2)
            '        Catch ex As Exception

            '        End Try
            '        Dim Proveedor As String = CType(rowsSelected(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Prov.Pedido")).First().LastNode, XElement).Value
            '        oChild.SetProperty("U_EXO_PROV", Proveedor)
            '        Dim fechaPrev As String = CType(rowsSelected(x).Descendants("Cell").Where(Function(attr) CType(attr.FirstNode, XElement).Name.ToString().Equals("ColumnUid") And CType(attr.FirstNode, XElement).Value.ToString.Equals("Fecha Prev.")).First().LastNode, XElement).Value
            '        If fechaPrev <> "" And Not IsNothing(fechaPrev) And fechaPrev <> "00000000" Then
            '            Dim dFecha As Date = New Date(CInt(Left(fechaPrev, 4)), CInt(Mid(fechaPrev, 5, 2)), CInt(Right(fechaPrev, 2)))
            '            oChild.SetProperty("U_EXO_FECHA", dFecha)
            '        End If

            '    Next
            'End If
#End Region
            oGeneralService.Add(oGeneralData)
            objGlobal.SBOApp.StatusBar.SetText("Filtro guardado correctamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            Throw ex
        Finally
            oChild = Nothing : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oChild, Object))
            oChildren = Nothing : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oChildren, Object))
        End Try
    End Sub
End Class
