Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OWTQ
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

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
                        Case "1250000940"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "1250000940"
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
                        Case "1250000940"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "1250000940"
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

            oForm.Freeze(True)
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oItem = oForm.Items.Add("btnPL", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("2").Left + 5
            oItem.Width = oForm.Items.Item("2").Width + 100
            oItem.Top = oForm.Items.Item("2").Top
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Enabled = False
            oItem.LinkTo = "2"
            Dim oBtnAct As SAPbouiCOM.Button
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Generar Packing List"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            oItem = oForm.Items.Add("btnTR", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("btnPL").Left + oForm.Items.Item("btnPL").Width + 5
            oItem.Width = oForm.Items.Item("2").Width + 100
            oItem.Top = oForm.Items.Item("2").Top
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Enabled = False
            oItem.LinkTo = "2"
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Crear Traslado"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            oForm.Freeze(False)

            EventHandler_Form_Load = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)

            If oForm IsNot Nothing Then oForm.Visible = True

            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)

            If oForm IsNot Nothing Then oForm.Visible = True

            Throw ex
        Finally
            oForm.Freeze(False)

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
#Region "variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim iTamMatrix As Integer = 0
        Dim bSelLinea As Boolean = False
        Dim sMensaje As String = ""
        Dim sEstado As String = ""
        Dim sCancelado As String = ""
        Dim sDocEntry As String = ""
        Dim sDocNumSolTraslado As String = ""
        Dim sObjType As String = ""
        Dim sIC As String = ""
        Dim sAlmOrigen As String = "" : Dim sAlmDestino As String = ""
#End Region


        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            sDocEntry = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("DocEntry", 0).ToString.Trim
            sDocNumSolTraslado = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("DocNum", 0).ToString.Trim
            sObjType = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("ObjType", 0).ToString.Trim
            sCancelado = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("CANCELED", 0).ToString.Trim
            sEstado = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("DocStatus", 0).ToString.Trim
            sIC = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("CardCode", 0).ToString.Trim
            sAlmOrigen = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("Filler", 0).ToString.Trim
            sAlmDestino = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("ToWhsCode", 0).ToString.Trim
            Select Case pVal.ItemUID
                Case "btnPL"
                    If sEstado = "O" And sCancelado = "N" Then
                        Dim sListaEmbalaje As String = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("U_EXO_LSTEMB", 0).ToString.Trim
                        If sListaEmbalaje = "" Then
                            sMensaje = "La Sol. de traslado no tiene asignado una Lista de embalaje. No se puede generar el packing list."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                        Else
                            'Cargamos el Packing list segín lista de embalaje
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Generando Packing List de la Sol. de traslado " & sDocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            GEN_PACKINGLIST(objGlobal.compañia, objGlobal, sListaEmbalaje, sDocEntry, sDocNumSolTraslado, sObjType, sIC)
                        End If
                    Else
                        If sCancelado = "Y" Then
                            sMensaje = "La sol. de traslado está cancelada, no se puede crear el Packing List."
                        Else
                            If sEstado = "C" Then
                                sMensaje = "La sol. de traslado está cerrada, no se puede crear el Packing List."
                            End If
                        End If
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox(sMensaje)
                    End If
                Case "btnTR"
                    If sEstado = "O" And sCancelado = "N" Then
                        Dim sPackingList As String = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("U_EXO_PACKING", 0).ToString.Trim
                        If sPackingList = "" Then
                            sMensaje = "La Sol. de traslado no tiene asignado un Packing List. No se puede generar el traslado desde esta opción."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                        Else
                            Generar_TR(objGlobal.compañia, objGlobal, sPackingList, sDocEntry, sIC, sAlmOrigen, sAlmDestino)
                        End If
                    Else
                        If sCancelado = "Y" Then
                            sMensaje = "La sol. de traslado está cancelada, no se puede generar el traslado."
                        Else
                            If sEstado = "C" Then
                                sMensaje = "La sol. de traslado está cerrada, no se puede generar el tarslado."
                            End If
                        End If
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox(sMensaje)
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))

        End Try
    End Function
    Public Shared Sub Generar_TR(ByRef oCompany As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sPacking_list As String,
                                 ByVal sSolTrasDocEntry As String, ByVal sCardCode As String, ByVal sAlmOrigen As String, ByVal sAlmDestino As String)
#Region "Variables"
        Dim oDtLinPacking As System.Data.DataTable = New System.Data.DataTable
        Dim oDtLin As System.Data.DataTable = New System.Data.DataTable
        Dim oOWTR As SAPbobsCOM.StockTransfer = Nothing
        Dim dfecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim sSQL As String = ""
        Dim sMensaje As String = "" : Dim sError As String = "" : Dim sComen As String = "" : Dim sEstado As String = ""
        Dim sDocEntry As String = "" : Dim sDocnum As String = ""
        Dim oRsLote As SAPbobsCOM.Recordset = Nothing : Dim oRsLocalizacion As SAPbobsCOM.Recordset = Nothing : Dim oRsLocalizacionDest As SAPbobsCOM.Recordset = Nothing
        Dim iAbsEntry As Integer = 0
        Dim dCantLotes As Double = 0 : Dim iLineaUbi As Integer = 0
        Dim sDocNumPedido As String = ""
#End Region

        Try
            oRsLote = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsLocalizacion = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsLocalizacionDest = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oOWTR = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer), SAPbobsCOM.StockTransfer)
            oOWTR.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer
            oOWTR.CardCode = sCardCode
            oOWTR.TaxDate = dfecha
            oOWTR.DocDate = dfecha
            oOWTR.FromWarehouse = sAlmOrigen
            oOWTR.ToWarehouse = sAlmDestino

            oOWTR.UserFields.Fields.Item("U_EXO_PACKING").Value = sPacking_list
            oDtLin.Clear()

            sSQL = "SELECT * FROM ""WTQ1"" where ""LineStatus""='O' and ""DocEntry""=" & sSolTrasDocEntry & " Order by ""LineNum"" "
            oDtLin = oObjGlobal.refDi.SQL.sqlComoDataTable(sSQL)
            If oDtLin.Rows.Count > 0 Then
                sDocNumPedido = oObjGlobal.refDi.SQL.sqlStringB1("SELECT ""DocNum"" FROM OWTQ WHERE ""DocEntry""=" & sSolTrasDocEntry)
                Dim bPlinea As Boolean = True
                For iLin As Integer = 0 To oDtLin.Rows.Count - 1
                    'buscamos en la tabla de ficheros
                    'Sólo Línea asignada
                    oDtLinPacking.Clear()
                    sSQL = "SELECT ""U_EXO_CODE"",sum(""U_EXO_CANT"") ""CANTIDAD"" FROM ""@EXO_PACKINGL"" 
                            where ""Code""='" & sPacking_list & "' and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' 
                            and ""U_EXO_LINEA""='" & oDtLin.Rows.Item(iLin).Item("LineNum").ToString & "' 
                            GROUP BY ""U_EXO_CODE"" "
                    oDtLinPacking = oObjGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                    If oDtLinPacking.Rows.Count > 0 Then
                        For iLinFich As Integer = 0 To oDtLinPacking.Rows.Count - 1
                            dCantLotes = 0
                            If bPlinea = False Then
                                oOWTR.Lines.Add()
                            Else
                                bPlinea = False
                            End If
                            oOWTR.Lines.ItemCode = oDtLin.Rows.Item(iLin).Item("ItemCode").ToString
                            oOWTR.Lines.ItemDescription = oDtLin.Rows.Item(iLin).Item("Dscription").ToString
                            Dim dCantFichero As Double = EXO_GLOBALES.DblTextToNumber(oCompany, oDtLinPacking.Rows.Item(iLinFich).Item("CANTIDAD").ToString)
                            Dim dCant As Double = EXO_GLOBALES.DblTextToNumber(oCompany, oDtLin.Rows.Item(iLin).Item("Quantity").ToString)
                            Dim sUnidad As String = oDtLin.Rows.Item(iLin).Item("UomCode").ToString.Trim
                            oOWTR.Lines.BaseEntry = CInt(oDtLin.Rows.Item(iLin).Item("DocEntry").ToString)
                            oOWTR.Lines.BaseType = SAPbobsCOM.InvBaseDocTypeEnum.InventoryTransferRequest
                            oOWTR.Lines.BaseLine = CInt(oDtLin.Rows.Item(iLin).Item("LineNum").ToString)

                            oOWTR.Lines.FromWarehouseCode = oDtLin.Rows.Item(iLin).Item("FromWhsCod").ToString
                            oOWTR.Lines.WarehouseCode = oDtLin.Rows.Item(iLin).Item("WhsCode").ToString
#Region "Lotes"
                            'Incluimos los Lotes y solo del pedido y la línea
                            sSQL = "SELECT ""U_EXO_CODE"",""U_EXO_LOTE"", sum(""U_EXO_CANT"") ""CANTIDAD"", ""U_EXO_FFAB"" FROM ""@EXO_PACKINGL""
                                    where ""Code""='" & sPacking_list & "' and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' 
                                    and ""U_EXO_LINEA""='" & oDtLin.Rows.Item(iLin).Item("LineNum").ToString & "'
                                    GROUP BY ""U_EXO_CODE"",""U_EXO_LOTE"", ""U_EXO_FFAB"" "
                            oRsLote.DoQuery(sSQL)
                            For iLote = 0 To oRsLote.RecordCount - 1
                                'Creamos el lote de la línea del artículo
                                If iLote <> 0 Then
                                    oOWTR.Lines.BatchNumbers.Add()
                                End If
                                oOWTR.Lines.BatchNumbers.BatchNumber = oRsLote.Fields.Item("U_EXO_LOTE").Value.ToString
                                oOWTR.Lines.BatchNumbers.Quantity = EXO_GLOBALES.DblTextToNumber(oCompany, oRsLote.Fields.Item("CANTIDAD").Value.ToString)
                                dCantLotes += oOWTR.Lines.BatchNumbers.Quantity
                                oOWTR.Lines.BatchNumbers.ManufacturingDate = CDate(oRsLote.Fields.Item("U_EXO_FFAB").Value.ToString)
                                sSQL = "Select IFNULL(OMRC.""FirmName"",'') FROM OCRD LEFT JOIN OMRC ON OCRD.""U_EXO_MARPRO""=OMRC.""FirmCode"" Where ""CardCode""='" & sCardCode & "' "
                                'oObjGlobal.SBOApp.StatusBar.SetText(sSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                oOWTR.Lines.BatchNumbers.ManufacturerSerialNumber = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
#Region "Localizacion"
                                sSQL = "SELECT ""U_EXO_CODE"",""U_EXO_LOTE"", sum(""U_EXO_CANT"") ""CANTIDAD"", ""U_EXO_FFAB"",""U_EXO_UBIORI""
                                        FROM ""@EXO_PACKINGL"" 
                                         WHERE ""Code""='" & sPacking_list & "'  
                                         and IFNULL(""U_EXO_UBIORI"",'')<>''
                                         and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "'
                                         and ""U_EXO_LOTE""='" & oRsLote.Fields.Item("U_EXO_LOTE").Value.ToString & "'  
                                         and ""U_EXO_LINEA""='" & oDtLin.Rows.Item(iLin).Item("LineNum").ToString & "'
                                        GROUP BY ""U_EXO_CODE"",""U_EXO_LOTE"", ""U_EXO_FFAB"",""U_EXO_UBIORI"" "
                                oRsLocalizacion.DoQuery(sSQL)

                                For iLoc = 0 To oRsLocalizacion.RecordCount - 1
                                    sSQL = "Select IFNULL(""AbsEntry"",0) from OBIN where ""BinCode"" = '" & oRsLocalizacion.Fields.Item("U_EXO_UBIORI").Value.ToString.Trim & "'"
                                    iAbsEntry = CInt(oObjGlobal.refDi.SQL.sqlStringB1(sSQL))
                                    If iAbsEntry <> 0 Then
                                        If iLoc <> 0 Then
                                            oOWTR.Lines.BinAllocations.Add()
                                        End If
                                        oOWTR.Lines.BinAllocations.BinAbsEntry = iAbsEntry
                                        oOWTR.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batFromWarehouse
                                        oOWTR.Lines.BinAllocations.Quantity = EXO_GLOBALES.DblTextToNumber(oCompany, oRsLocalizacion.Fields.Item("CANTIDAD").Value.ToString)
                                        'oOWTR.Lines.BinAllocations.BaseLineNumber = oOWTR.Lines.LineNum
                                        oOWTR.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = iLote
                                    Else
                                        sMensaje = "No se encuentra La ubicación origen: " & oRsLocalizacion.Fields.Item("U_EXO_UBIORI").Value.ToString.Trim
                                        sMensaje &= ". No se incluye en la Entrada de Mercancía."
                                        oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If

#Region "Localizacion Destino"
                                    sSQL = "SELECT ""U_EXO_CODE"",""U_EXO_LOTE"", sum(""U_EXO_CANT"") ""CANTIDAD"", ""U_EXO_FFAB"",""U_EXO_UBIRECEP""
                                        FROM ""@EXO_PACKINGL"" 
                                         WHERE ""Code""='" & sPacking_list & "' and IFNULL(""U_EXO_UBIRECEP"",'')<>'' 
                                         and IFNULL(""U_EXO_UBIORI"",'')<>''
                                         and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "'
                                         and ""U_EXO_LOTE""='" & oRsLote.Fields.Item("U_EXO_LOTE").Value.ToString & "' 
                                         and ""U_EXO_UBIORI""='" & oRsLocalizacion.Fields.Item("U_EXO_UBIORI").Value.ToString & "' 
                                         and ""U_EXO_LINEA""='" & oDtLin.Rows.Item(iLin).Item("LineNum").ToString & "'
                                        GROUP BY ""U_EXO_CODE"",""U_EXO_LOTE"", ""U_EXO_FFAB"",""U_EXO_UBIRECEP"" "
                                    oRsLocalizacionDest.DoQuery(sSQL)

                                    Dim dCantDestino As Double = 0
                                    If oRsLocalizacionDest.RecordCount = 0 Then
                                        dCantDestino = EXO_GLOBALES.DblTextToNumber(oCompany, oRsLocalizacion.Fields.Item("CANTIDAD").Value.ToString)
                                        sSQL = "SELECT TOP 1 ""BinCode"" ""U_EXO_UBIRECEP"" From OBIN WHERE""ReceiveBin""='Y' and ""WhsCode""='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'"
                                        oRsLocalizacionDest.DoQuery(sSQL)
                                    Else
                                        dCantDestino = EXO_GLOBALES.DblTextToNumber(oCompany, oRsLocalizacionDest.Fields.Item("CANTIDAD").Value.ToString)
                                    End If
                                    For iLocDest = 0 To oRsLocalizacionDest.RecordCount - 1
                                        sSQL = "Select IFNULL(""AbsEntry"",0) from OBIN where ""BinCode"" = '" & oRsLocalizacionDest.Fields.Item("U_EXO_UBIRECEP").Value.ToString.Trim & "'"
                                        iAbsEntry = CInt(oObjGlobal.refDi.SQL.sqlStringB1(sSQL))
                                        If iAbsEntry <> 0 Then
                                            oOWTR.Lines.BinAllocations.Add()

                                            oOWTR.Lines.BinAllocations.BinAbsEntry = iAbsEntry
                                            oOWTR.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batToWarehouse
                                            oOWTR.Lines.BinAllocations.Quantity = dCantDestino
                                            ' oOWTR.Lines.BinAllocations.BaseLineNumber = oOWTR.Lines.LineNum
                                            oOWTR.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = iLote
                                        Else
                                            sMensaje = "No se encuentra La ubicación destino: " & oRsLocalizacionDest.Fields.Item("U_EXO_UBIRECEP").Value.ToString.Trim
                                            sMensaje &= ". No se incluye en la Entrada de Mercancía."
                                            oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                        oRsLocalizacionDest.MoveNext()
                                    Next
#End Region
                                    oRsLocalizacion.MoveNext()
                                Next
#End Region

                                oRsLote.MoveNext()
                            Next
#End Region

                            oOWTR.Lines.Quantity = dCantLotes
                            ' oOPDN.Lines.InventoryQuantity = dCantLotes  
                        Next
                    Else
                        sMensaje = "No se encuentra la línea " & oDtLin.Rows.Item(iLin).Item("LineNum").ToString & " con el artículo " & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString
                        sMensaje &= ". No se incluye en la Entrada de Mercancía."
                        oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                Next
                If oOWTR.Add() <> 0 Then
                    sEstado = "Error"
                    sError = oCompany.GetLastErrorCode.ToString & " / " & oCompany.GetLastErrorDescription.Replace("'", "")
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    oCompany.GetNewObjectCode(sDocEntry)

                    sSQL = "SELECT ""DocNum"" FROM ""OWTR""  WHERE ""DocEntry""=" & sDocEntry
                    sDocnum = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    sMensaje = "Se ha generado correctamente el traslado con Nº " & sDocnum
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    oObjGlobal.SBOApp.MessageBox(sMensaje)


                    If oObjGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
                        oObjGlobal.SBOApp.ActivateMenuItem("1304")
                    End If

                End If
            Else
                sMensaje = "No se encuentra las líneas de la Sol. de Traslado interno Nº" & sSolTrasDocEntry & ". Se interrumpe el proceso."
                oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oObjGlobal.SBOApp.MessageBox(sMensaje)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oDtLin.Clear() : oDtLinPacking.Clear()
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLote, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLocalizacion, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLocalizacionDest, Object))
        End Try
    End Sub
    Private Sub GEN_PACKINGLIST(ByRef oCompany As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI,
                                ByVal sListaEmbalaje As String, ByVal sDocEntry As String, ByVal sDocNumSolTraslado As String,
                                ByVal sObjType As String, ByVal sIc As String)
#Region "Variables"
        Dim sSQL As String = ""
        Dim oDtLin As System.Data.DataTable = New System.Data.DataTable
        Dim dStock As Double = 0 : Dim dStockExt As Double = 0
        Dim dStockExt1 As Double = 0 : Dim dStockExt2 As Double = 0 : Dim dStockExt3 As Double = 0 : Dim dStockExt4 As Double = 0 : Dim dStockExt5 As Double = 0
        Dim sUBIDEF As String = "" : Dim sTIPOHUECODEF As String = "" : Dim dCANTMAXDEF As Double = 0 : Dim dSTOCKUBIDEF As Double = 0
        Dim dVMA As Double = 0 : Dim dVA As Double = 0 : Dim dCober As Double = 0
        Dim sCatalogo As String = ""
#End Region
        Try
            oDtLin.Clear()
            sSQL = "SELECT * FROM ""WTQ1"" where ""LineStatus""='O' and ""DocEntry""=" & sDocEntry & " Order by ""LineNum"" "
            oDtLin = oObjGlobal.refDi.SQL.sqlComoDataTable(sSQL)
            If oDtLin.Rows.Count > 0 Then
                sSQL = "DELETE FROM ""@EXO_PACKING"" WHERE ""Code""='" & sDocEntry & sObjType & "' "
                If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    sSQL = "DELETE FROM ""@EXO_PACKINGL"" WHERE ""Code""='" & sDocEntry & sObjType & "' "
                    oObjGlobal.refDi.SQL.executeNonQuery(sSQL)
                End If
                sSQL = "insert into ""@EXO_PACKING"" (""Code"",""Name"",""DocEntry"",""Object"",""U_EXO_OBJTYPE"") 
                                values('" & sDocEntry & sObjType & "','" & sDocNumSolTraslado & "'," & sDocEntry & Left(sObjType, 3) & "01" & ",'EXO_PACKING','" & sObjType & "')"
                If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    'Actualizamos el pedido con el Packing List
                    sSQL = "UPDATE OWTQ SET ""U_EXO_PACKING""='" & sDocEntry & sObjType & "' WHERE ""DocEntry""=" & sDocEntry
                    oObjGlobal.refDi.SQL.executeNonQuery(sSQL)
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Generando Packing List...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    For iLin As Integer = 0 To oDtLin.Rows.Count - 1
                        sSQL = " SELECT ""OnHand"" FROM OITW Where ""ItemCode""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' 
                                            and ""WhsCode""='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'"
                        dStock = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)

                        sSQL = "Select Sum(""OnHandQty"") as ""StockExterno""  from OIBQ T1
                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T1.""BinAbs"" 
                                    Where  T11.""SL1Code""  In (SELECT ""U_EXO_ZONA"" 
                                            							FROM ""@EXO_UBIEXTERNAS"" 
                                            							WHERE ""U_EXO_ALM"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "' )
                                    and T1.""WhsCode"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'
                                    and T1.""ItemCode""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "'"
                        dStockExt = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)

                        sSQL = "Select Sum(""OnHandQty"") as ""StockExterno""  from OIBQ T1
                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T1.""BinAbs"" 
                                    Where  T11.""SL1Code""  In (SELECT ""U_EXO_ZONA"" 
                                            							FROM ""@EXO_UBIEXTERNAS"" 
                                            							WHERE ""U_EXO_TIPOUBI""='Ext1' and ""U_EXO_ALM"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "' )
                                    and T1.""WhsCode"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'
                                    and T1.""ItemCode""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "'"
                        dStockExt1 = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)

                        sSQL = "Select Sum(""OnHandQty"") as ""StockExterno""  from OIBQ T1
                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T1.""BinAbs"" 
                                    Where  T11.""SL1Code""  In (SELECT ""U_EXO_ZONA"" 
                                            							FROM ""@EXO_UBIEXTERNAS"" 
                                            							WHERE ""U_EXO_TIPOUBI""='Ext2' and ""U_EXO_ALM"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "' )
                                    and T1.""WhsCode"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'
                                    and T1.""ItemCode""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "'"
                        dStockExt2 = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)

                        sSQL = "Select Sum(""OnHandQty"") as ""StockExterno""  from OIBQ T1
                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T1.""BinAbs"" 
                                    Where  T11.""SL1Code""  In (SELECT ""U_EXO_ZONA"" 
                                            							FROM ""@EXO_UBIEXTERNAS"" 
                                            							WHERE ""U_EXO_TIPOUBI""='Ext3' and ""U_EXO_ALM"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "' )
                                    and T1.""WhsCode"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'
                                    and T1.""ItemCode""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "'"
                        dStockExt3 = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)

                        sSQL = "Select Sum(""OnHandQty"") as ""StockExterno""  from OIBQ T1
                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T1.""BinAbs"" 
                                    Where  T11.""SL1Code""  In (SELECT ""U_EXO_ZONA"" 
                                            							FROM ""@EXO_UBIEXTERNAS"" 
                                            							WHERE ""U_EXO_TIPOUBI""='Ext4' and ""U_EXO_ALM"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "' )
                                    and T1.""WhsCode"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'
                                    and T1.""ItemCode""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "'"
                        dStockExt4 = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)

                        sSQL = "Select Sum(""OnHandQty"") as ""StockExterno""  from OIBQ T1
                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T1.""BinAbs"" 
                                    Where  T11.""SL1Code""  In (SELECT ""U_EXO_ZONA"" 
                                            							FROM ""@EXO_UBIEXTERNAS"" 
                                            							WHERE ""U_EXO_TIPOUBI""='Ext5' and ""U_EXO_ALM"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "' )
                                    and T1.""WhsCode"" = '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'
                                    and T1.""ItemCode""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "'"
                        dStockExt5 = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)

                        sSQL = "SELECT T1.""BinCode"" FROM OITW T0  
                                INNER JOIN OBIN T1 ON T0.""DftBinAbs"" = T1.""AbsEntry"" 
                                WHERE T0.""ItemCode"" ='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' 
                                  and T0.""WhsCode"" ='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'"
                        sUBIDEF = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)

                        sSQL = "SELECT T1.""Attr4Val"" FROM OITW T0  
                                INNER JOIN OBIN T1 ON T0.""DftBinAbs"" = T1.""AbsEntry"" 
                                WHERE T0.""ItemCode"" ='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' 
                                  and T0.""WhsCode"" ='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'"
                        sTIPOHUECODEF = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)

                        sSQL = "SELECT T1.""MaxLevel"" FROM OITW T0  
                                INNER JOIN OBIN T1 ON T0.""DftBinAbs"" = T1.""AbsEntry"" 
                                WHERE T0.""ItemCode"" ='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' 
                                  and T0.""WhsCode"" ='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'"
                        dCANTMAXDEF = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)

                        sSQL = "Select ""OnHandQty"" from OIBQ where ""ItemCode"" = '" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' 
                                and ""BinAbs"" = (Select T1.""AbsEntry"" FROM OITW T0  
                                                    INNER Join OBIN T1 ON T0.""DftBinAbs"" = T1.""AbsEntry"" 
                                                     WHERE T0.""ItemCode"" ='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' 
                                                    And T0.""WhsCode"" ='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'
                                                 )"
                        dSTOCKUBIDEF = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)
#Region "Nuevos Campos"
                        sSQL = " SELECT ""Ventas_Med_Año"" FROM (Select T0.""WhsCode"" as ""Almacen"", T0.""ItemCode"" as ""Artículo"", T0.""OnHand"",
                                    T0.""OnHand"" -  (coalesce(T3.""Stock"",0) + coalesce(T4.""Stock"",0) + coalesce(T5.""Stock"",0) + coalesce(T6.""Stock"",0) + coalesce(T7.""Stock"",0) ) AS ""STOCK_DENTRO""
                                    , coalesce(T3.""Stock"",0) as ""Stock_EXT1"" , Coalesce(T4.""Stock"" ,0) as ""Stock_EXT2"", Coalesce(T5.""Stock"" ,0) as ""Stock_EXT3"",
                                    Coalesce(T6.""Stock"" ,0) as ""Stock_EXT4"", Coalesce(T7.""Stock"" ,0) as ""Stock_EXT5"", 
                                    (Coalesce(T3.""STOCKCOBERTURA"", 0) + Coalesce(T4.""STOCKCOBERTURA"",0) + Coalesce(T5.""STOCKCOBERTURA"",0) + Coalesce(T6.""STOCKCOBERTURA"",0) + coalesce(T7.""STOCKCOBERTURA"",0) ) as ""ExternoSumaCobertura"",
                                    T1.""BinCode"" as ""Ubi_Defecto"" ,   T2.""Ventas_Med_Año"" , Coalesce( T2.""Ventas_Ult_Año"",0)as ""VA"",
                                    case when 
                                    (T0.""OnHand"" -  (coalesce(T3.""Stock"",0) + coalesce(T4.""Stock"",0) + coalesce(T5.""Stock"",0) + coalesce(T6.""Stock"",0) + coalesce(T7.""Stock"",0) ) + 
                                    (Coalesce(T3.""STOCKCOBERTURA"", 0) + Coalesce(T4.""STOCKCOBERTURA"",0) + Coalesce(T5.""STOCKCOBERTURA"",0) + Coalesce(T6.""STOCKCOBERTURA"",0) + coalesce(T7.""STOCKCOBERTURA"",0) ) )
                                    = 0 or T2.""Ventas_Med_Año"" = 0  then 0 else (T0.""OnHand"" -  (coalesce(T3.""Stock"",0) + coalesce(T4.""Stock"",0) + coalesce(T5.""Stock"",0) + coalesce(T6.""Stock"",0) + coalesce(T7.""Stock"",0) ) + 
                                    (Coalesce(T3.""STOCKCOBERTURA"", 0) + Coalesce(T4.""STOCKCOBERTURA"",0) + Coalesce(T5.""STOCKCOBERTURA"",0) + Coalesce(T6.""STOCKCOBERTURA"",0) + coalesce(T7.""STOCKCOBERTURA"",0) ) ) / T2.""Ventas_Med_Año"" end  as ""Cobertura""
                                    from OITW	T0
                                    LEFT JOIN  OBIN T1 ON T1.""AbsEntry"" = T0.""DftBinAbs""
                                    LEFT JOIN ""EXO_MRP_Ventas24Q"" T2 ON T2.""ItemCode"" = T0.""ItemCode"" and T1.""WhsCode"" = T2.""WhsCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA"", T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI"" = 'Ext1'
			                                    group by t12.""U_EXO_CALCOB"",T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T3 ON T3.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T3.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA"", T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext2'
			                                    group by t12.""U_EXO_CALCOB"",T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T4 ON T4.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T4.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA"", T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs""
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext3'
			                                    group by t12.""U_EXO_CALCOB"",T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T5 ON T5.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T5.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA""   ,T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext4'
			                                    group by t12.""U_EXO_CALCOB"", T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T6 ON T6.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T6.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA""  ,  T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext5'
			                                    group by t12.""U_EXO_CALCOB"", T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T7 ON T7.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T7.""ItemCode""
                                    )T WHERE ""Artículo""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' and ""Almacen""='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "' "
                        dVMA = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)

                        sSQL = " SELECT ""VA"" FROM (Select T0.""WhsCode"" as ""Almacen"", T0.""ItemCode"" as ""Artículo"", T0.""OnHand"",
                                    T0.""OnHand"" -  (coalesce(T3.""Stock"",0) + coalesce(T4.""Stock"",0) + coalesce(T5.""Stock"",0) + coalesce(T6.""Stock"",0) + coalesce(T7.""Stock"",0) ) AS ""STOCK_DENTRO""
                                    , coalesce(T3.""Stock"",0) as ""Stock_EXT1"" , Coalesce(T4.""Stock"" ,0) as ""Stock_EXT2"", Coalesce(T5.""Stock"" ,0) as ""Stock_EXT3"",
                                    Coalesce(T6.""Stock"" ,0) as ""Stock_EXT4"", Coalesce(T7.""Stock"" ,0) as ""Stock_EXT5"", 
                                    (Coalesce(T3.""STOCKCOBERTURA"", 0) + Coalesce(T4.""STOCKCOBERTURA"",0) + Coalesce(T5.""STOCKCOBERTURA"",0) + Coalesce(T6.""STOCKCOBERTURA"",0) + coalesce(T7.""STOCKCOBERTURA"",0) ) as ""ExternoSumaCobertura"",
                                    T1.""BinCode"" as ""Ubi_Defecto"" ,   T2.""Ventas_Med_Año"" , Coalesce( T2.""Ventas_Ult_Año"",0)as ""VA"",
                                    case when 
                                    (T0.""OnHand"" -  (coalesce(T3.""Stock"",0) + coalesce(T4.""Stock"",0) + coalesce(T5.""Stock"",0) + coalesce(T6.""Stock"",0) + coalesce(T7.""Stock"",0) ) + 
                                    (Coalesce(T3.""STOCKCOBERTURA"", 0) + Coalesce(T4.""STOCKCOBERTURA"",0) + Coalesce(T5.""STOCKCOBERTURA"",0) + Coalesce(T6.""STOCKCOBERTURA"",0) + coalesce(T7.""STOCKCOBERTURA"",0) ) )
                                    = 0 or T2.""Ventas_Med_Año"" = 0  then 0 else (T0.""OnHand"" -  (coalesce(T3.""Stock"",0) + coalesce(T4.""Stock"",0) + coalesce(T5.""Stock"",0) + coalesce(T6.""Stock"",0) + coalesce(T7.""Stock"",0) ) + 
                                    (Coalesce(T3.""STOCKCOBERTURA"", 0) + Coalesce(T4.""STOCKCOBERTURA"",0) + Coalesce(T5.""STOCKCOBERTURA"",0) + Coalesce(T6.""STOCKCOBERTURA"",0) + coalesce(T7.""STOCKCOBERTURA"",0) ) ) / T2.""Ventas_Med_Año"" end  as ""Cobertura""
                                    from OITW	T0
                                    LEFT JOIN  OBIN T1 ON T1.""AbsEntry"" = T0.""DftBinAbs""
                                    LEFT JOIN ""EXO_MRP_Ventas24Q"" T2 ON T2.""ItemCode"" = T0.""ItemCode"" and T1.""WhsCode"" = T2.""WhsCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA"", T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI"" = 'Ext1'
			                                    group by t12.""U_EXO_CALCOB"",T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T3 ON T3.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T3.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA"", T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext2'
			                                    group by t12.""U_EXO_CALCOB"",T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T4 ON T4.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T4.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA"", T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs""
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext3'
			                                    group by t12.""U_EXO_CALCOB"",T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T5 ON T5.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T5.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA""   ,T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext4'
			                                    group by t12.""U_EXO_CALCOB"", T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T6 ON T6.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T6.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA""  ,  T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext5'
			                                    group by t12.""U_EXO_CALCOB"", T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T7 ON T7.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T7.""ItemCode""
                                    )T WHERE ""Artículo""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' and ""Almacen""='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "' "
                        dVA = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)

                        sSQL = " SELECT ""Cobertura"" FROM (Select T0.""WhsCode"" as ""Almacen"", T0.""ItemCode"" as ""Artículo"", T0.""OnHand"",
                                    T0.""OnHand"" -  (coalesce(T3.""Stock"",0) + coalesce(T4.""Stock"",0) + coalesce(T5.""Stock"",0) + coalesce(T6.""Stock"",0) + coalesce(T7.""Stock"",0) ) AS ""STOCK_DENTRO""
                                    , coalesce(T3.""Stock"",0) as ""Stock_EXT1"" , Coalesce(T4.""Stock"" ,0) as ""Stock_EXT2"", Coalesce(T5.""Stock"" ,0) as ""Stock_EXT3"",
                                    Coalesce(T6.""Stock"" ,0) as ""Stock_EXT4"", Coalesce(T7.""Stock"" ,0) as ""Stock_EXT5"", 
                                    (Coalesce(T3.""STOCKCOBERTURA"", 0) + Coalesce(T4.""STOCKCOBERTURA"",0) + Coalesce(T5.""STOCKCOBERTURA"",0) + Coalesce(T6.""STOCKCOBERTURA"",0) + coalesce(T7.""STOCKCOBERTURA"",0) ) as ""ExternoSumaCobertura"",
                                    T1.""BinCode"" as ""Ubi_Defecto"" ,   T2.""Ventas_Med_Año"" , Coalesce( T2.""Ventas_Ult_Año"",0)as ""VA"",
                                    case when 
                                    (T0.""OnHand"" -  (coalesce(T3.""Stock"",0) + coalesce(T4.""Stock"",0) + coalesce(T5.""Stock"",0) + coalesce(T6.""Stock"",0) + coalesce(T7.""Stock"",0) ) + 
                                    (Coalesce(T3.""STOCKCOBERTURA"", 0) + Coalesce(T4.""STOCKCOBERTURA"",0) + Coalesce(T5.""STOCKCOBERTURA"",0) + Coalesce(T6.""STOCKCOBERTURA"",0) + coalesce(T7.""STOCKCOBERTURA"",0) ) )
                                    = 0 or T2.""Ventas_Med_Año"" = 0  then 0 else (T0.""OnHand"" -  (coalesce(T3.""Stock"",0) + coalesce(T4.""Stock"",0) + coalesce(T5.""Stock"",0) + coalesce(T6.""Stock"",0) + coalesce(T7.""Stock"",0) ) + 
                                    (Coalesce(T3.""STOCKCOBERTURA"", 0) + Coalesce(T4.""STOCKCOBERTURA"",0) + Coalesce(T5.""STOCKCOBERTURA"",0) + Coalesce(T6.""STOCKCOBERTURA"",0) + coalesce(T7.""STOCKCOBERTURA"",0) ) ) / T2.""Ventas_Med_Año"" end  as ""Cobertura""
                                    from OITW	T0
                                    LEFT JOIN  OBIN T1 ON T1.""AbsEntry"" = T0.""DftBinAbs""
                                    LEFT JOIN ""EXO_MRP_Ventas24Q"" T2 ON T2.""ItemCode"" = T0.""ItemCode"" and T1.""WhsCode"" = T2.""WhsCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA"", T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI"" = 'Ext1'
			                                    group by t12.""U_EXO_CALCOB"",T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T3 ON T3.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T3.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA"", T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext2'
			                                    group by t12.""U_EXO_CALCOB"",T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T4 ON T4.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T4.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA"", T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs""
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext3'
			                                    group by t12.""U_EXO_CALCOB"",T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T5 ON T5.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T5.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA""   ,T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext4'
			                                    group by t12.""U_EXO_CALCOB"", T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T6 ON T6.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T6.""ItemCode""
                                    LEFT JOIN (Select  case when t12.""U_EXO_CALCOB"" = 'Si' then SUM(T10.""OnHandQty"") else 0 end as ""STOCKCOBERTURA""  ,  T10.""WhsCode"", T10.""ItemCode"", 
                                                SUM(T10.""OnHandQty"") as ""Stock"", T12.""U_EXO_TIPOUBI"" as ""Zona_Ext"" from OIBQ T10
    		                                    LEFT JOIN  OBIN T11 ON T11.""AbsEntry"" = T10.""BinAbs"" 
			                                    INNER JOIN ""@EXO_UBIEXTERNAS"" T12 ON T11.""SL1Code"" = T12.""U_EXO_ZONA"" and T12.""U_EXO_ALM"" = T10.""WhsCode""
			                                    where T12.""U_EXO_TIPOUBI"" is not null and T12.""U_EXO_TIPOUBI""  = 'Ext5'
			                                    group by t12.""U_EXO_CALCOB"", T10.""WhsCode"", T10.""ItemCode"",T12.""U_EXO_TIPOUBI""
                                              ) T7 ON T7.""WhsCode"" = T0.""WhsCode"" and T0.""ItemCode"" = T7.""ItemCode""
                                    )T WHERE ""Artículo""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' and ""Almacen""='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "' "
                        dCober = oObjGlobal.refDi.SQL.sqlNumericaB1(sSQL)
#End Region
                        sSQL = "Select TOP 1 ""Substitute"" FROM ""OSCN"" WHERE ""ItemCode""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' and ""CardCode""='" & sIc & "' "
                        sCatalogo = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)

                        sSQL = "SELECT IFNULL(MAX(""LineId""),0) FROM ""@EXO_PACKINGL"" WHERE ""Code""='" & sDocEntry & sObjType & "'"
                        Dim sNLineas As String = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)

                        sSQL = "INSERT INTO ""@EXO_PACKINGL"" (""Code"", ""LineId"", ""U_EXO_LINEA"",""Object"", ""LogInst"", ""U_EXO_USUARIO"", ""U_EXO_CAT"", ""U_EXO_CODE"", ""U_EXO_CANT"", 
                                ""U_EXO_LOTE"", ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"",""U_EXO_ALM"",""U_EXO_STOCK"",""U_EXO_STOCKDENTRO"", 
                                ""U_EXO_EXT1"", ""U_EXO_EXT2"", ""U_EXO_EXT3"", ""U_EXO_EXT4"",""U_EXO_EXT5"",""U_EXO_UBIDEF"", ""U_EXO_TIPOHUECODEF"",""U_EXO_CANTMAXDEF"", ""U_EXO_STOCKUBIDEF"",
                                ""U_EXO_UBIORI"", ""U_EXO_VMA"", ""U_EXO_VA"", ""U_EXO_COBER"") ) 
                                Select '" & sDocEntry & sObjType & "' ""Code"", " & sNLineas & "+ ROW_NUMBER ( ) OVER( ORDER BY ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"",""U_EXO_ITEMCODE"",""U_EXO_LOTE"", ""U_EXO_UBICA"" ASC ) AS ""ROW_ID"",
                                '" & oDtLin.Rows.Item(iLin).Item("LineNum").ToString & "' ""U_EXO_LINEA"", 'EXO_PACKING', '0', 
                                '" & objGlobal.compañia.UserSignature.ToString & "' ""USUARIO"", '" & sCatalogo & "' ""U_EXO_CAT"", ""U_EXO_ITEMCODE"", SUM(""U_EXO_CANT"") ""U_EXO_CANT"", ""U_EXO_LOTE"",  
                                ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"", 
                                '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "', " & EXO_GLOBALES.DblNumberToText(oCompany, dStock, EXO_GLOBALES.FuenteInformacion.Otros) & " ""STOCK"" 
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStock - dStockExt, EXO_GLOBALES.FuenteInformacion.Otros) & " ""STOCKDENTRO""  
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStockExt1, EXO_GLOBALES.FuenteInformacion.Otros) & " ""EXT1""  
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStockExt2, EXO_GLOBALES.FuenteInformacion.Otros) & " ""EXT2""  
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStockExt3, EXO_GLOBALES.FuenteInformacion.Otros) & " ""EXT3""  
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStockExt4, EXO_GLOBALES.FuenteInformacion.Otros) & " ""EXT4""  
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStockExt5, EXO_GLOBALES.FuenteInformacion.Otros) & " ""EXT5""
                                , '" & sUBIDEF & "', '" & sTIPOHUECODEF & "', " & EXO_GLOBALES.DblNumberToText(oCompany, dCANTMAXDEF, EXO_GLOBALES.FuenteInformacion.Otros) & "
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dSTOCKUBIDEF, EXO_GLOBALES.FuenteInformacion.Otros) & "
                                , ""U_EXO_UBICA""
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dVMA, EXO_GLOBALES.FuenteInformacion.Otros) & "
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dVA, EXO_GLOBALES.FuenteInformacion.Otros) & "
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dCober, EXO_GLOBALES.FuenteInformacion.Otros) & "
                                FROM ""@EXO_LSTEMBL"" 
                                where ""DocEntry""=" & sListaEmbalaje & " and ""U_EXO_ITEMCODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' 
                                and ""U_EXO_DOCENTRY""=" & oDtLin.Rows.Item(iLin).Item("DocEntry").ToString & "
                                GROUP BY ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"",""U_EXO_ITEMCODE"",""U_EXO_LOTE"", ""U_EXO_UBICA"" "
                        oObjGlobal.refDi.SQL.sqlUpdB1(sSQL)

                    Next
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - Fin Proceso Packing List...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    If oObjGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
                        oObjGlobal.SBOApp.ActivateMenuItem("1304")
                    End If
                End If

            Else
                oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - No se ha podido insertar en la Sol. de traslado el Packing List.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oDtLin.Clear()

        End Try
    End Sub
End Class
