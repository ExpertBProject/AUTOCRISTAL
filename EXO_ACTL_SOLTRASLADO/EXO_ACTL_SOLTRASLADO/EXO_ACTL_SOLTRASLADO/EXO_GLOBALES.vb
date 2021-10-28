Imports SAPbouiCOM
Imports System.IO
Public Class EXO_GLOBALES

    Public Enum FuenteInformacion
        Visual = 1
        Otros = 2
    End Enum

#Region "Funciones formateos datos"
    Public Shared Function DblTextToNumber(ByRef oCompany As SAPbobsCOM.Company, ByVal sValor As String) As Double
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim cValor As Double = 0
        Dim sValorAux As String = "0"
        Dim sSeparadorMillarB1 As String = "."
        Dim sSeparadorDecimalB1 As String = ","
        Dim sSeparadorDecimalSO As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator

        DblTextToNumber = 0

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT COALESCE(""DecSep"", ',') ""DecSep"", COALESCE(""ThousSep"", '.') ""ThousSep"" " &
                   "FROM ""OADM"" " &
                   "WHERE ""Code"" = 1"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                sSeparadorMillarB1 = oRs.Fields.Item("ThousSep").Value.ToString
                sSeparadorDecimalB1 = oRs.Fields.Item("DecSep").Value.ToString
            End If

            sValorAux = sValor

            If sSeparadorDecimalSO = "," Then
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "." Then sValorAux = "0" & sValorAux

                    If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ""))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "").Replace(".", ","))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            Else
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "," Then sValorAux = "0" & sValorAux

                    If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", "").Replace(",", "."))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", ""))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            End If

            DblTextToNumber = cValor

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Shared Function DblNumberToText(ByRef oCompany As SAPbobsCOM.Company, ByVal cValor As Double, ByVal oDestino As FuenteInformacion) As String
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim sNumberDouble As String = "0"
        Dim sSeparadorMillarB1 As String = "."
        Dim sSeparadorDecimalB1 As String = ","
        Dim sSeparadorDecimalSO As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator

        DblNumberToText = "0"

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT COALESCE(""DecSep"", ',') ""DecSep"", COALESCE(""ThousSep"", '.') ""ThousSep"" " &
                   "FROM ""OADM"" " &
                   "WHERE ""Code"" = 1"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                sSeparadorMillarB1 = oRs.Fields.Item("ThousSep").Value.ToString
                sSeparadorDecimalB1 = oRs.Fields.Item("DecSep").Value.ToString
            End If

            If cValor.ToString <> "" Then
                If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
                    sNumberDouble = cValor.ToString
                Else 'Decimales USA
                    sNumberDouble = cValor.ToString.Replace(",", ".")
                End If
            End If

            If oDestino = FuenteInformacion.Visual Then
                If sSeparadorDecimalSO = "," Then
                    DblNumberToText = sNumberDouble
                Else
                    DblNumberToText = sNumberDouble.Replace(".", ",")
                End If
            Else
                DblNumberToText = sNumberDouble.Replace(",", ".")
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
#End Region

#Region "Solicitud de traslado"
    Public Shared Sub Crear_Sol_Traslado(ByRef oobjglobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form, ByRef sDocEntrySol As String, ByRef sDocNumSol As String)
#Region "Variables"
        Dim sExiste As String = "" : Dim sSQL As String = ""
        Dim oDtLin As System.Data.DataTable = New System.Data.DataTable
        Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sObjtype As String = "" : Dim sTabla As String = "" : Dim sTablaL As String = ""
        Dim sSerie As String = "" : Dim sIndicator As String = "" : Dim sSucursal As String = ""
        Dim oDoc As SAPbobsCOM.StockTransfer = Nothing : Dim iCuenta As Integer = 0
        Dim sMensaje As String = ""
        Dim oRsLote As SAPbobsCOM.Recordset = Nothing
#End Region
        Try

            oRsLote = CType(oobjglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Try
                sTablaL = CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("1").DataBind.TableName
                sTabla = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).DataBind.TableName
            Catch ex As Exception
                sTablaL = CType(oForm.Items.Item("23").Specific, SAPbouiCOM.Matrix).Columns.Item("1").DataBind.TableName
                sTabla = CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).DataBind.TableName
            End Try
            sDocEntry = oForm.DataSources.DBDataSources.Item(sTabla).GetValue("DocEntry", 0).ToString.Trim
            sDocNum = oForm.DataSources.DBDataSources.Item(sTabla).GetValue("DocNum", 0).ToString.Trim
            sObjtype = oForm.DataSources.DBDataSources.Item(sTabla).GetValue("ObjType", 0).ToString.Trim
            sSerie = oForm.DataSources.DBDataSources.Item(sTabla).GetValue("Series", 0).ToString.Trim
            If oobjglobal.SBOApp.MessageBox("¿Está seguro que quiere agregar el fichero en la lista activa?", 1, "Sí", "No") = 1 Then
                oobjglobal.SBOApp.StatusBar.SetText("(EXO) - Generando Sol. de Traslado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oForm.Freeze(True)
                If oobjglobal.compañia.InTransaction = True Then
                    oobjglobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                Else
                    oobjglobal.compañia.StartTransaction()
                End If
#Region "Crear Sol. de traslado"
                oDoc = CType(oobjglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest
                oDoc.CardCode = oForm.DataSources.DBDataSources.Item(sTabla).GetValue("CardCode", 0).ToString.Trim
                Select Case sTabla
                    Case "ORDN" : oDoc.Comments = oForm.DataSources.DBDataSources.Item(sTabla).GetValue("Comments", 0).ToString.Trim & ChrW(10) & ChrW(13) & "Basado en la Devolución del cliente Nº" & oForm.DataSources.DBDataSources.Item(sTabla).GetValue("DocNum", 0).ToString.Trim
                    Case "OWTR" : oDoc.Comments = oForm.DataSources.DBDataSources.Item(sTabla).GetValue("Comments", 0).ToString.Trim & ChrW(10) & ChrW(13) & "Basado en el Traslado Nº" & oForm.DataSources.DBDataSources.Item(sTabla).GetValue("DocNum", 0).ToString.Trim
                    Case Else
                        sMensaje = "Error no encuentra la tabla para buscar los datos.Se cancela el proceso."
                        oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oobjglobal.SBOApp.MessageBox(sMensaje)
                        Exit Sub
                End Select

                oDtLin.Clear()
                sSQL = "SELECT * FROM """ & sTablaL & """ "
                sSQL &= " where ""DocEntry""='" & sDocEntry & "' Order by ""LineNum"" "
                oDtLin = oobjglobal.refDi.SQL.sqlComoDataTable(sSQL)
                If oDtLin.Rows.Count > 0 Then
                    Dim bPlinea As Boolean = True
                    For iLin As Integer = 0 To oDtLin.Rows.Count - 1
#Region "Buscar la serie"
                        sSQL = "SELECT ""Indicator"" FROM ""NNM1"" WHERE ""ObjectCode""='" & sObjtype & "' and ""Series""='" & sSerie & "' "
                        sIndicator = oobjglobal.refDi.SQL.sqlStringB1(sSQL)
                        sSQL = "SELECT ""U_EXO_SUCURSAL"" FROM ""OWHS"" WHERE ""WhsCode""='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "' "
                        sSucursal = oobjglobal.refDi.SQL.sqlStringB1(sSQL)
                        sSQL = "SELECT ""Series"" FROM ""NNM1"" WHERE ""Indicator""='" & sIndicator & "' and ""Remark""='" & sSucursal & "' and ""ObjectCode""='1250000001' "
                        sSerie = oobjglobal.refDi.SQL.sqlStringB1(sSQL)
                        If sSerie = "" Then
                            If oobjglobal.SBOApp.MessageBox("No encuentra la serie del almacén " & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "¿Continuamos con la serie primaria?", 1, "Sí", "No") = 1 Then
                                sSQL = "SELECT ""Series"" FROM ""NNM1"" WHERE ""SeriesName""='Primario' and ""ObjectCode""='1250000001' "
                                sSerie = oobjglobal.refDi.SQL.sqlStringB1(sSQL)
                                oDoc.Series = sSerie
                            Else
                                sMensaje = "El usuario ha cancelado el proceso al no encontrar la serie correspondiente al almacén " & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString
                                oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                oobjglobal.SBOApp.MessageBox(sMensaje)
                                Exit Sub
                            End If
                        Else
                            oDoc.Series = sSerie
                        End If
#End Region
                        oDoc.FromWarehouse = oDtLin.Rows.Item(iLin).Item("WhsCode").ToString
                        oDoc.ToWarehouse = oDtLin.Rows.Item(iLin).Item("WhsCode").ToString
                        If bPlinea = False Then
                            oDoc.Lines.Add()
                        Else
                            bPlinea = False
                        End If
                        oDoc.Lines.ItemCode = oDtLin.Rows.Item(iLin).Item("ItemCode").ToString
                        oDoc.Lines.Quantity = EXO_GLOBALES.DblTextToNumber(oobjglobal.compañia, oDtLin.Rows.Item(iLin).Item("Quantity").ToString)
                        oDoc.Lines.UserFields.Fields.Item("U_EXO_LOT_ID").Value = oDtLin.Rows.Item(iLin).Item("U_EXO_LOT_ID").ToString
                        oDoc.Lines.UserFields.Fields.Item("U_EXO_TBULTO").Value = oDtLin.Rows.Item(iLin).Item("U_EXO_TBULTO").ToString

                        Select Case sTabla
                            Case "ORDN"
#Region "Lotes"
                                'Incluimos los Lotes
                                sSQL = "Select ""OBTN"".""DistNumber"",""ITL1"".* FROM ""OITL"" INNER JOIN ""ITL1"" On ""ITL1"".""LogEntry""=""OITL"".""LogEntry"" "
                                sSQL &= " Left Join ""OBTN"" On ""OBTN"".""AbsEntry""=""ITL1"".""SysNumber"" "
                                sSQL &= " WHERE ""DocEntry"" = " & oDtLin.Rows.Item(iLin).Item("DocEntry").ToString & " And ""DocLine"" =" & oDtLin.Rows.Item(iLin).Item("LineNum").ToString
                                sSQL &= " And ""DocType""='" & sObjtype & "' and ""LocCode""='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "'"
                                oRsLote.DoQuery(sSQL)
                                For iLote = 1 To oRsLote.RecordCount
                                    'Creamos el lote de la línea del artículo
                                    oDoc.Lines.BatchNumbers.BatchNumber = oRsLote.Fields.Item("DistNumber").Value.ToString
                                    oDoc.Lines.BatchNumbers.Quantity = EXO_GLOBALES.DblTextToNumber(oobjglobal.compañia, oRsLote.Fields.Item("Quantity").Value.ToString)
                                    oDoc.Lines.BatchNumbers.Add()
                                    oRsLote.MoveNext()
                                Next
#End Region
                            Case "OWTR"
                                oDoc.Lines.BaseType = SAPbobsCOM.InvBaseDocTypeEnum.WarehouseTransfers
                                oDoc.Lines.BaseEntry = CType(oDtLin.Rows.Item(iLin).Item("DocEntry").ToString, Integer)
                                oDoc.Lines.BaseLine = CType(oDtLin.Rows.Item(iLin).Item("LineNum").ToString, Integer)
                            Case Else
                                sMensaje = "Error no encuentra la tabla para indicar en la línea de donde proviene.Se cancela el proceso."
                                oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oobjglobal.SBOApp.MessageBox(sMensaje)
                                Exit Sub
                        End Select


                    Next
#Region "Doc. Enlazado o de referencia"
                    ''Buscamos si hay mas de una línea en los adjuntos
                    'sSQL = "SELECT COUNT(*) ""Cuenta"" FROM ""RDN21"" where ""DocEntry""=" & sDocEntry & " and ""ObjectType""='" & sObjtype & "' "
                    'iCuenta = oobjglobal.refDi.SQL.sqlNumericaB1(sSQL)
                    'If iCuenta > 0 Then
                    '    oDoc.DocumentReferences.Add() 'Solo si hay más de 2 líneas (A partir de la segunda línea..)
                    'End If
                    Select Case sTabla
                        Case "ORDN" : oDoc.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_Return
                        Case "OWTR" : oDoc.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_InventoryTransfer
                        Case Else
                            sMensaje = "Error no encuentra la tabla para indicar en el documento de referencia.Se cancela el proceso."
                            oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oobjglobal.SBOApp.MessageBox(sMensaje)
                            Exit Sub
                    End Select
                    oDoc.DocumentReferences.ReferencedDocEntry = sDocEntry
                    'oDoc.DocumentReferences.ReferencedDocNumber = sDocNum
                    oDoc.DocumentReferences.Remark = ""

#End Region
                    If oDoc.Add() <> 0 Then
                        sMensaje = oobjglobal.compañia.GetLastErrorCode.ToString & " / " & oobjglobal.compañia.GetLastErrorDescription.Replace("'", "")
                        oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Else
                        oobjglobal.compañia.GetNewObjectCode(sDocEntrySol)

                        sSQL = "SELECT ""DocNum"" FROM ""OWTQ""  WHERE ""DocEntry""=" & sDocEntrySol
                        sDocNumSol = oobjglobal.refDi.SQL.sqlStringB1(sSQL)
                        sMensaje = "Se ha generado correctamente la sol. de traslado con Nº " & sDocNumSol & " y Nº interno " & sDocEntrySol.ToString
                        oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                Else
                    sMensaje = "No se encuentra las líneas del documento Nº" & oForm.DataSources.DBDataSources.Item(sTabla).GetValue("DocNum", 0).ToString.Trim & ". Se interrumpe el proceso."
                    oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oobjglobal.SBOApp.MessageBox(sMensaje)
                End If
#End Region
                If oobjglobal.compañia.InTransaction = True Then
                    oobjglobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If

            Else
                oobjglobal.SBOApp.StatusBar.SetText("(EXO) - El usuario ha cancelado la generación de la solicitud de traslado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
            oDtLin.Clear()
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLote, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDoc, Object))

            If oobjglobal.compañia.InTransaction = True Then
                oobjglobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        End Try
    End Sub
    Public Shared Sub ASIGNAR_DOCREF(ByRef oobjglobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form, ByRef sDocEntrySol As String, ByRef sDocNumSol As String)
#Region "Variables"
        'Dim sExiste As String = "" : Dim sSQL As String = ""
        Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sObjtype As String = ""
        Dim sTabla As String = "" : Dim sTablaL As String = ""
        'Dim oDoc As SAPbobsCOM.Documents = Nothing
        Dim sMensaje As String = ""
        Dim sSQL As String = "" : Dim iCuenta As Integer = 0 : Dim sTabla_Adjunto As String = ""
        Dim oDoc As SAPbobsCOM.Documents = Nothing
        Dim oDocTransfer As SAPbobsCOM.StockTransfer = Nothing
#End Region
        Try
            If sDocEntrySol <> "" Then
                oForm.Freeze(True)
                Try
                    sTablaL = CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("1").DataBind.TableName
                    sTabla = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).DataBind.TableName
                Catch ex As Exception
                    sTablaL = CType(oForm.Items.Item("23").Specific, SAPbouiCOM.Matrix).Columns.Item("1").DataBind.TableName
                    sTabla = CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).DataBind.TableName
                End Try
                sDocEntry = oForm.DataSources.DBDataSources.Item(sTabla).GetValue("DocEntry", 0).ToString.Trim
                sDocNum = oForm.DataSources.DBDataSources.Item(sTabla).GetValue("DocNum", 0).ToString.Trim
                sObjtype = oForm.DataSources.DBDataSources.Item(sTabla).GetValue("ObjType", 0).ToString.Trim

                Select Case sTabla
                    Case "ORDN"
                        oDoc = CType(oobjglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns), SAPbobsCOM.Documents)
                        sTabla_Adjunto = "RDN21"
                        If oDoc.GetByKey(sDocEntry) = True Then
#Region "Doc. Enlazado o de referencia"
                            'Buscamos si hay mas de una línea en los adjuntos
                            sSQL = "SELECT COUNT(*) ""Cuenta"" FROM """ & sTabla_Adjunto & """ where ""DocEntry""=" & sDocEntry & " and ""ObjectType""='" & sObjtype & "' "
                            iCuenta = oobjglobal.refDi.SQL.sqlNumericaB1(sSQL)
                            If iCuenta > 0 Then
                                oDoc.DocumentReferences.Add() 'Solo si hay más de 2 líneas (A partir de la segunda línea..)
                            End If
                            oDoc.DocumentReferences.ReferencedObjectType = "1250000001"
                            oDoc.DocumentReferences.ReferencedDocEntry = sDocEntrySol
                            'oDoc.DocumentReferences.ReferencedDocNumber = sDocNum
                            oDoc.DocumentReferences.Remark = ""
#End Region
                            If oDoc.Update() <> 0 Then
                                sMensaje = oobjglobal.compañia.GetLastErrorCode.ToString & " / " & oobjglobal.compañia.GetLastErrorDescription.Replace("'", "")
                                oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else
                                sMensaje = "Se ha asignado la sol. de traslado con Nº " & sDocNumSol & " y al Documento Nº " & sDocNum
                                oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If
                        End If
                    Case "OWTR"
                        oDocTransfer = CType(oobjglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer), SAPbobsCOM.StockTransfer)
                        sTabla_Adjunto = "WTR21"
                        If oDocTransfer.GetByKey(sDocEntry) = True Then
#Region "Doc. Enlazado o de referencia"
                            'Buscamos si hay mas de una línea en los adjuntos
                            sSQL = "SELECT COUNT(*) ""Cuenta"" FROM """ & sTabla_Adjunto & """ where ""DocEntry""=" & sDocEntry & " and ""ObjectType""='" & sObjtype & "' "
                            iCuenta = oobjglobal.refDi.SQL.sqlNumericaB1(sSQL)
                            If iCuenta > 0 Then
                                oDoc.DocumentReferences.Add() 'Solo si hay más de 2 líneas (A partir de la segunda línea..)
                            End If
                            oDocTransfer.DocumentReferences.ReferencedObjectType = "1250000001"
                            oDocTransfer.DocumentReferences.ReferencedDocEntry = sDocEntrySol
                            oDocTransfer.DocumentReferences.Remark = ""
#End Region
                            If oDocTransfer.Update() <> 0 Then
                                sMensaje = oobjglobal.compañia.GetLastErrorCode.ToString & " / " & oobjglobal.compañia.GetLastErrorDescription.Replace("'", "")
                                oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else
                                sMensaje = "Se ha asignado la sol. de traslado con Nº " & sDocNumSol & " y al Documento Nº " & sDocNum
                                oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If
                        End If
                    Case Else
                        sMensaje = "Error no encuentra la tabla para buscar los datos.Se cancela el proceso."
                        oobjglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oobjglobal.SBOApp.MessageBox(sMensaje)
                        Exit Sub
                End Select
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDoc, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDocTransfer, Object))
            oForm.Freeze(False)
        End Try
    End Sub
#End Region
End Class
