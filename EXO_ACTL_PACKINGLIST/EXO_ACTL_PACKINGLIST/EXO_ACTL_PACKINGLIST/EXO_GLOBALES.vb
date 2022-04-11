Imports SAPbouiCOM
Imports SAPbobsCOM
Imports System.IO
Imports OfficeOpenXml

Public Class EXO_GLOBALES
#Region "Variables Globales"
    Public Shared _sPedido As String = ""
    Public Shared _sIc As String = ""
#End Region
    Public Enum FuenteInformacion
        Visual = 1
        Otros = 2
    End Enum
#Region "Funciones formateos datos"
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
    Public Shared Function FormateaString(ByVal dato As Object, ByVal tam As Integer) As String
        Dim retorno As String = String.Empty

        If dato IsNot Nothing Then
            retorno = dato.ToString
        End If

        If retorno.Length > tam Then
            retorno = retorno.Substring(0, tam)
        End If

        Return retorno.PadRight(tam, CChar(" "))
    End Function
    Public Shared Function FormateaNumero(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""

        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            retorno = retorno.Replace("-", "N")
            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If

        Return retorno
    End Function
    Public Shared Function FormateaNumeroSinPunto(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""
        Dim bEsNegativo As Boolean = False
        If Left(dato, 1) = "-" Then
            dato = dato.Replace("-", "")
            bEsNegativo = True
            posiciones = posiciones - 1
            totalNum = posiciones
        End If
        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            retorno = retorno.Replace(".", "")

            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If
        If bEsNegativo = True Then
            retorno = "N" & retorno
        End If
        Return retorno
    End Function
#End Region
#Region "Fichero"
    Public Shared Sub TratarFichero_TXT(ByVal sArchivo As String, ByVal sDelimitador As String, ByRef oForm As SAPbouiCOM.Form, ByRef oCompany As SAPbobsCOM.Company, ByRef oSboApp As Application, ByRef objglobal As EXO_UIAPI.EXO_UIAPI)
#Region "Variables"
        ' Apuntador libre a archivo
        Dim Apunt As Integer = FreeFile()
        Dim sMensaje As String = ""
        Dim sArticulo As String = "" : Dim sCatalogo As String = "" : Dim sCantidad As String = ""
        Dim sLote As String = "" : Dim sFFab As String = "" : Dim sIdBulto As String = "" : Dim sTBulto As String = ""
        Dim iLinea As Integer = 0
        Dim sExiste As String = ""
        Dim sSQL As String = ""
#End Region
        Try
            ' miramos si existe el fichero y cargamos
            If File.Exists(sArchivo) Then
                Using MyReader As New Microsoft.VisualBasic.
                        FileIO.TextFieldParser(sArchivo, System.Text.Encoding.UTF7)
                    MyReader.TextFieldType = FileIO.FieldType.Delimited
                    Select Case sDelimitador
                        Case "1" : MyReader.SetDelimiters(vbTab)
                        Case "2" : MyReader.SetDelimiters(";")
                        Case "3" : MyReader.SetDelimiters(",")
                        Case "4" : MyReader.SetDelimiters("-")
                        Case Else : MyReader.SetDelimiters(vbTab)
                    End Select

                    Dim currentRow As String()
                    iLinea = 0
                    While Not MyReader.EndOfData
                        Try
                            If iLinea = 0 Then ' Para quitar la cabecera
                                currentRow = MyReader.ReadFields()
                            End If
                            iLinea += 1
                            currentRow = MyReader.ReadFields()

                            Dim currentField As String
                            Dim scampos(1) As String
                            Dim iCampo As Integer = 0
                            For Each currentField In currentRow
                                iCampo += 1
                                ReDim Preserve scampos(iCampo)
                                scampos(iCampo) = currentField
                                'SboApp.MessageBox(scampos(iCampo))
                            Next
#Region "Lectura registros"
                            For i = 1 To iCampo
                                Select Case i
                                    Case 1 : sCatalogo = scampos(i)
                                    Case 2 : sArticulo = scampos(i)
                                    Case 3 : sCantidad = scampos(i)
                                    Case 4 : sLote = scampos(i)
                                    Case 5 : sFFab = scampos(i)
                                    Case 6 : sIdBulto = scampos(i)
                                    Case 7 : sTBulto = scampos(i)
                                End Select
                            Next
                            If sCatalogo = "" And sArticulo = "" Then
                                sMensaje = "En la línea " & iLinea.ToString & " no puede estar vacío el artículo y el catálogo. Revise el fichero."
                                objglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oSboApp.MessageBox(sMensaje)
                                Exit Sub
                            Else
                                If sCatalogo <> "" Then
                                    'Buscamos el artículo en un catálago de proveedor.
                                    sArticulo = objglobal.refDi.SQL.sqlStringB1("SELECT TOP 1 ""ItemCode"" FROM ""OSCN"" WHERE ""Substitute""='" & sCatalogo & "' and ""CardCode""='" & EXO_GLOBALES._sIc & "' ")

                                End If
                                If sArticulo = "" Then
                                    sMensaje = "En la línea " & iLinea.ToString & ", para el Catálogo " & sCatalogo & " no se encuentra el Artículo. Revise el fichero."
                                    objglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    oSboApp.MessageBox(sMensaje)
                                    Exit Sub
                                Else
                                    'Ponemos la fecha según sistema.
                                    If sFFab <> "" Then
                                        Dim dFecha As Date = New Date(CInt(Right(sFFab, 4)), CInt(Mid(sFFab, 4, 2)), CInt(Left(sFFab, 2)))
                                        sFFab = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                                    End If

                                    'Comprobamos que el bulto exista
                                    If sTBulto <> "" Then
                                        sExiste = objglobal.refDi.SQL.sqlStringB1("SELECT ""PkgType"" FROM ""OPKG"" Where ""PkgType""='" & sTBulto & "' ")
                                        If sExiste = "" Then
                                            sMensaje = "En la línea " & iLinea.ToString & ", el tipo de bulto " & sTBulto & " no está definido en SAP. Revise el fichero."
                                            objglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oSboApp.MessageBox(sMensaje)
                                            Exit Sub
                                        End If
                                    End If
                                    'Grabamos el registro
                                    sSQL = "insert into ""@EXO_TMPPACKINGL"" values(" & iLinea.ToString & ",'" & iLinea.ToString & "'," & iLinea.ToString & ",'N','',0,"
                                    sSQL &= objglobal.compañia.UserSignature.ToString & ",'','" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "',0,'',0,'',"
                                    sSQL &= "'" & objglobal.compañia.UserName.ToString & "','" & sCatalogo & "','" & sArticulo & "'," & sCantidad & ",'" & sLote & "','" & sFFab & "',"
                                    sSQL &= "'" & sIdBulto & "','" & sTBulto & "')"
                                    objglobal.refDi.SQL.sqlUpdB1(sSQL)
                                End If
                            End If
#End Region
                        Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                            objglobal.SBOApp.StatusBar.SetText("(EXO) - Línea " & ex.Message & " no es válida y se omitirá.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oSboApp.MessageBox("Línea " & ex.Message & " no es válida y se omitirá.")
                        End Try
                    End While
                End Using
            Else
                objglobal.SBOApp.StatusBar.SetText("(EXO) - No se ha encontrado el fichero txt a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            ' Cerramos el archivo
            FileClose(Apunt)
        End Try
    End Sub

    Public Shared Sub Generar_EM(ByRef oCompany As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI)
#Region "Variables"
        Dim oDtLinFichero As System.Data.DataTable = New System.Data.DataTable
        Dim oDtLin As System.Data.DataTable = New System.Data.DataTable
        Dim oOPDN As SAPbobsCOM.Documents = Nothing
        Dim dfecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim sSQL As String = "" : Dim sMensaje As String = "" : Dim sError As String = "" : Dim sComen As String = "" : Dim sEstado As String = ""
        Dim sSerie As String = "" : Dim sDocEntry As String = "" : Dim sDocnum As String = ""
        Dim oRsLote As SAPbobsCOM.Recordset = Nothing : Dim oRsArt As SAPbobsCOM.Recordset = Nothing
        Dim iTabla As Integer = 1
        Dim dCantLotes As Double = 0
#End Region

        Try
            oRsLote = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsArt = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            oCompany.StartTransaction()
            oOPDN = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts), SAPbobsCOM.Documents)
            oOPDN.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes
            oOPDN.CardCode = EXO_GLOBALES._sIc
            oOPDN.TaxDate = dfecha
            oOPDN.DocDueDate = dfecha
            oOPDN.NumAtCard = oObjGlobal.refDi.SQL.sqlStringB1("SELECT ""NumAtCard"" FROM ""OPOR"" WHERE ""DocEntry""=" & EXO_GLOBALES._sPedido)

            oDtLin.Clear()

            sSQL = "SELECT * FROM ""POR1"" where ""LineStatus""='O' and ""DocEntry""=" & EXO_GLOBALES._sPedido & " Order by ""LineNum"" "
            oDtLin = oObjGlobal.refDi.SQL.sqlComoDataTable(sSQL)

            If oDtLin.Rows.Count > 0 Then
                iTabla = 1
                Dim bPlinea As Boolean = True
                For iLin As Integer = 0 To oDtLin.Rows.Count - 1
                    'buscamos en la tabla de ficheros
                    oDtLinFichero.Clear()
                    sSQL = "SELECT ""U_EXO_CODE"",sum(""U_EXO_CANT"") ""CANTIDAD"",""U_EXO_TBULTO"" FROM ""@EXO_TMPPACKINGL"" "
                    sSQL &= " where ""U_EXO_USUARIO""='" & oCompany.UserName.ToString & "' and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' "
                    sSQL &= " GROUP BY ""U_EXO_CODE"",""U_EXO_TBULTO"" "
                    oDtLinFichero = oObjGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                    If oDtLinFichero.Rows.Count > 0 Then
                        If bPlinea = False Then
                            oOPDN.Lines.Add()
                        Else
                            bPlinea = False
                        End If
                        For iLinFich As Integer = 0 To oDtLinFichero.Rows.Count - 1
                            oOPDN.Lines.ItemCode = oDtLin.Rows.Item(iLin).Item("ItemCode").ToString
                            oOPDN.Lines.ItemDescription = oDtLin.Rows.Item(iLin).Item("Dscription").ToString
                            Dim dCantFichero As Double = EXO_GLOBALES.DblTextToNumber(oCompany, oDtLinFichero.Rows.Item(iLinFich).Item("CANTIDAD").ToString)
                            Dim dCant As Double = EXO_GLOBALES.DblTextToNumber(oCompany, oDtLin.Rows.Item(iLin).Item("Quantity").ToString)
                            Dim sUnidad As String = oDtLin.Rows.Item(iLin).Item("UomCode").ToString.Trim
                            Dim sUnidadFichero As String = oDtLinFichero.Rows.Item(iLinFich).Item("U_EXO_TBULTO").ToString

                            oOPDN.Lines.BaseEntry = CInt(oDtLin.Rows.Item(iLin).Item("DocEntry").ToString)
                            oOPDN.Lines.BaseType = 22
                            oOPDN.Lines.BaseLine = CInt(oDtLin.Rows.Item(iLin).Item("LineNum").ToString)
#Region "Lotes"
                            'Incluimos los Lotes
                            sSQL = "SELECT ""U_EXO_CODE"",""U_EXO_LOTE"", sum(""U_EXO_CANT"") ""CANTIDAD"",""U_EXO_TBULTO"",""U_EXO_FFAB"" FROM ""@EXO_TMPPACKINGL"" "
                            sSQL &= " where ""U_EXO_USUARIO""='" & oCompany.UserName.ToString & "' and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' "
                            sSQL &= " and ""U_EXO_TBULTO""='" & oDtLinFichero.Rows.Item(iLinFich).Item("U_EXO_TBULTO").ToString & "' "
                            sSQL &= " GROUP BY ""U_EXO_CODE"",""U_EXO_LOTE"", ""U_EXO_TBULTO"",""U_EXO_FFAB"" "
                            oRsLote.DoQuery(sSQL)
                            For iLote = 1 To oRsLote.RecordCount
                                'Creamos el lote de la línea del artículo
                                oOPDN.Lines.BatchNumbers.BatchNumber = oRsLote.Fields.Item("U_EXO_LOTE").Value.ToString
                                oOPDN.Lines.BatchNumbers.Quantity = EXO_GLOBALES.DblTextToNumber(oCompany, oRsLote.Fields.Item("CANTIDAD").Value.ToString)
                                dCantLotes += oOPDN.Lines.BatchNumbers.Quantity
                                oOPDN.Lines.BatchNumbers.ManufacturingDate = CDate(oRsLote.Fields.Item("U_EXO_FFAB").Value.ToString)
                                'sSQL = "SELECT STRING_AGG(""U_EXO_IDBULTO"",'; ') ""BULTOS"" FROM "
                                'sSQL &= " (SELECT DISTINCT ""U_EXO_IDBULTO"" FROM ""@EXO_TMPPACKINGL"" T0 "
                                'sSQL &= " WHERE ""U_EXO_USUARIO""='" & oCompany.UserName.ToString & "' and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' "
                                'sSQL &= " and ""U_EXO_TBULTO""='" & oDtLinFichero.Rows.Item(iLinFich).Item("U_EXO_TBULTO").ToString & "' "
                                'sSQL &= " and ""U_EXO_LOTE""='" & oRsLote.Fields.Item("U_EXO_LOTE").Value.ToString & "' "
                                'sSQL &= ") As T1"
                                'Dim sBultos As String = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                                ''oOPDN.Lines.BatchNumbers.Notes = "Bultos " & sBultos.Trim
                                'oOPDN.Lines.BatchNumbers.UserFields.Fields.Item("U_EXO_LOT_ID").Value = sBultos.Trim
                                'sSQL = "SELECT STRING_AGG(""U_EXO_CANT"",'; ') ""CANTIDADES"" FROM "
                                'sSQL &= " (SELECT DISTINCT ""U_EXO_CANT"" FROM ""@EXO_TMPPACKINGL"" T0 "
                                'sSQL &= " WHERE ""U_EXO_USUARIO""='" & oCompany.UserName.ToString & "' and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' "
                                'sSQL &= " and ""U_EXO_TBULTO""='" & oDtLinFichero.Rows.Item(iLinFich).Item("U_EXO_TBULTO").ToString & "' "
                                'sSQL &= " and ""U_EXO_LOTE""='" & oRsLote.Fields.Item("U_EXO_LOTE").Value.ToString & "' "
                                'sSQL &= ") As T1"
                                'Dim sCantidades As String = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                                'oOPDN.Lines.BatchNumbers.UserFields.Fields.Item("U_EXO_LOT_CAN").Value = sCantidades.Trim
                                'sSQL = "SELECT STRING_AGG(""U_EXO_TBULTO"",'; ') ""TBULTOS"" FROM "
                                'sSQL &= " (SELECT DISTINCT ""U_EXO_TBULTO"" FROM ""@EXO_TMPPACKINGL"" T0 "
                                'sSQL &= " WHERE ""U_EXO_USUARIO""='" & oCompany.UserName.ToString & "' and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' "
                                'sSQL &= " and ""U_EXO_TBULTO""='" & oDtLinFichero.Rows.Item(iLinFich).Item("U_EXO_TBULTO").ToString & "' "
                                'sSQL &= " and ""U_EXO_LOTE""='" & oRsLote.Fields.Item("U_EXO_LOTE").Value.ToString & "' "
                                'sSQL &= ") As T1"
                                'Dim sTBultos As String = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                                'oOPDN.Lines.BatchNumbers.UserFields.Fields.Item("U_EXO_LOT_CAN").Value = sTBultos.Trim

                                oOPDN.Lines.BatchNumbers.Add()
                                oRsLote.MoveNext()
                            Next
#End Region
                            If sUnidad = sUnidadFichero Then
                                If dCant <= dCantLotes Then
                                    oOPDN.Lines.Quantity = dCant
                                Else
                                    oOPDN.Lines.Quantity = dCantLotes
                                End If
                            Else
                                If dCant <= dCantLotes Then
                                    oOPDN.Lines.InventoryQuantity = dCant
                                Else
                                    oOPDN.Lines.InventoryQuantity = dCantLotes
                                End If
                            End If
                        Next
                    Else
                        sMensaje = "No se encuentra la línea " & oDtLin.Rows.Item(iLin).Item("LineNum").ToString & " con el artículo " & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString
                        sMensaje &= ". No se incluye en la Entrada de Mercancía."
                        oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                Next
                If oOPDN.Add() <> 0 Then
                    sEstado = "Error"
                    sError = oCompany.GetLastErrorCode.ToString & " / " & oCompany.GetLastErrorDescription.Replace("'", "")
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    sComen = sError : sDocEntry = "" : sSerie = "" : sDocnum = ""
                    'Creo línea para visualizar y el usuario sepa del error
                Else
                    oCompany.GetNewObjectCode(sDocEntry)

                    sSQL = "SELECT S.""SeriesName"" FROM ""ODRF"" D INNER JOIN  ""NNM1"" S ON S.""Series""=D.""Series"" WHERE D.""DocEntry""=" & sDocEntry
                    sSerie = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)

                    sSQL = "SELECT ""DocNum"" FROM ""ODRF""  WHERE ""DocEntry""=" & sDocEntry
                    sDocnum = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    sEstado = "OK"
                    sComen = "Se ha generado correctamente el borrador de la entrada de mercancía con Nº " & sDocnum
#Region "Pasamos los datos del fichero a la tabla real"
                    oDtLin.Clear()
                    sSQL = "SELECT * FROM ""DRF1""  WHERE ""DocEntry""=" & sDocEntry
                    oDtLin = oObjGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                    For iLin As Integer = 0 To oDtLin.Rows.Count - 1
                        sSQL = "INSERT INTO ""@EXO_PACKINGL "" (""Code"", ""LineId"", ""Object"", ""LogInst"", ""U_EXO_USUARIO"", ""U_EXO_CAT"", ""U_EXO_CODE"", ""U_EXO_CANT"", ""U_EXO_LOTE"", "
                        sSQL &= " ""U_EXO_FFAB"", ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"") "
                        sSQL &= "SELECT '" & sDocEntry & "', '" & oDtLin.Rows.Item(iLin).Item("LineNum").ToString & "', 'ODRF', '0', "
                        sSQL &= " ""U_EXO_USUARIO"", ""U_EXO_CAT"", ""U_EXO_CODE"", ""U_EXO_CANT"", ""U_EXO_LOTE"",  ""U_EXO_FFAB"", ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"" "
                        sSQL &= " FROM ""@EXO_TMPPACKINGL"" "
                        sSQL &= " where ""U_EXO_USUARIO""='" & oCompany.UserName.ToString & "' and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' "
                        sSQL &= " Order by ""Code"" "
                        oObjGlobal.refDi.SQL.sqlUpdB1(sSQL)
                    Next
#End Region
                End If
                'INSERT del registro creado
                sSQL = "insert into ""@EXO_TMPPACKING"" values('" & iTabla.ToString & "','" & iTabla.ToString & "'," & iTabla.ToString & ",'N','',0,"
                sSQL &= oObjGlobal.compañia.UserSignature.ToString & ",'','" & Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00") & "',0,'',0,'','" & sEstado & "', "
                sSQL &= "'" & sDocEntry & "', '" & sSerie & "', '" & sDocnum & "', '" & sComen & "', '" & oObjGlobal.compañia.UserName.ToString & "') "
                oObjGlobal.refDi.SQL.sqlUpdB1(sSQL)
                iTabla += 1
            Else
                sMensaje = "No se encuentra las líneas del pedido interno Nº" & EXO_GLOBALES._sPedido & ". Se interrumpe el proceso."
                oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oObjGlobal.SBOApp.MessageBox(sMensaje)
            End If
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
#Region "Pasamos los datos del fichero a la tabla real"
                oDtLin.Clear()
                sSQL = "SELECT * FROM ""DRF1""  WHERE ""DocEntry""=" & sDocEntry
                oDtLin = oObjGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                For iLin As Integer = 0 To oDtLin.Rows.Count - 1
                    sSQL = "INSERT INTO ""@EXO_PACKINGL"" (""Code"", ""LineId"", ""Object"", ""LogInst"", ""U_EXO_USUARIO"", ""U_EXO_CAT"", ""U_EXO_CODE"", ""U_EXO_CANT"", ""U_EXO_LOTE"", "
                    sSQL &= " ""U_EXO_FFAB"", ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"",""U_EXO_LINEA"") "
                    sSQL &= "SELECT '" & sDocEntry & "', ""Code"" , 'ODRF', '0', "
                    sSQL &= " ""U_EXO_USUARIO"", ""U_EXO_CAT"", ""U_EXO_CODE"", ""U_EXO_CANT"", ""U_EXO_LOTE"",  ""U_EXO_FFAB"", ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"", '" & oDtLin.Rows.Item(iLin).Item("LineNum").ToString & "' "
                    sSQL &= " FROM ""@EXO_TMPPACKINGL"" "
                    sSQL &= " where ""U_EXO_USUARIO""='" & oCompany.UserName.ToString & "' and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' "
                    sSQL &= " Order by ""Code"" "
                    oObjGlobal.refDi.SQL.sqlUpdB1(sSQL)
                Next
#End Region
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            oDtLin.Clear() : oDtLinFichero.Clear()
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLote, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsArt, Object))
        End Try
    End Sub

    Public Shared Sub Gen_Lista_Embalaje(ByRef oCompany As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oformE As SAPbouiCOM.Form)
#Region "Variables"
        Dim oDtLin As System.Data.DataTable = New System.Data.DataTable
        Dim oDoc As SAPbobsCOM.StockTransfer = Nothing
        Dim dfecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim sSQL As String = "" : Dim sMensaje As String = ""
        Dim oRsLote As SAPbobsCOM.Recordset = Nothing
        Dim sDocEntry As String = "" : Dim sDocnum As String = ""
#End Region

        Try
            oRsLote = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            oCompany.StartTransaction()
            oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
            oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest
            oDoc.CardCode = oformE.DataSources.DBDataSources.Item("OPDN").GetValue("CardCode", 0).ToString.Trim
            oDoc.Comments = oDoc.Comments.Trim & " " & "Basado en la Entrada de Mercancía Nº" & oformE.DataSources.DBDataSources.Item("OPDN").GetValue("DocNum", 0).ToString.Trim
            oDtLin.Clear()

            'sSQL = "SELECT * FROM ""PDN1"" where ""LineStatus""='O' and ""DocEntry""=" & oformE.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString.Trim & " Order by ""LineNum"" "
            sSQL = "SELECT ""PDN1"".""WhsCode"", ""EXO"".* FROM ""@EXO_PACKINGL"" ""EXO""  "
            sSQL &= " INNER JOIN ""PDN1"" ON ""PDN1"".""DocEntry""=""EXO"".""Code"" and ""PDN1"".""LineNum""=""EXO"".""U_EXO_LINEA"" "
            sSQL &= " where ""Object""='OPDN' and ""Code""='" & oformE.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString.Trim & "' Order by ""U_EXO_LINEA"", ""LineId"" "
            oDtLin = oObjGlobal.refDi.SQL.sqlComoDataTable(sSQL)

            If oDtLin.Rows.Count > 0 Then
                Dim bPlinea As Boolean = True
                For iLin As Integer = 0 To oDtLin.Rows.Count - 1
                    oDoc.FromWarehouse = oDtLin.Rows.Item(iLin).Item("WhsCode").ToString
                    oDoc.ToWarehouse = oDtLin.Rows.Item(iLin).Item("WhsCode").ToString
                    If bPlinea = False Then
                        oDoc.Lines.Add()
                    Else
                        bPlinea = False
                    End If
                    'oDoc.Lines.ItemCode = oDtLin.Rows.Item(iLin).Item("ItemCode").ToString
                    'oDoc.Lines.Quantity = EXO_GLOBALES.DblTextToNumber(oCompany, oDtLin.Rows.Item(iLin).Item("Quantity").ToString)
                    oDoc.Lines.ItemCode = oDtLin.Rows.Item(iLin).Item("U_EXO_CODE").ToString
                    oDoc.Lines.Quantity = EXO_GLOBALES.DblTextToNumber(oCompany, oDtLin.Rows.Item(iLin).Item("U_EXO_CANT").ToString)
                    oDoc.Lines.UserFields.Fields.Item("U_EXO_LOT_ID").Value = oDtLin.Rows.Item(iLin).Item("U_EXO_IDBULTO").ToString
                    oDoc.Lines.UserFields.Fields.Item("U_EXO_TBULTO").Value = oDtLin.Rows.Item(iLin).Item("U_EXO_TBULTO").ToString
                    'oDoc.Lines.WarehouseCode = oDtLin.Rows.Item(iLin).Item("WhsCode").ToString

                    oDoc.Lines.BaseType = InvBaseDocTypeEnum.PurchaseDeliveryNotes
                    oDoc.Lines.BaseEntry = CType(oDtLin.Rows.Item(iLin).Item("Code").ToString, Integer)
                    oDoc.Lines.BaseLine = CType(oDtLin.Rows.Item(iLin).Item("U_EXO_LINEA").ToString, Integer)
#Region "Lotes"
                    'Incluimos los Lotes
                    'sSQL = "SELECT ""OBTN"".""DistNumber"",""ITL1"".* FROM ""OITL"" INNER JOIN ""ITL1"" on ""ITL1"".""LogEntry""=""OITL"".""LogEntry"" "
                    'sSQL &= " Left Join ""OBTN"" on ""OBTN"".""SysNumber""=""ITL1"".""SysNumber"" "
                    'sSQL &= " WHERE ""DocEntry"" = " & oDtLin.Rows.Item(iLin).Item("DocEntry").ToString & " And ""DocLine"" =" & oDtLin.Rows.Item(iLin).Item("LineNum").ToString
                    'oRsLote.DoQuery(sSQL)
                    'For iLote = 1 To oRsLote.RecordCount
                    'Creamos el lote de la línea del artículo
                    'oDoc.Lines.BatchNumbers.BatchNumber = oRsLote.Fields.Item("DistNumber").Value.ToString
                    'oDoc.Lines.BatchNumbers.Quantity = EXO_GLOBALES.DblTextToNumber(oCompany, oRsLote.Fields.Item("Quantity").Value.ToString)
                    'oDoc.Lines.BatchNumbers.Add()                    
                    '    oRsLote.MoveNext()
                    'Next
                    oDoc.Lines.BatchNumbers.BatchNumber = oDtLin.Rows.Item(iLin).Item("U_EXO_LOTE").ToString()
                    oDoc.Lines.BatchNumbers.Quantity = EXO_GLOBALES.DblTextToNumber(oCompany, oDtLin.Rows.Item(iLin).Item("U_EXO_CANT").ToString)
                    oDoc.Lines.BatchNumbers.Add()
#End Region
                Next
                If oDoc.Add() <> 0 Then
                    sMensaje = oCompany.GetLastErrorCode.ToString & " / " & oCompany.GetLastErrorDescription.Replace("'", "")
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    oCompany.GetNewObjectCode(sDocEntry)

                    sSQL = "SELECT ""DocNum"" FROM ""OWTQ""  WHERE ""DocEntry""=" & sDocEntry
                    sDocnum = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    sMensaje = "Se ha generado correctamente la sol. de traslado con Nº " & sDocnum & " y Nº interno " & sDocEntry.ToString
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            Else
                sMensaje = "No se encuentra las líneas de la entrada de Mercancía Nº" & oformE.DataSources.DBDataSources.Item("OPDN").GetValue("DocNum", 0).ToString.Trim & ". Se interrumpe el proceso."
                oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oObjGlobal.SBOApp.MessageBox(sMensaje)
            End If
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLote, Object))

            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            oDtLin.Clear()
        End Try
    End Sub
#End Region
End Class


