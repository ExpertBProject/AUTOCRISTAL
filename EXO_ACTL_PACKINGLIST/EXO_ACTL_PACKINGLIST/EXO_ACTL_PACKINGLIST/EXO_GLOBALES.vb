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
                'Borramos todo del usuario activo del pedido
                sSQL = "DELETE  FROM  ""@EXO_TMPPACKING"" WHERE ""U_EXO_USUARIO""= '" & objglobal.compañia.UserName.ToString & "' "
                objglobal.refDi.SQL.sqlUpdB1(sSQL)
                sSQL = "DELETE  FROM  ""@EXO_TMPPACKINGL"" WHERE ""U_EXO_USUARIO""= '" & objglobal.compañia.UserName.ToString & "' "
                objglobal.refDi.SQL.sqlUpdB1(sSQL)

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
                                    sArticulo = objglobal.refDi.SQL.sqlStringB1("Select TOP 1 ""ItemCode"" FROM ""OSCN"" WHERE ""Substitute""='" & sCatalogo & "' and ""CardCode""='" & EXO_GLOBALES._sIc & "' ")

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

    Public Shared Sub Generar_PKList(ByRef oCompany As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI)
#Region "Variables"
        Dim oDtLinFichero As System.Data.DataTable = New System.Data.DataTable
        Dim oDtLin As System.Data.DataTable = New System.Data.DataTable
        Dim sSQL As String = "" : Dim sMensaje As String = "" '
        Dim iTabla As Integer = 1
        Dim sDocNumPedido As String = ""
        Dim sObjType As String = ""
        Dim dStock As Double = 0 : Dim dStockExt As Double = 0
        Dim dStockExt1 As Double = 0 : Dim dStockExt2 As Double = 0 : Dim dStockExt3 As Double = 0 : Dim dStockExt4 As Double = 0 : Dim dStockExt5 As Double = 0
        Dim sUBIDEF As String = "" : Dim sTIPOHUECODEF As String = "" : Dim dCANTMAXDEF As Double = 0
#End Region

        Try

#Region "Pasamos los datos del fichero a la tabla real"
            oDtLin.Clear()
            sSQL = "SELECT * FROM ""POR1"" where ""LineStatus""='O' and ""DocEntry""=" & EXO_GLOBALES._sPedido & " Order by ""LineNum"" "
            oDtLin = oObjGlobal.refDi.SQL.sqlComoDataTable(sSQL)
            If oDtLin.Rows.Count > 0 Then
                sSQL = "SELECT ""ObjType"" FROM OPOR WHERE ""DocEntry""=" & EXO_GLOBALES._sPedido
                sObjType = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                'Insertamos la cabecera
                sSQL = "DELETE FROM ""@EXO_PACKING"" WHERE ""Code""='" & EXO_GLOBALES._sPedido & "_" & sObjType & "' "
                If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    sSQL = "DELETE FROM ""@EXO_PACKINGL"" WHERE ""Code""='" & EXO_GLOBALES._sPedido & "_" & sObjType & "' "
                    oObjGlobal.refDi.SQL.executeNonQuery(sSQL)
                End If
                sSQL = "insert into ""@EXO_PACKING"" (""Code"",""Name"",""DocEntry"",""Object"",""U_EXO_OBJTYPE"") 
                                values('" & EXO_GLOBALES._sPedido & "_" & sObjType & "','" & sDocNumPedido & "'," & EXO_GLOBALES._sPedido & ",'EXO_PACKING','" & sObjType & "')"
                If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    'Actualizamos el pedido con el Packing List
                    sSQL = "UPDATE OPOR SET ""U_EXO_PACKING""='" & EXO_GLOBALES._sPedido & "' WHERE ""DocEntry""=" & EXO_GLOBALES._sPedido
                    oObjGlobal.refDi.SQL.executeNonQuery(sSQL)

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

                        sSQL = "INSERT INTO ""@EXO_PACKINGL"" (""Code"", ""LineId"", ""U_EXO_LINEA"",""Object"", ""LogInst"", ""U_EXO_USUARIO"", ""U_EXO_CAT"", ""U_EXO_CODE"", ""U_EXO_CANT"", 
                                ""U_EXO_LOTE"", ""U_EXO_FFAB"", ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"",""U_EXO_ALM"",""U_EXO_STOCK"",""U_EXO_STOCKDENTRO"", 
                                ""U_EXO_EXT1"", ""U_EXO_EXT2"", ""EXo_EXT3"", ""U_EXO_EXT4"",""U_EXO_EXT5"",""U_EXO_UBIDEF"", ""U_EXO_TIPOHUECODEF"",""U_EXO_CANTMAXDEF"") 
                                Select '" & EXO_GLOBALES._sPedido & "_" & sObjType & "', ""Code"", '" & oDtLin.Rows.Item(iLin).Item("LineNum").ToString & "', 'EXO_PACKING', '0', 
                                ""U_EXO_USUARIO"", ""U_EXO_CAT"", ""U_EXO_CODE"", ""U_EXO_CANT"", ""U_EXO_LOTE"",  ""U_EXO_FFAB"", ""U_EXO_IDBULTO"", ""U_EXO_TBULTO"", 
                                '" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "', " & EXO_GLOBALES.DblNumberToText(oCompany, dStock, EXO_GLOBALES.FuenteInformacion.Otros) & " ""STOCK"" 
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStock - dStockExt, EXO_GLOBALES.FuenteInformacion.Otros) & " ""STOCKDENTRO""  
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStockExt1, EXO_GLOBALES.FuenteInformacion.Otros) & " ""EXT1""  
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStockExt1, EXO_GLOBALES.FuenteInformacion.Otros) & " ""EXT2""  
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStockExt1, EXO_GLOBALES.FuenteInformacion.Otros) & " ""EXT3""  
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStockExt1, EXO_GLOBALES.FuenteInformacion.Otros) & " ""EXT4""  
                                , " & EXO_GLOBALES.DblNumberToText(oCompany, dStockExt1, EXO_GLOBALES.FuenteInformacion.Otros) & " ""EXT5""
                                , '" & sUBIDEF & "', '" & sTIPOHUECODEF & "', " & EXO_GLOBALES.DblNumberToText(oCompany, dCANTMAXDEF, EXO_GLOBALES.FuenteInformacion.Otros) & "
                        FROM ""@EXO_TMPPACKINGL"" 
                                where ""U_EXO_USUARIO""='" & oCompany.UserName.ToString & "' and ""U_EXO_CODE""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' 
                                Order by ""Code"" "
                        oObjGlobal.refDi.SQL.sqlUpdB1(sSQL)
                    Next
                End If
            Else
                oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - No se ha podido insertar en el pedido el Packing List del fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If


#End Region

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            oDtLin.Clear() : oDtLinFichero.Clear()
        End Try
    End Sub

    Public Shared Sub Gen_Lista_Embalaje(ByRef oCompany As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oformE As SAPbouiCOM.Form)
#Region "Variables"
        Dim oDtLin As System.Data.DataTable = New System.Data.DataTable
        Dim oDoc As SAPbobsCOM.StockTransfer = Nothing
        Dim dfecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim sSQL As String = "" : Dim sMensaje As String = ""
        Dim oRsLote As SAPbobsCOM.Recordset = Nothing
        Dim sDocEntry As String = "" : Dim sDocnum As String = "" : Dim sAlmacenDestino As String = ""
        Dim sPacking As String = ""
#End Region

        Try
            oRsLote = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            'If oCompany.InTransaction = True Then
            '    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            'End If
            'oCompany.StartTransaction()
            oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
            oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest
            oDoc.CardCode = oformE.DataSources.DBDataSources.Item("OPDN").GetValue("CardCode", 0).ToString.Trim
            oDoc.Comments = oDoc.Comments.Trim & " " & "Basado en la Entrada de Mercancía Nº" & oformE.DataSources.DBDataSources.Item("OPDN").GetValue("DocNum", 0).ToString.Trim
            sPacking = oformE.DataSources.DBDataSources.Item("OPDN").GetValue("U_EXO_PACKING", 0).ToString.Trim
            oDoc.UserFields.Fields.Item("U_EXO_PACKING").Value = sPacking
            oDtLin.Clear()

            sAlmacenDestino = oformE.DataSources.DBDataSources.Item("PDN1").GetValue("WhsCode", 0).ToString.Trim
            sDocEntry = oformE.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString.Trim
            ''sSQL = "SELECT Z.""BinCode"", Z.""Cantidad"", ""PDN1"".""WhsCode"", ""EXO"".* FROM ""@EXO_PACKINGL"" ""EXO""  "
            ''sSQL &= " INNER JOIN ""PDN1"" ON ""PDN1"".""DocEntry""=""EXO"".""Code"" and ""PDN1"".""LineNum""=""EXO"".""U_EXO_LINEA"" "
            ''sSQL &= " Left JOIN (Select  T1.""BinCode"", X.""DistNumber"", X.""ItemCode"", X.""Cantidad"",X.""DocEntry"", X.""DocLineNum"" "
            ''sSQL &= " from ""OBIN"" T1 inner join (  "
            ''sSQL &= " Select T1.""DocEntry"",T1.""DocLineNum"", T0.""BinAbs"", T0.""Quantity"" as ""Cantidad"" ,  "
            ''sSQL &= " T1.""ItemCode"", T1.""Quantity"",  T1.""EffectQty"" , "
            ''sSQL &= " T2.""DistNumber""  From OBTL T0 "
            ''sSQL &= " inner join OILM T1 on T0.""MessageID"" = T1.""MessageID"" And T1.""TransType"" = 20   And T1.""DocEntry"" =" & oformE.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString.Trim
            ''sSQL &= " Left join OBTN T2  ON T0.""SnBMDAbs"" = T2.""AbsEntry"" "
            ''sSQL &= " WHERE T1.""LocCode""='" & sAlmacenDestino & "' "
            ''sSQL &= " ) X on T1.""AbsEntry"" = X.""BinAbs"" "
            ''sSQL &= " )Z ON Z.""DocEntry""=""PDN1"".""DocEntry"" And Z.""DocLineNum""=""PDN1"".""LineNum"" "
            ''sSQL &= " where ""Object""='OPDN' and ""Code""='" & oformE.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).ToString.Trim & "' Order by ""U_EXO_LINEA"", ""LineId"" "
            'sSQL = "SELECT Z.""BinCode"", Z.""DistNumber"", Z.""ItemCode"", Z.""Cantidad"",Z.""DocEntry"", Z.""DocLineNum"", TT.* FROM PDN1 TT "
            'sSQL &= " Left JOIN (Select  T1.""BinCode"", X.""DistNumber"", X.""ItemCode"", X.""Cantidad"",X.""DocEntry"", X.""DocLineNum"" from obin T1 inner join ( "
            'sSQL &= " Select T1.""DocEntry"",T1.""DocLineNum"", T0.""BinAbs"", T0.""Quantity"" as ""Cantidad"" ,  T1.""ItemCode"", T1.""Quantity"",  T1.""EffectQty"" , T2.""DistNumber"" "
            'sSQL &= " From OBTL T0 "
            'sSQL &= " inner join OILM T1 on T0.""MessageID"" = T1.""MessageID"" And T1.""TransType"" = 20   And T1.""DocEntry"" = " & sDocEntry
            'sSQL &= " Left join OBTN T2  ON T0.""SnBMDAbs"" = T2.""AbsEntry"" WHERE T1.""LocCode""='" & sAlmacenDestino & "' "
            'sSQL &= " ) X on T1.""AbsEntry"" = X.""BinAbs"")Z ON Z.""DocEntry""=TT.""DocEntry"" And Z.""DocLineNum""=TT.""LineNum"" "
            'sSQL &= " where TT.""DocEntry""=" & sDocEntry & " Order by TT.""LineNum"" "
            sSQL = "SELECT Z.""BinCode"", Z.""ItemCode"", SUM(Z.""Cantidad"") ""Cantidad"",Z.""DocEntry"", Z.""DocLineNum"", TT.""LineNum"", TT.""WhsCode"",TT.""U_EXO_LOT_ID"",TT.""U_EXO_TBULTO"" "
            sSQL &= " FROM PDN1 TT "
            sSQL &= " Left JOIN (Select  T1.""BinCode"", X.""DistNumber"", X.""ItemCode"", X.""Cantidad"",X.""DocEntry"", X.""DocLineNum"" from obin T1 inner join ( "
            sSQL &= " Select T1.""DocEntry"",T1.""DocLineNum"", T0.""BinAbs"", T0.""Quantity"" as ""Cantidad"" ,  T1.""ItemCode"", T1.""Quantity"",  T1.""EffectQty"" , T2.""DistNumber"" "
            sSQL &= " From OBTL T0 "
            sSQL &= " inner join OILM T1 on T0.""MessageID"" = T1.""MessageID"" And T1.""TransType"" = 20   And T1.""DocEntry"" = " & sDocEntry
            sSQL &= " Left join OBTN T2  ON T0.""SnBMDAbs"" = T2.""AbsEntry"" " ' WHERE T1.""LocCode""='" & sAlmacenDestino & "' "
            sSQL &= " ) X on T1.""AbsEntry"" = X.""BinAbs"")Z ON Z.""DocEntry""=TT.""DocEntry"" And Z.""DocLineNum""=TT.""LineNum"" "
            sSQL &= " where TT.""DocEntry""=" & sDocEntry
            sSQL &= " GROUP BY Z.""BinCode"", Z.""ItemCode"", Z.""DocEntry"", Z.""DocLineNum"",TT.""LineNum"", TT.""WhsCode"",TT.""U_EXO_LOT_ID"",TT.""U_EXO_TBULTO"" "
            sSQL &= " Order by TT.""LineNum"" "

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
                    oDoc.Lines.ItemCode = oDtLin.Rows.Item(iLin).Item("ItemCode").ToString
                    oDoc.Lines.Quantity = EXO_GLOBALES.DblTextToNumber(oCompany, oDtLin.Rows.Item(iLin).Item("Cantidad").ToString)
                    oDoc.Lines.UserFields.Fields.Item("U_EXO_LOT_ID").Value = oDtLin.Rows.Item(iLin).Item("U_EXO_LOT_ID").ToString
                    oDoc.Lines.UserFields.Fields.Item("U_EXO_TBULTO").Value = oDtLin.Rows.Item(iLin).Item("U_EXO_TBULTO").ToString
                    'oDoc.Lines.WarehouseCode = oDtLin.Rows.Item(iLin).Item("WhsCode").ToString

                    oDoc.Lines.UserFields.Fields.Item("U_EXO_UBI_OR").Value = oDtLin.Rows.Item(iLin).Item("BinCode").ToString

                    sSQL = "SELECT ""OBIN"".""BinCode"" FROM ""OITW"" "
                    sSQL &= " INNER JOIN ""OBIN"" ON ""OBIN"".""AbsEntry""= ""OITW"".""DftBinAbs"" "
                    sSQL &= " WHERE ""OITW"".""ItemCode""='" & oDtLin.Rows.Item(iLin).Item("ItemCode").ToString & "' "
                    sSQL &= " and ""OITW"".""WhsCode""='" & oDtLin.Rows.Item(iLin).Item("WhsCode").ToString & "' "
                    Dim sUbiDes As String = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    oDoc.Lines.UserFields.Fields.Item("U_EXO_UBI_DE").Value = sUbiDes

                    oDoc.Lines.BaseType = InvBaseDocTypeEnum.PurchaseDeliveryNotes
                    oDoc.Lines.BaseEntry = CType(oDtLin.Rows.Item(iLin).Item("DocEntry").ToString, Integer)
                    oDoc.Lines.BaseLine = CType(oDtLin.Rows.Item(iLin).Item("LineNum").ToString, Integer)
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

#Region "Actualizamos tabla EXO_PACKINGL el campo EXO_UBIDES"
                    sSQL = "UPDATE ""PACKINGL""
                            SET ""PACKINGL"".""U_EXO_UBIDES"" = ""VISTA"".""UBICADESTINO""
                            FROM ""@EXO_PACKINGL"" AS ""PACKINGL""
                            INNER JOIN ""EXO_UbicacionDestinoEntradaCompra_2"" AS ""VISTA""
                            ON ""PACKINGL"".""Code"" = ""VISTA"".""Code""
                            and ""PACKINGL"".""LineId""=""VISTA"".""LineId""
                            WHERE ""PACKINGL"".""Code""= '" & sPacking & "' "
                    oObjGlobal.refDi.SQL.sqlUpdB1(sSQL)
#End Region
                End If
            Else
                sMensaje = "No se encuentra las líneas de la entrada de Mercancía Nº" & oformE.DataSources.DBDataSources.Item("OPDN").GetValue("DocNum", 0).ToString.Trim & ". Se interrumpe el proceso."
                oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oObjGlobal.SBOApp.MessageBox(sMensaje)
            End If
            'If oCompany.InTransaction = True Then
            '    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            'End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLote, Object))

            'If oCompany.InTransaction = True Then
            '    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            'End If

            oDtLin.Clear()
        End Try
    End Sub
#End Region
End Class


