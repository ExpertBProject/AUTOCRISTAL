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

#Region "Precios Agencia de Transporte"
    Public Shared Sub TratarFichero_TXT(ByVal sArchivo As String, ByVal sDelimitador As String, ByRef oForm As SAPbouiCOM.Form, ByRef oCompany As SAPbobsCOM.Company, ByRef oSboApp As Application, ByRef objglobal As EXO_UIAPI.EXO_UIAPI)
#Region "Variables"
        ' Apuntador libre a archivo
        Dim Apunt As Integer = FreeFile()
        Dim sMensaje As String = ""
        Dim sPais As String = "" : Dim sLocalidad As String = "" : Dim sZona As String = "" : Dim sTBulto As String = "" : Dim sVolumen As String = "0.00" : Dim sPeso As String = "0.00" : Dim sPrecio As String = "0.00"
        Dim iLinea As Integer = 0
        Dim sExiste As String = "" : Dim sSQL As String = ""
#End Region
        Try
            oForm.Freeze(True)
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
                                    Case 1 : sPais = scampos(i)
                                    Case 2 : sLocalidad = scampos(i)
                                    Case 3 : sZona = scampos(i)
                                    Case 4 : sTBulto = scampos(i)
                                    Case 5 : sVolumen = EXO_GLOBALES.DblNumberToText(oCompany, EXO_GLOBALES.DblTextToNumber(oCompany, scampos(i)), CType(2, FuenteInformacion))
                                    Case 6 : sPeso = EXO_GLOBALES.DblNumberToText(oCompany, EXO_GLOBALES.DblTextToNumber(oCompany, scampos(i)), CType(2, FuenteInformacion))
                                    Case 7 : sPrecio = EXO_GLOBALES.DblNumberToText(oCompany, EXO_GLOBALES.DblTextToNumber(oCompany, scampos(i)), CType(2, FuenteInformacion))
                                End Select
                            Next
                            If sPais.Trim = "" Then
                                sMensaje = "En la línea " & iLinea.ToString & " no puede estar vacío el país. Revise el fichero."
                                objglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oForm.Freeze(False)
                                oSboApp.MessageBox(sMensaje)
                            Else
                                If sLocalidad.Trim = "" Then
                                    sMensaje = "En la línea " & iLinea.ToString & " no puede estar vacía la localidad. Revise el fichero."
                                    objglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    oForm.Freeze(False)
                                    oSboApp.MessageBox(sMensaje)
                                Else
                                    'Buscamos que exista la localidad
                                    sSQL = "SELECT ""Code"" FROM OCST WHERE ""Country""='" & sPais.Trim & "' and UPPER(""Code"")='" & sLocalidad.ToUpper & "' "
                                    sExiste = objglobal.refDi.SQL.sqlStringB1(sSQL)
                                    If sExiste.Trim = "" Then
                                        sSQL = "SELECT ""Code"" FROM OCST WHERE ""Country""='" & sPais.Trim & "' and UPPER(""Name"")='" & sLocalidad.ToUpper & "' "
                                        sExiste = objglobal.refDi.SQL.sqlStringB1(sSQL)
                                        If sExiste.Trim = "" Then
                                            sMensaje = "En la línea " & iLinea.ToString & " no se encuentra la localidad. Revise el fichero."
                                            objglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oForm.Freeze(False)
                                            oSboApp.MessageBox(sMensaje)
                                            Exit Sub
                                        Else
                                            sLocalidad = sExiste
                                        End If
                                    End If

                                    If sTBulto.Trim = "" Then
                                        sTBulto = "-"
                                    End If
                                    If sTBulto = "-" Then
                                    Else
                                        'Debemos de buscar para ver si existe
                                        sSQL = "SELECT ""PkgCode"" FROM OPKG   WHERE UPPER(""PkgCode"")='" & sTBulto.ToUpper & "' "
                                        sExiste = objglobal.refDi.SQL.sqlStringB1(sSQL)
                                        If sExiste.Trim = "" Then
                                            sSQL = "SELECT ""PkgCode"" FROM OPKG   WHERE  UPPER(""PkgType"") like '" & sTBulto.ToUpper & "' "
                                            sExiste = objglobal.refDi.SQL.sqlStringB1(sSQL)
                                            If sExiste.Trim = "" Then
                                                sMensaje = "En la línea " & iLinea.ToString & " no se encuentra el Tipo de bulto. Revise el fichero."
                                                objglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                oForm.Freeze(False)
                                                oSboApp.MessageBox(sMensaje)
                                                Exit Sub
                                            Else
                                                sTBulto = sExiste
                                            End If
                                        End If
                                        'Revisamos que tipo de tarifa vamos a cargar
                                        Dim sTTarifa As String = ""
                                        If CType(oForm.Items.Item("13_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                                            sTTarifa = CType(oForm.Items.Item("13_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                                        End If
                                        If (sTTarifa = "B" And sTBulto <> "-") Or (sTTarifa = "V" And sTBulto = "-") Then
                                            'Grabamos la Línea
                                            Dim ilin As Integer = oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").Size
                                            oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").InsertRecord(ilin)
                                            oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").Offset = ilin
                                            oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").SetValue("U_EXO_PAIS", oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").Offset, sPais)
                                            oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").SetValue("U_EXO_LOCAL", oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").Offset, sLocalidad)
                                            oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").SetValue("U_EXO_ZONA", oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").Offset, sZona)
                                            oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").SetValue("U_EXO_TBULTO", oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").Offset, sTBulto)
                                            oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").SetValue("U_EXO_VOLUMEN", oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").Offset, sVolumen)
                                            oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").SetValue("U_EXO_PESO", oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").Offset, sPeso)
                                            oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").SetValue("U_EXO_PRECIO", oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").Offset, sPrecio)
                                            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSource()
                                            If oForm.Mode = BoFormMode.fm_OK_MODE Then
                                                oForm.Mode = BoFormMode.fm_UPDATE_MODE
                                            End If
                                        Else
                                            sMensaje = "En la línea " & iLinea.ToString & " El tipo de tarifa con el tipo de bulto no son compatibles. Revise los datos."
                                            objglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oForm.Freeze(False)
                                            oSboApp.MessageBox(sMensaje)
                                            oForm.Freeze(True)
                                        End If
                                    End If
                                End If
                            End If
#End Region

                        Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                            objglobal.SBOApp.StatusBar.SetText("(EXO) - Línea " & iLinea.ToString & " no es válida y se omitirá. " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oForm.Freeze(False)
                            oSboApp.MessageBox("Línea " & iLinea.ToString & " no es válida y se omitirá. " & ex.Message)
                        End Try
                    End While
                End Using
                'Eliminamos las línea 0 que LOCALIDAD esté Vacío
                Dim sLocal As String = ""
                sLocal = oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").GetValue("U_EXO_LOCAL", 0)
                If sLocal.Trim = "" Then
                    oForm.DataSources.DBDataSources.Item("@EXO_DTARIFASL").RemoveRecord(0)
                    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSource()
                End If

            Else
                objglobal.SBOApp.StatusBar.SetText("(EXO) - No se ha encontrado el fichero txt a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
            ' Cerramos el archivo
            FileClose(Apunt)
            CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = ""
        End Try
    End Sub

#End Region
End Class




