Imports SAPbouiCOM
Imports System.IO
Public Class EXO_GLOBALES
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
    Public Shared Sub CrearQuerys(ByRef oCompany As SAPbobsCOM.Company)
        Dim sSQL As String = ""
        Dim oOUQR As SAPbobsCOM.UserQueries = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            'Comprobamos si existe la query dentro de la categoría General, que muestra el log de los partes de trabajo automáticos
            oRs.DoQuery("SELECT t1.""IntrnalKey"" " &
                        "FROM ""OUQR"" t1 " &
                        "WHERE t1.""QCategory"" = -1 " &
                        "AND t1.""QName"" = 'Nueva Ubicación Principal'")

            If oRs.RecordCount = 0 Then
                'Creamos la query dentro de la categoría General, que muestra el log de de los partes de trabajo automáticos
                oOUQR = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries), SAPbobsCOM.UserQueries)

                sSQL = "SELECT ""BinCode"" FROM OBIN B 
                            LEFT JOIN (SELECT ""ItemCode"",""WhsCode"", ""BinAbs"", SUM(""OnHandQty"") ""OnHand""
                            FROM OBBQ GROUP BY ""ItemCode"",""WhsCode"",""BinAbs"")S ON S.""ItemCode""= $[$grd_DOC.Artículo.0]
                            and S.""WhsCode""=B.""WhsCode"" and S.""BinAbs""=B.""AbsEntry"" 
                            WHERE B.""WhsCode""= $[$txtALM.0] and B.""Attr2Val"" ='Picking'
                            And IFNULL(S.""OnHand"",0)>= 0
                            And ""BinCode"" Not In (SELECT  T1.""BinCode"" FROM OITW T0 
                            LEFT JOIN OBIN T1 ON T0.""DftBinAbs"" = T1.""AbsEntry""
                            Where T0.""WhsCode"" = $[$txtALM.0] and T1.""BinCode"" is not null)"
                oOUQR.Query = sSQL
                oOUQR.QueryCategory = -1 'General
                oOUQR.QueryDescription = "Nueva Ubicación Principal"
                oOUQR.QueryType = SAPbobsCOM.UserQueryTypeEnum.uqtWizard

                If oOUQR.Add <> 0 Then
                    Throw New Exception(oCompany.GetLastErrorCode & " " & oCompany.GetLastErrorDescription)
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOUQR, Object))
        End Try
    End Sub
    Public Shared Sub CrearConsultasFormateadas(ByRef oCompany As SAPbobsCOM.Company)
        Dim sSQL As String = ""
        Dim oXml As Xml.XmlDocument = Nothing
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim oOUQR As SAPbobsCOM.UserQueries = Nothing
        Dim oCSHS As SAPbobsCOM.FormattedSearches = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sIntrnalKey As String = "0"

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            oCompany.StartTransaction()

            'Comprobamos si existe la consulta formateada dentro de la categoría General, que devuelve las OTs abiertas
            oRs.DoQuery("SELECT t1.""IntrnalKey"" " &
                        "FROM ""OUQR"" t1 " &
                        "WHERE t1.""QCategory"" = -1 " &
                        "AND t1.""QName"" = 'Nueva Ubicación Principal'")

            If oRs.RecordCount = 0 Then
                'Creamos la consulta formateada dentro de la categoría General, que devuelve las OTs abiertas
                oOUQR = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries), SAPbobsCOM.UserQueries)

                sSQL = "SELECT ""BinCode"" FROM OBIN B 
                            LEFT JOIN (SELECT ""ItemCode"",""WhsCode"", ""BinAbs"", SUM(""OnHandQty"") ""OnHand""
                            FROM OBBQ GROUP BY ""ItemCode"",""WhsCode"",""BinAbs"")S ON S.""ItemCode""= $[$grd_DOC.Artículo.0]
                            and S.""WhsCode""=B.""WhsCode"" and S.""BinAbs""=B.""AbsEntry"" 
                            WHERE B.""WhsCode""= $[$txtALM.0] and B.""Attr2Val"" ='Picking'
                            And IFNULL(S.""OnHand"",0)>= 0
                            And ""BinCode"" Not In (SELECT  T1.""BinCode"" FROM OITW T0 
                            LEFT JOIN OBIN T1 ON T0.""DftBinAbs"" = T1.""AbsEntry""
                            Where T0.""WhsCode"" = $[$txtALM.0] and T1.""BinCode"" is not null)"
                oOUQR.Query = sSQL
                oOUQR.QueryCategory = -1 'General
                oOUQR.QueryDescription = "Nueva Ubicación Principal"
                oOUQR.QueryType = SAPbobsCOM.UserQueryTypeEnum.uqtWizard

                If oOUQR.Add <> 0 Then
                    Throw New Exception(oCompany.GetLastErrorCode & " " & oCompany.GetLastErrorDescription)
                End If

                oCompany.GetNewObjectCode(sIntrnalKey)
                sIntrnalKey = sIntrnalKey.Split(vbTab.ToCharArray)(0)
            Else
                sIntrnalKey = oRs.Fields.Item("IntrnalKey").Value.ToString
            End If

            If sIntrnalKey <> "" AndAlso sIntrnalKey <> "0" Then
                oCSHS = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches), SAPbobsCOM.FormattedSearches)

                'Comprobamos si hay una consulta formateada asiganada al campo Orden de trabajo
                oRs.DoQuery("SELECT t1.""IndexID"" " &
                            "FROM ""CSHS"" t1 " &
                            "WHERE t1.""FormID"" = 'EXO_OPUBIC' " &
                            "AND t1.""ItemID"" = 'grd_DOC'
                             AND t1.""ColID"" = 'Nueva Ub. Principal' ")

                'Si hay una consulta formateada asiganada al campo Orden de trabajo, la actualizamos por si a caso,
                'y si no la asignamos.
                If oRs.RecordCount = 0 Then
                    oCSHS.FormID = "EXO_OPUBIC"
                    oCSHS.ItemID = "grd_DOC"
                    oCSHS.ColumnID = "Nueva Ub. Principal"
                    oCSHS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                    oCSHS.QueryID = CInt(sIntrnalKey)
                    oCSHS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO

                    If oCSHS.Add <> 0 Then
                        Throw New Exception(oCompany.GetLastErrorCode & " " & oCompany.GetLastErrorDescription)
                    End If
                Else
                    If oCSHS.GetByKey(CInt(oRs.Fields.Item("IndexID").Value.ToString)) = True Then
                        oCSHS.FormID = "EXO_OPUBIC"
                        oCSHS.ItemID = "grd_DOC"
                        oCSHS.ColumnID = "Nueva Ub. Principal"
                        oCSHS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                        oCSHS.QueryID = CInt(sIntrnalKey)
                        oCSHS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO

                        If oCSHS.Update <> 0 Then
                            Throw New Exception(oCompany.GetLastErrorCode & " " & oCompany.GetLastErrorDescription)
                        End If
                    End If
                End If
            End If


            If sIntrnalKey <> "" AndAlso sIntrnalKey <> "0" Then
                oCSHS = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches), SAPbobsCOM.FormattedSearches)

                'Comprobamos si hay una consulta formateada asiganada al campo Orden de trabajo
                oRs.DoQuery("SELECT t1.""IndexID"" " &
                            "FROM ""CSHS"" t1 " &
                            "WHERE t1.""FormID"" = 'EXO_OPUBI' " &
                            "AND t1.""ItemID"" = 'grd_DOC'
                             AND t1.""ColID"" = 'Nueva Ub. Principal' ")

                'Si hay una consulta formateada asiganada al campo Orden de trabajo, la actualizamos por si a caso,
                'y si no la asignamos.
                If oRs.RecordCount = 0 Then
                    oCSHS.FormID = "EXO_OPUBI"
                    oCSHS.ItemID = "grd_DOC"
                    oCSHS.ColumnID = "Nueva Ub. Principal"
                    oCSHS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                    oCSHS.QueryID = CInt(sIntrnalKey)
                    oCSHS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO

                    If oCSHS.Add <> 0 Then
                        Throw New Exception(oCompany.GetLastErrorCode & " " & oCompany.GetLastErrorDescription)
                    End If
                Else
                    If oCSHS.GetByKey(CInt(oRs.Fields.Item("IndexID").Value.ToString)) = True Then
                        oCSHS.FormID = "EXO_OPUBI"
                        oCSHS.ItemID = "grd_DOC"
                        oCSHS.ColumnID = "Nueva Ub. Principal"
                        oCSHS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery
                        oCSHS.QueryID = CInt(sIntrnalKey)
                        oCSHS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO

                        If oCSHS.Update <> 0 Then
                            Throw New Exception(oCompany.GetLastErrorCode & " " & oCompany.GetLastErrorDescription)
                        End If
                    End If
                End If
            End If
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            Throw exCOM
        Catch ex As Exception
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            Throw ex
        Finally
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOUQR, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCSHS, Object))
        End Try
    End Sub
End Class
