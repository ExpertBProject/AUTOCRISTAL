Imports SAPbouiCOM
Imports System.IO
Imports System.Drawing.Printing
Imports System.Management
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing

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

    Public Shared Function TotalesBultosPdtes(ByVal sSQLUS As String, ByVal sCampo As String) As String
        Dim sSQL As String = ""
        TotalesBultosPdtes = sSQLUS
        Try
            sSQL = "SELECT Count(""Id Bulto"") ""Cuenta"" FROM ("
            sSQL &= "SELECT DISTINCT ""Id Bulto"" FROM ("
            sSQL &= sSQLUS
            sSQL &= ") T "
            Select Case sCampo
                Case "UDREC"
                    sSQL &= " WHERE ""Recibido""<>0 "
                Case "UDREU"
                    sSQL &= " WHERE ""Pdte. Reubicar""<>0 "
                Case "UDPDTE"
                    sSQL &= " WHERE ""Pdte. Recibir""<>0 "
            End Select
            sSQL &= " GROUP BY ""Tipo"", ""Nº Documento"", ""Id Bulto"" "
            sSQL &= ") T "

            TotalesBultosPdtes = sSQL
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function CargarUDO(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sUDO As String, ByVal sCode As String) As Boolean
        CargarUDO = False

        Try
            If sCode = "" Then
                oObjGlobal.funcionesUI.cargaFormUdoBD(sUDO)
            Else
                oObjGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_PAQ", sCode)
            End If

            CargarUDO = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
    Public Shared Sub CopiarRecurso(ByVal pAssembly As Reflection.Assembly, ByVal pNombreRecurso As String, ByVal pRuta As String)

        Dim s As Stream = pAssembly.GetManifestResourceStream(pAssembly.GetName().Name + "." + pNombreRecurso)
        If s.Length = 0 Then
            Throw New Exception("No se puede encontrar el recurso '" + pNombreRecurso + "'")
        Else
            Dim buffer(CInt(s.Length() - 1)) As Byte
            s.Read(buffer, 0, buffer.Length)

            Dim sw As BinaryWriter = New BinaryWriter(File.Open(pRuta, FileMode.Create))
            sw.Write(buffer)
            sw.Close()
        End If


    End Sub
    Public Shared Function IsPrinterOnline(ByVal printerName As String) As Boolean
        Dim Str As String = Nothing
        Dim online As Boolean = False

        '//set the scope of this search to the local machine
        Dim scope As ManagementScope = New ManagementScope(ManagementPath.DefaultPath)
        '//connect to the machine
        scope.Connect()

        '//query for the ManagementObjectSearcher
        Dim query As SelectQuery = New SelectQuery("select * from Win32_Printer")

        Dim m As ManagementClass = New ManagementClass("Win32_Printer")

        Dim obj As ManagementObjectSearcher = New ManagementObjectSearcher(scope, query)

        '//get each instance from the ManagementObjectSearcher object
        Dim printers As ManagementObjectCollection = m.GetInstances()

        '  //now loop through each printer instance returned
        For Each printer As ManagementObject In printers
            '    //first make sure we got something back
            If printer IsNot Nothing Then
                '      //get the current printer name in the loop
                Str = printer("Name").ToString().ToLower()

                '      //check if it matches the name provided
                If Str.ToLower = printerName.ToLower Then
                    '        //since we found a match check it's status
                    If (printer("WorkOffline").ToString().ToLower().Equals("true") And printer("PrinterStatus").Equals(7)) Then
                        '          //it's offline
                        online = False
                    Else
                        '         //it's online
                        online = True
                    End If
                    Exit For
                Else
                    'Throw New Exception("No printers were found")
                    online = False
                End If
            End If
        Next


        Return online
    End Function

    Public Shared Sub GetCrystalReportFile(ByRef oobjglobal As EXO_UIAPI.EXO_UIAPI, ByVal sOutFileName As String, ByVal sVariable As String)
        Dim oBlobParams As SAPbobsCOM.BlobParams = Nothing
        Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment = Nothing
        Dim oBlob As SAPbobsCOM.Blob = Nothing
        Dim sContent As String = ""
        Dim obuff() As Byte = Nothing

        Try
            oBlobParams = CType(oobjglobal.compañia.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams), SAPbobsCOM.BlobParams)

            oBlobParams.Table = "RDOC"
            oBlobParams.Field = "Template"

            oKeySegment = oBlobParams.BlobTableKeySegments.Add()
            oKeySegment.Name = "DocCode"

            oKeySegment.Value = sVariable

            oBlob = oobjglobal.compañia.GetCompanyService().GetBlob(oBlobParams)
            sContent = oBlob.Content

            obuff = Convert.FromBase64String(sContent)

            Using oFile As New System.IO.FileStream(sOutFileName, System.IO.FileMode.Create)
                oFile.Write(obuff, 0, obuff.Length)

                oFile.Close()
            End Using

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlobParams, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oKeySegment, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlob, Object))
        End Try
    End Sub
    Public Shared Sub GenerarImpCrystal(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal rutaCrystal As String, ByVal sCrystal As String,
                                        ByVal sDocNum As String, ByVal sDocEntry As String, ByVal sSchema As String, ByVal sTIPODOC As String,
                                        ByVal sDir As String, ByRef sReport As String, ByVal sTipoImp As String, ByVal sUsuario As String)

        Dim oCRReport As ReportDocument = Nothing
        Dim oFileDestino As DiskFileDestinationOptions = Nothing
        Dim sServer As String = ""
        Dim sDriver As String = ""
        Dim sBBDD As String = ""
        Dim sUser As String = ""
        Dim sPwd As String = ""
        Dim sConnection As String = ""
        Dim oLogonProps As NameValuePairs2 = Nothing

        Dim conrepor As DataSourceConnections = Nothing
        Dim sImpresora As String = "" : Dim nCopias As Integer = 1
        Dim sSQL As String = ""
        Try

            Select Case sTIPODOC
                Case "ALBVTA"
#Region "Entregas"
                    sTIPODOC = "15"
#End Region
                Case "SOLTRA" ' Sol. de Traslado                           
#Region "Sol de traslado"
                    sTIPODOC = "1250000001"
#End Region
                Case "DPROV" ' Dev. de Proveedor
#Region "Dev de proveedor"
                    sTIPODOC = "21"
#End Region
            End Select
            oCRReport = New ReportDocument()

            oCRReport.Load(rutaCrystal & "\" & sCrystal)

            oCRReport.DataSourceConnections.Clear()

            oObjGlobal.SBOApp.StatusBar.SetText("DocKey@: " & sDocEntry & " - ObjectID@: " & sTIPODOC & " - Schema@: " & sSchema, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)
            'Establecemos los parámetros para el report.
            oCRReport.SetParameterValue("DocKey@", sDocEntry)
            oCRReport.SetParameterValue("ObjectID@", sTIPODOC)
            oCRReport.SetParameterValue("Schema@", sSchema)

            'Establecemos las conexiones a la BBDD
            sServer = oObjGlobal.funcionesUI.refDi.OGEN.valorVariable("SERVIDOR_HANA") ' objGlobal.compañia.Server
            'sServer = objGlobal.refDi.SQL.dameCadenaConexion.ToString
            sBBDD = oObjGlobal.compañia.CompanyDB
            sUser = oObjGlobal.refDi.SQL.usuarioSQL
            sPwd = oObjGlobal.refDi.SQL.claveSQL

            sDriver = "HDBODBC"
            sConnection = "DRIVER={" & sDriver & "};UID=" & sUser & ";PWD=" & sPwd & ";SERVERNODE=" & sServer & ";DATABASE=" & sBBDD & ";"
            'sConnection = "DRIVER={" & sDriver & "};" & sServer & ";DATABASE=" & sBBDD & ";"
            oObjGlobal.SBOApp.StatusBar.SetText("Conectando: " & sConnection, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
            oLogonProps = oCRReport.DataSourceConnections(0).LogonProperties
            oLogonProps.Set("Provider", sDriver)
            oLogonProps.Set("Connection String", sConnection)

            oCRReport.DataSourceConnections(0).SetLogonProperties(oLogonProps)
            oCRReport.DataSourceConnections(0).SetConnection(sServer, sBBDD, False)
            oObjGlobal.SBOApp.StatusBar.SetText("Connection String: " & sConnection, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)

            For Each oSubReport As ReportDocument In oCRReport.Subreports
                For Each oConnection As IConnectionInfo In oSubReport.DataSourceConnections
                    oConnection.SetConnection(sServer, sBBDD, False)
                    oConnection.SetLogon(sUser, sPwd)
                Next
            Next
            oObjGlobal.SBOApp.StatusBar.SetText("Actualizado conect Subreport...", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)

            Select Case sTipoImp
                Case "PDF"
#Region "Exportar a PDF"
                    'Preparamos para la exportación
                    If IO.Directory.Exists(sDir) = False Then
                        IO.Directory.CreateDirectory(sDir)
                    End If
                    sReport = sDir & "sTIPODOC_" & sDocNum & ".pdf"
                    'Compruebo si existe y lo borro
                    If IO.File.Exists(sReport) Then
                        IO.File.Delete(sReport)
                    End If
                    oObjGlobal.SBOApp.StatusBar.SetText("Generando pdf para envio impresión...Espere por favor", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)

                    oCRReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat

                    oFileDestino = New CrystalDecisions.Shared.DiskFileDestinationOptions
                    oFileDestino.DiskFileName = sReport

                    'Le pasamos al reporte el parámetro destino del reporte (ruta)
                    oCRReport.ExportOptions.DestinationOptions = oFileDestino

                    'Le indicamos que el reporte no es para mostrarse en pantalla, sino, que es para guardar en disco
                    oCRReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile

                    'Finalmente exportamos el reporte a PDF
                    oCRReport.Export()
                    '            oCRReport.ExportToDisk(ExportFormatType.PortableDocFormat, sReport)
#End Region
                Case "IMP"
#Region "Imprimir a impresora"
                    'Buscamos la impresora por defecto
                    Dim instance As New Printing.PrinterSettings
                    sImpresora = instance.PrinterName
                    'oObjGlobal.SBOApp.StatusBar.SetText("Impresora: " & sImpresora, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)
                    If EXO_GLOBALES.IsPrinterOnline(sImpresora) = True Then
                        oObjGlobal.SBOApp.StatusBar.SetText("Imprimiendo en " & sImpresora & "...Espere por favor", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                        oCRReport.PrintOptions.NoPrinter = False
                        oCRReport.PrintOptions.PrinterName = sImpresora
                        oCRReport.PrintToPrinter(nCopias, False, 0, 9999)
                    Else
                        oObjGlobal.SBOApp.StatusBar.SetText("La impresora seleccionada en el usuario no se encuentra o está offline. Por favor verifique la parametrización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    End If
#End Region
            End Select

            'Cerramos
            oCRReport.Close()
            oCRReport.Dispose()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oObjGlobal.SBOApp.StatusBar.SetText("Fin del proceso de impresión.", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success)
            oCRReport = Nothing
            oFileDestino = Nothing
        End Try
    End Sub
    Public Shared Sub GenerarImpCrystal_Rangos(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal rutaCrystal As String, ByVal sCrystal As String,
                                               ByVal sDesde As String, ByVal sHasta As String, ByVal sFileName As String, ByRef sReport As String,
                                               ByVal sTipoImp As String, ByVal sUsuario As String)

        Dim oCRReport As ReportDocument = Nothing
        Dim oFileDestino As DiskFileDestinationOptions = Nothing
        Dim sServer As String = ""
        Dim sDriver As String = ""
        Dim sBBDD As String = ""
        Dim sUser As String = ""
        Dim sPwd As String = ""
        Dim sConnection As String = ""
        Dim oLogonProps As NameValuePairs2 = Nothing

        Dim conrepor As DataSourceConnections = Nothing
        Dim sImpresora As String = "" : Dim nCopias As Integer = 1
        Dim sSQL As String = ""
        Try
            oCRReport = New ReportDocument()

            oCRReport.Load(rutaCrystal & "\" & sCrystal)

            oCRReport.DataSourceConnections.Clear()

            'Establecemos los parámetros para el report.
            If sDesde = "" Then
                sDesde = "1"
            End If
            If sHasta = "" Then
                sHasta = "999999999999999"
            End If
            oCRReport.SetParameterValue("Desde", sDesde)
            oCRReport.SetParameterValue("Hasta", sHasta)

            'Establecemos las conexiones a la BBDD
            sServer = oObjGlobal.funcionesUI.refDi.OGEN.valorVariable("SERVIDOR_HANA") ' objGlobal.compañia.Server
            'sServer = objGlobal.refDi.SQL.dameCadenaConexion.ToString
            sBBDD = oObjGlobal.compañia.CompanyDB
            sUser = oObjGlobal.refDi.SQL.usuarioSQL
            sPwd = oObjGlobal.refDi.SQL.claveSQL

            sDriver = "HDBODBC"
            sConnection = "DRIVER={" & sDriver & "};UID=" & sUser & ";PWD=" & sPwd & ";SERVERNODE=" & sServer & ";DATABASE=" & sBBDD & ";"
            'sConnection = "DRIVER={" & sDriver & "};" & sServer & ";DATABASE=" & sBBDD & ";"
            oObjGlobal.SBOApp.StatusBar.SetText("Conectando: " & sConnection, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
            oLogonProps = oCRReport.DataSourceConnections(0).LogonProperties
            oLogonProps.Set("Provider", sDriver)
            oLogonProps.Set("Connection String", sConnection)

            oCRReport.DataSourceConnections(0).SetLogonProperties(oLogonProps)
            oCRReport.DataSourceConnections(0).SetConnection(sServer, sBBDD, False)

            For Each oSubReport As ReportDocument In oCRReport.Subreports
                For Each oConnection As IConnectionInfo In oSubReport.DataSourceConnections
                    oConnection.SetConnection(sServer, sBBDD, False)
                    oConnection.SetLogon(sUser, sPwd)
                Next
            Next

            Select Case sTipoImp
                Case "PDF"
#Region "Exportar a PDF"
                    'Preparamos para la exportación
                    sReport = sFileName & "\Et_Ubicaciones_" & sDesde & "_" & sHasta & ".pdf"
                    'Compruebo si existe y lo borro
                    If IO.File.Exists(sReport) Then
                        IO.File.Delete(sReport)
                    End If
                    oObjGlobal.SBOApp.StatusBar.SetText("Generando pdf para envio impresión...Espere por favor", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)

                    oCRReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat

                    oFileDestino = New CrystalDecisions.Shared.DiskFileDestinationOptions
                    oFileDestino.DiskFileName = sReport

                    'Le pasamos al reporte el parámetro destino del reporte (ruta)
                    oCRReport.ExportOptions.DestinationOptions = oFileDestino

                    'Le indicamos que el reporte no es para mostrarse en pantalla, sino, que es para guardar en disco
                    oCRReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile

                    'Finalmente exportamos el reporte a PDF
                    oCRReport.Export()
                    '            oCRReport.ExportToDisk(ExportFormatType.PortableDocFormat, sReport)
#End Region
                Case "IMP"
#Region "Imprimir a impresora"
                    'Buscamos la impresora
                    sSQL = "SELECT ""Fax"" FROM OUSR WHERE ""USERID""='" & oObjGlobal.compañia.UserSignature.ToString & "' "
                    sImpresora = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    If EXO_GLOBALES.IsPrinterOnline(sImpresora) = True Then
                        oObjGlobal.SBOApp.StatusBar.SetText("Imprimiendo en " & sImpresora & "...Espere por favor", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                        oCRReport.PrintOptions.NoPrinter = False
                        oCRReport.PrintOptions.PrinterName = sImpresora
                        oCRReport.PrintToPrinter(nCopias, False, 0, 9999)
                    Else
                        oObjGlobal.SBOApp.StatusBar.SetText("La impresora seleccionada en el usuario no se encuentra o está offline. Por favor verifique la parametrización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    End If
#End Region
            End Select

            'Cerramos
            oCRReport.Close()
            oCRReport.Dispose()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oObjGlobal.SBOApp.StatusBar.SetText("Fin del proceso de impresión.", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success)
            oCRReport = Nothing
            oFileDestino = Nothing
        End Try
    End Sub
End Class
