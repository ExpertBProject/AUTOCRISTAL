Imports Sap.Data.Hana
Imports System.IO
Imports SAPbobsCOM

Public Class Procesos
    Public Shared Sub OSHP()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOSHP As SAPbobsCOM.ShippingTypes = Nothing
        Dim oDB As HanaConnection = Nothing
        Dim olog As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sTrnspCode As String = ""
        Dim sWebSite As String = ""
        Dim oUserFields As SAPbobsCOM.UserFields = Nothing
        Dim refDI As EXO_DIAPI.EXO_DIAPI = Nothing
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim tipoServidor As SAPbobsCOM.BoDataServerTypes = SAPbobsCOM.BoDataServerTypes.dst_HANADB
        Dim tipocliente As EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico
        Dim sPass As String = Conexiones.Datos_Confi("DI", "Password")
        Try
            olog = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OSHP.txt", 1)
            olog.escribeMensaje("Antes de conectar sql", EXO_Log.EXO_Log.Tipo.informacion)
            Conexiones.Connect_SQLHANA(oDB, "HANA", olog)
            olog.escribeMensaje("Despues de conectar sql", EXO_Log.EXO_Log.Tipo.informacion)
            olog.escribeMensaje("Anmtes de conectar company", EXO_Log.EXO_Log.Tipo.informacion)
            Conexiones.Connect_Company(oCompany, "DI", Conexiones.sBBDD, Conexiones.sUser, Conexiones.sPwd, olog)
            olog.escribeMensaje("Despues de conectar company", EXO_Log.EXO_Log.Tipo.informacion)
            Try
                refDI = New EXO_DIAPI.EXO_DIAPI(oCompany, olog)

            Catch ex As Exception
                refDI = New EXO_DIAPI.EXO_DIAPI(tipoServidor, oCompany.Server.ToString, oCompany.LicenseServer.ToString, oCompany.CompanyDB.ToString, oCompany.UserName.ToString, sPass, tipocliente)
            End Try
            olog.escribeMensaje("Despues de refdi", EXO_Log.EXO_Log.Tipo.informacion)

            sSQL = "SELECT t1.""DBNAMEORIG"", t1.""DBNAMEDEST"", t1.""TABLENAME"", t1.""CODETABLE"", t1.""CODETABLE2"" " &
                   "FROM ""REPLICATE"" t1  " &
                   "WHERE t1.""TABLENAME"" = 'OSHP' " &
                   "ORDER BY t1.""DBNAMEORIG"", t1.""DBNAMEDEST"" "

            oDt = New System.Data.DataTable
            oDt = refDI.SQL.sqlComoDataTable(sSQL)
            'Conexiones.FillDtDB(oDB, oDt, sSQL)
            olog.escribeMensaje("Antes del count datatable", EXO_Log.EXO_Log.Tipo.informacion)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("DBNAMEORIG").ToString
                sDBD = oDt.Rows.Item(0).Item("DBNAMEDEST").ToString

                Conexiones.Connect_Company(oCompanyO, "DI", sDBO, Conexiones.sUser, Conexiones.sPwd, olog)
                Conexiones.Connect_Company(oCompanyD, "DI", sDBD, Conexiones.sUser, Conexiones.sPwd, olog)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("DBNAMEORIG").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)
                            Conexiones.Connect_Company(oCompanyO, "DI", oDt.Rows.Item(0).Item("DBNAMEORIG").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                            sDBO = oDt.Rows.Item(i).Item("DBNAMEORIG").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("DBNAMEDEST").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)
                            Conexiones.Connect_Company(oCompanyD, "DI", oDt.Rows.Item(0).Item("DBNAMEDEST").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                            sDBD = oDt.Rows.Item(i).Item("DBNAMEDEST").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOSHP = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oShippingTypes), SAPbobsCOM.ShippingTypes)

                        If oOSHP.GetByKey(CInt(oDt.Rows.Item(i).Item("CODETABLE").ToString)) = True Then
                            sXML = oOSHP.GetAsXML
                        Else
                            sXML = ""
                        End If

                        If sXML <> "" Then
                            'Porque en el modo Update no funciona por XML
                            sWebSite = oOSHP.Website
                            'oUserFields = oOSHP.UserFields
                            '''''''''''''''''''''''''''''''''''''''''''''

                            oOSHP = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oShippingTypes), SAPbobsCOM.ShippingTypes)

                            oOSHP = CType(oCompanyD.GetBusinessObjectFromXML(sXML, 0), SAPbobsCOM.ShippingTypes)

                            sTrnspCode = Conexiones.GetValueDB(oDB, """" & sDBD & """.""OSHP""", """TrnspCode""", """TrnspName"" = '" & oDt.Rows.Item(i).Item("CODETABLE2").ToString & "'")

                            If sTrnspCode = "" Then
                                'Añadir
                                olog.escribeMensaje("No existe", EXO_Log.EXO_Log.Tipo.informacion)
                                If oOSHP.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                olog.escribeMensaje(" existe", EXO_Log.EXO_Log.Tipo.informacion)
                                If oOSHP.GetByKey(CInt(sTrnspCode)) = True Then
                                    oOSHP.Website = sWebSite

                                    'For h As Integer = 0 To oUserFields.Fields.Count - 1
                                    '    If oOSHP.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).IsNull = SAPbobsCOM.BoYesNoEnum.tNO Then
                                    '        oOSHP.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value = oUserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value
                                    '    End If
                                    'Next

                                    If oOSHP.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOSHP.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM ""REPLICATE"" WHERE ""DBNAMEORIG"" = '" & sDBO & "' AND ""DBNAMEDEST"" = '" & sDBD & "' AND ""TABLENAME"" = '" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "' AND ""CODETABLE"" = '" & oDt.Rows.Item(i).Item("CODETABLE").ToString & "'"
                        refDI.SQL.executeNonQuery(sSQL)
                        'Conexiones.ExecuteSqlDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        olog.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "|" & oDt.Rows.Item(i).Item("CODETABLE").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        olog.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "|" & oDt.Rows.Item(i).Item("CODETABLE").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            olog.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            olog.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOSHP IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOSHP)

            Conexiones.Disconnect_SQLHANA(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
            Conexiones.Disconnect_Company(oCompany)
        End Try
    End Sub

    Public Shared Sub OITB()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOITB As SAPbobsCOM.ItemGroups = Nothing
        Dim oOITB2 As SAPbobsCOM.ItemGroups = Nothing
        Dim oDB As HanaConnection = Nothing
        Dim olog As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim oXml As Xml.XmlDocument = Nothing
        Dim oXmlNode As Xml.XmlNode = Nothing
        Dim sItmsGrpCod As String = ""
        Dim sUgpEntry As String = ""
        Dim sIUoMEntry As String = ""
        Dim oUserFields As SAPbobsCOM.UserFields = Nothing
        Dim refDI As EXO_DIAPI.EXO_DIAPI = Nothing
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim tipoServidor As SAPbobsCOM.BoDataServerTypes = SAPbobsCOM.BoDataServerTypes.dst_HANADB
        Dim tipocliente As EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico
        Dim sPass As String = Conexiones.Datos_Confi("DI", "Password")

        Try
            olog = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OITB.txt", 1)
            Conexiones.Connect_SQLHANA(oDB, "HANA", olog)
            Conexiones.Connect_Company(oCompany, "DI", Conexiones.sBBDD, Conexiones.sUser, Conexiones.sPwd, olog)

            Try
                refDI = New EXO_DIAPI.EXO_DIAPI(oCompany, olog)
            Catch ex As Exception
                refDI = New EXO_DIAPI.EXO_DIAPI(tipoServidor, oCompany.Server.ToString, oCompany.LicenseServer.ToString, oCompany.CompanyDB.ToString, oCompany.UserName.ToString, sPass, tipocliente)
            End Try


            'Conexiones.Connect_Company(oCompany, "DI", Conexiones.sBBDD, Conexiones.sUser, Conexiones.sPwd, olog))

            sSQL = "SELECT t1.""DBNAMEORIG"", t1.""DBNAMEDEST"", t1.""TABLENAME"", t1.""CODETABLE"", t1.""CODETABLE2"" " &
                   "FROM ""REPLICATE"" t1  " &
                   "WHERE t1.""TABLENAME"" = 'OITB' " &
                   "ORDER BY t1.""DBNAMEORIG"", t1.""DBNAMEDEST"" "

            oDt = New System.Data.DataTable
            oDt = refDI.SQL.sqlComoDataTable(sSQL)
            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("DBNAMEORIG").ToString
                sDBD = oDt.Rows.Item(0).Item("DBNAMEDEST").ToString

                'Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                'Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                Conexiones.Connect_Company(oCompanyO, "DI", oDt.Rows.Item(0).Item("DBNAMEORIG").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                Conexiones.Connect_Company(oCompanyD, "DI", oDt.Rows.Item(0).Item("DBNAMEDEST").ToString, Conexiones.sUser, Conexiones.sPwd, olog)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("DBNAMEORIG").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)
                            Conexiones.Connect_Company(oCompanyO, "DI", oDt.Rows.Item(0).Item("DBNAMEORIG").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                            sDBO = oDt.Rows.Item(i).Item("DBNAMEORIG").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("DBNAMEDEST").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)
                            Conexiones.Connect_Company(oCompanyD, "DI", oDt.Rows.Item(0).Item("DBNAMEDEST").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                            sDBD = oDt.Rows.Item(i).Item("DBNAMEDEST").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOITB = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemGroups), SAPbobsCOM.ItemGroups)

                        If oOITB.GetByKey(CInt(oDt.Rows.Item(i).Item("CODETABLE").ToString)) = True Then
                            sXML = oOITB.GetAsXML
                        Else
                            sXML = ""
                        End If

                        If sXML <> "" Then
                            'Porque en el modo Update no funciona por XML

                            oUserFields = oOITB.UserFields
                            '''''''''''''''''''''''''''''''''''''''''''''

                            'Esto es porque hay ciertos campos que son autonuméricos y no tienen por qué ser igual en todas las empresas
                            sUgpEntry = Conexiones.GetValueDB(oDB, """" & sDBO & """.""OUGP""", """UgpCode""", """UgpEntry"" = " & oOITB.DefaultUoMGroup & "")
                            sUgpEntry = Conexiones.GetValueDB(oDB, """" & sDBD & """.""OUGP""", """UgpEntry""", """UgpCode"" = '" & sUgpEntry & "'")
                            sIUoMEntry = Conexiones.GetValueDB(oDB, """" & sDBO & """.""OUOM""", """UomCode""", """UomEntry"" = " & oOITB.DefaultInventoryUoM & "")
                            sIUoMEntry = Conexiones.GetValueDB(oDB, """" & sDBD & """.""OUOM""", """UomEntry""", """UomCode"" = '" & sIUoMEntry & "'")

                            oOITB2 = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemGroups), SAPbobsCOM.ItemGroups)

                            sItmsGrpCod = Conexiones.GetValueDB(oDB, """" & sDBD & """.""OITB""", """ItmsGrpCod""", """ItmsGrpNam"" = '" & oDt.Rows.Item(i).Item("CODETABLE2").ToString & "'")

                            If sItmsGrpCod = "" Then
                                'Añadir
                                oOITB2.GroupName = oOITB.GroupName
                                If sUgpEntry <> "" Then
                                    oOITB2.DefaultUoMGroup = CInt(sUgpEntry)
                                End If

                                If sIUoMEntry <> "" Then
                                    oOITB2.DefaultInventoryUoM = CInt(sIUoMEntry)
                                End If

                                oOITB2.PlanningSystem = oOITB.PlanningSystem
                                oOITB2.ProcurementMethod = oOITB.ProcurementMethod
                                'oOITB2.OrderInterval = oOITB.OrderInterval
                                oOITB2.OrderMultiple = oOITB.OrderMultiple
                                oOITB2.MinimumOrderQuantity = oOITB.MinimumOrderQuantity
                                oOITB2.LeadTime = oOITB.LeadTime
                                oOITB2.ToleranceDays = oOITB.ToleranceDays
                                oOITB2.InventorySystem = oOITB.InventorySystem
                                'oOITB2.CycleCode = 1
                                'For h As Integer = 0 To oUserFields.Fields.Count - 1
                                '    If oOITB2.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).IsNull = SAPbobsCOM.BoYesNoEnum.tNO Then
                                '        oOITB2.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value = oUserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value
                                '    End If
                                'Next

                                If oOITB2.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If


                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML

                                If oOITB2.GetByKey(CInt(sItmsGrpCod)) = True Then
                                    If sIUoMEntry <> "" Then
                                        oOITB2.DefaultInventoryUoM = CInt(sIUoMEntry)
                                    Else
                                        oOITB2.DefaultInventoryUoM = 0
                                    End If

                                    If sUgpEntry <> "" Then
                                        oOITB2.DefaultUoMGroup = CInt(sUgpEntry)
                                    Else
                                        oOITB2.DefaultUoMGroup = 0
                                    End If

                                    oOITB2.PlanningSystem = oOITB.PlanningSystem
                                    oOITB2.ProcurementMethod = oOITB.ProcurementMethod
                                    'oOITB2.OrderInterval = oOITB.OrderInterval
                                    oOITB2.OrderMultiple = oOITB.OrderMultiple
                                    oOITB2.MinimumOrderQuantity = oOITB.MinimumOrderQuantity
                                    oOITB2.LeadTime = oOITB.LeadTime
                                    oOITB2.ToleranceDays = oOITB.ToleranceDays
                                    oOITB2.InventorySystem = oOITB.InventorySystem
                                    'oOITB2.ComponentWarehouse = oOITB.ComponentWarehouse

                                    'For h As Integer = 0 To oUserFields.Fields.Count - 1
                                    '    If oOITB2.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).IsNull = SAPbobsCOM.BoYesNoEnum.tNO Then
                                    '        oOITB2.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value = oUserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value
                                    '    End If
                                    'Next

                                    If oOITB2.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''
                            End If
                        End If

                        sSQL = "DELETE FROM ""REPLICATE"" WHERE ""DBNAMEORIG"" = '" & sDBO & "' AND ""DBNAMEDEST"" = '" & sDBD & "' AND ""TABLENAME"" = '" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "' AND ""CODETABLE"" = '" & oDt.Rows.Item(i).Item("CODETABLE").ToString & "'"
                        refDI.SQL.executeNonQuery(sSQL)
                        'Conexiones.ExecuteSqlDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        olog.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "|" & oDt.Rows.Item(i).Item("CODETABLE").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        olog.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "|" & oDt.Rows.Item(i).Item("CODETABLE").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            olog.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            olog.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOITB IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOITB)

            Conexiones.Disconnect_SQLHANA(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
            Conexiones.Disconnect_Company(oCompany)
        End Try
    End Sub
    Public Shared Sub OMRC()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOMRC As SAPbobsCOM.Manufacturers = Nothing
        Dim oDB As HanaConnection = Nothing
        Dim olog As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sFirmCode As String = ""
        Dim sFirmName As String = ""
        Dim oUserFields As SAPbobsCOM.UserFields = Nothing
        Dim refDI As EXO_DIAPI.EXO_DIAPI = Nothing
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim tipoServidor As SAPbobsCOM.BoDataServerTypes = SAPbobsCOM.BoDataServerTypes.dst_HANADB
        Dim tipocliente As EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico
        Dim sPass As String = Conexiones.Datos_Confi("DI", "Password")

        Try
            olog = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OMRC.txt", 1)
            Conexiones.Connect_SQLHANA(oDB, "HANA", olog)
            Conexiones.Connect_Company(oCompany, "DI", Conexiones.sBBDD, Conexiones.sUser, Conexiones.sPwd, olog)
            Try
                refDI = New EXO_DIAPI.EXO_DIAPI(oCompany, olog)
            Catch ex As Exception
                'refDI = New EXO_DIAPI.EXO_DIAPI(tipoServidor, oCompany.Server.ToString, oCompany.LicenseServer.ToString, oCompany.CompanyDB.ToString, oCompany.UserName.ToString, sPass, tipocliente)
            End Try


            sSQL = "SELECT t1.""DBNAMEORIG"", t1.""DBNAMEDEST"", t1.""TABLENAME"", t1.""CODETABLE"", t1.""CODETABLE2"" " &
                   "FROM ""REPLICATE"" t1  " &
                   "WHERE t1.""TABLENAME"" = 'OMRC' " &
                   "ORDER BY t1.""DBNAMEORIG"", t1.""DBNAMEDEST"" "

            oDt = New System.Data.DataTable
            oDt = refDI.SQL.sqlComoDataTable(sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("DBNAMEORIG").ToString
                sDBD = oDt.Rows.Item(0).Item("DBNAMEDEST").ToString
                Conexiones.Connect_Company(oCompanyO, "DI", oDt.Rows.Item(0).Item("DBNAMEORIG").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                Conexiones.Connect_Company(oCompanyD, "DI", oDt.Rows.Item(0).Item("DBNAMEDEST").ToString, Conexiones.sUser, Conexiones.sPwd, olog)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)
                            Conexiones.Connect_Company(oCompanyO, "DI", oDt.Rows.Item(0).Item("DBNAMEORIG").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                            sDBO = oDt.Rows.Item(i).Item("DBNAMEORIG").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("DBNAMEDEST").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, "DI", oDt.Rows.Item(0).Item("DBNAMEDEST").ToString, Conexiones.sUser, Conexiones.sPwd, olog)

                            sDBD = oDt.Rows.Item(i).Item("DBNAMEDEST").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOMRC = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oManufacturers), SAPbobsCOM.Manufacturers)

                        If oOMRC.GetByKey(CInt(oDt.Rows.Item(i).Item("CODETABLE").ToString)) = True Then
                            sXML = oOMRC.GetAsXML
                        Else
                            sXML = ""
                        End If

                        If sXML <> "" Then
                            'Porque en el modo Update no funciona por XML
                            sFirmName = oOMRC.ManufacturerName
                            oUserFields = oOMRC.UserFields
                            '''''''''''''''''''''''''''''''''''''''''''''

                            oOMRC = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oManufacturers), SAPbobsCOM.Manufacturers)

                            oOMRC = CType(oCompanyD.GetBusinessObjectFromXML(sXML, 0), SAPbobsCOM.Manufacturers)

                            sFirmCode = Conexiones.GetValueDB(oDB, """" & sDBD & """.""OMRC""", """FirmCode""", """FirmName"" = '" & oDt.Rows.Item(i).Item("CODETABLE2").ToString & "'")

                            If sFirmCode = "" Then
                                'Añadir
                                If oOMRC.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOMRC.GetByKey(CInt(sFirmCode)) = True Then
                                    oOMRC.ManufacturerName = sFirmName

                                    For h As Integer = 0 To oUserFields.Fields.Count - 1
                                        If oOMRC.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).IsNull = SAPbobsCOM.BoYesNoEnum.tNO Then
                                            oOMRC.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value = oUserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value
                                        End If
                                    Next

                                    If oOMRC.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOMRC.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM ""REPLICATE"" WHERE ""DBNAMEORIG"" = '" & sDBO & "' AND ""DBNAMEDEST"" = '" & sDBD & "' AND ""TABLENAME"" = '" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "' AND ""CODETABLE"" = '" & oDt.Rows.Item(i).Item("CODETABLE").ToString & "'"
                        refDI.SQL.executeNonQuery(sSQL)
                        'Conexiones.ExecuteSqlDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        olog.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "|" & oDt.Rows.Item(i).Item("CODETABLE").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        olog.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "|" & oDt.Rows.Item(i).Item("CODETABLE").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            olog.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            olog.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOMRC IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOMRC)

            Conexiones.Disconnect_SQLHANA(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
            Conexiones.Disconnect_Company(oCompany)
        End Try
    End Sub

    Public Shared Sub OITG()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOITG As SAPbobsCOM.ItemProperties = Nothing
        Dim oDB As HanaConnection = Nothing
        Dim olog As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sItmsGrpNam As String = ""
        Dim oUserFields As SAPbobsCOM.UserFields = Nothing
        Dim refDI As EXO_DIAPI.EXO_DIAPI = Nothing
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim tipoServidor As SAPbobsCOM.BoDataServerTypes = SAPbobsCOM.BoDataServerTypes.dst_HANADB
        Dim tipocliente As EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico
        Dim sPass As String = Conexiones.Datos_Confi("DI", "Password")

        Try
            olog = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OITG.txt", 1)
            Conexiones.Connect_SQLHANA(oDB, "HANA", olog)
            Conexiones.Connect_Company(oCompany, "DI", Conexiones.sBBDD, Conexiones.sUser, Conexiones.sPwd, olog)
            Try
                refDI = New EXO_DIAPI.EXO_DIAPI(oCompany, olog)
            Catch ex As Exception
                refDI = New EXO_DIAPI.EXO_DIAPI(tipoServidor, oCompany.Server.ToString, oCompany.LicenseServer.ToString, oCompany.CompanyDB.ToString, oCompany.UserName.ToString, sPass, tipocliente)
            End Try



            sSQL = "SELECT t1.""DBNAMEORIG"", t1.""DBNAMEDEST"", t1.""TABLENAME"", t1.""CODETABLE"", t1.""CODETABLE2"" " &
                   "FROM ""REPLICATE"" t1  " &
                   "WHERE t1.""TABLENAME"" = 'OITG' " &
                   "ORDER BY t1.""DBNAMEORIG"", t1.""DBNAMEDEST"" "

            oDt = New System.Data.DataTable
            oDt = refDI.SQL.sqlComoDataTable(sSQL)
            'Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("DBNAMEORIG").ToString
                sDBD = oDt.Rows.Item(0).Item("DBNAMEDEST").ToString

                Conexiones.Connect_Company(oCompanyO, "DI", oDt.Rows.Item(0).Item("DBNAMEORIG").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                Conexiones.Connect_Company(oCompanyD, "DI", oDt.Rows.Item(0).Item("DBNAMEDEST").ToString, Conexiones.sUser, Conexiones.sPwd, olog)



                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("DBNAMEORIG").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)
                            Conexiones.Connect_Company(oCompanyO, "DI", oDt.Rows.Item(0).Item("DBNAMEORIG").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                            sDBO = oDt.Rows.Item(i).Item("DBNAMEORIG").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("DBNAMEDEST").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)
                            Conexiones.Connect_Company(oCompanyD, "DI", oDt.Rows.Item(0).Item("DBNAMEDEST").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                            sDBD = oDt.Rows.Item(i).Item("DBNAMEDEST").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOITG = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemProperties), SAPbobsCOM.ItemProperties)

                        If oOITG.GetByKey(CInt(oDt.Rows.Item(i).Item("CODETABLE").ToString)) = True Then
                            sXML = oOITG.GetAsXML
                        Else
                            sXML = ""
                        End If

                        If sXML <> "" Then
                            'Porque en el modo Update no funciona por XML
                            sItmsGrpNam = oOITG.PropertyName
                            oUserFields = oOITG.UserFields
                            '''''''''''''''''''''''''''''''''''''''''''''

                            oOITG = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemProperties), SAPbobsCOM.ItemProperties)

                            oOITG = CType(oCompanyD.GetBusinessObjectFromXML(sXML, 0), SAPbobsCOM.ItemProperties)

                            'Modificar"
                            'Porque en el modo Update no funciona por XML
                            If oOITG.GetByKey(CInt(oDt.Rows.Item(i).Item("CODETABLE").ToString)) = True Then
                                oOITG.PropertyName = sItmsGrpNam

                                For h As Integer = 0 To oUserFields.Fields.Count - 1
                                    If oOITG.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).IsNull = SAPbobsCOM.BoYesNoEnum.tNO Then
                                        oOITG.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value = oUserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value
                                    End If
                                Next

                                If oOITG.Update() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            End If
                        End If

                        sSQL = "DELETE FROM ""REPLICATE"" WHERE ""DBNAMEORIG"" = '" & sDBO & "' AND ""DBNAMEDEST"" = '" & sDBD & "' AND ""TABLENAME"" = '" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "' AND ""CODETABLE"" = '" & oDt.Rows.Item(i).Item("CODETABLE").ToString & "'"
                        refDI.SQL.executeNonQuery(sSQL)
                        'Conexiones.ExecuteSqlDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        olog.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "|" & oDt.Rows.Item(i).Item("CODETABLE").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        olog.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "|" & oDt.Rows.Item(i).Item("CODETABLE").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            olog.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            olog.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOITG IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOITG)

            Conexiones.Disconnect_SQLHANA(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
            Conexiones.Disconnect_Company(oCompany)
        End Try
    End Sub

    Public Shared Sub OITM()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOITM As SAPbobsCOM.Items = Nothing
        Dim oOITM2 As SAPbobsCOM.Items = Nothing
        Dim oDB As HanaConnection = Nothing
        Dim olog As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim oDt2 As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim oXml As Xml.XmlDocument = Nothing
        Dim oXml2 As Xml.XmlDocument = Nothing
        Dim oXmlNode As Xml.XmlNode = Nothing
        Dim oXmlNode2 As Xml.XmlNode = Nothing
        Dim oXmlNodes As Xml.XmlNodeList = Nothing
        Dim oXmlNodes2 As Xml.XmlNodeList = Nothing
        Dim sSeries As String = ""
        Dim sItmsGrpCod As String = ""
        Dim sCstGrpCode As String = ""
        Dim sFirmCode As String = ""
        Dim sShipType As String = ""
        Dim sCodArt As String = ""

        Dim refDI As EXO_DIAPI.EXO_DIAPI = Nothing
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim tipoServidor As SAPbobsCOM.BoDataServerTypes = SAPbobsCOM.BoDataServerTypes.dst_HANADB
        Dim tipocliente As EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico
        Dim sPass As String = Conexiones.Datos_Confi("DI", "Password")

        Try
            olog = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OITM.txt", 1)
            Conexiones.Connect_SQLHANA(oDB, "HANA", olog)
            Conexiones.Connect_Company(oCompany, "DI", Conexiones.sBBDD, Conexiones.sUser, Conexiones.sPwd, olog)
            Try
                refDI = New EXO_DIAPI.EXO_DIAPI(oCompany, olog)
            Catch ex As Exception
                refDI = New EXO_DIAPI.EXO_DIAPI(tipoServidor, oCompany.Server.ToString, oCompany.LicenseServer.ToString, oCompany.CompanyDB.ToString, oCompany.UserName.ToString, sPass, tipocliente)
            End Try



            sSQL = "SELECT t1.""DBNAMEORIG"", t1.""DBNAMEDEST"", t1.""TABLENAME"", t1.""CODETABLE"", t1.""CODETABLE2"" " &
                   "FROM ""REPLICATE"" t1  " &
                   "WHERE t1.""TABLENAME"" = 'OITM' " &
                   "ORDER BY t1.""DBNAMEORIG"", t1.""DBNAMEDEST"" ASC , t1.""CODETABLE"" DESC "

            oDt = New System.Data.DataTable
            oDt = refDI.SQL.sqlComoDataTable(sSQL)
            'Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("DBNAMEORIG").ToString
                sDBD = oDt.Rows.Item(0).Item("DBNAMEDEST").ToString

                Conexiones.Connect_Company(oCompanyO, "DI", oDt.Rows.Item(0).Item("DBNAMEORIG").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                Conexiones.Connect_Company(oCompanyD, "DI", oDt.Rows.Item(0).Item("DBNAMEDEST").ToString, Conexiones.sUser, Conexiones.sPwd, olog)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("DBNAMEORIG").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)
                            Conexiones.Connect_Company(oCompanyO, "DI", oDt.Rows.Item(0).Item("DBNAMEORIG").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                            sDBO = oDt.Rows.Item(i).Item("DBNAMEORIG").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("DBNAMEDEST").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)
                            Conexiones.Connect_Company(oCompanyD, "DI", oDt.Rows.Item(0).Item("DBNAMEDEST").ToString, Conexiones.sUser, Conexiones.sPwd, olog)
                            sDBD = oDt.Rows.Item(i).Item("DBNAMEDEST").ToString
                        End If

                        'If sDBD = sBDQuitar Then



                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOITM = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems), SAPbobsCOM.Items)
                        oOITM2 = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems), SAPbobsCOM.Items)

                        If oOITM.GetByKey(oDt.Rows.Item(i).Item("CODETABLE").ToString) = True Then
                            sXML = oOITM.GetAsXML
                        Else
                            sXML = ""
                        End If

                        If sXML <> "" Then
                            'Esto es porque hay ciertos campos que son autonuméricos y no tienen por qué ser igual en todas las empresas
                            sItmsGrpCod = Conexiones.GetValueDB(oDB, """" & sDBO & """.""OITB""", """ItmsGrpNam""", """ItmsGrpCod"" = " & oOITM.ItemsGroupCode & "")
                            sItmsGrpCod = Conexiones.GetValueDB(oDB, """" & sDBD & """.""OITB""", """ItmsGrpCod""", """ItmsGrpNam"" = '" & sItmsGrpCod & "'")
                            sCstGrpCode = Conexiones.GetValueDB(oDB, """" & sDBO & """.""OARG""", """CstGrpName""", """CstGrpCode"" = " & oOITM.CustomsGroupCode & "")
                            sCstGrpCode = Conexiones.GetValueDB(oDB, """" & sDBD & """.""OARG""", """CstGrpCode""", """CstGrpName"" = '" & sCstGrpCode & "'")
                            sFirmCode = Conexiones.GetValueDB(oDB, """" & sDBO & """.""OMRC""", """FirmName""", """FirmCode"" = " & oOITM.Manufacturer & "")
                            sFirmCode = Conexiones.GetValueDB(oDB, """" & sDBD & """.""OMRC""", """FirmCode""", """FirmName"" = '" & sFirmCode & "'")


                            'log.escribeMensaje("2", EXO_Log.EXO_Log.Tipo.informacion)

                            'La serie de Items en InterCompany siempre la manual al añadir y al modificar la que tenga en destino
                            sSeries = Conexiones.GetValueDB(oDB, """" & sDBD & """.""OITM""", """Series""", """ItemCode"" = '" & oDt.Rows.Item(i).Item("CODETABLE").ToString & "'")
                            If sSeries = "" Then
                                sSeries = Conexiones.GetValueDB(oDB, """" & sDBD & """.""NNM1""", """Series""", """ObjectCode"" = '4' AND ""SeriesName"" = 'Manual'")
                            End If


                            If Conexiones.GetValueDB(oDB, """" & sDBD & """.""OITM""", """ItemCode""", """ItemCode"" = '" & oDt.Rows.Item(i).Item("CODETABLE").ToString & "'") = "" Then
                                'Añadir
                                oOITM2.ItemCode = oDt.Rows.Item(i).Item("CODETABLE").ToString
                                oOITM2.ItemName = oOITM.ItemName
                                'grupo de articulos
                                If sItmsGrpCod <> "" Then
                                    oOITM2.ItemsGroupCode = CInt(sItmsGrpCod)
                                End If

                                oOITM2.InventoryItem = BoYesNoEnum.tYES
                                oOITM2.SalesItem = BoYesNoEnum.tYES
                                oOITM2.PurchaseItem = BoYesNoEnum.tYES
                                'longitud compras 
                                If oOITM.PurchaseLengthUnit <> 0 Then
                                    oOITM2.PurchaseLengthUnit = oOITM.PurchaseLengthUnit
                                End If
                                If oOITM.PurchaseUnitLength <> 0 Then
                                    oOITM2.PurchaseUnitLength = oOITM.PurchaseUnitLength
                                End If

                                'ancho compras
                                If oOITM.PurchaseUnitWidth <> 0 Then
                                    oOITM2.PurchaseUnitWidth = oOITM.PurchaseUnitWidth
                                End If

                                'peso compras
                                If oOITM.PurchaseUnitWeight <> 0 Then
                                    oOITM2.PurchaseUnitWeight = oOITM.PurchaseUnitWeight
                                End If


                                'longitud ventas
                                oOITM2.SalesUnit = oOITM.SalesUnit
                                oOITM2.SalesUnitLength = oOITM.SalesUnitLength
                                'ancho ventas
                                oOITM2.SalesUnitWidth = oOITM.SalesUnitWidth
                                'peso ventas
                                oOITM2.SalesUnitWeight = oOITM.SalesUnitWeight


                                'fabricante
                                If sFirmCode <> "" Then
                                    oOITM2.Manufacturer = CInt(sFirmCode)
                                End If

                                'clase de expedecion
                                If sShipType <> "" Then
                                    oOITM2.ShipType = CInt(sShipType)
                                End If

                                For j As Integer = 1 To 64
                                    If oOITM.Properties(j) = SAPbobsCOM.BoYesNoEnum.tYES Then
                                        oOITM2.Properties(j) = SAPbobsCOM.BoYesNoEnum.tYES
                                    End If
                                Next

                                If oOITM2.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                If oOITM2.GetByKey(oDt.Rows.Item(i).Item("CODETABLE").ToString) = True Then
                                    oOITM2.ItemName = oOITM.ItemName
                                    If sItmsGrpCod <> "" Then
                                        oOITM2.ItemsGroupCode = CInt(sItmsGrpCod)
                                    End If

                                    oOITM2.InventoryItem = BoYesNoEnum.tYES
                                    oOITM2.SalesItem = BoYesNoEnum.tYES
                                    oOITM2.PurchaseItem = BoYesNoEnum.tYES
                                    'longitud compras 
                                    oOITM2.PurchaseLengthUnit = oOITM.PurchaseLengthUnit
                                    oOITM2.PurchaseUnitLength = oOITM.PurchaseUnitLength
                                    'ancho compras
                                    oOITM2.PurchaseUnitWidth = oOITM.PurchaseUnitWidth
                                    'peso compras
                                    oOITM2.PurchaseUnitWeight = oOITM.PurchaseUnitWeight

                                    'longitud ventas
                                    oOITM2.SalesUnit = oOITM.SalesUnit
                                    oOITM2.SalesUnitLength = oOITM.SalesUnitLength
                                    'ancho ventas
                                    oOITM2.SalesUnitWidth = oOITM.SalesUnitWidth
                                    'peso ventas
                                    oOITM2.SalesUnitWeight = oOITM.SalesUnitWeight


                                    'fabricante
                                    If sFirmCode <> "" Then
                                        oOITM2.Manufacturer = CInt(sFirmCode)
                                    End If

                                    'clase de expedecion
                                    If sShipType <> "" Then
                                        oOITM2.ShipType = CInt(sShipType)
                                    End If

                                    For j As Integer = 1 To 64
                                        If oOITM.Properties(j) = SAPbobsCOM.BoYesNoEnum.tYES Then
                                            oOITM2.Properties(j) = SAPbobsCOM.BoYesNoEnum.tYES
                                        End If
                                    Next


                                    If oOITM2.Update <> 0 Then
                                        olog.escribeMensaje("error update", EXO_Log.EXO_Log.Tipo.informacion)
                                        olog.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "|" & oDt.Rows.Item(i).Item("CODETABLE").ToString, EXO_Log.EXO_Log.Tipo.error)
                                    End If

                                End If
                            End If

                            sSQL = "DELETE FROM ""REPLICATE"" WHERE ""DBNAMEORIG"" = '" & sDBO & "' AND ""DBNAMEDEST"" = '" & sDBD & "' AND ""TABLENAME"" = '" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "' AND ""CODETABLE"" = '" & oDt.Rows.Item(i).Item("CODETABLE").ToString & "'"
                            refDI.SQL.executeNonQuery(sSQL)

                        End If

                        'End If
                    Catch exCOM As System.Runtime.InteropServices.COMException
                        olog.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "|" & oDt.Rows.Item(i).Item("CODETABLE").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        olog.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("TABLENAME").ToString & "|" & oDt.Rows.Item(i).Item("CODETABLE").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            olog.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            olog.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOITM IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOITM)
            If oOITM2 IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOITM2)

            Conexiones.Disconnect_SQLHANA(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
            Conexiones.Disconnect_Company(oCompany)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
