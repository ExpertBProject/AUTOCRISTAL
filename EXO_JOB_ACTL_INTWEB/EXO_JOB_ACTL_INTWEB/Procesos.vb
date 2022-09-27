Imports Sap.Data.Hana
Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml
Imports SAPbobsCOM

Public Class Procesos
    Public Shared Sub LecturaTabla(ByRef db As HanaConnection, ByRef dbWEB As HanaConnection, ByRef oCompany As SAPbobsCOM.Company, ByRef oLog As EXO_Log.EXO_Log)
#Region "Variables"
        Dim sError As String = ""
        Dim sSQL As String = ""
        Dim sBBDDWEB As String = Conexiones.Datos_Confi("HANAWEB", "databaseName")
        Dim odtDatosWeb As System.Data.DataTable = New System.Data.DataTable
        Dim sCliente As String = "" : Dim sID As String = ""
        Dim oORDR As SAPbobsCOM.Documents = Nothing : Dim iLin As Integer = 0

        Dim sSubject As String = ""
        Dim sTipo As String = ""
        Dim sComen As String = ""
        Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sDELALMACEN As String = "" : Dim sDELEGACION As String = ""
#End Region
        Try
            oLog.escribeMensaje("Lectura del carrito...", EXO_Log.EXO_Log.Tipo.informacion)
            sSQL = "SELECT * FROM """ & sBBDDWEB & """.""CARRITO""  WHERE ""CONFIRMADO""=1 AND ""NPEDIDO""=0 AND ""REPROCESAR""=0 ORDER BY ""ID"" "
            oLog.escribeMensaje("SQL: " & sSQL, EXO_Log.EXO_Log.Tipo.informacion)
            Conexiones.FillDtDB(dbWEB, odtDatosWeb, sSQL)
            If odtDatosWeb.Rows.Count > 0 Then
                For iCab As Integer = 0 To odtDatosWeb.Rows.Count - 1
                    If sCliente <> odtDatosWeb.Rows.Item(iCab).Item("USUARIO").ToString Then
                        iLin = 0
                        If iCab <> 0 Then
                            If oORDR.Add() <> 0 Then
                                sError = oCompany.GetLastErrorCode.ToString & " / " & oCompany.GetLastErrorDescription.Replace("'", "")
                                oLog.escribeMensaje("Se ha producido una error al crear el pedido web del cliente " & sCliente & vbCrLf & sError & "", EXO_Log.EXO_Log.Tipo.error)

                                'Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas
                                sSubject = "Pedido WEB del Cliente " & sCliente & " con error: " & sError
                                sTipo = "Pedido WEB"
                                sComen = sError
                                EnviarAlerta(oLog, oCompany, "", "", "", sSubject, sTipo, sComen, "", sDELEGACION)
                            Else
                                oCompany.GetNewObjectCode(sDocEntry)
                                sDocNum = Conexiones.GetValueDB(db, " """ & oCompany.CompanyDB & """.""ORDR""", """DocNum""", """DocEntry"" = " & sDocEntry & "", oLog)

                                'udpate BBDD
                                sSQL = "UPDATE """ & sBBDDWEB & """.""CARRITO"" SET ""NPEDIDO""='" & sDocNum & "',""NUMPEDIDO""='" & sDocEntry & "' WHERE ""USUARIO""='" & sCliente & "' and ""ID"" IN(" & sID & ") "
                                oLog.escribeMensaje("SQL: " & sSQL, EXO_Log.EXO_Log.Tipo.informacion)
                                Conexiones.ExecuteSqlDB(dbWEB, sSQL)
                                oLog.escribeMensaje("Se ha Actualizado la tabla de la BBDD " & sBBDDWEB, EXO_Log.EXO_Log.Tipo.informacion)
                                'Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas
                                sSubject = "Pedido WEB de Venta " & sDocNum & " se ha registrado correctamente con el cliente " & sCliente
                                sTipo = "Pedido de Cliente WEB"
                                oLog.escribeMensaje(sSubject, EXO_Log.EXO_Log.Tipo.advertencia)
                                sComen = ""
                                EnviarAlerta(oLog, oCompany, sDocNum, sDocEntry, "17", sSubject, sTipo, sComen, "", sDELEGACION)
                            End If
                            sID = odtDatosWeb.Rows.Item(iCab).Item("ID").ToString
                        Else
                            sID = odtDatosWeb.Rows.Item(iCab).Item("ID").ToString
                        End If


                        sCliente = odtDatosWeb.Rows.Item(iCab).Item("USUARIO").ToString
                        oORDR = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
#Region "Serie"
                        Dim sSerieName As String = Conexiones.GetValueDB(db, " """ & oCompany.CompanyDB & """.""@EXO_OGEN1""", """U_EXO_INFV""", """U_EXO_NOMV"" ='EXO_SERIEPEDWEB' and ""Code""='EXO_KERNEL'", oLog)
                        Dim sSerie As String = Conexiones.GetValueDB(db, " """ & oCompany.CompanyDB & """.""NNM1""", """Series""", """ObjectCode""='17' and ""SeriesName""='" & sSerieName & "'", oLog)
                        If sSerie <> "" Then
                            oORDR.Series = CInt(sSerie)
                        End If
#End Region
                        oORDR.CardCode = sCliente
#Region "Dirección"
                        If odtDatosWeb.Rows.Item(iCab).Item("DIRECCION_ENVIO").ToString <> "" Then
                            oORDR.AddressExtension.ShipToStreet = odtDatosWeb.Rows.Item(iCab).Item("DIRECCION_ENVIO").ToString
                            oORDR.AddressExtension.ShipToZipCode = odtDatosWeb.Rows.Item(iCab).Item("CP_ENVIO").ToString
                            oORDR.AddressExtension.ShipToCity = odtDatosWeb.Rows.Item(iCab).Item("MUNICIPIO_ENVIO").ToString
                            oORDR.AddressExtension.ShipToCounty = odtDatosWeb.Rows.Item(iCab).Item("PROVINCIA_ENVIO").ToString
                            oORDR.AddressExtension.ShipToCountry = odtDatosWeb.Rows.Item(iCab).Item("PAIS_ENVIO").ToString
                        End If
#End Region
#Region "Autorizado"
                        If odtDatosWeb.Rows.Item(iCab).Item("TPV").ToString.Trim = "0" Then
                            oORDR.Confirmed = BoYesNoEnum.tYES
                        ElseIf odtDatosWeb.Rows.Item(iCab).Item("TPV").ToString.Trim = "1" And odtDatosWeb.Rows.Item(iCab).Item("PAGADO").ToString.Trim = "1" Then
                            oORDR.Confirmed = BoYesNoEnum.tNO
                            'Condición de pago
                            oORDR.GroupNumber = -1
                            Dim sPMethod As String = Conexiones.GetValueDB(db, " """ & oCompany.CompanyDB & """.""@EXO_OGEN1""", """U_EXO_INFV""", """U_EXO_NOMV"" ='EXO_VIAPAGO' and ""Code""='EXO_KERNEL'", oLog)
                            If sPMethod <> "" Then
                                oORDR.PaymentMethod = sPMethod
                            End If

                        End If
#End Region
                        oORDR.TaxDate = CDate(odtDatosWeb.Rows.Item(iCab).Item("FECHA").ToString)
                        oORDR.DocDueDate = CDate(odtDatosWeb.Rows.Item(iCab).Item("FECHA").ToString)
                        Dim sAgencia As String = Conexiones.GetValueDB(db, " """ & oCompany.CompanyDB & """.""OCRD""", """U_EXO_AGENCIA""", """CardCode"" ='" & odtDatosWeb.Rows.Item(iCab).Item("CLIENTE").ToString & "'", oLog)
                        Dim sTransporte As String = ""
                        sDELALMACEN = odtDatosWeb.Rows.Item(iCab).Item("TRANSPORTE").ToString
                        'If sAgencia = "" Or sAgencia = "-" Then
                        '    sTransporte = Conexiones.GetValueDB(db, " """ & oCompany.CompanyDB & """.""OSHP""", """TrnspCode""", """U_EXO_SERVIC"" ='" & odtDatosWeb.Rows.Item(iCab).Item("TRANSPORTE").ToString & "'", oLog)
                        'Else
                        '    sTransporte = Conexiones.GetValueDB(db, " """ & oCompany.CompanyDB & """.""OSHP""", """TrnspCode""", """U_EXO_SERVIC"" = '" & odtDatosWeb.Rows.Item(iCab).Item("TRANSPORTE").ToString & "' and ""U_EXO_AGE""='" & sAgencia & "' ", oLog)
                        'End If
                        sTransporte = odtDatosWeb.Rows.Item(iCab).Item("TRANSPORTE_F").ToString
                        If IsNumeric(sTransporte) Then
                            oORDR.TransportationCode = CInt(sTransporte)
                        End If

                        oORDR.Comments = "Pedido creado desde WEB. " & ChrW(13) & ChrW(10) & odtDatosWeb.Rows.Item(iCab).Item("OBSERVACIONES").ToString
                        oLog.escribeMensaje("Tratando Documento de Cliente " & sCliente & "...", EXO_Log.EXO_Log.Tipo.informacion)
                    Else
                        iLin += 1
                        sID &= "," & odtDatosWeb.Rows.Item(iCab).Item("ID").ToString
                    End If
                    If iLin <> 0 Then
                        oORDR.Lines.Add()
                    End If
                    oORDR.Lines.ItemCode = odtDatosWeb.Rows.Item(iCab).Item("CREF").ToString
                    oORDR.Lines.Quantity = CDbl(odtDatosWeb.Rows.Item(iCab).Item("NUNIDADES").ToString)
                    oORDR.Lines.UnitPrice = CDbl(odtDatosWeb.Rows.Item(iCab).Item("PRECIO").ToString)
                    oORDR.Lines.UserFields.Fields.Item("U_EXO_DCT001").Value = CDbl(odtDatosWeb.Rows.Item(iCab).Item("DTO").ToString)
                    oORDR.Lines.UserFields.Fields.Item("U_EXO_DCT002").Value = CDbl(odtDatosWeb.Rows.Item(iCab).Item("DTO_WEB").ToString)
                    oORDR.Lines.DiscountPercent = (CDbl(odtDatosWeb.Rows.Item(iCab).Item("DTO").ToString) + CDbl(odtDatosWeb.Rows.Item(iCab).Item("DTO_WEB").ToString) - ((CDbl(odtDatosWeb.Rows.Item(iCab).Item("DTO").ToString) * CDbl(odtDatosWeb.Rows.Item(iCab).Item("DTO_WEB").ToString) / 100)))
                    sDELEGACION = odtDatosWeb.Rows.Item(iCab).Item("ALMACEN").ToString
                    Dim sAlmacen As String = Conexiones.GetValueDB(db, " """ & oCompany.CompanyDB & """.""OWHS""", """WhsCode""", """U_EXO_SUCURSAL"" = " & odtDatosWeb.Rows.Item(iCab).Item("ALMACEN").ToString & " AND ""U_EXO_PRINCIPAL""='Y' ", oLog)
                    oORDR.Lines.WarehouseCode = sAlmacen
                Next
                If oORDR.Add() <> 0 Then
                    sError = oCompany.GetLastErrorCode.ToString & " / " & oCompany.GetLastErrorDescription.Replace("'", "")
                    oLog.escribeMensaje("Se ha producido una error al crear el pedido web del cliente " & sCliente & vbCrLf & sError & "", EXO_Log.EXO_Log.Tipo.error)

                    'Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas
                    sSubject = "Pedido WEB del Cliente " & sCliente & " ha tenido un error"
                    sTipo = "Pedido WEB"
                    sComen = sError
                    EnviarAlerta(oLog, oCompany, "", "", "", sSubject, sTipo, sComen, "", sDELEGACION)
                Else
                    oCompany.GetNewObjectCode(sDocEntry)
                    sDocNum = Conexiones.GetValueDB(db, " """ & oCompany.CompanyDB & """.""ORDR""", """DocNum""", """DocEntry"" = " & sDocEntry & "", oLog)

                    'udpate BBDD
                    sSQL = "UPDATE """ & sBBDDWEB & """.""CARRITO"" SET ""NPEDIDO""='" & sDocNum & "',""NUMPEDIDO""='" & sDocEntry & "' WHERE ""USUARIO""='" & sCliente & "' and ""ID"" IN(" & sID & ") "
                    Conexiones.ExecuteSqlDB(dbWEB, sSQL)

                    'Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas
                    sSubject = "Pedido WEB de Venta " & sDocNum & " se ha registrado correctamente con el cliente " & sCliente
                    sTipo = "Pedido de Cliente WEB"
                    oLog.escribeMensaje(sSubject, EXO_Log.EXO_Log.Tipo.advertencia)
                    sComen = ""
                    EnviarAlerta(oLog, oCompany, sDocNum, sDocEntry, "17", sSubject, sTipo, sComen, "", sDELEGACION)
                End If
            Else
                oLog.escribeMensaje("##### No existen registros para crear pedidos.", EXO_Log.EXO_Log.Tipo.advertencia)
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally

        End Try
    End Sub
    Public Shared Sub REPROCESAR(ByRef db As HanaConnection, ByRef dbWEB As HanaConnection, ByRef oCompany As SAPbobsCOM.Company, ByRef oLog As EXO_Log.EXO_Log)
#Region "Variables"
        Dim sError As String = ""
        Dim sSQL As String = ""
        Dim sBBDDWEB As String = Conexiones.Datos_Confi("HANAWEB", "databaseName")
        Dim odtDatosWeb As System.Data.DataTable = New System.Data.DataTable
        Dim sCliente As String = "" : Dim sID As String = ""
        Dim oORDR As SAPbobsCOM.Documents = Nothing : Dim iLin As Integer = 0
        Dim sPAGADO As String = ""
        Dim sSubject As String = ""
        Dim sTipo As String = ""
        Dim sComen As String = ""
        Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sDELALMACEN As String = "" : Dim sDELEGACION As String = ""
#End Region
        Try
            oLog.escribeMensaje("Reprocesar el carrito...", EXO_Log.EXO_Log.Tipo.informacion)
            sSQL = "SELECT * FROM """ & sBBDDWEB & """.""CARRITO""  WHERE ""NPEDIDO""<>0 AND ""REPROCESAR""=1 ORDER BY ""ID"" "
            oLog.escribeMensaje("SQL: " & sSQL, EXO_Log.EXO_Log.Tipo.informacion)
            Conexiones.FillDtDB(dbWEB, odtDatosWeb, sSQL)
            If odtDatosWeb.Rows.Count > 0 Then
                For iCab As Integer = 0 To odtDatosWeb.Rows.Count - 1
                    If sCliente <> odtDatosWeb.Rows.Item(iCab).Item("USUARIO").ToString Then
                        sCliente = odtDatosWeb.Rows.Item(iCab).Item("USUARIO").ToString
                        sDELEGACION = odtDatosWeb.Rows.Item(iCab).Item("ALMACEN").ToString
                        sDocNum = odtDatosWeb.Rows.Item(iCab).Item("NPEDIDO").ToString
                        sDocEntry = odtDatosWeb.Rows.Item(iCab).Item("NUMPEDIDO").ToString
                        sID = odtDatosWeb.Rows.Item(iCab).Item("ID").ToString
                        oLog.escribeMensaje("Tratando pedido Nº " & sDocNum & " de Cliente " & sCliente & "...", EXO_Log.EXO_Log.Tipo.informacion)

                        oORDR = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                        If oORDR.GetByKey(CInt(sDocEntry)) = True Then
                            sPAGADO = odtDatosWeb.Rows.Item(iCab).Item("PAGADO").ToString
                            Select Case sPAGADO.Trim
                                Case "2" 'Autorizar pedido
                                    oORDR.Confirmed = BoYesNoEnum.tYES
                                    If oORDR.Update() <> 0 Then
                                        sError = oCompany.GetLastErrorCode.ToString & " / " & oCompany.GetLastErrorDescription.Replace("'", "")
                                        oLog.escribeMensaje("Se ha producido un error al autorizar el pedido web Nº Interno " & sDocEntry & " del cliente " & sCliente & vbCrLf & sError & "", EXO_Log.EXO_Log.Tipo.error)

                                        'Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas
                                        sSubject = "Pedido WEB del Cliente " & sCliente & " no se ha podido autorizar debido a un error"
                                        sTipo = "Pedido WEB"
                                        sComen = sError
                                        EnviarAlerta(oLog, oCompany, sDocNum, sDocEntry, "17", sSubject, sTipo, sComen, "", sDELEGACION)
                                    Else
                                        'udpate BBDD
                                        sSQL = "UPDATE """ & sBBDDWEB & """.""CARRITO"" SET ""REPROCESAR""=0 WHERE ""USUARIO""='" & sCliente & "' and ""ID"" IN(" & sID & ") "
                                        Conexiones.ExecuteSqlDB(dbWEB, sSQL)

                                        'Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas
                                        sSubject = "Pedido WEB de Venta " & sDocNum & " se ha autorizado correctamente con el cliente " & sCliente
                                        sTipo = "Pedido de Cliente WEB"
                                        oLog.escribeMensaje(sSubject, EXO_Log.EXO_Log.Tipo.advertencia)
                                        sComen = ""
                                        EnviarAlerta(oLog, oCompany, sDocNum, sDocEntry, "17", sSubject, sTipo, sComen, "", sDELEGACION)
                                        'Crear Cobro a cuenta por transferencia  a la cta de un parámetro

                                    End If
                                Case "3" 'Cancelar pedido
                                    'oORDR.Comments &= ChrW(13) & ChrW(10) & "CANCELADO POR FALTA DE PAGO VIA WEB."
                                    If oORDR.Cancel() <> 0 Then
                                        'Error
                                        sError = oCompany.GetLastErrorCode.ToString & " / " & oCompany.GetLastErrorDescription.Replace("'", "")
                                        oLog.escribeMensaje("Se ha producido un error al cancelar el pedido web Nº Interno " & sDocEntry & " del cliente " & sCliente & vbCrLf & sError & "", EXO_Log.EXO_Log.Tipo.error)

                                        'Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas
                                        sSubject = "Pedido WEB del Cliente " & sCliente & " no se ha podido cancelar debido a un error"
                                        sTipo = "Pedido WEB"
                                        sComen = sError
                                        EnviarAlerta(oLog, oCompany, sDocNum, sDocEntry, "17", sSubject, sTipo, sComen, "", sDELEGACION)
                                    Else
                                        'udpate BBDD
                                        sSQL = "UPDATE """ & sBBDDWEB & """.""CARRITO"" SET ""REPROCESAR""=0 WHERE ""USUARIO""='" & sCliente & "' and ""ID"" IN(" & sID & ") "
                                        Conexiones.ExecuteSqlDB(dbWEB, sSQL)

                                        'OK
                                        sSQL = "UPDATE """ & oCompany.CompanyDB & """.""ORDR"" SET ""Comments""= ""Comments"" || '" & ChrW(13) & ChrW(10) & "CANCELADO POR FALTA DE PAGO VIA WEB." & "' WHERE ""DocEntry""= " & sDocEntry
                                        oLog.escribeMensaje("SQL: " & sSQL, EXO_Log.EXO_Log.Tipo.informacion)
                                        Conexiones.ExecuteSqlDB(dbWEB, sSQL)
                                        'Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas
                                        sSubject = "Pedido WEB de Venta " & sDocNum & " se ha cancelado por falta de pago con el cliente " & sCliente
                                        sTipo = "Pedido de Cliente WEB"
                                        oLog.escribeMensaje(sSubject, EXO_Log.EXO_Log.Tipo.advertencia)
                                        sComen = ""
                                        EnviarAlerta(oLog, oCompany, sDocNum, sDocEntry, "17", sSubject, sTipo, sComen, "", sDELEGACION)
                                    End If
                            End Select
                        Else
                            oLog.escribeMensaje("Nº Interno: " & sDocEntry & ". No se ha encontrado el pedido web " & sDocNum & " del cliente " & sCliente & ". No se puede procesar.", EXO_Log.EXO_Log.Tipo.error)

                            'Enviamos alerta a los usuarios que estén marcados en la ficha del usuario con el campo Alertas
                            sSubject = "Pedido WEB Nº" & sDocNum & " del Cliente " & sCliente & " no se ha podido reprocesar. No se encuentra"
                            sTipo = "Pedido WEB"
                            sComen = sError
                            EnviarAlerta(oLog, oCompany, "", "", "", sSubject, sTipo, sComen, "", sDELEGACION)
                        End If
                    Else
                        oLog.escribeMensaje("##### No existen registros para crear pedidos.", EXO_Log.EXO_Log.Tipo.advertencia)
                    End If
                Next
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally

        End Try
    End Sub
#Region "Actualizar campos"
#Region "SAP"
    Public Shared Function exiteCampoUsuario(ByVal tabla As String, ByVal campo As String, ByRef interfazDatos As SAPbobsCOM.Company) As Boolean
        Dim rs As SAPbobsCOM.Recordset = CType(interfazDatos.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        If interfazDatos.DbServerType = BoDataServerTypes.dst_HANADB Then
            rs.DoQuery("SELECT COUNT('A') FROM ""CUFD"" WHERE ""TableID"" = '" + tabla + "' AND ""AliasID"" = '" + campo + "'")
        Else
            rs.DoQuery("SELECT COUNT('A') FROM CUFD WHERE TableID = '" + tabla + "' AND AliasID = '" + campo + "'")
        End If
        Dim num As Integer = Int32.Parse(rs.Fields.Item(0).Value.ToString())
        System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
        Return num = 1
    End Function
    Public Shared Function actualizaUDO(ByRef udo As SAPbobsCOM.UserObjectsMD, udoB1Aux As SAPbobsCOM.UserObjectsMD, ByRef oCompany As Company) As Integer
        Dim res As Integer = 0
        Dim xmlasstring As Boolean = oCompany.XMLAsString
        oCompany.XMLAsString = True
        If udo.Code = udoB1Aux.Code Then
            udo.Name = udoB1Aux.Name
            udo.CanArchive = udoB1Aux.CanArchive
            udo.CanCancel = udoB1Aux.CanCancel
            udo.CanClose = udoB1Aux.CanClose
            udo.CanCreateDefaultForm = udoB1Aux.CanCreateDefaultForm
            udo.CanDelete = udoB1Aux.CanDelete
            udo.CanFind = udoB1Aux.CanFind
            udo.CanLog = udoB1Aux.CanLog
            udo.CanYearTransfer = udoB1Aux.CanYearTransfer
            For indiceTablas As Integer = 0 To udoB1Aux.ChildTables.Count - 1
                Dim encontrada As Boolean = False
                udoB1Aux.ChildTables.SetCurrentLine(indiceTablas)
                For indiceTablasOriginales As Integer = 0 To udo.ChildTables.Count - 1
                    udo.ChildTables.SetCurrentLine(indiceTablasOriginales)
                    If udo.ChildTables.TableName = udoB1Aux.ChildTables.TableName Then
                        encontrada = True
                        Exit For
                    End If
                Next
                If Not encontrada Then
                    udo.ChildTables.Add()
                    udo.ChildTables.SetCurrentLine(udo.ChildTables.Count - 1)
                    udo.ChildTables.TableName = udoB1Aux.ChildTables.TableName
                    udo.ChildTables.LogTableName = udoB1Aux.ChildTables.LogTableName
                End If
            Next
            udo.EnableEnhancedForm = udoB1Aux.EnableEnhancedForm
            For indiceForm As Integer = 0 To udoB1Aux.EnhancedFormColumns.Count - 1
                Dim encontrada As Boolean = False
                udoB1Aux.EnhancedFormColumns.SetCurrentLine(indiceForm)
                For indiceFormOriginal As Integer = 0 To udo.EnhancedFormColumns.Count - 1
                    udo.EnhancedFormColumns.SetCurrentLine(indiceFormOriginal)
                    If udo.EnhancedFormColumns.ColumnAlias = udoB1Aux.EnhancedFormColumns.ColumnAlias Then
                        encontrada = True
                        udo.EnhancedFormColumns.ColumnDescription = udoB1Aux.EnhancedFormColumns.ColumnDescription
                        Try
                            udo.EnhancedFormColumns.ColumnIsUsed = udoB1Aux.EnhancedFormColumns.ColumnIsUsed
                        Catch
                        End Try
                        udo.EnhancedFormColumns.ColumnNumber = udoB1Aux.EnhancedFormColumns.ColumnNumber
                        Try
                            udo.EnhancedFormColumns.Editable = udoB1Aux.EnhancedFormColumns.Editable
                        Catch
                        End Try
                        udo.EnhancedFormColumns.ChildNumber = udoB1Aux.EnhancedFormColumns.ChildNumber
                        Exit For
                    End If
                Next
                If Not encontrada Then
                    udo.EnhancedFormColumns.Add()
                    udo.EnhancedFormColumns.SetCurrentLine(udo.EnhancedFormColumns.Count - 1)
                    udo.EnhancedFormColumns.ColumnAlias = udoB1Aux.EnhancedFormColumns.ColumnAlias
                    udo.EnhancedFormColumns.ColumnDescription = udoB1Aux.EnhancedFormColumns.ColumnDescription
                    udo.EnhancedFormColumns.ColumnIsUsed = udoB1Aux.EnhancedFormColumns.ColumnIsUsed
                    udo.EnhancedFormColumns.ColumnNumber = udoB1Aux.EnhancedFormColumns.ColumnNumber
                    udo.EnhancedFormColumns.Editable = udoB1Aux.EnhancedFormColumns.Editable
                    udo.EnhancedFormColumns.ChildNumber = udoB1Aux.EnhancedFormColumns.ChildNumber
                End If
            Next
            udo.ExtensionName = udoB1Aux.ExtensionName
            udo.FatherMenuID = udoB1Aux.FatherMenuID
            For indiceBucar As Integer = 0 To udoB1Aux.FindColumns.Count - 1
                Dim encontrada As Boolean = False
                udoB1Aux.FindColumns.SetCurrentLine(indiceBucar)
                For indiceBuscarOriginal As Integer = 0 To udo.FindColumns.Count - 1
                    udo.FindColumns.SetCurrentLine(indiceBuscarOriginal)
                    If udo.FindColumns.ColumnAlias = udoB1Aux.FindColumns.ColumnAlias Then
                        encontrada = True
                        udo.FindColumns.ColumnDescription = udoB1Aux.FindColumns.ColumnDescription
                        Exit For
                    End If
                Next
                If Not encontrada Then
                    udo.FindColumns.Add()
                    udo.FindColumns.SetCurrentLine(udo.FindColumns.Count - 1)
                    udo.FindColumns.ColumnAlias = udoB1Aux.FindColumns.ColumnAlias
                    udo.FindColumns.ColumnDescription = udoB1Aux.FindColumns.ColumnDescription
                End If
            Next
            For indiceFormB As Integer = 0 To udoB1Aux.FormColumns.Count - 1
                Dim encontrada As Boolean = False
                udoB1Aux.FormColumns.SetCurrentLine(indiceFormB)
                For indiceFormBOriginal As Integer = 0 To udo.FormColumns.Count - 1
                    udo.FormColumns.SetCurrentLine(indiceFormBOriginal)
                    If udo.FormColumns.FormColumnAlias = udoB1Aux.FormColumns.FormColumnAlias Then
                        encontrada = True
                        udo.FormColumns.Editable = udoB1Aux.FormColumns.Editable
                        udo.FormColumns.FormColumnDescription = udoB1Aux.FormColumns.FormColumnDescription
                        udo.FormColumns.SonNumber = udoB1Aux.FormColumns.SonNumber
                        Exit For
                    End If
                Next
                If Not encontrada Then
                    udo.FormColumns.Add()
                    udo.FormColumns.SetCurrentLine(udo.FormColumns.Count - 1)
                    udo.FormColumns.FormColumnAlias = udoB1Aux.FormColumns.FormColumnAlias
                    udo.FormColumns.Editable = udoB1Aux.FormColumns.Editable
                    udo.FormColumns.FormColumnDescription = udoB1Aux.FormColumns.FormColumnDescription
                    udo.FormColumns.SonNumber = udoB1Aux.FormColumns.SonNumber
                End If
            Next
            udo.FormSRF = udoB1Aux.FormSRF
            udo.RebuildEnhancedForm = udoB1Aux.RebuildEnhancedForm
            udo.LogTableName = udoB1Aux.LogTableName
            udo.ManageSeries = udoB1Aux.ManageSeries
            udo.MenuCaption = udoB1Aux.MenuCaption
            udo.MenuItem = udoB1Aux.MenuItem
            udo.MenuUID = udoB1Aux.MenuUID
            udo.Name = udoB1Aux.Name
            udo.OverwriteDllfile = udoB1Aux.OverwriteDllfile
            udo.Position = udoB1Aux.Position
            udo.TableName = udoB1Aux.TableName
            udo.UseUniqueFormType = udoB1Aux.UseUniqueFormType
            System.Runtime.InteropServices.Marshal.ReleaseComObject(udoB1Aux)
            udoB1Aux = Nothing
            GC.Collect()
            res = udo.Update()
        End If
        oCompany.XMLAsString = xmlasstring
        Return res
    End Function
#End Region
    Public Shared Sub Actualizar_Campos(ByRef oLog As EXO_Log.EXO_Log, ByRef oDBSAP As HanaConnection, ByRef oCompany As SAPbobsCOM.Company, ByVal errorSBO As Boolean)
#Region "Variables"
        Dim oDBSAPSQL As SqlConnection = Nothing
        Dim sError As String = ""
        Dim sSQL As String = ""
        Dim OdtDatos As System.Data.DataTable = Nothing
        Dim sPass As String = "" : Dim sVSQL As String = ""
        Dim oXML As String = ""
        Dim sDir As String = Application.StartupPath
#End Region
        Try
#Region "CREAR EN SAP EN BBDD"
            Dim sruta As String = ""
            For F = 0 To 2
                Select Case F
                    Case 0 : sruta = sDir & "\01.XML\XML_BD\UDFs_OUSR.xml"
                    Case 1 : sruta = sDir & "\01.XML\XML_BD\UDFs_OWHS.xml"
                    Case 2 : sruta = sDir & "\01.XML\XML_BD\UDFs_RDR1.xml"
                End Select
#Region "Importación"
                oLog.escribeMensaje("######                                                      ###### ", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("##################################################################", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("###### Actualizando:  " & Path.GetFileNameWithoutExtension(sruta), EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("##################################################################", EXO_Log.EXO_Log.Tipo.informacion)

                If sruta <> "" Then
                    Dim i As Integer = 4000
                    Dim elementos As Integer
                    Dim codError As Integer

                    If Not errorSBO Then
                        oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                        oCompany.XMLAsString = True
                        If System.IO.File.Exists(sruta) Then
                            Dim docXML As Xml.XmlDocument = New Xml.XmlDocument()
                            docXML.Load(sruta)

                            elementos = oCompany.GetXMLelementCount(docXML.InnerXml)
                            For i = 0 To elementos - 1
                                Select Case oCompany.GetXMLobjectType(docXML.InnerXml, i)
                                    Case SAPbobsCOM.BoObjectTypes.oUserFields
                                        Dim campoUsuario As SAPbobsCOM.UserFieldsMD
                                        campoUsuario = CType(oCompany.GetBusinessObjectFromXML(docXML.InnerXml, i), SAPbobsCOM.UserFieldsMD)
                                        oLog.escribeMensaje("Campo: " + campoUsuario.Name, EXO_Log.EXO_Log.Tipo.informacion)
                                        If Not exiteCampoUsuario(campoUsuario.TableName, campoUsuario.Name, oCompany) Then
                                            codError = campoUsuario.Add()
                                        End If
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(campoUsuario)
                                        campoUsuario = Nothing
                                    Case SAPbobsCOM.BoObjectTypes.oUserTables
                                        Dim tablaUsuario As SAPbobsCOM.UserTablesMD
                                        tablaUsuario = CType(oCompany.GetBusinessObjectFromXML(docXML.InnerXml, i), SAPbobsCOM.UserTablesMD)
                                        oLog.escribeMensaje("Tabla: " + tablaUsuario.TableName, EXO_Log.EXO_Log.Tipo.informacion)
                                        If Not tablaUsuario.GetByKey(tablaUsuario.TableName) Then
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(tablaUsuario)
                                            tablaUsuario = Nothing
                                            tablaUsuario = CType(oCompany.GetBusinessObjectFromXML(docXML.InnerXml, i), UserTablesMD)
                                            codError = tablaUsuario.Add()
                                        End If
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(tablaUsuario)
                                        tablaUsuario = Nothing
                                'UDOS
                                    Case SAPbobsCOM.BoObjectTypes.oUserObjectsMD
                                        Dim oUDO As SAPbobsCOM.UserObjectsMD = CType(oCompany.GetBusinessObjectFromXML(docXML.InnerXml, i), UserObjectsMD)
                                        '               gProgressBar.Value = gProgressBar.Value + 1
                                        oLog.escribeMensaje("UDO: " + oUDO.Code, EXO_Log.EXO_Log.Tipo.informacion)
                                        Dim oUDO2 As SAPbobsCOM.UserObjectsMD = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD), UserObjectsMD)
                                        If oUDO2.GetByKey(oUDO.Code) Then
                                            Dim xmlUDO As String = oUDO.GetAsXML
                                            codError = actualizaUDO(oUDO2, oUDO, oCompany)
                                            If codError <> 0 Then
                                                oLog.escribeMensaje("Error: " + oCompany.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                                                System.Threading.Thread.Sleep(3000)
                                            End If
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO2)
                                            oUDO2 = Nothing
                                            Continue For
                                        Else
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO2)
                                            oUDO2 = Nothing
                                            GC.Collect()
                                            Dim xmlUDO As String = oUDO.GetAsXML
                                            codError = oUDO.Add
                                            If codError <> 0 And codError <> -2035 Then
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO)
                                                oUDO = Nothing
                                                oLog.escribeMensaje("Error: " + oCompany.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                                                System.Threading.Thread.Sleep(3000)
                                                Exit For
                                            ElseIf codError = -2035 Then
                                            End If
                                            If Not oUDO Is Nothing Then
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO)
                                                oUDO = Nothing
                                            End If
                                        End If
                                    Case SAPbobsCOM.BoObjectTypes.oUserKeys
                                        Dim oKeys As SAPbobsCOM.UserKeysMD = CType(oCompany.GetBusinessObjectFromXML(docXML.InnerXml, i), UserKeysMD)
                                        codError = oKeys.Add
                                        If codError <> 0 And codError <> -1 Then
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oKeys)
                                            oKeys = Nothing
                                            oLog.escribeMensaje("Error: " + oCompany.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                                            System.Threading.Thread.Sleep(3000)
                                            Exit For
                                        End If
                                        If Not oKeys Is Nothing Then
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oKeys)
                                            oKeys = Nothing
                                        End If
                                End Select
                            Next i
                        Else
                            oLog.escribeMensaje("No existe el fichero indicado", EXO_Log.EXO_Log.Tipo.error)
                        End If
                    End If


                Else
                    oLog.escribeMensaje("Debe indicar un fichero", EXO_Log.EXO_Log.Tipo.error)
                End If
#End Region
            Next

#End Region
#Region "Parámetros"
            oLog.escribeMensaje("######                                                      ###### ", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("##################################################################", EXO_Log.EXO_Log.Tipo.informacion)

            Dim oGeneralService As SAPbobsCOM.GeneralService = Nothing
            Dim oGeneralData As SAPbobsCOM.GeneralData = Nothing
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams = Nothing
            Dim genserdataint As SAPbobsCOM.GeneralServiceDataInterfaces = Nothing
            Dim oCompService As SAPbobsCOM.CompanyService = oCompany.GetCompanyService()

            oGeneralService = CType(oCompany.GetCompanyService().GetGeneralService("EXO_OGEN"), SAPbobsCOM.GeneralService)
            oGeneralParams = CType(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams), SAPbobsCOM.GeneralDataParams)
            oGeneralParams.SetProperty("Code", "EXO_KERNEL")
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            Dim oChild As SAPbobsCOM.GeneralData = Nothing
            Dim oChildren As SAPbobsCOM.GeneralDataCollection = oGeneralData.Child("EXO_OGEN1")
            oLog.escribeMensaje("###### CREANDO PARÁMETRO:  EXO_VIAPAGO ", EXO_Log.EXO_Log.Tipo.informacion)
            oChild = oChildren.Add()
            oChild.SetProperty("U_EXO_NOMV", "EXO_VIAPAGO")
            oChild.SetProperty("U_EXO_INFV", "CO-TA")
            oLog.escribeMensaje("###### CREANDO PARÁMETRO:  EXO_SERIEPEDWEB", EXO_Log.EXO_Log.Tipo.informacion)
            oChild = oChildren.Add()
            oChild.SetProperty("U_EXO_NOMV", "EXO_SERIEPEDWEB")
            oChild.SetProperty("U_EXO_INFV", "P_WEB")
            oGeneralService.Update(oGeneralData)
            oLog.escribeMensaje("###### PARÁMETROS CREADOS ", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("##################################################################", EXO_Log.EXO_Log.Tipo.informacion)
#End Region
        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
            End If

        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
            If oCompany.InTransaction = True Then
                oCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
            End If
        Finally


        End Try
    End Sub
#End Region
#Region "Alertas"
    Public Shared Sub EnviarAlerta(ByRef oLog As EXO_Log.EXO_Log, oCompany As SAPbobsCOM.Company, ByVal sDocNum As String, ByVal sDocEntry As String, ByVal sObject As String, ByVal sSubject As String, ByVal sTipo As String, ByVal sComentarios As String, ByVal Sfile As String, ByVal sDelAlmacen As String)
        Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing
        Dim oMessageService As SAPbobsCOM.MessagesService = Nothing
        Dim oMessage As SAPbobsCOM.Message = Nothing
        Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns = Nothing
        Dim pMessageDataColumnT As SAPbobsCOM.MessageDataColumn = Nothing
        Dim pMessageDataColumnD As SAPbobsCOM.MessageDataColumn = Nothing
        Dim pMessageDataColumnC As SAPbobsCOM.MessageDataColumn = Nothing
        Dim oLines As SAPbobsCOM.MessageDataLines = Nothing
        Dim oLine As SAPbobsCOM.MessageDataLine = Nothing
        Dim oRecipientCollection As SAPbobsCOM.RecipientCollection = Nothing
        Dim sSQL As String = ""
        Dim oXmlAux As XmlDocument = Nothing
        Dim oNodesAux As Xml.XmlNodeList = Nothing
        Dim oNodeAux As Xml.XmlNode = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "Select t1.""USER_CODE"" FROM OUSR t1 WHERE IFNULL(t1.""U_EXO_ALERTAWEB"", 'N') = 'Y' and ""Branch""='" & sDelAlmacen & "' "
            oRs.DoQuery(sSQL)

            oXmlAux = New XmlDocument
            oXmlAux.LoadXml(oRs.GetAsXML())
            oNodesAux = oXmlAux.SelectNodes("//row")

            If oRs.RecordCount > 0 Then
                oCmpSrv = oCompany.GetCompanyService()

                oMessageService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService), SAPbobsCOM.MessagesService)
                oMessage = CType(oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage), SAPbobsCOM.Message)

                'Añadimos los destinatarios de la alerta
                oRecipientCollection = oMessage.RecipientCollection

                For k As Integer = 0 To oNodesAux.Count - 1
                    oNodeAux = oNodesAux.Item(k)

                    oRecipientCollection.Add()
                    oRecipientCollection.Item(k).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                    oRecipientCollection.Item(k).UserCode = oNodeAux.SelectSingleNode("USER_CODE").InnerText
                Next

                pMessageDataColumns = oMessage.MessageDataColumns

                pMessageDataColumnT = pMessageDataColumns.Add()
                pMessageDataColumnT.ColumnName = "Tipo"

                pMessageDataColumnD = pMessageDataColumns.Add()
                pMessageDataColumnD.ColumnName = "Num. doc."
                pMessageDataColumnD.Link = SAPbobsCOM.BoYesNoEnum.tYES

                pMessageDataColumnC = pMessageDataColumns.Add()
                pMessageDataColumnC.ColumnName = "Concepto"

                oMessage.Subject = Left(sSubject, 254)

                oLines = pMessageDataColumnT.MessageDataLines
                oLine = oLines.Add()
                oLine.Value = sTipo

                If sDocEntry <> "" And sDocNum <> "" Then
                    oLines = pMessageDataColumnD.MessageDataLines
                    oLine = oLines.Add()
                    oLine.Value = sDocNum
                    oLine.Object = sObject
                    oLine.ObjectKey = sDocEntry
                End If

                'CONCEPTO
                oLines = pMessageDataColumnC.MessageDataLines
                oLine = oLines.Add()
                oLine.Value = Left(sComentarios, 254)

                oMessageService.SendMessage(oMessage)
                oLog.escribeMensaje("Alerta enviada...", EXO_Log.EXO_Log.Tipo.informacion)
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            oLog.escribeMensaje(sTipo & " " & Sfile & " - " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            oLog.escribeMensaje(sTipo & " " & Sfile & " - " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally

            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)
            oRs = Nothing

            oXmlAux = Nothing

            If oMessageService IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oMessageService)
            oMessageService = Nothing

            If oMessage IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oMessage)
            oMessage = Nothing

            If pMessageDataColumns IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumns)
            pMessageDataColumns = Nothing

            If pMessageDataColumnT IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumnT)
            pMessageDataColumnT = Nothing

            If pMessageDataColumnD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumnD)
            pMessageDataColumnD = Nothing

            If pMessageDataColumnC IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumnC)
            pMessageDataColumnC = Nothing

            If oLines IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLines)
            oLines = Nothing

            If oLine IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLine)
            oLine = Nothing

            If oRecipientCollection IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRecipientCollection)
            oRecipientCollection = Nothing

            If oCmpSrv IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrv)
            oCmpSrv = Nothing
        End Try
    End Sub
#End Region
End Class
