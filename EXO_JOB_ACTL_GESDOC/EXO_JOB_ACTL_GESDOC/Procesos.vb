Imports Sap.Data.Hana
Imports System.IO
Imports SAPbobsCOM

Public Class Procesos
#Region "Especifico"
    Public Shared Sub LecturaFicheros(ByRef db As HanaConnection, ByRef oCompany As SAPbobsCOM.Company, ByRef oLog As EXO_Log.EXO_Log)
        Dim refDI As EXO_DIAPI.EXO_DIAPI = Nothing
        Dim tipoServidor As SAPbobsCOM.BoDataServerTypes = SAPbobsCOM.BoDataServerTypes.dst_HANADB
        Dim tipocliente As EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico
        Dim sPass As String = Conexiones.Datos_Confi("DI", "Password")
        Dim sError As String = ""
        Dim sRutaLectura As String = ""
        Dim sObjType As String = ""
        Dim sDocEntry As String = ""

        Try

            Try
                refDI = New EXO_DIAPI.EXO_DIAPI(oCompany, oLog)
            Catch ex As Exception
                refDI = New EXO_DIAPI.EXO_DIAPI(tipoServidor, oCompany.Server.ToString, oCompany.LicenseServer.ToString, oCompany.CompanyDB.ToString, oCompany.UserName.ToString, sPass, tipocliente)
            End Try

            sRutaLectura = refDI.OGEN.valorVariable("EXO_RUTAGESTORDOC")
            If sRutaLectura <> "" Then
                Dim sDirectorio As DirectoryInfo = New DirectoryInfo(sRutaLectura)
                For Each sArchivo As FileInfo In sDirectorio.GetFiles
                    Dim sExtension As String = sArchivo.Extension
                    If sExtension.ToLower = ".pdf" Then
                        'adjuntar
                        If sArchivo.Name.Length = 15 Then
                            sObjType = Mid(sArchivo.Name, 1, 2)
                            sDocEntry = Mid(sArchivo.Name, 4, 8)
                            Procesos.Attach_SAP(oCompany.CompanyDB, oCompany, sDocEntry, sObjType, sRutaLectura & "\" & sArchivo.ToString, db, oLog)
                            'borrar

                        End If
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
            refDI = Nothing
        End Try
    End Sub


#End Region

#Region "Metodos SAP"
    Public Shared Sub ParametrizacionGeneral(ByRef oCompany As SAPbobsCOM.Company, ByRef oLog As EXO_Log.EXO_Log)
        Dim refDI As EXO_DIAPI.EXO_DIAPI = Nothing
        Dim tipoServidor As SAPbobsCOM.BoDataServerTypes = SAPbobsCOM.BoDataServerTypes.dst_HANADB
        Dim tipocliente As EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico
        Dim sPass As String = Conexiones.Datos_Confi("DI", "Password")
        Dim sError As String = ""
        Try

            Try
                refDI = New EXO_DIAPI.EXO_DIAPI(oCompany, oLog)
            Catch ex As Exception
                refDI = New EXO_DIAPI.EXO_DIAPI(tipoServidor, oCompany.Server.ToString, oCompany.LicenseServer.ToString, oCompany.CompanyDB.ToString, oCompany.UserName.ToString, sPass, tipocliente)
            End Try

            If Not refDI.OGEN.existeVariable("EXO_RUTAGESTORDOC") Then
                refDI.OGEN.fijarValorVariable("EXO_RUTAGESTORDOC", "")
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            refDI = Nothing
        End Try
    End Sub

    Public Shared Function Attach_SAP(BaseDatos As String, oCompany As SAPbobsCOM.Company, DocEntry As String, objType As String, Fichero As String, oDBSAP As HanaConnection, log As EXO_Log.EXO_Log) As String


        Dim jRes As String = ""

        Dim res As String = ""
        Dim bPrimero As Boolean = True

        Dim oDtCab As System.Data.DataTable = New System.Data.DataTable

        Dim oDocuments As SAPbobsCOM.Documents = Nothing
        Dim oAtt As SAPbobsCOM.Attachments2 = Nothing
        Dim sTabla As String = ""
        Dim sSql As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Try

            oAtt = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2), SAPbobsCOM.Attachments2)


            Dim Sql As String = "SELECT ""AttachPath"" from  """ & Conexiones.sBBDD & """.""OADP"""
            Dim Attachpath As String = ""

            Conexiones.FillDtDB(oDBSAP, oDtCab, Sql)

            If oDtCab.Rows.Count > 0 Then
                For iCab As Integer = 0 To oDtCab.Rows.Count - 1
                    Attachpath = oDtCab.Rows.Item(iCab).Item("attachpath").ToString
                Next
            End If

            If Attachpath = "" Then
                'RUTA NO EXISTE
                log.escribeMensaje("Defina la ruta de anexos en las parametrizaciones generales para poder anexar documentos", EXO_Log.EXO_Log.Tipo.error)
            Else

                Dim NomFichero As String = ""
                Dim str As String = ""

                NomFichero = Fichero
                Select Case objType
                    Case "18"
                        sTabla = "OPCH"
                    Case "19"
                        sTabla = "ORPC"
                End Select
                'comprobamos si ya tenemos algun attach en el docentry
                Sql = "SELECT COALESCE(""AtcEntry"",'0') ""AtcEntry"" FROM  """ & Conexiones.sBBDD & """.""" & sTabla & """ WHERE ""DocEntry""=" & Convert.ToInt16(DocEntry) & ""
                Dim TieneAtcEntry As String = ""

                Conexiones.FillDtDB(oDBSAP, oDtCab, Sql)

                If oDtCab.Rows.Count > 0 Then
                    For iCab As Integer = 0 To oDtCab.Rows.Count - 1
                        TieneAtcEntry = oDtCab.Rows.Item(iCab).Item("AtcEntry").ToString
                    Next
                End If

                If TieneAtcEntry <> "" And TieneAtcEntry <> "0" Then

                    oAtt.GetByKey(TieneAtcEntry)
                    oAtt.Lines.Add()
                    oAtt.Lines.SourcePath = System.IO.Path.GetDirectoryName(NomFichero)
                    oAtt.Lines.FileName = System.IO.Path.GetFileNameWithoutExtension(NomFichero)
                    oAtt.Lines.FileExtension = System.IO.Path.GetExtension(NomFichero).Substring(1)
                    oAtt.Lines.Override = BoYesNoEnum.tYES
                    If oAtt.Update() = 0 Then
                        jRes = "OK"
                    Else
                        jRes = "Error: " + oCompany.GetLastErrorDescription()
                    End If

                Else
                    'hacerlo por sql

                    Select Case objType
                        Case "18"
                            oDocuments = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices), SAPbobsCOM.Documents)
                        Case "19"
                            oDocuments = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes), SAPbobsCOM.Documents)
                    End Select


                    oAtt.Lines.SourcePath = System.IO.Path.GetDirectoryName(NomFichero)
                    oAtt.Lines.FileName = System.IO.Path.GetFileNameWithoutExtension(NomFichero)
                    oAtt.Lines.FileExtension = System.IO.Path.GetExtension(NomFichero).Substring(1)
                    oAtt.Lines.Override = BoYesNoEnum.tYES

                    Dim AttEntry As Integer
                    If oAtt.Add() <> 0 Then
                        jRes = "Error: " + oCompany.GetLastErrorDescription
                        log.escribeMensaje("error pdf 1" + oCompany.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                    Else
                        AttEntry = CInt(oCompany.GetNewObjectKey())
                        oDocuments.GetByKey(DocEntry)
                        oDocuments.AttachmentEntry = AttEntry
                        If oDocuments.Update() = 0 Then
                            log.escribeMensaje("OK ATTACH Adjuntado Documento ObjType " & objType & " DocEntry: " & DocEntry & "", EXO_Log.EXO_Log.Tipo.error)
                            jRes = "OK"
                        Else
                            log.escribeMensaje("ERROR ATTACH " + oCompany.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                            jRes = "Error: " + oCompany.GetLastErrorDescription
                        End If

                        '''update 
                        'oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                        'sSql = "UPDATE " & sTabla & " SET AtcEntry =" & AttEntry & " WHERE DocEntry=" & DocEntry & ""
                        'oRs.DoQuery(sSql)
                    End If
                End If

                'borramos el fichero
                If jRes = "OK" Then
                    My.Computer.FileSystem.DeleteFile(NomFichero)
                End If
            End If

        Catch ex As Exception
            log.escribeMensaje("attach: " + ex.Message, EXO_Log.EXO_Log.Tipo.error)
            jRes = "Error: " + ex.Message

        Finally

            If oDocuments IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDocuments)
                oDocuments = Nothing
            End If

            If oAtt IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oAtt)
                oAtt = Nothing
            End If
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try

        Return res

    End Function
#End Region
End Class
