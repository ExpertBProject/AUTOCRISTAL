Imports Sap.Data.Hana
Module JOB
    Public Sub Main()
#Region "Variables"
        Dim iCountExeJOB As Integer = 0
        Dim oLog As EXO_Log.EXO_Log = Nothing
        Dim sError As String
        Dim sPath As String = ""
        Dim oDBSAP As HanaConnection = Nothing
        Dim oDBWEB As HanaConnection = Nothing
        Dim oCompany As SAPbobsCOM.Company = Nothing
#End Region
        Try
            'Comprobamos si el JOB está en ejecución y en caso afirmativo no lanzamos ningún proceso del JOB.
            For Each oProcess As Process In Process.GetProcesses()
                If oProcess.ProcessName.ToString = "EXO_JOB_ACTL_INTWEB" Then
                    iCountExeJOB += 1
                End If
            Next
            If iCountExeJOB = 1 Or iCountExeJOB = 0 Then
                sPath = My.Application.Info.DirectoryPath.ToString

                If Not System.IO.Directory.Exists(sPath & "\Logs") Then
                    System.IO.Directory.CreateDirectory(sPath & "\Logs")
                End If
                oLog = New EXO_Log.EXO_Log(sPath & "\Logs\LOG_", 10, EXO_Log.EXO_Log.Nivel.todos, 4, "", EXO_Log.EXO_Log.GestionFichero.dia)
                oLog.escribeMensaje("", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#####          INICIO INTEGRACIÓN WEB             #####", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)

                Conexiones.Connect_SQLHANA(oDBWEB, "HANAWEB", oLog) ' Primero la WEB para que los parámetros publicos se queden con la Válida de SAP
                Conexiones.Connect_SQLHANA(oDBSAP, "HANA", oLog)

                Conexiones.Connect_Company(oCompany, "DI", Conexiones.sBBDD, Conexiones.sUser, Conexiones.sPwd, oLog)
                If Conexiones.Datos_Confi("ACTUALIZAR", "CAMPOS") = "Y" Then
                    oLog.escribeMensaje(" ", EXO_Log.EXO_Log.Tipo.informacion)
                    oLog.escribeMensaje("Procedimiento. ACTUALIZAR CAMPO", EXO_Log.EXO_Log.Tipo.informacion)
                    oLog.escribeMensaje("##################################################################", EXO_Log.EXO_Log.Tipo.informacion)
                    Procesos.Actualizar_Campos(oLog, oDBSAP, oCompany, False)
                Else
                    Procesos.LecturaTabla(oDBSAP, oDBWEB, oCompany, oLog)
                End If
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#####                 FIN PROCESO                 #####", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("", EXO_Log.EXO_Log.Tipo.informacion)

            End If

        Catch ex As Exception
            If ex.InnerException IsNot Nothing AndAlso ex.InnerException.Message <> "" Then
                sError = ex.InnerException.Message
            Else
                sError = ex.Message
            End If
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            Conexiones.Disconnect_SQLHANA(oDBSAP)
            Conexiones.Disconnect_Company(oCompany)
            Conexiones.Disconnect_SQLHANA(oDBWEB)
        End Try
    End Sub
End Module

