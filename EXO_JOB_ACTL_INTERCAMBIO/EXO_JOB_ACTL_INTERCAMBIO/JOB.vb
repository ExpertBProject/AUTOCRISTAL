Imports Sap.Data.Hana
Module JOB
    Public Sub Main()
#Region "Variables"
        Dim iCountExeJOB As Integer = 0
        Dim oLog As EXO_Log.EXO_Log = Nothing
        Dim sError As String
        Dim sPath As String = ""
        Dim sSQL As String = ""
        Dim oDBSAP As HanaConnection = Nothing : Dim odtDatos As System.Data.DataTable = Nothing
        Dim sBBDD As String = "" : Dim sUsuario As String = "" : Dim sPassword As String = ""
        Dim oCompany As SAPbobsCOM.Company = Nothing

#End Region
        Try
            ''Comprobamos si el JOB está en ejecución y en caso afirmativo no lanzamos ningún proceso del JOB.
            For Each oProcess As Process In Process.GetProcesses()
                If oProcess.ProcessName.ToString = "EXO_JOB_ACTL_INTERCAMBIO" Then
                    iCountExeJOB += 1
                End If
            Next
            If iCountExeJOB = 1 Then

                sPath = My.Application.Info.DirectoryPath.ToString

                If Not System.IO.Directory.Exists(sPath & "\Logs") Then
                    System.IO.Directory.CreateDirectory(sPath & "\Logs")
                End If
                oLog = New EXO_Log.EXO_Log(sPath & "\Logs\LOG_", 10, EXO_Log.EXO_Log.Nivel.todos, 4, "", EXO_Log.EXO_Log.GestionFichero.dia)
                oLog.escribeMensaje("", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#####           INICIO INTERCOMPANY ARTICULOS         #####", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)

                'Conexiones.Connect_SQLHANA(oDBSAP, "HANA", oLog)
                'Conexiones.Connect_Company(oCompany, "DI", Conexiones.sBBDD, Conexiones.sUser, Conexiones.sPwd, oLog)

                ''Articulos - Clases de expedicion
                Procesos.OSHP()

                ''Articulos - Fabricantes
                Procesos.OMRC()

                ''Articulos - atributos propiedades del articulo
                Procesos.OITG()

                ''Articulos - grupos de articulos /familias
                Procesos.OITB()

                ''Articulos
                Procesos.OITM()



            End If

        Catch ex As Exception
            If ex.InnerException IsNot Nothing AndAlso ex.InnerException.Message <> "" Then
                sError = ex.InnerException.Message
            Else
                sError = ex.Message
            End If
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("#####                 FIN PROCESO                 #####", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("", EXO_Log.EXO_Log.Tipo.informacion)
            Conexiones.Disconnect_SQLHANA(oDBSAP)
        End Try
    End Sub
End Module
