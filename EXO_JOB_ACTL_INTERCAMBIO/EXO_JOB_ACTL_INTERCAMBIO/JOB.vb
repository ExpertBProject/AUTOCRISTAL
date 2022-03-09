Imports Sap.Data.Hana
Module JOB
    Public Sub Main()
#Region "Variables"
        Dim iCountExeJOB As Integer = 0
        Dim oLog As EXO_Log.EXO_Log = Nothing
        Dim sError As String
        Dim sPath As String = ""
        Dim sSQL As String = ""


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
                oLog.escribeMensaje("#####           INICIO INTERCOMPANY               #####", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)


                ''Articulos - Clases de expedicion
                oLog.escribeMensaje("Antes de tratar clases de expedición", EXO_Log.EXO_Log.Tipo.informacion)
                Procesos.OSHP()

                ''Articulos - Fabricantes
                oLog.escribeMensaje("Antes de tratar fabricantes", EXO_Log.EXO_Log.Tipo.informacion)
                Procesos.OMRC()

                ''Articulos - atributos propiedades del articulo
                oLog.escribeMensaje("Antes de tratar propiedades artículos", EXO_Log.EXO_Log.Tipo.informacion)
                Procesos.OITG()

                ''Articulos - grupos de articulos /familias
                oLog.escribeMensaje("Antes de tratar grupos de articulos", EXO_Log.EXO_Log.Tipo.informacion)
                Procesos.OITB()

                ''Articulos
                oLog.escribeMensaje("Antes de tratar artículos", EXO_Log.EXO_Log.Tipo.informacion)
                Procesos.OITM()

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

            'Conexiones.Disconnect_SQLHANA(oDBSAP)
        End Try
    End Sub
End Module
