Imports SAPbouiCOM
Imports System.Xml
Imports EXO_UIAPI.EXO_UIAPI

Public Class INICIO
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            CargaProcedures()
            CargaViews()
            CargaFunctions()

        End If
    End Sub
    Private Sub CargaProcedures()
        Dim sXML As String = ""
        Dim res As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        If objGlobal.refDi.comunes.esAdministrador Then
            objGlobal.SBOApp.StatusBar.SetText("Carga de Procedures...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

#Region "EXO_GET_INTEGRACION_SYNC"
            'objGlobal.SBOApp.StatusBar.SetText("Creando Procedure: EXO_GET_INTEGRACION_SYNC", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_GET_INTEGRACION_SYNC.sql")
            'sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            'If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
            '    objGlobal.SBOApp.StatusBar.SetText("Procedure Creado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'Else
            '    objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '    sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_GET_INTEGRACION_SYNC.sql")
            'sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            '    If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
            '        objGlobal.SBOApp.StatusBar.SetText("Procedure Actualizado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '    Else
            '        objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    End If
            'End If
#End Region

#Region "EXO_GP_PROPONGO_LOTE"
            objGlobal.SBOApp.StatusBar.SetText("Creando Procedure: EXO_GP_PROPONGO_LOTE", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_GP_PROPONGO_LOTE.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Procedure Creado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_GP_PROPONGO_LOTE.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("Procedure Actualizado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    Try
                        oRs.DoQuery(sSQL)
                    Catch ex As Exception
                        objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente. " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
            End If
#End Region

#Region "EXO_GP_TRABAJO_LISTA_PICKING"
            objGlobal.SBOApp.StatusBar.SetText("Creando Procedure: EXO_GP_TRABAJO_LISTA_PICKING", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_GP_TRABAJO_LISTA_PICKING.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Procedure Creado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_GP_TRABAJO_LISTA_PICKING.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("Procedure Actualizado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    Try
                        oRs.DoQuery(sSQL)
                    Catch ex As Exception
                        objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente. " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
            End If
#End Region

#Region "EXO_GP_TRABAJO_LISTA_TRASLADO"
            objGlobal.SBOApp.StatusBar.SetText("Creando Procedure: EXO_GP_TRABAJO_LISTA_TRASLADO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_GP_TRABAJO_LISTA_TRASLADO.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Procedure Creado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_GP_TRABAJO_LISTA_TRASLADO.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("Procedure Actualizado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    Try
                        oRs.DoQuery(sSQL)
                    Catch ex As Exception
                        objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente. " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
            End If
#End Region

#Region "EXO_RECAL_CONTROL_STOCK"
            objGlobal.SBOApp.StatusBar.SetText("Creando Procedure: EXO_RECAL_CONTROL_STOCK", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_RECAL_CONTROL_STOCK.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Procedure Creado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_RECAL_CONTROL_STOCK.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("Procedure Actualizado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Try
                        oRs.DoQuery(sSQL)
                    Catch ex As Exception
                        Throw ex
                    End Try
                End If
            End If
#End Region


            objGlobal.SBOApp.StatusBar.SetText("Fin de la Carga de Procedures.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
    End Sub

    Private Sub CargaViews()
        Dim sXML As String = ""
        Dim res As String = ""
        Dim sSQL As String = ""
        If objGlobal.refDi.comunes.esAdministrador Then
            objGlobal.SBOApp.StatusBar.SetText("Carga de Views...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

#Region "EXO_A"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_A", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_A.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_A.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_A_PARRILLA"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_A_PARRILLA", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_A_PARRILLA.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_A_PARRILLA.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_Clasificacion_Artículos"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_Clasificacion_Artículos", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_Clasificacion_Artículos.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_Clasificacion_Artículos.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_DetalleBultosEnvioTransporte"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_DetalleBultosEnvioTransporte", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_DetalleBultosEnvioTransporte.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_DetalleBultosEnvioTransporte.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_DetalleBultosExpediciones"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_DetalleBultosExpediciones", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_DetalleBultosExpediciones.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_DetalleBultosExpediciones.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_EtiquetaBultos"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_EtiquetaBultos", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_EtiquetaBultos.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_EtiquetaBultos.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MANIFIESTO_TTE"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MANIFIESTO_TTE", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MANIFIESTO_TTE.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MANIFIESTO_TTE.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MRP_Clasificacion_Artículos"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MRP_Clasificacion_Artículos", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MRP_Clasificacion_Artículos.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MRP_Clasificacion_Artículos.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MRP_Compras8Q"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MRP_Compras8Q", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MRP_Compras8Q.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MRP_Compras8Q.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MRP_ComprasSemestre"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MRP_ComprasSemestre", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MRP_ComprasSemestre.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MRP_ComprasSemestre.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MRP_Pdte"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MRP_Pdte", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MRP_Pdte.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MRP_Pdte.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MRP_Pdte_Desglose"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MRP_Pdte_Desglose", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MRP_Pdte_Desglose.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MRP_Pdte_Desglose.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MRP_Proveedores"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MRP_Proveedores", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MRP_Proveedores.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MRP_Proveedores.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MRP_StocksActuales"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MRP_StocksActuales", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MRP_StocksActuales.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MRP_StocksActuales.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MRP_Ventas24Q"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MRP_Ventas24Q", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MRP_Ventas24Q.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MRP_Ventas24Q.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MRP_Ventas8Q"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MRP_Ventas8Q", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MRP_Ventas8Q.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MRP_Ventas8Q.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MRP_VentasA_8Q"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MRP_VentasA_8Q", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MRP_VentasA_8Q.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MRP_VentasA_8Q.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_MRP_Ventas_MED_24Q"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_MRP_Ventas_MED_24Q", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_MRP_Ventas_MED_24Q.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_MRP_Ventas_MED_24Q.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_PedidoCompraEnvioTransporte"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_PedidoCompraEnvioTransporte", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_PedidoCompraEnvioTransporte.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_PedidoCompraEnvioTransporte.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_PEDIDOS_VENTA"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_PEDIDOS_VENTA", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_PEDIDOS_VENTA.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_PEDIDOS_VENTA.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_PesoBultos_Agencia"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_PesoBultos_Agencia", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_PesoBultos_Agencia.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_PesoBultos_Agencia.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_ResumenBultosExpedicion"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_ResumenBultosExpedicion", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_ResumenBultosExpedicion.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_ResumenBultosExpedicion.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_ROTURA"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_ROTURA", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_ROTURA.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_ROTURA.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_ROTURA_DETAILS"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_ROTURA_DETAILS", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_ROTURA_DETAILS.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_ROTURA_DETAILS.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_SITUACION"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_SITUACION", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_SITUACION.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_SITUACION.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_SOL_DEVOLUCION"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_SOL_DEVOLUCION", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_SOL_DEVOLUCION.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_SOL_DEVOLUCION.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_SOL_TRASLADO"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_SOL_TRASLADO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_SOL_TRASLADO.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_SOL_TRASLADO.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_UbicacionDestinoEntradaCompra"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_UbicacionDestinoEntradaCompra", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_UbicacionDestinoEntradaCompra.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_UbicacionDestinoEntradaCompra.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_UbicacionDestinoEntradaCompra_2"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_UbicacionDestinoEntradaCompra_2", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_UbicacionDestinoEntradaCompra_2.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_UbicacionDestinoEntradaCompra_2.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_UbicaciónDestinoEntradaCompra"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_UbicaciónDestinoEntradaCompra", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_UbicaciónDestinoEntradaCompra.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_UbicaciónDestinoEntradaCompra.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "EXO_UbicaciónDestinoEntradaCompra_2"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: EXO_UbicaciónDestinoEntradaCompra_2", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_UbicaciónDestinoEntradaCompra_2.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_UbicaciónDestinoEntradaCompra_2.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "VEXO_PARRILLA_ESTADO_SALCOMP"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: VEXO_PARRILLA_ESTADO_SALCOMP", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_VEXO_PARRILLA_ESTADO_SALCOMP.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_VEXO_PARRILLA_ESTADO_SALCOMP.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "VEXO_PICKING"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: VEXO_PICKING", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_VEXO_PICKING.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_VEXO_PICKING.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "VEXO_TRASLADOS"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: VEXO_TRASLADOS", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_VEXO_TRASLADOS.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_VEXO_TRASLADOS.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "VEXO_USUARIO_ALMACENES"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: VEXO_USUARIO_ALMACENES", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_VEXO_USUARIO_ALMACENES.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_VEXO_USUARIO_ALMACENES.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region

#Region "VTSP"
            objGlobal.SBOApp.StatusBar.SetText("Creando View: VTSP", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_VTSP.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("View Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_VTSP.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("View Actualizada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
#End Region
            objGlobal.SBOApp.StatusBar.SetText("Fin de la Carga de Views.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
    End Sub

    Private Sub CargaFunctions()
        Dim sXML As String = ""
        Dim res As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        If objGlobal.refDi.comunes.esAdministrador Then
            objGlobal.SBOApp.StatusBar.SetText("Carga de Functions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)



#Region "EXO_GP_PROPONGO_UBICACION"
            objGlobal.SBOApp.StatusBar.SetText("Creando Functions: EXO_GP_PROPONGO_UBICACION", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_EXO_GP_PROPONGO_UBICACION.sql")
            sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Functions Creado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear, se intenta actualizar...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "ALTER_EXO_GP_PROPONGO_UBICACION.sql")
                sSQL = sSQL.Replace("""BBDD""", """" & objGlobal.compañia.CompanyDB & """")
                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    objGlobal.SBOApp.StatusBar.SetText("Function Actualizado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    Try
                        oRs.DoQuery(sSQL)
                    Catch ex As Exception
                        objGlobal.SBOApp.StatusBar.SetText("No se ha podido Actualizar, Tendrá que revisarlo manualmente. " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
            End If
#End Region

            objGlobal.SBOApp.StatusBar.SetText("Fin de la Carga de Functions.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
    End Sub
    Public Overrides Function filtros() As Global.SAPbouiCOM.EventFilters
        Return Nothing
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
End Class
