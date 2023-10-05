Imports SAPbouiCOM
Imports System.Xml
Imports EXO_UIAPI.EXO_UIAPI

Public Class INICIO
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaDatos()
            InsertarReport()
        End If
        cargamenu()
    End Sub
    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""
        Dim sSQL As String = ""
        If objGlobal.refDi.comunes.esAdministrador Then
            ParametrizacionGeneral()

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OCRD.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OCRD", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults


            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_ORDR.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_ORDR", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OWTQ.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OWTQ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_ROTURA.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_ROTURA", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_SITUACION.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_SITUACION", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_A.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_A", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_PEDIDOS_VENTA.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_PEDIDOS_VENTA", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_SOL_TRASLADO.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_SOL_TRASLADO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_SOL_DEVOLUCION.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_SOL_DEVOLUCION", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "VEXO_PICKING.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: VEXO_PICKING", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "VEXO_TRASLADOS.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: VEXO_TRASLADOS", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "VEXO_PARRILLA_ESTADO_SALCOMP.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: VEXO_PARRILLA_ESTADO_SALCOMP", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        End If
    End Sub
    Private Sub ParametrizacionGeneral()
        If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("EXO_ETPARRILLA") Then
            objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("EXO_ETPARRILLA", "RCRI0015")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable: EXO_ETPARRILLA", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objGlobal.SBOApp.StatusBar.SetText("Ya existe Variable: EXO_ETPARRILLA", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
    End Sub

    Private Sub InsertarReport()
#Region "Variables"
        Dim sArchivo As String = ""
        Dim sReport As String = ""
#End Region
        Try
            sArchivo = objGlobal.path & "\05.Rpt\PARRILLADOC\"
#Region "ENTREGAS"
            sReport = "ENTREGAS.rpt"
            'Si no existe lo importamos
            If IO.File.Exists(sArchivo & sReport) = False Then
                If IO.Directory.Exists(sArchivo) = False Then
                    IO.Directory.CreateDirectory(sArchivo)
                End If
                EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), sReport, sArchivo & sReport)
            End If
#End Region
#Region "DEV. PROVEEDOR"
            sReport = "DEVPROVEEDOR.rpt"
            'Si no existe lo importamos
            If IO.File.Exists(sArchivo & sReport) = False Then
                If IO.Directory.Exists(sArchivo) = False Then
                    IO.Directory.CreateDirectory(sArchivo)
                End If
                EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), sReport, sArchivo & sReport)
            End If
#End Region

#Region "SOL. TRASLADO"
            sReport = "SOLTRASLADO.rpt"
            'Si no existe lo importamos
            If IO.File.Exists(sArchivo & sReport) = False Then
                If IO.Directory.Exists(sArchivo) = False Then
                    IO.Directory.CreateDirectory(sArchivo)
                End If
                EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), sReport, sArchivo & sReport)
            End If
#End Region

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
        'objGlobal.SBOApp.StatusBar.SetText(res, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
    Public Overrides Function filtros() As Global.SAPbouiCOM.EventFilters
        Dim fXML As String = ""
        Try
            fXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml")
            Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
            filtro.LoadFromXML(fXML)
            Return filtro
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        Finally

        End Try
    End Function

    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim Clase As Object = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.MenuUID
                    Case ""
                End Select
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnPARC"
                        Clase = New EXO_PARRILLA(objGlobal)
                        Return CType(Clase, EXO_PARRILLA).SBOApp_MenuEvent(infoEvento)
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            Clase = Nothing
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Dim res As Boolean = True
        Dim Clase As Object = Nothing

        Try
            Select Case infoEvento.FormTypeEx
                Case "EXO_PARRILLA"
                    Clase = New EXO_PARRILLA(objGlobal)
                    Return CType(Clase, EXO_PARRILLA).SBOApp_ItemEvent(infoEvento)
                Case "EXO_RSTOCK"
                    Clase = New EXO_RSTOCK(objGlobal)
                    Return CType(Clase, EXO_RSTOCK).SBOApp_ItemEvent(infoEvento)
                Case "1250000940"
                    Clase = New EXO_OWTQ(objGlobal)
                    Return CType(Clase, EXO_OWTQ).SBOApp_ItemEvent(infoEvento)
                Case "134"
                    Clase = New EXO_OCRD(objGlobal)
                    Return CType(Clase, EXO_OCRD).SBOApp_ItemEvent(infoEvento)
                Case "139"
                    Clase = New EXO_ORDR(objGlobal)
                    Return CType(Clase, EXO_ORDR).SBOApp_ItemEvent(infoEvento)
                Case "EXO_BTOS"
                    Clase = New EXO_BTOS(objGlobal)
                    Return CType(Clase, EXO_BTOS).SBOApp_ItemEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Finally
            Clase = Nothing
        End Try
    End Function
    Public Overrides Function SBOApp_FormDataEvent(infoEvento As BusinessObjectInfo) As Boolean
        Dim Res As Boolean = True
        Dim Clase As Object = Nothing
        Try
            Select Case infoEvento.FormTypeEx
                Case "134"
                    Clase = New EXO_OCRD(objGlobal)
                    Return CType(Clase, EXO_OCRD).SBOApp_FormDataEvent(infoEvento)
                Case "139"
                    Clase = New EXO_ORDR(objGlobal)
                    Return CType(Clase, EXO_ORDR).SBOApp_FormDataEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Finally
            Clase = Nothing
        End Try

    End Function
End Class
