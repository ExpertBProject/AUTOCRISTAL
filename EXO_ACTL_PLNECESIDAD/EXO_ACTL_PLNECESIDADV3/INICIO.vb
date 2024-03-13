Imports SAPbouiCOM
Imports System.Xml
Imports EXO_UIAPI.EXO_UIAPI
Public Class INICIO
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "Variables Globales"
    Public Shared _WidthART As Integer = 215
    Public Shared _HeightART As Integer = 160

    Public Shared _WidthALM As Integer = 208
    Public Shared _HeightALM As Integer = 160

    Public Shared _WidthCLAS As Integer = 95
    Public Shared _HeightCLAS As Integer = 160

    Public Shared _WidthGRU As Integer = 208
    Public Shared _HeightGRU As Integer = 160

    Public Shared _WidthCOMP As Integer = 170
    Public Shared _HeightCOMP As Integer = 160

    Public Shared _WidthVENT As Integer = 170
    Public Shared _HeightVENT As Integer = 160

    Public Shared _dtDatos As New System.Data.DataTable

    Public Shared _iRowGrid As Integer = -1
#End Region
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaDatos()
        End If
        cargamenu()
    End Sub
    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""
        Dim sSQL As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OCRD.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OCRD", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OSCN.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OSCN", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MRP_Clasif_Art.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_MRP_Clasificacion_Artículos", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MRP_Compras8Q.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_MRP_Compras8Q", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MRP_ComprasSemestre.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_MRP_ComprasSemestre", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MRP_Pdte.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_MRP_Pdte", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MRP_StocksActuales.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_MRP_StocksActuales", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MRP_Ventas24Q.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_MRP_Ventas24Q", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MRP_Ventas8Q.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_MRP_Ventas8Q", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MRP_VentasA_8Q.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_MRP_VentasA_8Q", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MRP_Proveedores.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_MRP_Proveedores", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MRP_Ventas_MED_24Q.sql")
            objGlobal.SBOApp.StatusBar.SetText("Creando Vista: EXO_MRP_Ventas_MED_24Q", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Vista Creada...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido crear la vista...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_PLNHCO.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_PLNHCO", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If
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
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        Finally

        End Try
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Dim res As Boolean = True
        Dim Clase As Object = Nothing

        Try
            Select Case infoEvento.FormTypeEx
                Case "EXO_PNEC"
                    Clase = New EXO_PNEC(objGlobal)
                    Return CType(Clase, EXO_PNEC).SBOApp_ItemEvent(infoEvento)
                Case "EXO_IMPART"
                    Clase = New EXO_IMPART(objGlobal)
                    Return CType(Clase, EXO_IMPART).SBOApp_ItemEvent(infoEvento)
                Case "EXO_PLNSAVE"
                    Clase = New EXO_PLNSAVE(objGlobal)
                    Return CType(Clase, EXO_PLNSAVE).SBOApp_ItemEvent(infoEvento)
                Case "EXO_PLNREC"
                    Clase = New EXO_PLNREC(objGlobal)
                    Return CType(Clase, EXO_PLNREC).SBOApp_ItemEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Finally
            Clase = Nothing
        End Try
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
                    Case "EXO-MnPNEC"
                        Clase = New EXO_PNEC(objGlobal)
                        Return CType(Clase, EXO_PNEC).SBOApp_MenuEvent(infoEvento)

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
End Class

