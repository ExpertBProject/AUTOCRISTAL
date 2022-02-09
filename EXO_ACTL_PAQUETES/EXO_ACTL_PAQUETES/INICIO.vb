Imports SAPbouiCOM
Imports System.Xml
Imports EXO_UIAPI.EXO_UIAPI
Public Class INICIO
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "Variables Globales"
    Public Shared _sOPKG As String = ""
    Public Shared _sTipo As String = ""
    Public Shared _sDes As String = ""
    Public Shared _sLineaSel As String = ""
#End Region
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaDatos()
            ParametrizacionGeneral()
            InsertarReport()
        End If

        cargamenu()
    End Sub

    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_PAQ.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_PAQ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OPKG.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OPKG", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_ABULTOS.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_ABULTOS", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

        End If
    End Sub
    Private Sub ParametrizacionGeneral()

        If Not objGlobal.refDi.OGEN.existeVariable("SERVIDOR_HANA") Then
            objGlobal.refDi.OGEN.fijarValorVariable("SERVIDOR_HANA", "10.10.1.13:30015")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""SERVIDOR_HANA"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If

    End Sub
    Private Sub InsertarReport()
#Region "Variables"
        Dim sArchivo As String = ""
        Dim sReport As String = ""
#End Region
        Try
            sArchivo = objGlobal.path & "\05.Rpt\ETIQUETAS\"
#Region "Report Etiqueta Agrupaciones de Bultos"
            sReport = "Et_ABultos.rpt"
            'Si no existe lo importamos
            If IO.File.Exists(sArchivo & sReport) = False Then
                If IO.Directory.Exists(sArchivo) = False Then
                    IO.Directory.CreateDirectory(sArchivo)
                End If
                EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), sReport, sArchivo & sReport)
            End If
#End Region
#Region "Report Etiqueta Paquete"
            sReport = "Et_Paquete.rpt"
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
                Case "60016"
                    Clase = New EXO_60016(objGlobal)
                    Return CType(Clase, EXO_60016).SBOApp_ItemEvent(infoEvento)
                Case "UDO_FT_EXO_PAQ"
                    Clase = New EXO_PAQ(objGlobal)
                    Return CType(Clase, EXO_PAQ).SBOApp_ItemEvent(infoEvento)
                Case "UDO_FT_EXO_ABULTOS"
                    Clase = New EXO_ABULTOS(objGlobal)
                    Return CType(Clase, EXO_ABULTOS).SBOApp_ItemEvent(infoEvento)

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
                'Case "UDO_FT_EXO_SUPLEMENTO"
                '    Clase = New EXO_SUPLEMENTO(objGlobal)
                '    Return CType(Clase, EXO_SUPLEMENTO).SBOApp_FormDataEvent(infoEvento)

            End Select

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

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
                    Case "EXO-MnABUL"
                        Clase = New EXO_ABULTOS(objGlobal)
                        Return CType(Clase, EXO_ABULTOS).SBOApp_MenuEvent(infoEvento)

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
