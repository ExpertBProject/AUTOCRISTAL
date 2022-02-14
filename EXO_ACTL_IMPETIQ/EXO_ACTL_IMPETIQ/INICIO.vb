Imports SAPbouiCOM
Imports System.Xml
Imports EXO_UIAPI.EXO_UIAPI
Public Class INICIO
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaDatos()
            Cambiar_Nombre_Propiedades()
            ParametrizacionGeneral()
            InsertarReport()
        End If
        cargamenu()
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
        'objGlobal.SBOApp.StatusBar.SetText(res, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UT_EXO_ETIQUETA.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UT_EXO_ETIQUETA", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

        End If
    End Sub
    Private Sub Cambiar_Nombre_Propiedades()
        Dim sSQL As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then
            sSQL = "UPDATE ""OITG"" SET ""ItmsGrpNam""='Impresión agrupada' WHERE ""ItmsTypCod""=7"
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = False Then
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido actualizar la propiedad 7 del artículo como ""Impresión agrupada"" ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                objGlobal.SBOApp.StatusBar.SetText("Se ha actualizado la propiedad 7 del artículo como ""Impresión agrupada"" ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If
        End If
    End Sub
    Private Sub ParametrizacionGeneral()

        If Not objGlobal.refDi.OGEN.existeVariable("SERVIDOR_HANA") Then
            'objGlobal.refDi.OGEN.fijarValorVariable("SERVIDOR_HANA", "xper-hanades02:30015")
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
#Region "Logo"
            sReport = "Logo.png"
            'Si no existe lo importamos
            If IO.File.Exists(sArchivo & sReport) = False Then
                If IO.Directory.Exists(sArchivo) = False Then
                    IO.Directory.CreateDirectory(sArchivo)
                End If
                EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), sReport, sArchivo & sReport)
            End If
#End Region

#Region "Etiquetas"
            sReport = "Etiquetas.rpt"
            'Si no existe lo importamos
            If IO.File.Exists(sArchivo & sReport) = False Then
                If IO.Directory.Exists(sArchivo) = False Then
                    IO.Directory.CreateDirectory(sArchivo)
                End If
                EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), sReport, sArchivo & sReport)
            End If
#End Region
#Region "Et. Ubicaciones"
            sReport = "Et_Ubicaciones.rpt"
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
    Public Overrides Function SBOApp_RightClickEvent(infoEvento As ContextMenuInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim Clase As Object = Nothing

        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            Select Case oForm.TypeEx
                Case "142", "143", "1250000940", "940", "180", "150"
                    Clase = New EXO_IMPET(objGlobal)
                    Return CType(Clase, EXO_IMPET).SBOApp_RightClickEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_RightClickEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            If objGlobal.SBOApp.ClientType = BoClientType.ct_Desktop Then
                EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            End If
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
                    Case "EXO-ETIMP"
                        Clase = New EXO_IMPET(objGlobal)
                        Return CType(Clase, EXO_IMPET).SBOApp_MenuEvent(infoEvento)
                    Case "EXO-MnETUB"
                        Clase = New EXO_ETUBI(objGlobal)
                        Return CType(Clase, EXO_ETUBI).SBOApp_MenuEvent(infoEvento)
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
                Case "EXO_ETUBI"
                    Clase = New EXO_ETUBI(objGlobal)
                    Return CType(Clase, EXO_ETUBI).SBOApp_ItemEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Finally
            Clase = Nothing
        End Try
    End Function
End Class
