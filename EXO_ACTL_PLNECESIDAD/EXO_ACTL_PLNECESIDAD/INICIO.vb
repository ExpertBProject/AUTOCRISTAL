Imports SAPbouiCOM
Imports System.Xml
Imports EXO_UIAPI.EXO_UIAPI
Public Class INICIO
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "Variables Globales"

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

        If objGlobal.refDi.comunes.esAdministrador Then

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OCRD.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OCRD", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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

