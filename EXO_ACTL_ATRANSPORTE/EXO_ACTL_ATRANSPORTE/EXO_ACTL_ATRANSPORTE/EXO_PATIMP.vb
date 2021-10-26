Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_PATIMP
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "Variables globales"
    Public Shared _oformTAgencia As SAPbouiCOM.Form
#End Region
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PATIMP"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PATIMP"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PATIMP"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PATIMP"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sArchivo As String = objGlobal.refDi.OGEN.pathGeneral & "\08.Historico\DOC_CARGADOS\" & objGlobal.compañia.CompanyDB & "\PAGTRANSPORTE\"
        Dim sTipoArchivo As String = "" : Dim sNomFICH As String = ""
        Dim sArchivoOrigen As String = ""
        Dim sSQL As String = ""
#End Region
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btn_Fich"
#Region "Coger y leer fichero"
                    sTipoArchivo = "Ficheros CSV|*.csv|Texto|*.txt"
                    'Tenemos que controlar que es cliente o web
                    If objGlobal.SBOApp.ClientType = SAPbouiCOM.BoClientType.ct_Browser Then
                        sArchivoOrigen = objGlobal.SBOApp.GetFileFromBrowser() 'Modificar
                    Else
                        'Controlar el tipo de fichero que vamos a abrir según campo de formato
                        sArchivoOrigen = objGlobal.funciones.OpenDialogFiles("Abrir archivo como", sTipoArchivo)
                    End If

                    If Len(sArchivoOrigen) = 0 Then
                        CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = ""
                        objGlobal.SBOApp.MessageBox("Debe indicar un archivo a importar.")
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Debe indicar un archivo a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        oForm.Items.Item("btn_Carga").Enabled = False
                        Exit Function
                    Else
                        CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = sArchivoOrigen
                        sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                        sArchivo = sArchivo & sNomFICH
                        'Hacemos copia de seguridad para tratarlo
                        Copia_Seguridad(sArchivoOrigen, sArchivo)
#Region "Tarifa de Agencia de transporte"
                        'Ahora abrimos el fichero para tratarlo
                        TratarFichero(sArchivo, oForm)
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Fin de la Actualización...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
                    End If
#End Region
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub TratarFichero(ByVal sArchivo As String, ByRef oForm As SAPbouiCOM.Form)
        Dim myStream As StreamReader = Nothing
        Dim Reader As XmlTextReader = New XmlTextReader(myStream)
        Dim sSQL As String = ""
        Dim sExiste As String = "" ' Para comprobar si existen los datos
        Dim sDelimitador As String = "2"
        Try
            objGlobal.SBOApp.StatusBar.SetText("Cargando datos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            EXO_GLOBALES.TratarFichero_TXT(sArchivo, sDelimitador, oForm, objGlobal.compañia, objGlobal.SBOApp, objGlobal)


            objGlobal.SBOApp.StatusBar.SetText("Se terminado de leer el fichero. Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.SBOApp.MessageBox("Se terminado de leer el fichero. Fin del proceso")
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub
    Private Sub Copia_Seguridad(ByVal sArchivoOrigen As String, ByVal sArchivo As String)
        'Comprobamos el directorio de copia que exista
        Dim sPath As String = ""
        sPath = IO.Path.GetDirectoryName(sArchivo)
        If IO.Directory.Exists(sPath) = False Then
            IO.Directory.CreateDirectory(sPath)
        End If
        If IO.File.Exists(sArchivo) = True Then
            IO.File.Delete(sArchivo)
        End If
        'Subimos el archivo
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Comienza la Copia de seguridad del fichero - " & sArchivoOrigen & " -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        If objGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
            Dim fs As FileStream = New FileStream(sArchivoOrigen, FileMode.Open, FileAccess.Read)
            Dim b(CInt(fs.Length() - 1)) As Byte
            fs.Read(b, 0, b.Length)
            fs.Close()
            Dim fs2 As New System.IO.FileStream(sArchivo, IO.FileMode.Create, IO.FileAccess.Write)
            fs2.Write(b, 0, b.Length)
            fs2.Close()
        Else
            My.Computer.FileSystem.CopyFile(sArchivoOrigen, sArchivo)
        End If
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Copia de Seguridad realizada Correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
End Class
