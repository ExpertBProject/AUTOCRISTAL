Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_IMPART
    Private objGlobal As EXO_UIAPI.EXO_UIAPI

    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            'Apaño por un error que da EXO_Basic.dll al consultar infoEvento.FormTypeEx
            Try
                If infoEvento.FormTypeEx <> "" Then

                End If
            Catch ex As Exception
                Return False
            End Try

            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_IMPART"
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
                        Case "EXO_IMPART"

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
                        Case "EXO_IMPART"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_IMPART"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sArchivo As String = objGlobal.refDi.OGEN.pathGeneral & "\08.Historico\DOC_CARGADOS\" & objGlobal.compañia.CompanyDB & "\PLNECESIDAD\"
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
                    'Limpiar_Grid(oForm)
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
#Region "Grabar en la tabla temporal para tratar los datos"
                        'Ahora abrimos el fichero para tratarlo
                        TratarFichero(sArchivo, oForm)
#End Region
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Fin de la Importación de datos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
            oForm.Close()
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            If objGlobal.SBOApp.Menus.Item("1304").Enabled = True Then
                objGlobal.SBOApp.ActivateMenuItem("1304")
            End If

        End Try
    End Function
    Private Sub Limpiar_Grid(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            'Limpiamos grid
            'Borrar tablas temporales por usuario activo
            sSQL = "DELETE FROM ""@EXO_TMPPACKINGL"" where ""U_EXO_USUARIO""='" & objGlobal.compañia.UserName.ToString & "'  "
            oRs.DoQuery(sSQL)
            sSQL = "DELETE FROM ""@EXO_TMPPACKING"" where ""U_EXO_US""='" & objGlobal.compañia.UserName.ToString & "'  "
            oRs.DoQuery(sSQL)
            'Ahora cargamos el Grid con los datos guardados
            objGlobal.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' ""Sel."", ""U_EXO_ESTADO"" ""Estado"", ""U_EXO_DOCENTRY"" ""Nº Int."",""U_EXO_SERIE"" ""Serie"", ""U_EXO_DOCNUM"" ""Nº Documento"", "
            sSQL &= " ""U_EXO_COMENT"" as ""Descripción Estado"" "
            sSQL &= " From ""@EXO_TMPPACKING"" "
            sSQL &= " WHERE ""U_EXO_US""='" & objGlobal.compañia.UserName.ToString & "' "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Private Sub FormateaGrid(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            For i = 1 To 5
                CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                If i = 2 Then
                    oColumnTxt.LinkedObjectType = "112"
                End If
            Next
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
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
    Private Sub TratarFichero(ByVal sArchivo As String, ByRef oForm As SAPbouiCOM.Form)
        Dim myStream As StreamReader = Nothing
        Dim Reader As XmlTextReader = New XmlTextReader(myStream)
        Dim sSQL As String = ""
        Dim sExiste As String = "" ' Para comprobar si existen los datos
        Dim sDelimitador As String = "2"
        Dim oFormOrigen As SAPbouiCOM.Form = Nothing
        Try
            objGlobal.SBOApp.StatusBar.SetText("Cargando datos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oForm.Freeze(True)
            oFormOrigen = objGlobal.SBOApp.Forms.Item(oForm.DataSources.UserDataSources.Item("UDFRM").Value)
            TratarFichero_TXT(sArchivo, sDelimitador, oFormOrigen, objGlobal.compañia, objGlobal.SBOApp, objGlobal)


            oForm.Freeze(False)


            objGlobal.SBOApp.StatusBar.SetText("Se terminado de leer el fichero. Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.SBOApp.MessageBox("Se terminado de leer el fichero. Fin del proceso")
        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub
    Public Shared Sub TratarFichero_TXT(ByVal sArchivo As String, ByVal sDelimitador As String, ByRef oForm As SAPbouiCOM.Form, ByRef oCompany As SAPbobsCOM.Company, ByRef oSboApp As Application, ByRef objglobal As EXO_UIAPI.EXO_UIAPI)
#Region "Variables"
        ' Apuntador libre a archivo
        Dim Apunt As Integer = FreeFile()
        Dim sMensaje As String = ""
        Dim sArticulo As String = "" : Dim sDescripcion As String = ""
        Dim sExiste As String = ""
        Dim sSQL As String = "" : Dim sSQLART As String = ""
        Dim iLinea As Integer = 0

#End Region
        Try
            ' miramos si existe el fichero y cargamos
            If File.Exists(sArchivo) Then
                objglobal.SBOApp.StatusBar.SetText("Actualizando lista de artículos... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Using MyReader As New Microsoft.VisualBasic.
                        FileIO.TextFieldParser(sArchivo, System.Text.Encoding.UTF7)
                    MyReader.TextFieldType = FileIO.FieldType.Delimited
                    Select Case sDelimitador
                        Case "1" : MyReader.SetDelimiters(vbTab)
                        Case "2" : MyReader.SetDelimiters(";")
                        Case "3" : MyReader.SetDelimiters(",")
                        Case "4" : MyReader.SetDelimiters("-")
                        Case Else : MyReader.SetDelimiters(vbTab)
                    End Select

                    Dim currentRow As String()
                    iLinea = 0
                    'Buscamos el numero code max
                    While Not MyReader.EndOfData
                        Try
                            If iLinea = 0 Then ' Para quitar la cabecera
                                currentRow = MyReader.ReadFields()
                            End If
                            currentRow = MyReader.ReadFields()

                            Dim currentField As String
                            Dim scampos(1) As String
                            Dim iCampo As Integer = 0
                            For Each currentField In currentRow
                                iCampo += 1
                                ReDim Preserve scampos(iCampo)
                                scampos(iCampo) = currentField
                                'SboApp.MessageBox(scampos(iCampo))
                            Next
#Region "Lectura registros"
                            For i = 1 To iCampo
                                Select Case i
                                    Case 1 : sArticulo = scampos(i)
                                    Case 2
                                        Try
                                            sDescripcion = scampos(i)
                                        Catch ex As Exception
                                            sDescripcion = ""
                                        End Try

                                End Select
                            Next
                            If sArticulo = "" Then
                                sMensaje = "En la línea " & iLinea.ToString & " no puede estar vacío el artículo y se omitirá."
                                objglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Else
                                'Buscamos el artículo para ver si existe
                                sSQL = "SELECT ""ItemName"" FROM OITM WHERE ""ItemCode""='" & sArticulo & "' "
                                sExiste = objglobal.refDi.SQL.sqlStringB1(sSQL)
                                If sExiste = "" Then
                                    sMensaje = "En la línea " & iLinea.ToString & " no existe el artículo y se omitirá."
                                    objglobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Else
                                    sDescripcion = sExiste
                                    'Insertamos en el Grid
                                    If sSQLART = "" Then
                                        sSQLART = " SELECT '" & sArticulo & "' ""Código"", '" & sDescripcion & "' ""Descripción"" FROM DUMMY "
                                    Else
                                        sSQLART &= vbCrLf & " UNION ALL SELECT '" & sArticulo & "' ""Código"", '" & sDescripcion & "' ""Descripción"" FROM DUMMY "
                                    End If
                                End If
                            End If
#End Region
                        Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                            objglobal.SBOApp.StatusBar.SetText("(EXO) - Línea " & ex.Message & " no es válida y se omitirá.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                        iLinea += 1
                    End While
                End Using

            Else
                objglobal.SBOApp.StatusBar.SetText("(EXO) - No se ha encontrado el fichero txt a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            ' Cerramos el archivo
            FileClose(Apunt)
            If oForm IsNot Nothing Then
                oForm.DataSources.DataTables.Item("DTART").ExecuteQuery(sSQLART)
                EXO_PNEC.FormateaGridART(oForm)
            Else
                objglobal.SBOApp.StatusBar.SetText("(EXO) - No se puede actualizar la lista de artículos cargados. Vuelva a intentarlo.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        End Try
    End Sub
End Class
