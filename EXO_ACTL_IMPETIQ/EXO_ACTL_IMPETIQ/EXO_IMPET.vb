Imports SAPbouiCOM

Public Class EXO_IMPET
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_RightClickEvent(ByVal infoEvento As ContextMenuInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams

        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            If infoEvento.BeforeAction = True Then
                Select Case oForm.TypeEx
                    Case "142", "143"
                        If objGlobal.SBOApp.Menus.Exists("EXO-ETIMS") Then
                            objGlobal.SBOApp.Menus.RemoveEx("EXO-ETIMS")
                        End If

                        If objGlobal.SBOApp.Menus.Exists("EXO-ETIMP") Then
                            objGlobal.SBOApp.Menus.RemoveEx("EXO-ETIMP")
                        End If

                        oCreationPackage = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams), SAPbouiCOM.MenuCreationParams)
                        Dim oMenuItem As SAPbouiCOM.MenuItem = objGlobal.SBOApp.Menus.Item("1280") 'Data'
                        Dim oMenus As SAPbouiCOM.Menus = oMenuItem.SubMenus
                        If Not objGlobal.SBOApp.Menus.Exists("EXO-ETIMS") Then
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_SEPERATOR
                            oCreationPackage.Position = oMenuItem.SubMenus.Count + 1
                            oCreationPackage.UniqueID = "EXO-ETIMS"
                            oCreationPackage.Enabled = True
                            oMenus = oMenuItem.SubMenus
                            oMenus.AddEx(oCreationPackage)
                        End If

                        If Not objGlobal.SBOApp.Menus.Exists("EXO-ETIMP") Then
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                            oCreationPackage.Position = oMenuItem.SubMenus.Count + 1
                            oCreationPackage.UniqueID = "EXO-ETIMP"
                            oCreationPackage.String = "Imprimir Etiquetas de Artículos"
                            oCreationPackage.Enabled = True
                            oMenus = oMenuItem.SubMenus
                            oMenus.AddEx(oCreationPackage)
                        End If
                End Select
            Else
                Select Case oForm.TypeEx
                    Case "142", "143"
                        If infoEvento.ItemUID = "" Then
                            'If infoEvento.Row > 0 Then
                            '    _iLineNumRightClick = infoEvento.Row
                            'End If

                        End If
                End Select
            End If

            Return True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim sMensaje As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.ActiveForm
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-ETIMP"
                        Select Case oForm.TypeEx
                            Case "142", "143"
                                Return Menu_Imprimir_Etiquetas(oForm)
                        End Select

                End Select
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally

        End Try
    End Function
    Private Function Menu_Imprimir_Etiquetas(ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Variables"
        Dim rutaCrystal As String = "" : Dim sRutaFicheros As String = "" : Dim sReport As String = "" : Dim sTipoImp As String = ""
        Dim sCrystal As String = "Etiquetas.rpt"
        Dim sCode As String = ""
        Dim sTable_Cab As String = CType(oForm.Items.Item("8").Specific, SAPbouiCOM.EditText).DataBind.TableName.ToString
        Dim sTable_Lin As String = Right(sTable_Cab, sTable_Cab.Length - 1) & "1"
        Dim sSQL As String = ""
        Dim odtArticulos As Data.DataTable = Nothing : Dim iCopias As Integer = 1
        Dim sCodeET As String = ""
#End Region
        Menu_Imprimir_Etiquetas = False

        Try
            rutaCrystal = objGlobal.path & "\05.Rpt\ETIQUETAS\"
            sCode = oForm.DataSources.DBDataSources.Item(sTable_Cab).GetValue("DocEntry", 0).ToUpper
#Region "Rellena datos Tabla TMP"
            sSQL = "DELETE FROM ""@EXO_ETIQUETA"" WHERE ""U_EXO_USUARIO""='" & objGlobal.compañia.UserSignature.ToString & "' "
            objGlobal.refDi.SQL.sqlStringB1(sSQL)

            sSQL = "SELECT L.""ItemCode"", ifnull(BT.""Quantity"",L.""Quantity"") ""Cantidad"", I.""QryGroup7"", ifnull(BT.""BatchNum"",'') ""Lote"", "
            sSQL &= " I.""SuppCatNum"" ""Fabricante"" "
            sSQL &= " From " & sTable_Lin & " L INNER JOIN OITM I ON L.""ItemCode""=I.""ItemCode"" "
            sSQL &= " LEFT JOIN " & sTable_Cab & " C ON C.""DocEntry""=L.""DocEntry"" "
            sSQL &= " LEFT JOIN IBT1 BT ON BT.""BaseEntry"" = L.""DocEntry"" And BT.""BaseLinNum"" = L.""LineNum"" And  BT.""BaseType"" = C.""ObjType"" "
            sSQL &= " WHERE L.""DocEntry""=" & sCode
            sSQL &= " ORDER BY L.""LineNum"" "
            odtArticulos = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
            For Each MiDataRow As DataRow In odtArticulos.Rows
                If MiDataRow("QryGroup7").ToString = "Y" Then
                    iCopias = 1
                Else
                    iCopias = CType(MiDataRow("Cantidad").ToString, Integer)
                End If
                For i = 1 To iCopias
                    sSQL = "Select ifnull(MAX(CAST(""Code"" As Integer)),0)+1 FROM ""@EXO_ETIQUETA"" "
                    sCodeET = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    sSQL = "insert into ""@EXO_ETIQUETA"" values('" & sCodeET & "','" & i.ToString & " de " & iCopias & "','" & MiDataRow("ItemCode").ToString & "',"
                    sSQL &= "'" & MiDataRow("Lote").ToString & "','" & MiDataRow("Fabricante").ToString & "','" & objGlobal.compañia.UserSignature & "')"
                    objGlobal.refDi.SQL.sqlUpdB1(sSQL)
                Next
            Next
#End Region
            sTipoImp = "IMP"
            'Imprimimos la etiqueta
            EXO_GLOBALES.GenerarImpCrystal(objGlobal, rutaCrystal, sCrystal, sCode, sRutaFicheros, sReport, sTipoImp, objGlobal.compañia.UserSignature.ToString)


            Menu_Imprimir_Etiquetas = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(odtArticulos, Object))
        End Try
    End Function

End Class
