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
                    Case "142", "143", "1250000940", "940", "180", "150", "1470000002", "65053"
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
                    Case Else
                        If objGlobal.SBOApp.Menus.Exists("EXO-ETIMS") Then
                            objGlobal.SBOApp.Menus.RemoveEx("EXO-ETIMS")
                        End If

                        If objGlobal.SBOApp.Menus.Exists("EXO-ETIMP") Then
                            objGlobal.SBOApp.Menus.RemoveEx("EXO-ETIMP")
                        End If
                End Select
            Else
                Select Case oForm.TypeEx
                    Case "142", "143", "1250000940", "940", "180", "150", "1470000002", "65053"
                        If infoEvento.ItemUID = "" Then
                            'If infoEvento.Row > 0 Then
                            '    _iLineNumRightClick = infoEvento.Row
                            'End If

                        End If
                    Case Else
                        If objGlobal.SBOApp.Menus.Exists("EXO-ETIMS") Then
                            objGlobal.SBOApp.Menus.RemoveEx("EXO-ETIMS")
                        End If

                        If objGlobal.SBOApp.Menus.Exists("EXO-ETIMP") Then
                            objGlobal.SBOApp.Menus.RemoveEx("EXO-ETIMP")
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
                            Case "142", "1250000940", "940", "180"
                                Return Menu_Imprimir_Etiquetas(oForm)
                            Case "150"
                                ' Return Menu_Imprimir_Etiquetas_ART(oForm)
                                Return Menu_Imprimir_Etiquetas_ART2(oForm)
                            Case "1470000002"
                                Return Menu_Imprimir_Etiquetas_ART3(oForm)
                            Case "143"
                                Return Menu_Imprimir_Etiquetas_ART4(oForm)
                            Case "65053"
                                Return Menu_Imprimir_Etiquetas_ART5(oForm)
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
        Dim sTable_Cab As String = ""
        Try
            sTable_Cab = CType(oForm.Items.Item("8").Specific, SAPbouiCOM.EditText).DataBind.TableName.ToString
        Catch ex As Exception
            sTable_Cab = CType(oForm.Items.Item("11").Specific, SAPbouiCOM.EditText).DataBind.TableName.ToString
        End Try
        Dim sTable_Lin As String = Right(sTable_Cab, sTable_Cab.Length - 1) & "1"
        Dim sSQL As String = ""
        Dim odtArticulos As Data.DataTable = Nothing : Dim iCopias As Integer = 1
        Dim sCodeET As String = ""
#End Region
        Menu_Imprimir_Etiquetas = False

        Try
            If oForm.Mode = BoFormMode.fm_OK_MODE Then
                rutaCrystal = objGlobal.path & "\05.Rpt\ETIQUETAS\"
                sCode = oForm.DataSources.DBDataSources.Item(sTable_Cab).GetValue("DocEntry", 0).ToUpper
                objGlobal.SBOApp.StatusBar.SetText("Imprimiendo: " & sCode, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
#Region "Rellena datos Tabla TMP"
                sSQL = "DELETE FROM ""@EXO_ETIQUETA"" WHERE ""U_EXO_USUARIO""='" & objGlobal.compañia.UserSignature.ToString & "' "
                objGlobal.refDi.SQL.sqlStringB1(sSQL)

                sSQL = "SELECT L.""ItemCode"", I.""ItemName"", ifnull(BT.""Quantity"",L.""Quantity"") ""Cantidad"", I.""QryGroup7"", ifnull(BT.""BatchNum"",'') ""Lote"", "
                sSQL &= " ifnull(LT.""MnfSerial"",'') ""Fabricante"", ifnull(B.""BinCode"",'') ""Ubicacion"",L.""DocEntry"", L.""LineNum"" "
                sSQL &= " From " & sTable_Lin & " L INNER JOIN OITM I ON L.""ItemCode""=I.""ItemCode"" "
                sSQL &= " LEFT JOIN " & sTable_Cab & " C ON C.""DocEntry""=L.""DocEntry"" "
                sSQL &= " LEFT JOIN IBT1 BT ON BT.""BaseEntry"" = L.""DocEntry"" And BT.""BaseLinNum"" = L.""LineNum"" And  BT.""BaseType"" = C.""ObjType"" "
                sSQL &= " LEFT JOIN OBTN LT ON LT.""ItemCode"" =BT.""ItemCode"" and LT.""DistNumber""=BT.""BatchNum"" "
                sSQL &= " LEFT JOIN OWHS OW ON OW.""WhsCode""=L.""WhsCode"" "
                sSQL &= " LEFT JOIN OBIN B ON B.""AbsEntry""=OW.""DftBinAbs"" "
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
                        sSQL = "insert into ""@EXO_ETIQUETA"" values('" & sCodeET & "','" & sCodeET & " - " & i.ToString & " de " & iCopias & "','" & MiDataRow("ItemCode").ToString & "',"
                        sSQL &= "'" & MiDataRow("Lote").ToString & "','" & MiDataRow("Fabricante").ToString & "','" & objGlobal.compañia.UserSignature & "',"
                        sSQL &= "'" & MiDataRow("Ubicacion").ToString & "','" & MiDataRow("ItemName").ToString & "', " & MiDataRow("DocEntry").ToString & ", " & MiDataRow("LineNum").ToString & ")"
                        objGlobal.refDi.SQL.sqlUpdB1(sSQL)
                    Next
                Next
#End Region
                sTipoImp = "IMP"
                'Imprimimos la etiqueta
                EXO_GLOBALES.GenerarImpCrystal(objGlobal, rutaCrystal, sCrystal, sCode, sRutaFicheros, sReport, sTipoImp, objGlobal.compañia.UserSignature.ToString)
            Else
                objGlobal.SBOApp.StatusBar.SetText("Antes de imprimir, guarde los cambios...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            End If



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
    Private Function Menu_Imprimir_Etiquetas_ART(ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Variables"
        Dim rutaCrystal As String = "" : Dim sRutaFicheros As String = "" : Dim sReport As String = "" : Dim sTipoImp As String = ""
        Dim sCrystal As String = "Etiquetas.rpt"
        Dim sCode As String = ""
        Dim sTable As String = "OITM"
        Dim sSQL As String = ""
        Dim odtArticulos As Data.DataTable = Nothing : Dim iCopias As Integer = 1
        Dim sCodeET As String = ""
#End Region
        Menu_Imprimir_Etiquetas_ART = False

        Try
            If oForm.Mode = BoFormMode.fm_OK_MODE Then
                rutaCrystal = objGlobal.path & "\05.Rpt\ETIQUETAS\"
                sCode = oForm.DataSources.DBDataSources.Item(sTable).GetValue("ItemCode", 0).ToString
#Region "Rellena datos Tabla TMP"
                sSQL = "DELETE FROM ""@EXO_ETIQUETA"" WHERE ""U_EXO_USUARIO""='" & objGlobal.compañia.UserSignature.ToString & "' "
                objGlobal.refDi.SQL.sqlStringB1(sSQL)

                sSQL = "SELECT I.""ItemCode"", I.""ItemName"", '1' ""Cantidad"", I.""QryGroup7"", ifnull(CAST(LT.""DistNumber"" as Varchar),'') ""Lote"", "
                sSQL &= " ifnull(LT.""MnfSerial"",'') ""Fabricante"", ifnull(B.""BinCode"",'') ""Ubicacion""  "
                sSQL &= " From OITM I "
                sSQL &= " LEFT JOIN OBTN LT ON LT.""ItemCode"" =I.""ItemCode""  "
                sSQL &= " LEFT JOIN OWHS OW ON OW.""WhsCode""=I.""DfltWH"" "
                sSQL &= " LEFT JOIN OBIN B ON B.""AbsEntry""=OW.""DftBinAbs"" "
                sSQL &= " WHERE I.""ItemCode""='" & sCode & "' "
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
                        sSQL = "insert into ""@EXO_ETIQUETA"" values('" & sCodeET & "','" & sCodeET & " - " & i.ToString & " de " & iCopias & "','" & MiDataRow("ItemCode").ToString & "',"
                        sSQL &= "'" & MiDataRow("Lote").ToString & "','" & MiDataRow("Fabricante").ToString & "','" & objGlobal.compañia.UserSignature & "',"
                        sSQL &= "'" & MiDataRow("ItemName").ToString & "','" & MiDataRow("Ubicacion").ToString & "')"
                        objGlobal.refDi.SQL.sqlUpdB1(sSQL)
                    Next
                Next
#End Region
                sTipoImp = "IMP"
                'Imprimimos la etiqueta
                EXO_GLOBALES.GenerarImpCrystal(objGlobal, rutaCrystal, sCrystal, sCode, sRutaFicheros, sReport, sTipoImp, objGlobal.compañia.UserSignature.ToString)
            Else
                objGlobal.SBOApp.StatusBar.SetText("Antes de imprimir, guarde los cambios...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            End If



            Menu_Imprimir_Etiquetas_ART = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(odtArticulos, Object))
        End Try
    End Function

    Private Function Menu_Imprimir_Etiquetas_ART2(ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Variables"
        Dim rutaCrystal As String = "" : Dim sRutaFicheros As String = "" : Dim sReport As String = "" : Dim sTipoImp As String = ""
        Dim sCrystal As String = "Et_Art_Art.rpt"
        Dim sCode As String = "" : Dim sTable As String = "OITM"
        Dim sSQL As String = ""
#End Region
        Menu_Imprimir_Etiquetas_ART2 = False

        Try
            If oForm.Mode = BoFormMode.fm_OK_MODE Then
                rutaCrystal = objGlobal.path & "\05.Rpt\ETIQUETAS\"
                sCode = oForm.DataSources.DBDataSources.Item(sTable).GetValue("ItemCode", 0).ToString
                sTipoImp = "IMP"
                'Imprimimos la etiqueta
                EXO_GLOBALES.GenerarImpCrystal2(objGlobal, rutaCrystal, sCrystal, sCode, sRutaFicheros, sReport, sTipoImp, objGlobal.compañia.UserSignature.ToString)
            Else
                objGlobal.SBOApp.StatusBar.SetText("Antes de imprimir, guarde los cambios...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            End If



            Menu_Imprimir_Etiquetas_ART2 = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)

        End Try
    End Function
    Private Function Menu_Imprimir_Etiquetas_ART3(ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Variables"
        Dim rutaCrystal As String = "" : Dim sRutaFicheros As String = "" : Dim sReport As String = "" : Dim sTipoImp As String = ""
        Dim sCrystal As String = "Et_Art_Ubi.rpt"
        Dim sCode As String = "" : Dim sTable As String = "OBIN"
        Dim sSQL As String = ""
#End Region
        Menu_Imprimir_Etiquetas_ART3 = False

        Try
            If oForm.Mode = BoFormMode.fm_OK_MODE Then
                rutaCrystal = objGlobal.path & "\05.Rpt\ETIQUETAS\"
                sCode = oForm.DataSources.DBDataSources.Item(sTable).GetValue("BinCode", 0).ToString
                sTipoImp = "IMP"
                'Imprimimos la etiqueta
                EXO_GLOBALES.GenerarImpCrystal3(objGlobal, rutaCrystal, sCrystal, sCode, sRutaFicheros, sReport, sTipoImp, objGlobal.compañia.UserSignature.ToString)
            Else
                objGlobal.SBOApp.StatusBar.SetText("Antes de imprimir, guarde los cambios...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            End If



            Menu_Imprimir_Etiquetas_ART3 = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)

        End Try
    End Function
    Private Function Menu_Imprimir_Etiquetas_ART4(ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Variables"
        Dim rutaCrystal As String = "" : Dim sRutaFicheros As String = "" : Dim sReport As String = "" : Dim sTipoImp As String = ""
        Dim sCrystal As String = "Et_Art_Compra.rpt"
        Dim sCode As String = "" : Dim sObjType As String = "" : Dim sTable As String = CType(oForm.Items.Item("8").Specific, SAPbouiCOM.EditText).DataBind.TableName
        Dim sSQL As String = ""
#End Region
        Menu_Imprimir_Etiquetas_ART4 = False

        Try
            If oForm.Mode = BoFormMode.fm_OK_MODE Then
                rutaCrystal = objGlobal.path & "\05.Rpt\ETIQUETAS\"
                sCode = oForm.DataSources.DBDataSources.Item(sTable).GetValue("DocEntry", 0).ToString
                sObjType = oForm.DataSources.DBDataSources.Item(sTable).GetValue("ObjType", 0).ToString
                sTipoImp = "IMP"
                'Imprimimos la etiqueta
                EXO_GLOBALES.GenerarImpCrystal4(objGlobal, rutaCrystal, sCrystal, sCode, sObjType, sRutaFicheros, sReport, sTipoImp, objGlobal.compañia.UserSignature.ToString)
            Else
                objGlobal.SBOApp.StatusBar.SetText("Antes de imprimir, guarde los cambios...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            End If



            Menu_Imprimir_Etiquetas_ART4 = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)

        End Try
    End Function
    Private Function Menu_Imprimir_Etiquetas_ART5(ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Variables"
        Dim rutaCrystal As String = "" : Dim sRutaFicheros As String = "" : Dim sReport As String = "" : Dim sTipoImp As String = ""
        Dim sCrystal As String = "Et_Art_Art_Lote.rpt"
        Dim sCode As String = "" : Dim sLote As String = "" : Dim sTable As String = "OBTN"
        Dim sSQL As String = ""
#End Region
        Menu_Imprimir_Etiquetas_ART5 = False

        Try
            If oForm.Mode = BoFormMode.fm_OK_MODE Then
                rutaCrystal = objGlobal.path & "\05.Rpt\ETIQUETAS\"
                sCode = oForm.DataSources.DBDataSources.Item(sTable).GetValue("ItemCode", 0).ToString
                sLote = oForm.DataSources.DBDataSources.Item(sTable).GetValue("DistNumber", 0).ToString
                sTipoImp = "IMP"
                'Imprimimos la etiqueta
                EXO_GLOBALES.GenerarImpCrystal5(objGlobal, rutaCrystal, sCrystal, sCode, sLote, sRutaFicheros, sReport, sTipoImp, objGlobal.compañia.UserSignature.ToString)
            Else
                objGlobal.SBOApp.StatusBar.SetText("Antes de imprimir, guarde los cambios...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            End If



            Menu_Imprimir_Etiquetas_ART5 = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)

        End Try
    End Function
End Class
