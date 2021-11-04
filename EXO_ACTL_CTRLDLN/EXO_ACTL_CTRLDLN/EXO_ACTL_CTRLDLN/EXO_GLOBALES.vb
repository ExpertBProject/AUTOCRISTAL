Public Class EXO_GLOBALES
#Region "Datos Pedidos Pdtes de asignar"
    Public Shared Sub Cargar_Grid_Ped_Pdte_Asignar(ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oform As SAPbouiCOM.Form)
#Region "Variables"
        Dim sSQL As String = ""
#End Region
        Try
            oform.Freeze(True)

#Region "Cargar Datos Grid"
            oobjGlobal.SBOApp.StatusBar.SetText("Cargando en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' ""Sel."", ""U_EXO_ITEMCODE"" ""Cod. Artículo"", ""U_EXO_ITEMNAME"" ""Descripción"",""U_EXO_CANT"" ""Cantidad"", ""U_EXO_UBACT"" ""Ubicación actual"", "
            sSQL &= " ""U_EXO_ZONAACT"" as ""Zona almacén de Rotación actual"", ""U_EXO_CLAACT"" as ""Clasificación actual de Rotación"" , ""U_EXO_TRASLADO"" as ""Traslado"", "
            sSQL &= " ""U_EXO_UBIDES"" ""Ubicación destino"" "
            sSQL &= " From ""@EXO_TMPOPUBI"" "
            sSQL &= " WHERE ""U_EXO_USUARIO""='" & oobjGlobal.compañia.UserName.ToString & "' "
            sSQL &= " ORDER BY ""Code"", ""LineId"" "
            'Cargamos grid
            oform.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid_Ped_Pdte_Asignar(oobjGlobal, oform)
#End Region

            oform.Freeze(False)
            oobjGlobal.SBOApp.StatusBar.SetText("Fin del proceso de carga.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch exCOM As System.Runtime.InteropServices.COMException
            oform.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oform.Freeze(False)
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub

    Public Shared Sub FormateaGrid_Ped_Pdte_Asignar(ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            CType(oform.Items.Item("grdPDTE").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdPDTE").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            'For i = 1 To 5
            '    CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            '    oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
            '    oColumnTxt.Editable = False
            '    If i = 2 Then
            '        oColumnTxt.LinkedObjectType = "112"
            '    End If
            'Next
            CType(oform.Items.Item("grdPDTE").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
        End Try
    End Sub
#End Region


End Class
