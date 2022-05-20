Imports SAPbouiCOM

Public Class EXO_GLOBALES
    Public Shared Sub Modo_Anadir(ByRef oForm As SAPbouiCOM.Form, ByRef oObjglobal As EXO_UIAPI.EXO_UIAPI)
#Region "variables"
        Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
        Dim sFecha As String = ""
        Dim sSQL As String = ""
        Dim sSerieDef As String = ""
#End Region

        Try
            Select Case oForm.TypeEx
                Case "UDO_FT_EXO_LSTEMB"
                    'Poner fecha
                    sFecha = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                    oForm.DataSources.DBDataSources.Item("@EXO_LSTEMB").SetValue("U_EXO_DOCDATE", 0, sFecha)

                    'Series 
                    sSQL = "SELECT ""Series"",""SeriesName"" FROM NNM1 WHERE ""ObjectCode""='EXO_LSTEMB' "
                    oObjglobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                    oForm.Items.Item("4_U_Cb").DisplayDesc = True

                    'Poner serie por defecto y el num. de documento
                    sSQL = " SELECT ""DfltSeries"" FROM ONNM WHERE ""ObjectCode""='EXO_LSTEMB' "
                    sSerieDef = oObjglobal.refDi.SQL.sqlStringB1(sSQL)
                    CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(sSerieDef, BoSearchKey.psk_ByValue)
                    Poner_DocNum(oForm, sSerieDef, oObjglobal)
                Case "UDO_FT_EXO_ENVTRANS"
                    sFecha = dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")
                    oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("U_EXO_DOCDATE", 0, sFecha)

                    'Series 
                    sSQL = "SELECT ""Series"",""SeriesName"" FROM NNM1 WHERE ""ObjectCode""='EXO_ENVTRANS' "
                    oObjglobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                    oForm.Items.Item("4_U_Cb").DisplayDesc = True

                    'Poner serie por defecto y el num. de documento
                    sSQL = " SELECT ""DfltSeries"" FROM ONNM WHERE ""ObjectCode""='EXO_ENVTRANS' "
                    sSerieDef = oObjglobal.refDi.SQL.sqlStringB1(sSQL)
                    CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Select(sSerieDef, BoSearchKey.psk_ByValue)
                    Dim iNum As Integer
                    iNum = oForm.BusinessObject.GetNextSerialNumber(CType(oForm.Items.Item("4_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, oForm.BusinessObject.Type.ToString)
                    oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("DocNum", 0, iNum.ToString)
                    'Poner_DocNum(oForm, sSerieDef, oObjglobal)
            End Select



        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Sub Poner_DocNum(ByRef oForm As SAPbouiCOM.Form, ByVal sSerie As String, ByRef oObjglobal As EXO_UIAPI.EXO_UIAPI)
#Region "Variables"
        Dim sDocNum As String = ""
        Dim sSQL As String = ""
#End Region
        Try
            Select Case oForm.TypeEx
                Case "UDO_FT_EXO_LSTEMB"
                    sSQL = "SELECT ""NextNumber"" FROM NNM1 WHERE ""ObjectCode""='EXO_LSTEMB' and ""Series""='" & sSerie & "' "
                    sDocNum = oObjglobal.refDi.SQL.sqlStringB1(sSQL)
                    oForm.DataSources.DBDataSources.Item("@EXO_LSTEMB").SetValue("DocNum", 0, sDocNum)
                Case "UDO_FT_EXO_ENVTRANS"
                    sSQL = "SELECT ""NextNumber"" FROM NNM1 WHERE ""ObjectCode""='EXO_ENVTRANS' and ""Series""='" & sSerie & "' "
                    sDocNum = oObjglobal.refDi.SQL.sqlStringB1(sSQL)
                    oForm.DataSources.DBDataSources.Item("@EXO_ENVTRANS").SetValue("DocNum", 0, sDocNum)
            End Select

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
