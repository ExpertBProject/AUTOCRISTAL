Imports SAPbouiCOM
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class EXO_PARRILLA
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Public Shared _ClaseExp As String = ""

    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        SBOApp_MenuEvent = False
        Dim sSQL As String = ""
        Try
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnPARC"
                        If CargarForm() = False Then
                            Exit Function
                        End If
                End Select
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally

        End Try
    End Function
    Public Function CargarForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oColumnCb As SAPbouiCOM.ComboBox = Nothing
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)

        CargarForm = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_PARRILLA.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                Else
                    objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
                    Exit Function
                End If
            End Try
            'Ini_Grid(oForm)

            Try
                sSQL = "SELECT T2.""WhsCode"",T2.""WhsName"" "
                sSQL &= " From OWHS T2 "
                sSQL &= " Order by T2.""WhsName"" "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)
            Catch ex As Exception

            End Try


            Try
                sSQL = " SELECT CAST('-' as NVARCHAR(50)) ""TrnspCode"", CAST(' ' AS NVARCHAR(150))  ""TrnspName"" "
                sSQL &= " FROM DUMMY "
                sSQL &= " UNION ALL "
                sSQL &= " SELECT CAST(""TrnspCode"" as NVARCHAR(50)) ,""TrnspName"" "
                sSQL &= " From OSHP  "
                sSQL &= " ORDER By  ""TrnspName"" "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbEXPE").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                CType(oForm.Items.Item("cbEXPE").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)
            Catch ex As Exception

            End Try


            Try
                sSQL = " SELECT CAST('-' as NVARCHAR(50)) ""territryID"", CAST(' ' AS NVARCHAR(150))  ""descript"" "
                sSQL &= " FROM DUMMY "
                sSQL &= " UNION ALL "
                sSQL &= "SELECT CAST(""territryID"" as NVARCHAR(50)),""descript"" "
                sSQL &= " From OTER  "
                sSQL &= " ORDER By  ""descript"" "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbTERRI").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                CType(oForm.Items.Item("cbTERRI").Specific, SAPbouiCOM.ComboBox).Select(0, BoSearchKey.psk_Index)
            Catch ex As Exception

            End Try

            CType(oForm.Items.Item("cbSAL").Specific, SAPbouiCOM.ComboBox).ValidValues.Remove("-", BoSearchKey.psk_ByValue)
            CType(oForm.Items.Item("cbENT").Specific, SAPbouiCOM.ComboBox).ValidValues.Remove("-", BoSearchKey.psk_ByValue)

            CType(oForm.Items.Item("cbSAL").Specific, SAPbouiCOM.ComboBox).Select("TODOS", BoSearchKey.psk_ByValue)
            CType(oForm.Items.Item("cbENT").Specific, SAPbouiCOM.ComboBox).Select("TODOS", BoSearchKey.psk_ByValue)

            oForm.Items.Item("btCCEXPC").Visible = False
            oForm.State = BoFormStateEnum.fs_Maximized


            CargarForm = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Function CargarFormPSTOCKCOMP(ByRef oFormParrilla As SAPbouiCOM.Form) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)
        Dim dtDatos As System.Data.DataTable = Nothing
        Dim dt As SAPbouiCOM.DataTable = Nothing
        CargarFormPSTOCKCOMP = False

        Try
            'Rellenar grid
            If oFormParrilla.DataSources.DataTables.Item("DTSCOM").Rows.Count > 0 Then
                oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
                oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_PSTOCKCOMP.srf")

                Try
                    oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
                Catch ex As Exception
                    If ex.Message.StartsWith("Form - already exists") = True Then
                        objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                        Exit Function
                    Else
                        objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
                        Exit Function
                    End If
                End Try

                dt = Nothing : dt = oFormParrilla.DataSources.DataTables.Item("DTSCOM")
                dtDatos = New System.Data.DataTable
                Dim responseDataSel = ComprobarDOCSELLIB(oFormParrilla, "DTSCOM", dtDatos, dt)
                Dim sSalidas As String = oFormParrilla.DataSources.UserDataSources.Item("UDSAL").Value.ToString
                Select Case sSalidas
                    Case "-"
                        sSQL = "SELECT
	                        ""DocNum"",
	                        ""ItemCode"",
	                        0 AS ""Cantidad Pendiente""
                        FROM dummy "
                    Case "TODOS"
                        sSQL = "SELECT * FROM ("
                        sSQL &= "(Select
	                        ODLN.""DocNum"",
	                        DLN1.""ItemCode"",
	                        (DLN1.""Quantity""-IFNULL(""EMBL"".""Cant"",0)) AS ""Cantidad Pendiente""
                        FROM ODLN 
                        JOIN DLN1 ON ODLN.""DocEntry"" = DLN1.""DocEntry""
                        JOIN OITM ON DLN1.""ItemCode"" = OITM.""ItemCode""
                        LEFT JOIN 
                        (
                            SELECT 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN"",
                                 SUM(""U_EXO_CANT"") AS ""Cant""
                            FROM ""@EXO_LSTEMBL""
                           GROUP BY 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN""
                        ) ""EMBL"" ON ODLN.""DocEntry"" =  ""EMBL"".""U_EXO_DOCENTRY"" AND  DLN1.""ItemCode"" = ""EMBL"".""U_EXO_ITEMCODE"" AND ""EMBL"".""U_EXO_LINNUM"" = DLN1.""LineNum"" AND ""U_EXO_ORIGEN"" = 'ALBVTA'
                        WHERE 
	                        OITM.""InvntItem"" = 'Y' AND
	                        IFNULL(""EMBL"".""Cant"",0) < DLN1.""Quantity"" AND
	                        ODLN.""DocEntry"" IN "

                        If responseDataSel.Rows.Count > 0 Then
                            sSQL &= "("
                            Dim bComa As Boolean = False
                            For Each MiDataRow As DataRow In responseDataSel.Rows
                                If bComa = True Then
                                    sSQL &= ", "
                                Else
                                    bComa = True
                                End If
                                sSQL &= "'" & MiDataRow("Nº INTERNO").ToString & "' "
                            Next
                            sSQL &= ") ORDER BY 1, 2"
                        Else
                            sSQL &= "(-1)"
                        End If
                        sSQL &= ") UNION ALL "
                        sSQL &= "(Select
	                        OWTQ.""DocNum"",
	                        WTQ1.""ItemCode"",
	                        (WTQ1.""Quantity""-IFNULL(""EMBL"".""Cant"",0)) AS ""Cantidad Pendiente""
                        FROM OWTQ 
                        JOIN WTQ1 ON OWTQ.""DocEntry"" = WTQ1.""DocEntry""
                        JOIN OITM ON WTQ1.""ItemCode"" = OITM.""ItemCode""
                        LEFT JOIN 
                        (
                            SELECT 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN"",
                                 SUM(""U_EXO_CANT"") AS ""Cant""
                            FROM ""@EXO_LSTEMBL""
                           GROUP BY 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN""
                        ) ""EMBL"" ON OWTQ.""DocEntry"" =  ""EMBL"".""U_EXO_DOCENTRY"" AND  WTQ1.""ItemCode"" = ""EMBL"".""U_EXO_ITEMCODE"" AND ""EMBL"".""U_EXO_LINNUM"" = WTQ1.""LineNum"" 
                        WHERE 
	                        OITM.""InvntItem"" = 'Y' AND
	                        IFNULL(""EMBL"".""Cant"",0) < WTQ1.""Quantity"" AND
	                        WTQ1.""DocEntry"" IN "

                        If responseDataSel.Rows.Count > 0 Then
                            sSQL &= "("
                            Dim bComa As Boolean = False
                            For Each MiDataRow As DataRow In responseDataSel.Rows
                                If bComa = True Then
                                    sSQL &= ", "
                                Else
                                    bComa = True
                                End If
                                sSQL &= "'" & MiDataRow("Nº INTERNO").ToString & "' "
                            Next
                            sSQL &= ") ORDER BY 1, 2"
                        Else
                            sSQL &= "(-1)"
                        End If
                        sSQL &= ") UNION ALL "
                        sSQL &= "(Select
	                        ORPD.""DocNum"",
	                        RPD1.""ItemCode"",
	                        (RPD1.""Quantity""-IFNULL(""EMBL"".""Cant"",0)) AS ""Cantidad Pendiente""
                        FROM ORPD 
                        JOIN RPD1 ON ORPD.""DocEntry"" = RPD1.""DocEntry""
                        JOIN OITM ON RPD1.""ItemCode"" = OITM.""ItemCode""
                        LEFT JOIN 
                        (
                            SELECT 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN"",
                                 SUM(""U_EXO_CANT"") AS ""Cant""
                            FROM ""@EXO_LSTEMBL""
                           GROUP BY 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN""
                        ) ""EMBL"" ON ORPD.""DocEntry"" =  ""EMBL"".""U_EXO_DOCENTRY"" AND  RPD1.""ItemCode"" = ""EMBL"".""U_EXO_ITEMCODE"" AND ""EMBL"".""U_EXO_LINNUM"" = RPD1.""LineNum"" 
                        WHERE 
	                        OITM.""InvntItem"" = 'Y' AND
	                        IFNULL(""EMBL"".""Cant"",0) < RPD1.""Quantity"" AND
	                        RPD1.""DocEntry"" IN "

                        If responseDataSel.Rows.Count > 0 Then
                            sSQL &= "("
                            Dim bComa As Boolean = False
                            For Each MiDataRow As DataRow In responseDataSel.Rows
                                If bComa = True Then
                                    sSQL &= ", "
                                Else
                                    bComa = True
                                End If
                                sSQL &= "'" & MiDataRow("Nº INTERNO").ToString & "' "
                            Next
                            sSQL &= ") ORDER BY 1, 2"
                        Else
                            sSQL &= "(-1)"
                        End If
                        sSQL &= ")"
                        sSQL &= ")T "
                    Case "PEDVTA"
#Region "PEDVTA"
                        sSQL = "Select
	                        ODLN.""DocNum"",
	                        DLN1.""ItemCode"",
	                        (DLN1.""Quantity""-IFNULL(""EMBL"".""Cant"",0)) AS ""Cantidad Pendiente""
                        FROM ODLN 
                        JOIN DLN1 ON ODLN.""DocEntry"" = DLN1.""DocEntry""
                        JOIN OITM ON DLN1.""ItemCode"" = OITM.""ItemCode""
                        LEFT JOIN 
                        (
                            SELECT 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN"",
                                 SUM(""U_EXO_CANT"") AS ""Cant""
                            FROM ""@EXO_LSTEMBL""
                           GROUP BY 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN""
                        ) ""EMBL"" ON ODLN.""DocEntry"" =  ""EMBL"".""U_EXO_DOCENTRY"" AND  DLN1.""ItemCode"" = ""EMBL"".""U_EXO_ITEMCODE"" AND ""EMBL"".""U_EXO_LINNUM"" = DLN1.""LineNum"" AND ""U_EXO_ORIGEN"" = 'ALBVTA'
                        WHERE 
	                        OITM.""InvntItem"" = 'Y' AND
	                        IFNULL(""EMBL"".""Cant"",0) < DLN1.""Quantity"" AND
	                        ODLN.""DocEntry"" IN "

                        If responseDataSel.Rows.Count > 0 Then
                            sSQL &= "("
                            Dim bComa As Boolean = False
                            For Each MiDataRow As DataRow In responseDataSel.Rows
                                If bComa = True Then
                                    sSQL &= ", "
                                Else
                                    bComa = True
                                End If
                                sSQL &= "'" & MiDataRow("Nº INTERNO").ToString & "' "
                            Next
                            sSQL &= ") ORDER BY 1, 2"
                        Else
                            sSQL &= "(-1)"
                        End If
#End Region
                    Case "SOLTRA"
#Region "SOLTRA"
                        sSQL = "Select
	                        OWTQ.""DocNum"",
	                        WTQ1.""ItemCode"",
	                        (WTQ1.""Quantity""-IFNULL(""EMBL"".""Cant"",0)) AS ""Cantidad Pendiente""
                        FROM OWTQ 
                        JOIN WTQ1 ON OWTQ.""DocEntry"" = WTQ1.""DocEntry""
                        JOIN OITM ON WTQ1.""ItemCode"" = OITM.""ItemCode""
                        LEFT JOIN 
                        (
                            SELECT 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN"",
                                 SUM(""U_EXO_CANT"") AS ""Cant""
                            FROM ""@EXO_LSTEMBL""
                           GROUP BY 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN""
                        ) ""EMBL"" ON OWTQ.""DocEntry"" =  ""EMBL"".""U_EXO_DOCENTRY"" AND  WTQ1.""ItemCode"" = ""EMBL"".""U_EXO_ITEMCODE"" AND ""EMBL"".""U_EXO_LINNUM"" = WTQ1.""LineNum"" 
                        WHERE 
	                        OITM.""InvntItem"" = 'Y' AND
	                        IFNULL(""EMBL"".""Cant"",0) < WTQ1.""Quantity"" AND
	                        WTQ1.""DocEntry"" IN "

                        If responseDataSel.Rows.Count > 0 Then
                            sSQL &= "("
                            Dim bComa As Boolean = False
                            For Each MiDataRow As DataRow In responseDataSel.Rows
                                If bComa = True Then
                                    sSQL &= ", "
                                Else
                                    bComa = True
                                End If
                                sSQL &= "'" & MiDataRow("Nº INTERNO").ToString & "' "
                            Next
                            sSQL &= ") ORDER BY 1, 2"
                        Else
                            sSQL &= "(-1)"
                        End If
#End Region

                    Case "SDPROV"
#Region "SDPROV"
                        sSQL = "Select
	                        ORPD.""DocNum"",
	                        RPD1.""ItemCode"",
	                        (RPD1.""Quantity""-IFNULL(""EMBL"".""Cant"",0)) AS ""Cantidad Pendiente""
                        FROM ORPD 
                        JOIN RPD1 ON ORPD.""DocEntry"" = RPD1.""DocEntry""
                        JOIN OITM ON RPD1.""ItemCode"" = OITM.""ItemCode""
                        LEFT JOIN 
                        (
                            SELECT 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN"",
                                 SUM(""U_EXO_CANT"") AS ""Cant""
                            FROM ""@EXO_LSTEMBL""
                           GROUP BY 
                                 ""U_EXO_DOCENTRY"",
                                 ""U_EXO_ITEMCODE"",
                                 ""U_EXO_LINNUM"",
                                 ""U_EXO_ORIGEN""
                        ) ""EMBL"" ON ORPD.""DocEntry"" =  ""EMBL"".""U_EXO_DOCENTRY"" AND  RPD1.""ItemCode"" = ""EMBL"".""U_EXO_ITEMCODE"" AND ""EMBL"".""U_EXO_LINNUM"" = RPD1.""LineNum"" 
                        WHERE 
	                        OITM.""InvntItem"" = 'Y' AND
	                        IFNULL(""EMBL"".""Cant"",0) < RPD1.""Quantity"" AND
	                        RPD1.""DocEntry"" IN "

                        If responseDataSel.Rows.Count > 0 Then
                            sSQL &= "("
                            Dim bComa As Boolean = False
                            For Each MiDataRow As DataRow In responseDataSel.Rows
                                If bComa = True Then
                                    sSQL &= ", "
                                Else
                                    bComa = True
                                End If
                                sSQL &= "'" & MiDataRow("Nº INTERNO").ToString & "' "
                            Next
                            sSQL &= ") ORDER BY 1, 2"
                        Else
                            sSQL &= "(-1)"
                        End If
#End Region

                End Select

                oForm.DataSources.DataTables.Item("DTSTOCK").ExecuteQuery(sSQL)

                If (oForm.DataSources.DataTables.Item("DTSTOCK").Rows.Count = 1) Then
                    If (oForm.DataSources.DataTables.Item("DTSTOCK").GetValue(1, 0).ToString.Equals("0")) Then
                        sSQL = "SELECT 'No hay articulos pendientes de Packing' AS ""Message"" FROM DUMMY"
                        oForm.DataSources.DataTables.Item("DTSTOCK").ExecuteQuery(sSQL)
                    End If
                End If

                FormateaGrid_PSTOCKCOMP(oForm)
                CargarFormPSTOCKCOMP = True
            Else
                objGlobal.SBOApp.StatusBar.SetText("No hay datos para mostrar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox("No hay datos para mostrar.")
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            If oForm IsNot Nothing Then
                oForm.Visible = True
            End If

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Function CargarFormPSTOCKBTOS(ByRef oFormParrilla As SAPbouiCOM.Form) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)
        Dim dtDatos As System.Data.DataTable = Nothing
        Dim dt As SAPbouiCOM.DataTable = Nothing
        CargarFormPSTOCKBTOS = False

        Try
            'Rellenar grid
            If oFormParrilla.DataSources.DataTables.Item("DTE").Rows.Count > 0 Then
                'Comprobamos si tenemos marcado alguna línea
                If ComprobarDOC(oFormParrilla, "DTE") = True Then
                    oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
                    oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_BTOS.srf")

                    Try
                        oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
                    Catch ex As Exception
                        If ex.Message.StartsWith("Form - already exists") = True Then
                            objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Function
                        ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                            Exit Function
                        Else
                            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
                            Exit Function
                        End If
                    End Try

                    dt = Nothing : dt = oFormParrilla.DataSources.DataTables.Item("DTE")
                    dtDatos = New System.Data.DataTable
                    Dim responseDataSel = ComprobarDOCENT(oFormParrilla, "DTE", dtDatos, dt, "PED")
                    sSQL = "SELECT * FROM ("
#Region "Pedidos"
                    sSQL &= "(SELECT 'Pedido' ""Tipo"",
	                        X4.""DocEntry"" AS ""Nº Interno"", X4.""DocNum"" AS ""Nº Documento"",	                       
	                        X0.""DocEntry"" AS ""Interno Entrada"", X0.""DocNum"" AS ""Nº Entrada"",
	                         0 AS ""Interno Emb"",
	                        IFNULL(T0.""Code"", 'Sin PackingList') AS ""Nº Embalaje"",   IFNULL(T1.""U_EXO_IDBULTO"",'Sin Bulto') AS ""Id Bulto"",
	                        IFNULL(T1.""U_EXO_CODE"", X3.""ItemCode"") AS ""Articulo"",  
	                        case when X3.""OpenQty"" = 0 then 0 else IFNULL(T1.""U_EXO_CANT"", X3.""OpenQty"") end AS ""Pdte. Recibir"" ,                                             
                            CASE WHEN  X1.""DocEntry"" is null then 0 else T1.""U_EXO_CANT"" end  AS ""Recibido"",
	                        CASE WHEN (X1.""DocEntry"" is null) THEN 0 ELSE  (IFNULL(T1.""U_EXO_CANT"",0) - IFNULL(T.""Cant_REUBICADA"" ,0) ) END AS ""Pdte. Reubicar""                  
                        FROM OPOR X4
                        LEFT JOIN POR1 X3 ON X4.""DocEntry"" = X3.""DocEntry"" 
	                    LEFT JOIN ""@EXO_PACKING""  T0 On T0.""U_EXO_OBJTYPE""='22' and T0.""Code"" = X4.""U_EXO_PACKING""
	                    Left Join ""@EXO_PACKINGL"" T1 On T0.""Code"" = T1.""Code"" And T1.""U_EXO_LINEA"" = X3.""LineNum"" 
	                    left JOIN OPDN X0 ON X0.""U_EXO_PACKING"" = T0.""Code"" and X0.""CANCELED"" = 'N'      
                        LEFT JOIN PDN1 X1 ON X1.""DocEntry"" = X0.""DocEntry"" AND X1.""BaseLine"" = X3.""LineNum"" AND X1.""BaseType"" = X3.""ObjType""
                        LEFT JOIN (select  T0.""U_EXO_PACKING"", T1.""ItemCode"", T1.""U_EXO_LOT_ID"", IFNULL(SUM(T1.""Quantity""), 0) as ""Cant_REUBICADA""  
                                           from OWTR               T0
                                           LEFT JOIN WTR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" 
                                           where T1.""U_EXO_LOT_ID"" is not null 
                                           GROUP by  T0.""U_EXO_PACKING"", T1.""ItemCode"", T1.""U_EXO_LOT_ID"") T
                                  ON  T.""U_EXO_PACKING"" = T0.""Code"" and  T.""ItemCode"" = T1.""U_EXO_CODE"" AND T.""U_EXO_LOT_ID"" = T1.""U_EXO_IDBULTO""  
                        WHERE X4.""DocEntry"" IN   "
                    If responseDataSel.Rows.Count > 0 Then


                        sSQL &= "("
                        Dim bComa As Boolean = False
                        For Each MiDataRow As DataRow In responseDataSel.Rows
                            If bComa = True Then
                                sSQL &= ", "
                            Else
                                bComa = True
                            End If
                            sSQL &= "'" & MiDataRow("Nº INTERNO").ToString & "' "
                        Next
                        sSQL &= ") ORDER BY 1, 2) "
                    Else
                        sSQL &= " (-1) )"
                    End If
#End Region
#Region "Devoluciones"
                    dt = Nothing : dt = oFormParrilla.DataSources.DataTables.Item("DTE")
                    dtDatos = New System.Data.DataTable
                    Dim responseDataSel2 = ComprobarDOCENT(oFormParrilla, "DTE", dtDatos, dt, "SDE")

                    sSQL &= "UNION ALL
                            (SELECT 'Sol. Devolución' ""Tipo"",
	                        X4.""DocEntry"" AS ""Nº Interno"", X4.""DocNum"" AS ""Nº Documento"",	                       
	                        X0.""DocEntry"" AS ""Interno Entrada"", X0.""DocNum"" AS ""Nº Entrada"",
	                         0 AS ""Interno Emb"",
	                        IFNULL(T0.""Code"", 'Sin PackingList') AS ""Nº Embalaje"",   IFNULL(T1.""U_EXO_IDBULTO"",'Sin Bulto') AS ""Id Bulto"",
	                        IFNULL(T1.""U_EXO_CODE"", X3.""ItemCode"") AS ""Articulo"",  
	                        case when X3.""OpenQty"" = 0 then 0 else IFNULL(T1.""U_EXO_CANT"", X3.""OpenQty"") end AS ""Pdte. Recibir"" ,                                             
                            CASE WHEN  X1.""DocEntry"" is null then 0 else T1.""U_EXO_CANT"" end  AS ""Recibido"",
	                        CASE WHEN (X1.""DocEntry"" is null) THEN 0 ELSE  (IFNULL(T1.""U_EXO_CANT"",0) - IFNULL(T.""Cant_REUBICADA"" ,0) ) END AS ""Pdte. Reubicar""                  
                        FROM ORRR X4
                        LEFT JOIN RRR1 X3 ON X4.""DocEntry"" = X3.""DocEntry"" 
	                    LEFT JOIN ""@EXO_PACKING""  T0 On T0.""U_EXO_OBJTYPE""='234000031' and T0.""Code"" = X4.""U_EXO_PACKING""
	                    Left Join ""@EXO_PACKINGL"" T1 On T0.""Code"" = T1.""Code"" And T1.""U_EXO_LINEA"" = X3.""LineNum"" 
	                    left JOIN ORDN X0 ON X0.""U_EXO_PACKING"" = T0.""Code"" and X0.""CANCELED"" = 'N'      
                        LEFT JOIN RDN1 X1 ON X1.""DocEntry"" = X0.""DocEntry"" AND X1.""BaseLine"" = X3.""LineNum"" AND X1.""BaseType"" = X3.""ObjType""
                        LEFT JOIN (select  T0.""U_EXO_PACKING"", T1.""ItemCode"", T1.""U_EXO_LOT_ID"", IFNULL(SUM(T1.""Quantity""), 0) as ""Cant_REUBICADA""  
                                           from OWTR               T0
                                           LEFT JOIN WTR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" 
                                           where T1.""U_EXO_LOT_ID"" is not null 
                                           GROUP by  T0.""U_EXO_PACKING"", T1.""ItemCode"", T1.""U_EXO_LOT_ID"") T
                                  ON  T.""U_EXO_PACKING"" = T0.""Code"" and  T.""ItemCode"" = T1.""U_EXO_CODE"" AND T.""U_EXO_LOT_ID"" = T1.""U_EXO_IDBULTO""  
                        WHERE X4.""DocEntry"" IN   "

                    If responseDataSel2.Rows.Count > 0 Then
                        sSQL &= "("
                        Dim bComa As Boolean = False
                        For Each MiDataRow As DataRow In responseDataSel2.Rows
                            If bComa = True Then
                                sSQL &= ", "
                            Else
                                bComa = True
                            End If
                            sSQL &= "'" & MiDataRow("Nº INTERNO").ToString & "' "
                        Next
                        sSQL &= ") ORDER BY 1, 2) "
                    Else
                        sSQL &= " (-1) )"
                    End If
#End Region
#Region "Solicitud"
                    dt = Nothing : dt = oFormParrilla.DataSources.DataTables.Item("DTE")
                    dtDatos = New System.Data.DataTable
                    Dim responseDataSel3 = ComprobarDOCENT(oFormParrilla, "DTE", dtDatos, dt, "STR")
                    sSQL &= " UNION ALL
                            (SELECT 'Sol. Traslado' ""Tipo"",
	                        X4.""DocEntry"" AS ""Nº Interno"", X4.""DocNum"" AS ""Nº Documento"",	                       
	                        X0.""DocEntry"" AS ""Interno Entrada"", X0.""DocNum"" AS ""Nº Entrada"",
	                         0 AS ""Interno Emb"",
	                        IFNULL(T0.""Code"", 'Sin PackingList') AS ""Nº Embalaje"",   IFNULL(T1.""U_EXO_IDBULTO"",'Sin Bulto') AS ""Id Bulto"",
	                        IFNULL(T1.""U_EXO_CODE"", X3.""ItemCode"") AS ""Articulo"",  
	                        case when X3.""OpenQty"" = 0 then 0 else IFNULL(T1.""U_EXO_CANT"", X3.""OpenQty"") end AS ""Pdte. Recibir"" ,                                             
                            CASE WHEN  X1.""DocEntry"" is null then 0 else T1.""U_EXO_CANT"" end  AS ""Recibido"",
	                        CASE WHEN (X1.""DocEntry"" is null) THEN 0 ELSE  (IFNULL(T1.""U_EXO_CANT"",0) - IFNULL(T.""Cant_REUBICADA"" ,0) ) END AS ""Pdte. Reubicar""                  
                        FROM OWTQ X4
                        LEFT JOIN WTQ1 X3 ON X4.""DocEntry"" = X3.""DocEntry"" 
	                    LEFT JOIN ""@EXO_PACKING""  T0 On T0.""U_EXO_OBJTYPE""='1250000001' and T0.""Code"" = X4.""U_EXO_PACKING""
	                    Left Join ""@EXO_PACKINGL"" T1 On T0.""Code"" = T1.""Code"" And T1.""U_EXO_LINEA"" = X3.""LineNum"" 
	                    left JOIN OWTR X0 ON X0.""U_EXO_PACKING"" = T0.""Code"" and X0.""CANCELED"" = 'N'      
                        LEFT JOIN WTR1 X1 ON X1.""DocEntry"" = X0.""DocEntry"" AND X1.""BaseLine"" = X3.""LineNum"" AND X1.""BaseType"" = X3.""ObjType""
                        LEFT JOIN (select  T0.""U_EXO_PACKING"", T1.""ItemCode"", T1.""U_EXO_LOT_ID"", IFNULL(SUM(T1.""Quantity""), 0) as ""Cant_REUBICADA""  
                                           from OWTR               T0
                                           LEFT JOIN WTR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" 
                                           where T1.""U_EXO_LOT_ID"" is not null 
                                           GROUP by  T0.""U_EXO_PACKING"", T1.""ItemCode"", T1.""U_EXO_LOT_ID"") T
                                  ON  T.""U_EXO_PACKING"" = T0.""Code"" and  T.""ItemCode"" = T1.""U_EXO_CODE"" AND T.""U_EXO_LOT_ID"" = T1.""U_EXO_IDBULTO""  
                        WHERE X4.""DocEntry"" IN   "


                    If responseDataSel3.Rows.Count > 0 Then
                        sSQL &= "("
                        Dim bComa As Boolean = False
                        For Each MiDataRow As DataRow In responseDataSel3.Rows
                            If bComa = True Then
                                sSQL &= ", "
                            Else
                                bComa = True
                            End If
                            sSQL &= "'" & MiDataRow("Nº INTERNO").ToString & "' "
                        Next
                        sSQL &= ") ORDER BY 1, 2) "
                    Else
                        sSQL &= " (-1) ) "
                    End If

#End Region


                    sSQL &= ")T "

                    oForm.DataSources.DataTables.Item("DTSTOCK").ExecuteQuery(sSQL)

                    If (oForm.DataSources.DataTables.Item("DTSTOCK").Rows.Count = 1) Then
                        If (oForm.DataSources.DataTables.Item("DTSTOCK").GetValue(1, 0).ToString.Equals("0")) Then
                            sSQL = "SELECT 'No hay articulos para bultos' AS ""Message"" FROM DUMMY"
                            oForm.DataSources.DataTables.Item("DTSTOCK").ExecuteQuery(sSQL)
                        End If
                    End If

                    FormateaGrid_PSTOCKBULTOS(oForm)

                End If

            Else
                objGlobal.SBOApp.StatusBar.SetText("No hay datos para mostrar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox("No hay datos para mostrar.")
            End If

            CargarFormPSTOCKBTOS = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            If oForm IsNot Nothing Then
                oForm.Visible = True
            End If

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Function CargarFormPSTOCK(ByRef oFormParrilla As SAPbouiCOM.Form) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)
        Dim dtDatos As System.Data.DataTable = Nothing
        Dim dt As SAPbouiCOM.DataTable = Nothing
        CargarFormPSTOCK = False

        Try
            'Rellenar grid
            If oFormParrilla.DataSources.DataTables.Item("DTSLIB").Rows.Count > 0 Then
                oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
                oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_PSTOCK.srf")

                Try
                    oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
                Catch ex As Exception
                    If ex.Message.StartsWith("Form - already exists") = True Then
                        objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                        Exit Function
                    Else
                        objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
                        Exit Function
                    End If
                End Try

                dt = Nothing : dt = oFormParrilla.DataSources.DataTables.Item("DTSLIB")
                dtDatos = New System.Data.DataTable
                Dim responseDataSel = ComprobarDOCSELLIB(oFormParrilla, "DTSLIB", dtDatos, dt)

                sSQL = "SELECT
                            OPKL.""AbsEntry"" AS ""Numero Picking"",
                            IFNULL(ORDR.""DocNum"", OWTQ.""DocNum"") AS ""Num. Documento"",
                            IFNULL(RDR1.""ItemCode"", WTQ1.""ItemCode"") AS ""Articulo"",
                            IFNULL(RDR1.""LineNum"", WTQ1.""LineNum"") AS ""Num. Linea"",
                            PKL1.""RelQtty"" - IFNULL(WTR1.""Quantity"",0) AS ""Pendiente"" 
                        FROM OPKL
                        JOIN PKL1 ON OPKL.""AbsEntry"" = PKL1.""AbsEntry"" 
                        LEFT JOIN ORDR ON PKL1.""OrderEntry"" = ORDR.""DocEntry"" AND PKL1.""BaseObject"" = ORDR.""ObjType""
                        LEFT JOIN RDR1 ON ORDR.""DocEntry"" = RDR1.""DocEntry"" AND RDR1.""LineNum"" = PKL1.""OrderLine""
                        LEFT JOIN OWTQ ON PKL1.""OrderEntry"" = OWTQ.""DocEntry"" AND PKL1.""BaseObject"" = OWTQ.""ObjType""
                        LEFT JOIN WTQ1 ON OWTQ.""DocEntry"" = WTQ1.""DocEntry"" AND WTQ1.""LineNum"" = PKL1.""OrderLine""
                        LEFT JOIN OWTR ON OPKL.""AbsEntry"" = CAST((CASE WHEN IFNULL(OWTR.""U_EXO_NUMPIC"",'') = '' THEN '0' ELSE OWTR.""U_EXO_NUMPIC"" END) AS INTEGER) AND 
                                          PKL1.""PickEntry"" = CAST((CASE WHEN IFNULL(OWTR.""U_EXO_LINPIC"",'') = '' THEN '0' ELSE OWTR.""U_EXO_LINPIC"" END) AS INTEGER)  
                        LEFT JOIN WTR1 ON OWTR.""DocEntry"" = WTR1.""DocEntry""    
                        WHERE OPKL.""AbsEntry"" IN "

                If responseDataSel.Rows.Count > 0 Then
                    sSQL &= "("
                    Dim bComa As Boolean = False
                    For Each MiDataRow As DataRow In responseDataSel.Rows
                        If bComa = True Then
                            sSQL &= ", "
                        Else
                            bComa = True
                        End If
                        sSQL &= "'" & MiDataRow("Picking").ToString & "' "
                    Next
                    sSQL &= ")  AND PKL1.""RelQtty"" - IFNULL(WTR1.""Quantity"",0) > 0 ORDER BY 2 "
                Else
                    sSQL &= "(-1)"
                End If
                oForm.DataSources.DataTables.Item("DTSTOCK").ExecuteQuery(sSQL)

                If (oForm.DataSources.DataTables.Item("DTSTOCK").Rows.Count = 1) Then
                    If (oForm.DataSources.DataTables.Item("DTSTOCK").GetValue(1, 0).ToString.Equals("0")) Then
                        sSQL = "SELECT 'No hay articulos pendientes' AS ""Message"" FROM DUMMY"
                        oForm.DataSources.DataTables.Item("DTSTOCK").ExecuteQuery(sSQL)
                    End If
                End If

                FormateaGrid_PSTOCK(oForm)
                CargarFormPSTOCK = True
            Else
                objGlobal.SBOApp.StatusBar.SetText("No hay datos para mostrar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox("No hay datos para mostrar.")
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            If oForm IsNot Nothing Then
                oForm.Visible = True
            End If

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Function CargarFormRSTOCK(ByRef oFormParrilla As SAPbouiCOM.Form) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)
        Dim dtDatos As System.Data.DataTable = Nothing
        Dim dt As SAPbouiCOM.DataTable = Nothing
        CargarFormRSTOCK = False

        Try
            'Rellenar grid
            If oFormParrilla.DataSources.DataTables.Item("DTSPTE").Rows.Count > 0 Then
                oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
                oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_RSTOCK.srf")

                Try
                    oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
                Catch ex As Exception
                    If ex.Message.StartsWith("Form - already exists") = True Then
                        objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                        Exit Function
                    Else
                        objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
                        Exit Function
                    End If
                End Try

                dt = Nothing : dt = oFormParrilla.DataSources.DataTables.Item("DTSPTE")
                dtDatos = New System.Data.DataTable
                ComprobarDOCSEL(oFormParrilla, "DTSPTE", dtDatos, dt)
                sSQL = "SELECT ""ObjType"" ""TIPO"", ""DocEntry"" ""Nº INTERNO"", ""DocNum"" ""Documento"", ""LineNum"" ""Nº LINEA"", ""ItemCode"" ""ARTÍCULO"", ""ALMACEN"" ""ALMACÉN"", ""OpenQty"" ""CANTIDAD"", ""OnHand"", 0 AS ""Pendiente"", 0 AS ""Reservado"" FROM ""EXO_ROTURA_DETAILS"" "
                sSQL &= " WHERE ""ALMACEN""='" & CType(oFormParrilla.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                If dtDatos.Rows.Count > 0 Then
                    sSQL &= " and ""DocEntry"" in ("
                    Dim bComa As Boolean = False
                    For Each MiDataRow As DataRow In dtDatos.Rows
                        If bComa = True Then
                            sSQL &= ", "
                        Else
                            bComa = True
                        End If
                        sSQL &= "'" & MiDataRow("Nº INTERNO").ToString & "' "
                    Next
                    sSQL &= ")"
                End If

                sSQL &= " ORDER BY ""U_EXO_DOCPRI"" DESC, ""CreateDate"" ASC, ""LineNum"" ASC, ""DocEntry"" ASC"
                oForm.DataSources.DataTables.Item("DTSTOCK").ExecuteQuery(sSQL)
                If (oForm.DataSources.DataTables.Item("DTSTOCK").Rows.Count = 1) Then
                    If (oForm.DataSources.DataTables.Item("DTSTOCK").GetValue(2, 0) = 0) Then
                        oForm.DataSources.DataTables.Item("DTSTOCK").ExecuteQuery("SELECT 'No existen registros coincidentes' AS ""Mensaje"" FROM DUMMY")
                        CType(oForm.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item("Mensaje").Editable = False
                        Return True
                    End If
                End If

                FormateaGrid_RSTOCK(oForm)

                Dim dtRotStock = oForm.DataSources.DataTables.Item("DTSTOCK")
                Dim rotStock = New Dictionary(Of String, Integer)
                For row = 0 To dtRotStock.Rows.Count - 1

                    Dim itemCode = dtRotStock.GetValue("ARTÍCULO", row).ToString()
                    Dim wareHouse = dtRotStock.GetValue("ALMACÉN", row).ToString()

                    If (rotStock.ContainsKey(itemCode + "_" + wareHouse)) Then
                        dtRotStock.SetValue("OnHand", row, rotStock(itemCode + "_" + wareHouse))
                    End If

                    Dim quantityAvailable = dtRotStock.GetValue("OnHand", row).ToString()
                    Dim quantity = dtRotStock.GetValue("CANTIDAD", row).ToString()
                    Dim reservado = 0
                    Dim pendiente = 0

                    If (rotStock.ContainsKey(itemCode + "_" + wareHouse)) Then
                        If (quantityAvailable > 0) Then
                            If (rotStock(itemCode + "_" + wareHouse) > 0) Then
                                If (quantityAvailable > quantity) Then
                                    pendiente = 0
                                    reservado = quantity
                                Else
                                    pendiente = quantityAvailable - quantity
                                    reservado = quantityAvailable
                                End If
                            End If
                        Else
                            pendiente = quantity - quantityAvailable
                            reservado = 0
                        End If

                        rotStock(itemCode + "_" + wareHouse) = quantityAvailable - reservado
                        dtRotStock.SetValue("Pendiente", row, Math.Abs(pendiente))
                        dtRotStock.SetValue("Reservado", row, reservado)
                    Else
                        If (quantityAvailable > 0) Then
                            If (quantityAvailable > quantity) Then
                                pendiente = 0
                                reservado = quantity
                            Else
                                pendiente = quantityAvailable - quantity
                                reservado = quantityAvailable
                            End If
                        Else
                            pendiente = quantityAvailable - quantity
                            reservado = 0
                        End If

                        rotStock.Add(itemCode + "_" + wareHouse, quantityAvailable - reservado)
                        dtRotStock.SetValue("Reservado", row, reservado)
                        dtRotStock.SetValue("Pendiente", row, Math.Abs(pendiente))
                    End If
                Next row

                Dim grid = CType(oForm.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid)
                grid.Columns.Item("Nº INTERNO").Visible = False
                CType(grid.Columns.Item("Documento"), EditTextColumn).LinkedObjectType = "2"

                CargarFormRSTOCK = True
            Else
                objGlobal.SBOApp.StatusBar.SetText("No hay datos para mostrar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox("No hay datos para mostrar.")
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            If oForm IsNot Nothing Then
                oForm.Visible = True
            End If

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
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
                        Case "EXO_PARRILLA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                    If EventHandler_FORM_RESIZE_After(infoEvento) = False Then
                                        Return False
                                    End If
                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PARRILLA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                    If EventHandler_MATRIX_LINK_PRESSED(infoEvento) = False Then
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PARRILLA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PARRILLA"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

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
    Private Function EventHandler_FORM_RESIZE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        EventHandler_FORM_RESIZE_After = False
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Items.Item("grdSPTE").Height = 140
            oForm.Items.Item("Item_5").Top = oForm.Items.Item("grdSPTE").Top - 15
            oForm.Items.Item("btLPicking").Top = oForm.Items.Item("grdSPTE").Top
            oForm.Items.Item("btCPed").Top = oForm.Items.Item("grdSPTE").Top + 25
            oForm.Items.Item("btCALM").Top = oForm.Items.Item("grdSPTE").Top + 50
            oForm.Items.Item("cbEXPCB").Top = oForm.Items.Item("grdSPTE").Top + 75
            oForm.Items.Item("btCCEXP").Top = oForm.Items.Item("grdSPTE").Top + 90
            oForm.Items.Item("btASS").Top = oForm.Items.Item("grdSPTE").Top + 115

            oForm.Items.Item("grdSLIB").Height = 140
            oForm.Items.Item("Item_6").Top = oForm.Items.Item("grdSLIB").Top - 15
            oForm.Items.Item("btGENALB").Top = oForm.Items.Item("grdSLIB").Top + 10
            oForm.Items.Item("cbEXPCBL").Top = oForm.Items.Item("btGENALB").Top + 70
            oForm.Items.Item("btCCEXPL").Top = oForm.Items.Item("btGENALB").Top + 88

            oForm.Items.Item("grdSCOM").Height = 140
            oForm.Items.Item("Item_12").Top = oForm.Items.Item("grdSCOM").Top - 15
            oForm.Items.Item("btCCEXPC").Top = oForm.Items.Item("grdSCOM").Top + 5
            oForm.Items.Item("btImpD").Top = oForm.Items.Item("btCCEXPC").Top + 50
            oForm.Items.Item("btIMPE").Top = oForm.Items.Item("btImpD").Top + 50
            oForm.Items.Item("btnPndCm").Top = oForm.Items.Item("btIMPE").Top - 25


            oForm.Items.Item("grdE").Height = 140
            oForm.Items.Item("Item_18").Top = oForm.Items.Item("grdE").Top - 15
            oForm.Items.Item("btnBtos").Top = oForm.Items.Item("grdE").Top + 10
            EventHandler_FORM_RESIZE_After = True
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_COMBO_SELECT_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_COMBO_SELECT_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                If pVal.ItemUID = "grdSPTE" And pVal.ColUID = "CLASE EXP." Then
                    _ClaseExp = oForm.DataSources.DataTables.Item("DTSPTE").GetValue("CLASE EXP.", pVal.Row).ToString

                    If objGlobal.SBOApp.MessageBox("Desea cambiar la clase de expedicion ?", 1, "Si", "No") <> 1 Then
                        Return False
                    End If
                ElseIf pVal.ItemUID = "grdSLIB" And pVal.ColUID = "CLASE EXP." Then
                    _ClaseExp = oForm.DataSources.DataTables.Item("DTSLIB").GetValue("CLASE EXP.", pVal.Row).ToString
                End If
            End If



            EventHandler_COMBO_SELECT_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)

        End Try
    End Function
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        EventHandler_COMBO_SELECT_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                If pVal.ItemUID = "grdSCOM" And pVal.ColUID = "CLASE EXP." Then
                    Dim sExpe As String = CType(CType(oForm.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item("CLASE EXP."), SAPbouiCOM.ComboBoxColumn).GetSelectedValue(pVal.Row).Value.ToString
                    'Buscamos la agencia
                    sSQL = "SELECT ""U_EXO_AGE"" FROM OSHP WHERE ""TrnspCode""='" & sExpe & "' "
                    Dim sAGE As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    If sAGE = "" Then
                        sAGE = "-1"
                    End If

                    CType(oForm.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).DataTable.SetValue("AG. TRANSPORTE", pVal.Row, sAGE)
                ElseIf pVal.ItemUID = "grdSPTE" And pVal.ColUID = "CLASE EXP." Then
                    Dim grid = CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.Grid)
                    objGlobal.SBOApp.StatusBar.SetText("Cambiando clase de expedición... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    CambiarClaseExpedicionCombo(oForm, grid.DataTable.UniqueID, objGlobal, grid.GetDataTableRowIndex(pVal.Row))

                    objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                ElseIf pVal.ItemUID = "grdSLIB" And pVal.ColUID = "CLASE EXP." Then
                    Dim grid = CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.Grid)
                    objGlobal.SBOApp.StatusBar.SetText("Cambiando clase de expedición... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    CambiarClaseExpedicionCombo(oForm, grid.DataTable.UniqueID, objGlobal, grid.GetDataTableRowIndex(pVal.Row))

                    objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                End If
            End If

            EventHandler_COMBO_SELECT_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_MATRIX_LINK_PRESSED(ByVal pVal As ItemEvent) As Boolean

        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim sTipo As String = ""
        EventHandler_MATRIX_LINK_PRESSED = False

        Try
            oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            Dim grid = CType(oForm.Items.Item(pVal.ItemUID.ToString.Trim).Specific, SAPbouiCOM.Grid)

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                Return True
            End If

            Select Case pVal.ItemUID
                Case "grdSPTE"

                    oColumnTxt = CType(CType(oForm.Items.Item(pVal.ItemUID.ToString.Trim).Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.EditTextColumn)
                    sTipo = CType(oForm.Items.Item(pVal.ItemUID.ToString.Trim).Specific, SAPbouiCOM.Grid).DataTable.GetValue("T. SALIDA", grid.GetDataTableRowIndex(pVal.Row)).ToString
                    Dim objType = CType(oForm.Items.Item("grdSPTE").Specific, Grid).DataTable.GetValue("ObjectType", grid.GetDataTableRowIndex(pVal.Row)).ToString
                    Dim nroInter = CType(oForm.Items.Item("grdSPTE").Specific, Grid).DataTable.GetValue("Nº INTERNO", grid.GetDataTableRowIndex(pVal.Row)).ToString

                    If (pVal.ColUID = "Nº DOCUMENTO") Then
                        Select Case objType
                            Case "17" 'Albaranes de ventas
                                objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_Order, "", nroInter)
                                Return False
                            Case "1250000001" 'Sol de traslado
                                objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_StockTransfersRequest, "", nroInter)
                                Return False
                            Case "234000032" 'Devolución de proveedor
                                objGlobal.SBOApp.OpenForm("234000032", "", nroInter)
                                Return False
                        End Select
                    ElseIf (pVal.ColUID = "Picking" Or pVal.ColUID = "CÓDIGO") Then
                        Return True
                    End If
                    Return False
                Case "grdSLIB"
                    oColumnTxt = CType(CType(oForm.Items.Item(pVal.ItemUID.ToString.Trim).Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.EditTextColumn)
                    sTipo = grid.DataTable.GetValue("T. SALIDA", grid.GetDataTableRowIndex(pVal.Row)).ToString
                    Dim nroInter = grid.DataTable.GetValue("Nº INTERNO", grid.GetDataTableRowIndex(pVal.Row)).ToString

                    If (pVal.ColUID = "Nº DOCUMENTO") Then
                        Select Case sTipo
                            Case "PEDVTA" 'Albaranes de ventas
                                objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_Order, "", nroInter)
                                Return False
                            Case "SOLTRA" 'Sol de traslado
                                objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_StockTransfersRequest, "", nroInter)
                                Return False
                            Case "SDPROV" 'Devolución de proveedor
                                objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_GoodsReturns, "", nroInter)
                                Return False
                        End Select
                    ElseIf (pVal.ColUID = "Picking") Then
                        Return True
                    End If
                    Return False
                Case "grdSCOM"
                    oColumnTxt = CType(CType(oForm.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.EditTextColumn)
                    sTipo = CType(oForm.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).DataTable.GetValue("T. SALIDA", grid.GetDataTableRowIndex(pVal.Row)).ToString
                    Dim nroInt = CType(oForm.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).DataTable.GetValue("Nº INTERNO", grid.GetDataTableRowIndex(pVal.Row)).ToString

                    If (pVal.ColUID = "Nº DOCUMENTO") Then
                        Select Case sTipo
                            Case "ALBVTA" 'Albaranes de ventas
                                objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_DeliveryNotes, "", nroInt)
                                Return False
                            Case "SDPROV" 'Sol de traslado
                                objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_GoodsReturns, "", nroInt)
                                Return False
                            Case "SOLTRA" 'Devolución de proveedor
                                objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_StockTransfersRequest, "", nroInt)
                                Return False
                        End Select
                        Return False
                    Else
                        Return True
                    End If
                Case "grdE"
                    oColumnTxt = CType(CType(oForm.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.EditTextColumn)
                    sTipo = CType(oForm.Items.Item("grdE").Specific, SAPbouiCOM.Grid).DataTable.GetValue("T. ENTRADA", grid.GetDataTableRowIndex(pVal.Row)).ToString
                    Dim nroInt = CType(oForm.Items.Item("grdE").Specific, SAPbouiCOM.Grid).DataTable.GetValue("Nº INTERNO", grid.GetDataTableRowIndex(pVal.Row)).ToString

                    If (pVal.ColUID = "Nº DOCUMENTO") Then
                        Select Case sTipo
                            Case "PED" 'Albaranes de ventas
                                objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_PurchaseOrder, "", nroInt)
                                Return False
                            Case "STR" 'Sol de traslado
                                objGlobal.SBOApp.OpenForm("1250000001", "", nroInt)
                                Return False
                            Case "SDE" 'Sol de Devolución de cliente
                                objGlobal.SBOApp.OpenForm("234000031", "", nroInt)
                                Return False
                        End Select
                        'Return False
                    ElseIf (pVal.ColUID = "DOC. ENTRADA") Then
                        Dim nroIntDevol = CType(oForm.Items.Item("grdE").Specific, SAPbouiCOM.Grid).DataTable.GetValue("ID DOC. ENTRADA", grid.GetDataTableRowIndex(pVal.Row)).ToString
                        'objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_GoodsReceiptPO, "", nroIntDevol)
                        Select Case sTipo
                            Case "PED" '
                                objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_GoodsReceiptPO, "17", nroIntDevol)
                                Return False
                            Case "STR" 'Sol de traslado
                                'CType(CType(oForm.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item("DOC. ENTRADA"), SAPbouiCOM.EditTextColumn).LinkedObjectType = "67"
                                objGlobal.SBOApp.OpenForm("67", "", nroIntDevol)
                                Return False
                            Case "SDE" 'Sol de Devolución de cliente
                                objGlobal.SBOApp.OpenForm("16", "", nroIntDevol)
                                Return False
                        End Select
                        'Return False
                    Else
                        Return True
                    End If
                Case "grdRSTOCK"
                    Dim nroInt = CType(oForm.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).DataTable.GetValue("Nº INTERNO", grid.GetDataTableRowIndex(pVal.Row)).ToString
                    Dim tipo = CType(oForm.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).DataTable.GetValue("TIPO", grid.GetDataTableRowIndex(pVal.Row)).ToString

                    Select Case tipo
                        Case "17" 'Albaranes de ventas
                            objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_Order, "", nroInt)
                            Return False
                        Case "1250000001" 'Sol de traslado
                            objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_StockTransfersRequest, "", nroInt)
                            Return False
                        Case "234000032" 'Devolución de proveedor
                            objGlobal.SBOApp.OpenForm("234000032", "", nroInt)
                            Return False
                    End Select

                    Return False
            End Select

            EventHandler_MATRIX_LINK_PRESSED = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByVal pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oDataTable As DataTable = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        EventHandler_Choose_FromList_After = False

        Try
            oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                Return True
            End If

            oCFLEvento = CType(pVal, IChooseFromListEvent)

            oDataTable = oCFLEvento.SelectedObjects
            If Not oDataTable Is Nothing Then
                Select Case oCFLEvento.ChooseFromListUID
                    Case "CFLICD"
                        oDataTable = oCFLEvento.SelectedObjects

                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "txtICD" Then
                                Try
                                    CType(oForm.Items.Item("txtICD").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                                Catch ex As Exception

                                End Try
                            End If
                        End If
                    Case "CFLICH"
                        oDataTable = oCFLEvento.SelectedObjects

                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "txtICH" Then
                                Try
                                    CType(oForm.Items.Item("txtICH").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                                Catch ex As Exception

                                End Try
                            End If
                        End If
                End Select
            End If

            EventHandler_Choose_FromList_After = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.FormDatatable(oDataTable)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnFIL" ' Filtro
                    FiltrarPDTE(oForm)
                    FiltrarLIB(oForm)
                    FiltrarCOM(oForm)
                    FiltrarENT(oForm)
                    SetComboExp(oForm)
                Case "btLPicking" ' Liberar picking
                    If ComprobarDOC(oForm, "DTSPTE") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Liberando para picking... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If LiberarparaPicking(oForm, "DTSPTE", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarPDTE(oForm)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btnPndCm"
                    CargarFormPSTOCKCOMP(oForm)
                Case "btnBtos"
                    CargarFormPSTOCKBTOS(oForm)
                Case "btCPed" ' Cerrar Documentos
                    If ComprobarDOC(oForm, "DTSPTE") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Cerrando documentos... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If CerrarDocumentos(oForm, "DTSPTE", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarPDTE(oForm)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btCCEXP" 'Cambio clase de expedición
                    If ComprobarDOC(oForm, "DTSPTE") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Cambiando clase de expedición... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If CambiarClaseExpedicionMasiva(oForm, "DTSPTE", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarPDTE(oForm)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btCCEXPL" ' Cambio de clase de exp. liberadas
                    If ComprobarDOC(oForm, "DTSLIB") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Cambiando clase de expedición... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If CambiarClaseExpedicionMasiva(oForm, "DTSLIB", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btASS" 'Acceso a Art. sin Stocks
                    CargarFormRSTOCK(oForm)
                Case "btnPend" 'Acceso a Art pendiente
                    CargarFormPSTOCK(oForm)
                Case "btCALM" 'Cambio de almacén
                    If ComprobarDOCPED(oForm, "DTSPTE") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Cambiando almacén de documentos... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If CambiarAlmacen(oForm, "DTSPTE", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarPDTE(oForm)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btGENALB" 'Generar Albaranes
                    If ComprobarDOC(oForm, "DTSLIB") = True Then
                        objGlobal.SBOApp.StatusBar.SetText("Generando Documentos... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If Gen_DOC(oForm, "DTSLIB", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarPDTE(oForm)
                        FiltrarLIB(oForm)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btCCEXPC" ' Cambio de clase de exp. completadas
                    If ComprobarDOC(oForm, "DTSCOM") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Cambiando clase de expedición... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If CambiarClaseExpedicion(oForm, "DTSCOM", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        FiltrarCOM(oForm)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btImpD" ' Impresión de documentos
                    If ComprobarDOC(oForm, "DTSCOM") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Imprimiendo documentos... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        If Impresion_Doc(oForm, "DTSCOM", objGlobal) = False Then
                            Exit Function
                        End If
                        oForm.Freeze(False)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
                Case "btIMPE" 'Impresión de etiquetas
                    If ComprobarDOC(oForm, "DTSCOM") = True Then
                        'Calculando datos
                        objGlobal.SBOApp.StatusBar.SetText("Imprimiendo Etiquetas... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        If Impresion_ET(oForm, "DTSCOM", objGlobal) = False Then
                            Exit Function
                        End If
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log del sistema para ver las operaciones realizadas.")
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Shared Function Impresion_ET(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Impresion_ET = False
#Region "VARIABLES"
        Dim oCmpSrv As SAPbobsCOM.CompanyService = oobjGlobal.compañia.GetCompanyService()
        Dim oReportLayoutService As SAPbobsCOM.ReportLayoutsService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), SAPbobsCOM.ReportLayoutsService)
        Dim oPrintParam As SAPbobsCOM.ReportLayoutPrintParams = CType(oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams), SAPbobsCOM.ReportLayoutPrintParams)
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim sLayout As String = "" : Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
#End Region

        Try
            oRs = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString

                    sLayout = oobjGlobal.funcionesUI.refDi.OGEN.valorVariable("EXO_ETPARRILLA")
                    If sLayout = "" Then
                        oobjGlobal.SBOApp.StatusBar.SetText("Parámetro [EXO_ETPARRILLA] no tiene valor. Revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Else
                        sSQL = "SELECT DISTINCT ""DocEntry"",""U_EXO_IDBULTO"" FROM ""@EXO_LSTEMBL"" WHERE ""U_EXO_ORIGEN""='" & sTIPODOC & "' And ""U_EXO_DOCENTRY""='" & sDocEntry & "'"
                        oRs.DoQuery(sSQL)
                        If oRs.RecordCount > 0 Then
                            Dim sDirExportar As String = oobjGlobal.path & "\05.Rpt\PARRILLADOC\"
                            Dim sRutaFicheros As String = oobjGlobal.path & "\05.Rpt\PARRILLADOC\ET_CREADAS\"
                            If IO.Directory.Exists(sDirExportar) = False Then
                                IO.Directory.CreateDirectory(sDirExportar)
                            End If
                            If IO.Directory.Exists(sRutaFicheros) = False Then
                                IO.Directory.CreateDirectory(sRutaFicheros)
                            End If
                            Dim sCrystal As String = "ETIQUETASBULTOS.rpt"
                            EXO_GLOBALES.GetCrystalReportFile(oobjGlobal, sDirExportar & sCrystal, sLayout)
                            oobjGlobal.SBOApp.StatusBar.SetText("Layout " & sDirExportar & sCrystal, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)

                            For p = 0 To oRs.RecordCount - 1
                                Dim sDocEntryLstEmb As String = oRs.Fields.Item("DocEntry").Value.ToString.Trim
                                Dim sIDBulto As String = oRs.Fields.Item("U_EXO_IDBULTO").Value.ToString.Trim

                                Dim sTipoImp As String = "IMP"
                                'Imprimimos la etiqueta
                                GenerarImpCrystalET(oobjGlobal, sDirExportar, sCrystal, sDocEntryLstEmb, sIDBulto, sTipoImp, sRutaFicheros, "")

                                oRs.MoveNext()
                            Next
                        Else
                            oobjGlobal.SBOApp.StatusBar.SetText("No tiene Lista de embalajes. No puede imprimir la etiqueta.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                        End If
                        'oPrintParam.LayoutCode = sLayout 'codigo del formato importado en SAP
                        'oPrintParam.DocEntry = sDocEntryLstEmb 'parametro que se envia al crystal, DocEntry de la transaccion

                        'oReportLayoutService.Print(oPrintParam)


                    End If

                End If
            Next

            Impresion_ET = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oReportLayoutService = Nothing
            oCmpSrv = Nothing
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Shared Sub GenerarImpCrystalET(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal rutaCrystal As String, ByVal sCrystal As String,
                                       ByVal sDocEntry As String, ByVal sIDBULTO As String, ByVal sTIPOIMP As String, ByVal sDir As String, ByRef sReport As String)

        Dim oCRReport As ReportDocument = Nothing
        Dim oFileDestino As DiskFileDestinationOptions = Nothing
        Dim sServer As String = ""
        Dim sDriver As String = ""
        Dim sBBDD As String = ""
        Dim sUser As String = ""
        Dim sPwd As String = ""
        Dim sConnection As String = ""
        Dim oLogonProps As NameValuePairs2 = Nothing

        Dim conrepor As DataSourceConnections = Nothing
        Dim sImpresora As String = "" : Dim nCopias As Integer = 1
        Dim sSQL As String = ""
        Try

            oCRReport = New ReportDocument()

            oCRReport.Load(rutaCrystal & sCrystal)

            oCRReport.DataSourceConnections.Clear()

            oObjGlobal.SBOApp.StatusBar.SetText("DocEntry: " & sDocEntry & " - IDBULTO: " & sIDBULTO, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)

            'Establecemos las conexiones a la BBDD
            sServer = oObjGlobal.funcionesUI.refDi.OGEN.valorVariable("SERVIDOR_HANA") ' objGlobal.compañia.Server
            'sServer = objGlobal.refDi.SQL.dameCadenaConexion.ToString
            sBBDD = oObjGlobal.compañia.CompanyDB
            sUser = oObjGlobal.refDi.SQL.usuarioSQL
            sPwd = oObjGlobal.refDi.SQL.claveSQL

            sDriver = "HDBODBC"
            sConnection = "DRIVER={" & sDriver & "};UID=" & sUser & ";PWD=" & sPwd & ";SERVERNODE=" & sServer & ";DATABASE=" & sBBDD & ";"
            'sConnection = "DRIVER={" & sDriver & "};" & sServer & ";DATABASE=" & sBBDD & ";"
            oObjGlobal.SBOApp.StatusBar.SetText("Conectando: " & sConnection, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
            oLogonProps = oCRReport.DataSourceConnections(0).LogonProperties
            oLogonProps.Set("Provider", sDriver)
            oLogonProps.Set("Connection String", sConnection)


            'Establecemos los parámetros para el report.
            oCRReport.SetParameterValue("DocEntry", sDocEntry)
            oCRReport.SetParameterValue("ID_Bulto", sIDBULTO)
            'oCRReport.SetParameterValue("Schema@", sBBDD)


            oCRReport.DataSourceConnections(0).SetLogonProperties(oLogonProps)
            oCRReport.DataSourceConnections(0).SetConnection(sServer, sBBDD, False)
            oObjGlobal.SBOApp.StatusBar.SetText("Connection String: " & sConnection, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)

            For Each oSubReport As ReportDocument In oCRReport.Subreports
                For Each oConnection As IConnectionInfo In oSubReport.DataSourceConnections
                    oConnection.SetConnection(sServer, sBBDD, False)
                    oConnection.SetLogon(sUser, sPwd)
                Next
            Next
            oObjGlobal.SBOApp.StatusBar.SetText("Actualizado conect Subreport...", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)

            Select Case sTIPOIMP
                Case "PDF"
#Region "Exportar a PDF"
                    'Preparamos para la exportación
                    If IO.Directory.Exists(sDir) = False Then
                        IO.Directory.CreateDirectory(sDir)
                    End If
                    sReport = sDir & "Et_Bultos_" & sDocEntry & ".pdf"
                    'Compruebo si existe y lo borro
                    If IO.File.Exists(sReport) Then
                        IO.File.Delete(sReport)
                    End If
                    oObjGlobal.SBOApp.StatusBar.SetText("Generando pdf para envio impresión...Espere por favor", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)

                    oCRReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat

                    oFileDestino = New CrystalDecisions.Shared.DiskFileDestinationOptions
                    oFileDestino.DiskFileName = sReport

                    'Le pasamos al reporte el parámetro destino del reporte (ruta)
                    oCRReport.ExportOptions.DestinationOptions = oFileDestino

                    'Le indicamos que el reporte no es para mostrarse en pantalla, sino, que es para guardar en disco
                    oCRReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile

                    'Finalmente exportamos el reporte a PDF
                    oCRReport.Export()
                    '            oCRReport.ExportToDisk(ExportFormatType.PortableDocFormat, sReport)
#End Region
                Case "IMP"
#Region "Imprimir a impresora"
                    'Buscamos la impresora por defecto
                    'Dim instance As New Printing.PrinterSettings
                    'sImpresora = instance.PrinterName
                    sImpresora = oObjGlobal.refDi.SQL.sqlStringB1("SELECT ""U_EXO_IMPBUL"" FROM OUSR WHERE ""USERID""='" & oObjGlobal.compañia.UserSignature.ToString & "' ")
                    'oObjGlobal.SBOApp.StatusBar.SetText("Impresora: " & sImpresora, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)
                    oObjGlobal.SBOApp.StatusBar.SetText("Buscando Impresora " & sImpresora & "...Espere por favor", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                    If EXO_GLOBALES.IsPrinterOnline(sImpresora) = True Then
                        oObjGlobal.SBOApp.StatusBar.SetText("Imprimiendo en " & sImpresora & "...Espere por favor", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                        oCRReport.PrintOptions.NoPrinter = False
                        oCRReport.PrintOptions.PrinterName = sImpresora
                        oCRReport.PrintToPrinter(nCopias, False, 0, 9999)
                    Else
                        oObjGlobal.SBOApp.StatusBar.SetText("La impresora seleccionada en el usuario no se encuentra o está offline. Por favor verifique la parametrización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    End If
#End Region
            End Select

            'Cerramos
            oCRReport.Close()
            oCRReport.Dispose()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oObjGlobal.SBOApp.StatusBar.SetText("Fin del proceso de impresión.", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success)
            oCRReport = Nothing
            oFileDestino = Nothing
        End Try
    End Sub
    '    Public Shared Function Impresion_Doc(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
    '        Impresion_Doc = False
    '#Region "VARIABLES"
    '        Dim sTIPODOC As String = "" : Dim sTIPODOCDES As String = ""
    '        Dim sDocEntry As String = "" : Dim sDocNum As String = ""
    '        Dim rutaCrystal As String = "" : Dim sRutaFicheros As String = "" : Dim sCrystal As String = "" : Dim sReport As String = "" : Dim sTipoImp As String = ""
    '#End Region

    '        Try
    '            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
    '                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
    '                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
    '                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
    '                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString

    '                    rutaCrystal = oobjGlobal.path & "\05.Rpt\PARRILLADOC\"
    '                    sRutaFicheros = My.Computer.FileSystem.SpecialDirectories.Temp

    '                    Select Case sTIPODOC
    '                        Case "ALBVTA"
    '#Region "Entregas"
    '                            sCrystal = "ENTREGAS.rpt"
    '                            sTIPODOCDES = " entrega "
    '#End Region
    '                        Case "SOLTRA" ' Sol. de Traslado                           
    '#Region "Sol de traslado"
    '                            sCrystal = "SOLTRASLADO.rpt"
    '                            sTIPODOCDES = " sol. de traslado "
    '#End Region
    '                        Case "DPROV" ' Dev. de Proveedor
    '#Region "Dev de proveedor"
    '                            sCrystal = "DEVPROVEEDOR.rpt"
    '                            sTIPODOCDES = " dev. de proveedor "
    '#End Region
    '                    End Select
    '                    oobjGlobal.SBOApp.StatusBar.SetText("Imprimiendo " & sTIPODOCDES & ": " & sDocNum, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success)
    '                    sTipoImp = "IMP"
    '                    'Imprimimos la etiqueta
    '                    EXO_GLOBALES.GenerarImpCrystal(oobjGlobal, rutaCrystal, sCrystal, sDocNum, sDocEntry, oobjGlobal.compañia.CompanyDB, sTIPODOC, sRutaFicheros, sReport, sTipoImp, oobjGlobal.compañia.UserSignature.ToString)
    '                End If
    '            Next

    '            Impresion_Doc = True
    '        Catch exCOM As System.Runtime.InteropServices.COMException
    '            Throw exCOM
    '        Catch ex As Exception
    '            Throw ex
    '        Finally

    '        End Try
    '    End Function
    Public Shared Function Impresion_Doc(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Impresion_Doc = False
#Region "VARIABLES"
        Dim oCmpSrv As SAPbobsCOM.CompanyService = oobjGlobal.compañia.GetCompanyService()
        Dim oReportLayoutService As SAPbobsCOM.ReportLayoutsService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), SAPbobsCOM.ReportLayoutsService)
        Dim oPrintParam As SAPbobsCOM.ReportLayoutPrintParams = CType(oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams), SAPbobsCOM.ReportLayoutPrintParams)
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim sLayout As String = "" : Dim sSQL As String = ""
#End Region

        Try
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                    Select Case sTIPODOC
                        Case "ALBVTA"
#Region "Entregas"
                            sSQL = "SELECT ""DEFLT_REP"" FROM RTYP WHERE left(""CODE"",4)='DLN2' "

#End Region
                        Case "SOLTRA" ' Sol. de Traslado                           
#Region "Sol de traslado"
                            sSQL = "SELECT ""DEFLT_REP"" FROM RTYP WHERE left(""CODE"",4)='WTQ1' "
#End Region
                        Case "DPROV" ' Dev. de Proveedor
#Region "Dev de proveedor"
                            sSQL = "SELECT ""DEFLT_REP"" FROM RTYP WHERE left(""CODE"",4)='RPD2' "
#End Region
                    End Select
                    sLayout = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    oPrintParam.LayoutCode = sLayout 'codigo del formato importado en SAP
                    oPrintParam.DocEntry = CType(sDocEntry, Integer) 'parametro que se envia al crystal, DocEntry de la transaccion

                    oReportLayoutService.Print(oPrintParam)
                End If
            Next

            Impresion_Doc = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oReportLayoutService = Nothing
            oCmpSrv = Nothing
        End Try
    End Function
    Public Shared Function Gen_DOC(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Gen_DOC = False
#Region "VARIABLES"
        Dim oRs As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLinPICK As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsPedido As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sDocEntryFinal As String = "" : Dim sDocNumFinal As String = ""
        Dim oDocuments As SAPbobsCOM.Documents = Nothing : Dim oDocument_Lines As SAPbobsCOM.Document_Lines = Nothing
        Dim oDocFinal As SAPbobsCOM.Documents = Nothing : Dim oDocFinal_Lines As SAPbobsCOM.Document_Lines = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing : Dim oDocStockTransfer_Lines As SAPbobsCOM.StockTransfer_Lines = Nothing
        Dim oPicking As SAPbobsCOM.PickLists = Nothing
        Dim oRsLote As SAPbobsCOM.Recordset = Nothing
        Dim dCantidad As Double = 0 : Dim dCantPdte As Double = 0
        Dim grid = CType(oForm.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid)
#End Region

        Try
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                    dCantidad = EXO_GLOBALES.DblTextToNumber(oobjGlobal.compañia, oForm.DataSources.DataTables.Item(sData).GetValue("Cant.", i).ToString)
                    dCantPdte = EXO_GLOBALES.DblTextToNumber(oobjGlobal.compañia, oForm.DataSources.DataTables.Item(sData).GetValue("Cant. Pdte.", i).ToString)
                    If dCantidad = dCantPdte Then
                        oobjGlobal.SBOApp.StatusBar.SetText("No se ha movido ningún artículo a la zona de packing del Documento Nº: " & sDocNum & ". ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Else
#Region "Generar Documento"
                        Select Case sTIPODOC
                            Case "PEDVTA" ' Pedido 
                                sSQL = "SELECT P.""AbsEntry"", P.""OrderEntry"", P.""OrderLine"", P.""BaseObject"", ifnull(WTR1.""Quantity"",0) ""Cantidad"",
                                    ifnull(WTR1.""DocEntry"",0) ""Traslado"", WTR1.""LineNum"" ""LinTraslado""
                                            FROM PKL1 P
                                            INNER JOIN OPKL ON OPKL.""AbsEntry""=P.""AbsEntry"" 
                                            LEFT JOIN OWTR ON CAST((CASE WHEN IFNULL(OWTR.""U_EXO_NUMPIC"",'') = '' THEN '0' ELSE OWTR.""U_EXO_NUMPIC"" END) AS INTEGER) = P.""AbsEntry"" and CAST((CASE WHEN IFNULL(OWTR.""U_EXO_LINPIC"",'') = '' THEN '0' ELSE OWTR.""U_EXO_LINPIC"" END) AS INTEGER) = P.""PickEntry""
                                            LEFT JOIN WTR1 ON OWTR.""DocEntry""=WTR1.""DocEntry""
                                        Where P.""OrderEntry""=" & sDocEntry & " and P.""BaseObject""='17' and ifnull(WTR1.""Quantity"",0)>0 and ifnull(WTR1.""DocEntry"",0)<>0 
                                        Order By P.""OrderLine"""
                                oRsLinPICK.DoQuery(sSQL)
                                If oRsLinPICK.RecordCount > 0 Then
                                    oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)

                                    If oDocuments.GetByKey(CType(sDocEntry, Integer)) = True Then
                                        oDocument_Lines = oDocuments.Lines
                                        oDocFinal = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes), SAPbobsCOM.Documents)
                                        'oDocFinal_Lines = oDocFinal.Lines
                                        oDocFinal.CardCode = oDocuments.CardCode
                                        oDocFinal.DocDate = Now.Date
                                        oDocFinal.TaxDate = Now.Date

                                        For cu As Integer = 0 To oDocuments.UserFields.Fields.Count - 1
                                            If oDocuments.UserFields.Fields.Item(oDocuments.UserFields.Fields.Item(cu).Name).IsNull = SAPbobsCOM.BoYesNoEnum.tNO Then
                                                oDocFinal.UserFields.Fields.Item(oDocuments.UserFields.Fields.Item(cu).Name).Value = oDocuments.UserFields.Fields.Item(oDocuments.UserFields.Fields.Item(cu).Name).Value
                                            Else
                                                oDocFinal.UserFields.Fields.Item(oDocuments.UserFields.Fields.Item(cu).Name).SetNullValue()
                                            End If
                                        Next
                                        'Estatus a c
                                        oDocFinal.UserFields.Fields.Item("U_EXO_STATUSP").Value = "C"

                                        For J = 0 To oRsLinPICK.RecordCount - 1
                                            If (J > 0) Then
                                                oDocFinal.Lines.Add()
                                            End If
                                            oDocFinal.Lines.BaseType = oDocuments.DocObjectCode
                                            oDocFinal.Lines.BaseEntry = oDocuments.DocEntry
                                            oDocFinal.Lines.BaseLine = oRsLinPICK.Fields.Item("Orderline").Value

#Region "Lotes y ubicacion"
                                            oRsLote = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                                            dCantidad = 0
                                            'Incluimos los Lotes
                                            sSQL = "Select sum(T0.""Quantity"") ""Quantity"", T6.""DistNumber""    ,IFNULL(t0.""OcrCode"",'') ""OcrCode"",IFNULL(t0.""OcrCode2"",'') ""OcrCode2""            
                                                FROM  WTR1 T0 INNER JOIN OWTR T1 On T0.""DocEntry""=T1.""DocEntry""
                                                INNER JOIN OITM T3 ON T0.""ItemCode""=T3.""ItemCode"" and T3.""ManBtchNum""='Y'
                                                INNER Join OITL T4 On T4.""DocEntry""=T0.""DocEntry"" And T4.""DocLine""=T0.""LineNum"" And T4.""DocType""=67 
                                                INNER JOIN ITL1 T5 ON T5.""LogEntry"" = T4.""LogEntry""
                                                INNER JOIN OBTN T6 ON  T6.""SysNumber"" = T5.""SysNumber"" AND T6.""ItemCode"" = T5.""ItemCode"" And T6.""AbsEntry""=T5.""MdAbsEntry""
                                                WHERE T1.""U_EXO_NUMPIC""='" & oRsLinPICK.Fields.Item("Traslado").Value.ToString & "' 
                                                And t1.U_EXO_LINPIC='" & oRsLinPICK.Fields.Item("LinTraslado").Value.ToString & "' And t5.""Quantity"" > 0
                                                Group by T6.""DistNumber""  ,t0.""OcrCode"",t0.""OcrCode2"" "
                                            oRsLote.DoQuery(sSQL)
                                            If oRsLote.RecordCount > 0 Then
                                                If oRsLote.Fields.Item("OcrCode").Value.ToString() <> "" Then
                                                    oDocFinal.Lines.DistributionRule = oRsLote.Fields.Item("OcrCode").Value.ToString()
                                                End If

                                                If oRsLote.Fields.Item("OcrCode2").Value.ToString() <> "" Then
                                                    oDocFinal.Lines.DistributionRule2 = oRsLote.Fields.Item("OcrCode2").Value.ToString()
                                                End If


                                                For iLote = 1 To oRsLote.RecordCount
                                                    'Creamos el lote de la línea del artículo
                                                    oDocFinal.Lines.BatchNumbers.BatchNumber = oRsLote.Fields.Item("DistNumber").Value.ToString
                                                    oDocFinal.Lines.BatchNumbers.Quantity = EXO_GLOBALES.DblTextToNumber(oobjGlobal.compañia, oRsLote.Fields.Item("Quantity").Value.ToString)
                                                    dCantidad += EXO_GLOBALES.DblTextToNumber(oobjGlobal.compañia, oRsLote.Fields.Item("Quantity").Value.ToString)
                                                    oDocFinal.Lines.BatchNumbers.Add()
                                                    oRsLote.MoveNext()
                                                Next
                                            End If
                                            If dCantidad = 0 Then
                                                oDocFinal.Lines.Quantity = oRsLinPICK.Fields.Item("Cantidad").Value
                                            Else
                                                oDocFinal.Lines.Quantity = dCantidad
                                            End If
#End Region
                                            oRsLinPICK.MoveNext()
                                        Next
                                        If oDocFinal.Add() <> 0 Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al generar el doc. de salida del pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Else
                                            oobjGlobal.compañia.GetNewObjectCode(sDocEntryFinal)
                                            sSQL = "SELECT ""DocNum"" FROM """ & oobjGlobal.compañia.CompanyDB & """.""ODLN"" WHERE ""DocEntry"" = " & sDocEntryFinal
                                            oRs.DoQuery(sSQL)
                                            If oRs.RecordCount > 0 Then
                                                sDocNumFinal = oRs.Fields.Item("DocNum").Value.ToString
                                                oobjGlobal.SBOApp.StatusBar.SetText("Entrega Nº: " & sDocNumFinal & " del Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

#Region "Preguntamos si dejamos abierto el pedido"
                                                Dim bCerrarPedido As Boolean = True
                                                ' Comprobamos si el pedido sigue abierto
                                                sSQL = "SELECT ""DocStatus"" FROM ORDR WHERE ""DocEntry""='" & oDocuments.DocEntry & "' "
                                                oRsPedido.DoQuery(sSQL)
                                                If oRsPedido.RecordCount > 0 Then
                                                    If oRsPedido.Fields.Item("DocStatus").Value.ToString <> "C" Then
                                                        If oobjGlobal.SBOApp.MessageBox("El Pedido sigue abierto. ¿Deseas cerrarlo?", 1, "Sí", "No") = 1 Then
                                                            bCerrarPedido = True
                                                        Else
                                                            bCerrarPedido = False
                                                        End If
                                                    End If
                                                End If
                                                If bCerrarPedido Then
#Region "Cierro Pedido"
                                                    oDocuments = Nothing
                                                    oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)

                                                    If oDocuments.GetByKey(CType(sDocEntry, Integer)) = True Then
                                                        oDocuments.Close()
                                                        If oDocuments.Close() <> 0 Then
                                                            oobjGlobal.SBOApp.StatusBar.SetText("El pedido " & sDocNum & " se ha cerrado con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                        End If
                                                    Else
                                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido " & sDocNum & " para cerrarlo.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    End If
                                                    sSQL = "UPDATE ORDR SET ""U_EXO_STATUSP""='C' WHERE ""DocEntry""=" & sDocEntry
                                                    If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                                        oobjGlobal.SBOApp.StatusBar.SetText("Actualizado Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                    Else
                                                        oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    End If
#End Region
#Region "Cierro Picking"
                                                    oRsLinPICK.MoveFirst()
                                                    oPicking = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists), SAPbobsCOM.PickLists)
                                                    If oPicking.GetByKey(CType(oRsLinPICK.Fields.Item("AbsEntry").Value, Integer)) = True Then
                                                        If oPicking.Close() <> 0 Then
                                                        Else
                                                            oobjGlobal.SBOApp.StatusBar.SetText("Picking " & oRsLinPICK.Fields.Item("AbsEntry").Value.ToString & " se ha cerrarlo con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                        End If

                                                    Else
                                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el picking " & oRsLinPICK.Fields.Item("AbsEntry").Value.ToString & " para cerrarlo.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    End If
#End Region
                                                Else
#Region "Cierro Picking"
                                                    oRsLinPICK.MoveFirst()
                                                    oPicking = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists), SAPbobsCOM.PickLists)
                                                    If oPicking.GetByKey(CType(oRsLinPICK.Fields.Item("AbsEntry").Value, Integer)) = True Then
                                                        If oPicking.Close() <> 0 Then
                                                        Else
                                                            oobjGlobal.SBOApp.StatusBar.SetText("Picking " & oRsLinPICK.Fields.Item("AbsEntry").Value.ToString & " se ha cerrarlo con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                        End If

                                                    Else
                                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el picking " & oRsLinPICK.Fields.Item("AbsEntry").Value.ToString & " para cerrarlo.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    End If
#End Region

#Region "Volvemos a poner las líneas a pdte"
                                                    sSQL = "UPDATE RDR1 SET ""PickStatus""='N' WHERE ""LineStatus""<>'C' and ""DocEntry""=" & sDocEntry
                                                    If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                                        oobjGlobal.SBOApp.StatusBar.SetText("Actualizado Líneas pdtes. Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                    Else
                                                        oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar Líneas pdtes. Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    End If
#End Region
                                                End If
#End Region
                                            Else
                                                sDocNumFinal = "0"
                                                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la entrega para el pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            End If
                                        End If
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para para generar el documento de salida con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra traslado del picking en el documento Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End If

                            Case "SOLTRA" ' Sol. de Traslado
                                oDocStockTransfer = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                                If oDocStockTransfer.GetByKey(CType(sDocEntry, Integer)) = True Then
                                    oDocStockTransfer_Lines = oDocStockTransfer.Lines
                                    oDocFinal = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer), SAPbobsCOM.StockTransfer)
                                    oDocFinal_Lines = oDocFinal.Lines
                                    For J = 0 To oDocStockTransfer_Lines.Count - 1
                                        If (J > 0) Then
                                            oDocFinal_Lines.Add()
                                        End If
                                        oDocStockTransfer_Lines.SetCurrentLine(J)
                                        oDocFinal_Lines.BaseObjectType = oDocStockTransfer.DocObjectCode
                                        oDocFinal_Lines.OrderEntry = oDocStockTransfer.DocEntry
                                        oDocFinal_Lines.OrderRowID = J
                                        oDocFinal_Lines.ReleasedQuantity = oDocStockTransfer_Lines.RemainingOpenQuantity
                                    Next
                                    If oDocFinal.Add() <> 0 Then
                                        oobjGlobal.SBOApp.StatusBar.SetText("Error al generar traslado de la Sol. de Traslado Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Else
                                        oobjGlobal.compañia.GetNewObjectCode(sDocEntryFinal)
                                        sSQL = "SELECT ""DocNum"" FROM """ & oobjGlobal.compañia.CompanyDB & """.""OWTR"" WHERE ""DocEntry"" = " & sDocEntryFinal
                                        oRs.DoQuery(sSQL)
                                        If oRs.RecordCount > 0 Then
                                            sSQL = "UPDATE OWTQ SET ""U_EXO_STATUSP""='C' WHERE ""DocEntry""=" & sDocEntry
                                            If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                                oobjGlobal.SBOApp.StatusBar.SetText("Actualizado Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            Else
                                                oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            End If
                                            sDocNumFinal = oRs.Fields.Item("DocNum").Value.ToString
                                            oobjGlobal.SBOApp.StatusBar.SetText("Traslado Nº: " & sDocNumFinal & " de la Sol. de Traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        Else
                                            sDocNumFinal = "0"
                                            oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el traslado generado para la Sol. de Traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If

                                    End If
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Traslado para generar el traslado con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            Case "SDPROV" ' Sol. de dev. 
                                sSQL = "UPDATE OPRR SET ""U_EXO_STATUSP""='C' WHERE ""DocEntry""=" & sDocEntry
                                If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                    oobjGlobal.SBOApp.StatusBar.SetText("Picking Liberado Sol. de dev. de Proveedor con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("Error en Picking Liberado Sol. de dev. de Proveedor con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                        End Select
#End Region
                    End If
                End If
            Next

            Gen_DOC = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLinPICK, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsPedido, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLote, Object))
            oDocFinal = Nothing : oDocFinal_Lines = Nothing
            oDocuments = Nothing : oDocument_Lines = Nothing
            oDocStockTransfer = Nothing : oDocStockTransfer_Lines = Nothing
        End Try
    End Function
    Public Shared Function CambiarClaseExpedicionCombo(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal i As Integer) As Boolean
        CambiarClaseExpedicionCombo = False
#Region "VARIABLES"
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sIC As String = "" : Dim sClaseExp As String = ""
        Dim oDocuments As SAPbobsCOM.Documents = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing
#End Region

        Try
            sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
            sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
            sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
            sIC = oForm.DataSources.DataTables.Item(sData).GetValue("CÓDIGO", i).ToString
            sClaseExp = oForm.DataSources.DataTables.Item(sData).GetValue("CLASE EXP.", i).ToString

            Select Case sTIPODOC
                Case "PEDVTA" ' Pedido de venta
                    oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                    If oDocuments.GetByKey(CType(sDocEntry, Integer)) = True Then
                        For i = 0 To oDocuments.Lines.Count - 1
                            oDocuments.Lines.SetCurrentLine(i)
                            oDocuments.Lines.ShippingMethod = CType(sClaseExp, Integer)
                        Next

                        If oDocuments.Update() <> 0 Then
                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar el pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        Else
                            oobjGlobal.SBOApp.StatusBar.SetText("Se ha actualizado la clase de expedición en el pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    Else
                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para cambiar la clase de expedición con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                Case "SOLTRA" ' Sol. de Traslado                           
                    oDocStockTransfer = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                    If oDocStockTransfer.GetByKey(sDocEntry) = True Then
                        oDocStockTransfer.UserFields.Fields.Item("U_EXO_CLASEE").Value = sClaseExp
                        If oDocStockTransfer.Update() <> 0 Then
                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar la Sol. de traslado Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        Else
                            oobjGlobal.SBOApp.StatusBar.SetText("Se ha actualizado la clase de expedición en la Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    Else
                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Traslado  Nº: " & sDocNum & ". No s epuede cerrar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                Case "SDPROV" ' Sol. de dev. de Proveedor
                    oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oGoodsReturnRequest), SAPbobsCOM.Documents)
                    If oDocuments.GetByKey(sDocEntry) = True Then
                        oDocuments.ShippingMethod = sClaseExp
                        If oDocuments.Update() <> 0 Then
                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar la clase de expedición de la Sol. de Dev de proveedor Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        Else
                            oobjGlobal.SBOApp.StatusBar.SetText("Sol. de Dev de proveedor actualizada Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    Else
                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Dev de proveedor para cerrarla con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
            End Select

            CambiarClaseExpedicionCombo = True
        Catch exCOM As Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oDocStockTransfer = Nothing
            oDocuments = Nothing
        End Try
    End Function
    Public Shared Function CambiarClaseExpedicionMasiva(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        CambiarClaseExpedicionMasiva = False

        Dim oRs As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sIC As String = "" : Dim sClaseExp As String = "" : Dim sClaseExpNew As String = ""
        Dim sAgenenClase As String = "" : Dim sAgenciaListaNegra As String = ""
        Dim oDocuments As SAPbobsCOM.Documents = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing


        Try
            Select Case sData
                Case "DTSPTE"
                    sClaseExpNew = oForm.DataSources.UserDataSources.Item("UDEXPCB").Value.ToString
                Case "DTSLIB"
                    sClaseExpNew = oForm.DataSources.UserDataSources.Item("UDEXPCBL").Value.ToString
            End Select

            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString

                    Select Case sTIPODOC
                        Case "PEDVTA" ' Pedido de venta
                            oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                            If oDocuments.GetByKey(CType(sDocEntry, Integer)) = True Then
                                For l = 0 To oDocuments.Lines.Count - 1
                                    oDocuments.Lines.SetCurrentLine(l)
                                    If oDocuments.Lines.ShippingMethod <> CType(sClaseExpNew, Integer) Then
                                        oDocuments.Lines.ShippingMethod = CType(sClaseExpNew, Integer)
                                    End If
                                Next
                                If oDocuments.Update() <> 0 Then
                                    oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar el pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("Se ha actualizado la clase de expedición en el pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para cambiar la clase de expedición con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        Case "SOLTRA" ' Sol. de Traslado                           
                            oDocStockTransfer = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                            If oDocStockTransfer.GetByKey(sDocEntry) = True Then
                                oDocStockTransfer.UserFields.Fields.Item("U_EXO_CLASEE").Value = sClaseExpNew
                                If oDocStockTransfer.Update() <> 0 Then
                                    oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar la Sol. de traslado Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("Se ha actualizado la clase de expedición en la Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Traslado  Nº: " & sDocNum & ". No s epuede cerrar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        Case "SDPROV" ' Sol. de dev. de Proveedor
                            oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oGoodsReturnRequest), SAPbobsCOM.Documents)
                            If oDocuments.GetByKey(sDocEntry) = True Then
                                oDocuments.TransportationCode = sClaseExpNew
                                If oDocuments.Update() <> 0 Then
                                    oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar la clase de expedición de la Sol. de Dev de proveedor Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("Sol. de Dev de proveedor actualizada Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Dev de proveedor para cerrarla con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                    End Select


                End If
            Next

            CambiarClaseExpedicionMasiva = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

            oDocStockTransfer = Nothing
            oDocuments = Nothing
        End Try
    End Function
    Public Shared Function CambiarClaseExpedicion(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        CambiarClaseExpedicion = False
#Region "VARIABLES"
        Dim oRs As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sIC As String = "" : Dim sClaseExp As String = ""
        Dim sAgenenClase As String = "" : Dim sAgenciaListaNegra As String = ""
        Dim oDocuments As SAPbobsCOM.Documents = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing
#End Region

        Try
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                    sIC = oForm.DataSources.DataTables.Item(sData).GetValue("CÓDIGO", i).ToString
                    sClaseExp = oForm.DataSources.DataTables.Item(sData).GetValue("CLASE EXP.", i).ToString
#Region "Comprobamos que la clase de expedicion sea permitida y no este en la lista negra"
                    sSQL = " SELECT ""U_EXO_AGE"" FROM OSHP WHERE ""TrnspCode""='" & sClaseExp & "' "
                    sAgenenClase = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    sSQL = " SELECT ""U_EXO_COD"" FROM ""@EXO_LNEGRAL"" WHERE ""Code""='" & sIC & "' and ""U_EXO_COD""='" & sAgenenClase & "' "
                    sAgenciaListaNegra = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                    Dim bActualiza As Boolean = True
                    If sAgenciaListaNegra <> "" Then
                        oobjGlobal.SBOApp.StatusBar.SetText("En el documento Nº: " & sDocNum & ", la clase de expedición tiene asignada la agencia """ & sAgenciaListaNegra & """ en la lista negra. No puede actualizarlo." & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        bActualiza = False
                    End If

                    If bActualiza = True Then
                        sSQL = " SELECT ""U_EXO_COD"" FROM ""@EXO_LNEGRAL"" WHERE ""Code""='" & sIC & "' and ""U_EXO_COD""='" & sAgenenClase & "' "
                        sAgenciaListaNegra = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                        If sAgenciaListaNegra <> "" Then
                            oobjGlobal.SBOApp.StatusBar.SetText("En el documento Nº: " & sDocNum & ", la clase de expedición tiene asignada la agencia """ & sAgenciaListaNegra & """ en la lista negra. No puede actualizarlo." & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Select Case sTIPODOC
                                Case "PEDVTA" ' Pedido de venta
                                    oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                                    If oDocuments.GetByKey(sDocEntry) = True Then
                                        oDocuments.TransportationCode = sClaseExp
                                        If oDocuments.Update() <> 0 Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar el pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Se ha actualizado la clase de expedición en el pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        End If
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para cambiar la clase de expedición con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                Case "SOLTRA" ' Sol. de Traslado                           
                                    oDocStockTransfer = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                                    If oDocStockTransfer.GetByKey(sDocEntry) = True Then
                                        oDocStockTransfer.UserFields.Fields.Item("U_EXO_CLASEE").Value = sClaseExp
                                        If oDocStockTransfer.Update() <> 0 Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar la Sol. de traslado Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Se ha actualizado la clase de expedición en la Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        End If
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Traslado  Nº: " & sDocNum & ". No s epuede cerrar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                Case "SDPROV" ' Sol. de dev. de Proveedor
                                    oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oGoodsReturnRequest), SAPbobsCOM.Documents)
                                    If oDocuments.GetByKey(sDocEntry) = True Then
                                        oDocuments.TransportationCode = sClaseExp
                                        If oDocuments.Update() <> 0 Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar la clase de expedición de la Sol. de Dev de proveedor Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Sol. de Dev de proveedor actualizada Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        End If
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Dev de proveedor para cerrarla con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                            End Select
                        End If
                    End If
#End Region
                End If
            Next

            CambiarClaseExpedicion = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

            oDocStockTransfer = Nothing
            oDocuments = Nothing
        End Try
    End Function
    Public Shared Function CerrarDocumentos(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        CerrarDocumentos = False
#Region "VARIABLES"
        Dim oRs As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim oDocuments As SAPbobsCOM.Documents = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing
#End Region

        Try
            If oobjGlobal.SBOApp.MessageBox("¿Está seguro que quiere cerrar los Documentos seleccionados?", 1, "Sí", "No") = 1 Then
                For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                    If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                        sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                        sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                        sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                        Select Case sTIPODOC
                            Case "PEDVTA" ' Pedido de venta
                                oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                                If oDocuments.GetByKey(sDocEntry) = True Then
                                    If oDocuments.Close() <> 0 Then
                                        oobjGlobal.SBOApp.StatusBar.SetText("Error al cerrar el pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("Pedido cerrado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para cerrarlo con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            Case "SOLTRA" ' Sol. de Traslado                           
                                oDocStockTransfer = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                                If oDocStockTransfer.GetByKey(sDocEntry) = True Then
                                    If oDocStockTransfer.Close() <> 0 Then
                                        oobjGlobal.SBOApp.StatusBar.SetText("Error al cerrar la Sol. de traslado Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("Sol. de traslado cerrada con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Traslado  Nº: " & sDocNum & ". No s epuede cerrar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            Case "SDPROV" ' Sol. de dev. de Proveedor
                                oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oGoodsReturnRequest), SAPbobsCOM.Documents)
                                If oDocuments.GetByKey(sDocEntry) = True Then
                                    If oDocuments.Close() <> 0 Then
                                        oobjGlobal.SBOApp.StatusBar.SetText("Error al cerrar la Sol. de Dev de proveedor Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("Sol. de Dev de proveedor cerrada Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If
                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Dev de proveedor para cerrarla con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                        End Select
                    End If
                Next
            End If
            CerrarDocumentos = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

            oDocStockTransfer = Nothing
            oDocuments = Nothing
        End Try
    End Function
    Public Shared Function LiberarparaPicking(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        LiberarparaPicking = False
#Region "VARIABLES"
        Dim oRs As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sDocEntryPicking As String = "" : Dim sDocNumPicking As String = ""
        Dim oPickLists As SAPbobsCOM.PickLists = Nothing : Dim oPickLists_Lines As SAPbobsCOM.PickLists_Lines = Nothing
        Dim oDocuments As SAPbobsCOM.Documents = Nothing : Dim oDocument_Lines As SAPbobsCOM.Document_Lines = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing : Dim oDocStockTransfer_Lines As SAPbobsCOM.StockTransfer_Lines = Nothing
        Dim sAlm As String = ""
#End Region

        Try
            sAlm = CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", i).ToString
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", i).ToString
                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", i).ToString
                    Select Case sTIPODOC

                        Case "PEDVTA" ' Pedido de venta
#Region "Pedido Venta"
                            oPickLists = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists), SAPbobsCOM.PickLists)
                            oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                            Dim oItems = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems), SAPbobsCOM.Items)
                            If oDocuments.GetByKey(sDocEntry) = True Then
                                Dim existsOneLine = False
                                oDocument_Lines = oDocuments.Lines
                                oPickLists = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists), SAPbobsCOM.PickLists)
                                oPickLists.PickDate = oDocuments.DocDate

                                For J = 0 To oDocument_Lines.Count - 1
                                    oDocument_Lines.SetCurrentLine(J)
                                    Dim sarticulo As String = oDocument_Lines.ItemCode
                                    If oDocument_Lines.WarehouseCode = sAlm Then
                                        sSQL = "SELECT ""OnHand"" FROM OITW Where ""ItemCode""='" & sarticulo & "' and ""WhsCode""='" & sAlm & "'"
                                        Dim dStock As Double = oobjGlobal.refDi.SQL.sqlNumericaB1(sSQL)
                                        If dStock > 0 And oDocument_Lines.LineStatus <> SAPbobsCOM.BoStatus.bost_Close Then
                                            'If dStock >= oDocument_Lines.Quantity And oDocument_Lines.LineStatus <> SAPbobsCOM.BoStatus.bost_Close Then
                                            oDocument_Lines.SetCurrentLine(J)
                                            oItems.GetByKey(oDocuments.Lines.ItemCode)
                                            If (oItems.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO) Then
                                                Continue For
                                            End If
                                            oPickLists.Lines.BaseObjectType = "17"
                                            oPickLists.Lines.OrderEntry = oDocuments.DocEntry
                                            oPickLists.Lines.OrderRowID = oDocument_Lines.LineNum
                                            oPickLists.Lines.ReleasedQuantity = If(dStock >= oDocument_Lines.Quantity, oDocument_Lines.Quantity, dStock)
                                            oPickLists.Lines.Add()

                                            existsOneLine = True
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Liberar Picking del pedido Nº: " & sDocNum & ". Artículo " & sarticulo & " en Stock " & dStock.ToString & " y en pedido " & oDocument_Lines.Quantity.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        End If
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("Liberar Picking del pedido Nº: " & sDocNum & ". Línea " & oDocument_Lines.LineNum & " - Artículo " & sarticulo & " con almacén " & oDocument_Lines.WarehouseCode & " no se tiene en cuenta.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If
                                Next

                                If (Not existsOneLine) Then
                                    oobjGlobal.SBOApp.StatusBar.SetText("No hay stock para liberar picking", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Continue For
                                End If

                                If oPickLists.Add() <> 0 Then
                                    oobjGlobal.SBOApp.StatusBar.SetText("Error al liberar Picking del pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    oobjGlobal.compañia.GetNewObjectCode(sDocEntryPicking)
                                    sSQL = "SELECT ""AbsEntry"" FROM """ & oobjGlobal.compañia.CompanyDB & """.""OPKL"" WHERE ""AbsEntry"" = " & sDocEntryPicking
                                    oRs.DoQuery(sSQL)
                                    If oRs.RecordCount > 0 Then
                                        sSQL = "UPDATE ORDR SET ""U_EXO_STATUSP""='L' WHERE ""DocEntry""=" & sDocEntry
                                        If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Actualizado Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                        sDocNumPicking = oRs.Fields.Item("AbsEntry").Value.ToString
                                        oobjGlobal.SBOApp.StatusBar.SetText("Picking Liberado Nº: " & sDocNumPicking & " del Pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    Else
                                        sDocNumPicking = "0"
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el Picking generado para el pedido Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If

                                End If
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para liberar Picking con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
#End Region
                        Case "SOLTRA" ' Sol. de Traslado
#Region "Sol. Traslado"
                            oPickLists = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists), SAPbobsCOM.PickLists)
                            oDocStockTransfer = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest), SAPbobsCOM.StockTransfer)
                            If oDocStockTransfer.GetByKey(sDocEntry) = True Then
                                oDocStockTransfer_Lines = oDocStockTransfer.Lines
                                oPickLists = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists), SAPbobsCOM.PickLists)

                                For J = 0 To oDocStockTransfer_Lines.Count - 1
                                    If oDocStockTransfer_Lines.FromWarehouseCode = sAlm Then
                                        oDocStockTransfer_Lines.SetCurrentLine(J)
                                        oPickLists.Lines.BaseObjectType = "1250000001"
                                        oPickLists.Lines.OrderEntry = oDocStockTransfer.DocEntry
                                        oPickLists.Lines.OrderRowID = J
                                        oPickLists.Lines.ReleasedQuantity = oDocStockTransfer_Lines.RemainingOpenQuantity
                                        oPickLists.Lines.Add()
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("Liberar Picking de la Sol. de Traslado Nº: " & sDocNum & ". Línea " & oDocStockTransfer_Lines.LineNum & " - Artículo " & oDocStockTransfer_Lines.ItemCode & " con almacén " & oDocStockTransfer_Lines.WarehouseCode & " no se tiene en cuenta.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If

                                Next

                                If oPickLists.Add() <> 0 Then
                                    oobjGlobal.SBOApp.StatusBar.SetText("Error al liberar Picking de la Sol. de Traslado Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    oDocStockTransfer.UserFields.Fields.Item("U_EXO_STATUSP").Value = "L"
                                    oDocStockTransfer.Update()

                                    oobjGlobal.compañia.GetNewObjectCode(sDocEntryPicking)
                                    sSQL = "SELECT ""AbsEntry"" FROM """ & oobjGlobal.compañia.CompanyDB & """.""OPKL"" WHERE ""AbsEntry"" = " & sDocEntryPicking
                                    oRs.DoQuery(sSQL)
                                    If oRs.RecordCount > 0 Then
                                        sSQL = "UPDATE OWTQ SET ""U_EXO_STATUSP""='L' WHERE ""DocEntry""=" & sDocEntry
                                        If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                            oobjGlobal.SBOApp.StatusBar.SetText("Actualizado Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("Error al actualizar Sol. de traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                        sDocNumPicking = oRs.Fields.Item("AbsEntry").Value.ToString
                                        oobjGlobal.SBOApp.StatusBar.SetText("Picking Liberado Nº: " & sDocNumPicking & " de la Sol. de Traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    Else
                                        sDocNumPicking = "0"
                                        oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el Picking generado para la Sol. de Traslado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                End If
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra la Sol. de Traslado para liberar Picking con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
#End Region
                        Case "SDPROV" ' Sol. de dev. de Proveedor
#Region "Sol. Dev. Proveedor"
                            sSQL = "UPDATE OPRR SET ""U_EXO_STATUSP""='L' WHERE ""DocEntry""=" & sDocEntry
                            If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                oobjGlobal.SBOApp.StatusBar.SetText("Picking Liberado Sol. de dev. de Proveedor con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("Error en Picking Liberado Sol. de dev. de Proveedor con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
#End Region

                    End Select
                End If
            Next

            LiberarparaPicking = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            oPickLists = Nothing : oPickLists_Lines = Nothing
            oDocuments = Nothing : oDocument_Lines = Nothing
        End Try
    End Function
    Public Shared Function CambiarAlmacen(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        CambiarAlmacen = False
#Region "VARIABLES"
        Dim oRs As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sTIPODOC As String = "" : Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim cardCode As String = "" : Dim deleg As String = ""
        Dim oDocuments As SAPbobsCOM.Documents = Nothing
        Dim oDocStockTransfer As SAPbobsCOM.StockTransfer = Nothing
        Dim bActualiza As Boolean = False
        Dim sDelPedido As String = "" : Dim sALMPedido As String = "" : Dim sALM As String = "" : Dim sDelALM As String = ""
        Dim sDocEntryTraslado As String = "" : Dim sDocNumTraslado As String = ""
        Dim grid = CType(oForm.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid)
#End Region

        Try
            If oobjGlobal.SBOApp.MessageBox("¿Está seguro de cambiar el almacén a los documentos seleccionados?", 1, "Sí", "No") = 1 Then
                For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                    If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", grid.GetDataTableRowIndex(i)).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                        sTIPODOC = oForm.DataSources.DataTables.Item(sData).GetValue("T. SALIDA", grid.GetDataTableRowIndex(i)).ToString
                        sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº INTERNO", grid.GetDataTableRowIndex(i)).ToString
                        sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº DOCUMENTO", grid.GetDataTableRowIndex(i)).ToString
                        cardCode = oForm.DataSources.DataTables.Item(sData).GetValue("CÓDIGO", grid.GetDataTableRowIndex(i)).ToString
                        deleg = oForm.DataSources.DataTables.Item(sData).GetValue("DELEGACIÓN", grid.GetDataTableRowIndex(i)).ToString

                        bActualiza = False
                        Select Case sTIPODOC
                            Case "PEDVTA" ' Pedido de venta
                                oDocuments = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                                If oDocuments.GetByKey(sDocEntry) = True Then

                                    For lin = 0 To oDocuments.Lines.Count - 1
                                        oDocuments.Lines.SetCurrentLine(lin)
                                        sSQL = "SELECT OWHS.""WhsCode"" FROM OCRD JOIN OWHS ON OCRD.""U_EXO_DELE"" = OWHS.""U_EXO_SUCURSAL"" WHERE OCRD.""CardCode"" = '" & cardCode & "'"
                                        sDelALM = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                                        oDocuments.Lines.WarehouseCode = sDelALM

                                        sSQL = "SELECT ""OcrCode"" FROm OOCR Where ""OcrName""='" & deleg & "' "
                                        Dim sOcrdCode = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                                        oDocuments.Lines.CostingCode = sOcrdCode
                                        oDocuments.Lines.ShippingMethod = -1
                                    Next

                                    If oDocuments.Update() <> 0 Then
                                        oobjGlobal.SBOApp.StatusBar.SetText("Error modificar  el pedido Nº: " & sDocNum & ". " & oobjGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("Pedido modificado Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If

                                Else
                                    oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el pedido para cambiar el almacén con Nº: " & sDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                        End Select
                    End If
                Next
            End If

            CambiarAlmacen = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

            oDocStockTransfer = Nothing
            oDocuments = Nothing
        End Try
    End Function
    Private Function ComprobarDOC(ByRef oForm As SAPbouiCOM.Form, ByVal sTABLA As String) As Boolean
        Dim bLineasSel As Boolean = False

        ComprobarDOC = False

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sTABLA).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sTABLA).GetValue("Sel", i).ToString = "Y" Then
                    bLineasSel = True
                    Exit For
                End If
            Next

            If bLineasSel = False Then
                objGlobal.SBOApp.MessageBox("Debe seleccionar al menos una línea.")
                Exit Function
            End If

            ComprobarDOC = bLineasSel

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ComprobarDOCPED(ByRef oForm As SAPbouiCOM.Form, ByVal sTABLA As String) As Boolean
        Dim bLineasSel As Boolean = False

        ComprobarDOCPED = False

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sTABLA).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sTABLA).GetValue("Sel", i).ToString = "Y" And oForm.DataSources.DataTables.Item(sTABLA).GetValue("T. SALIDA", i).ToString = "PEDVTA" Then
                    bLineasSel = True
                    Exit For
                End If
            Next

            If bLineasSel = False Then
                objGlobal.SBOApp.MessageBox("Debe seleccionar al menos una línea de pedido.")
                Exit Function
            End If

            ComprobarDOCPED = bLineasSel

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ComprobarDOCSEL(ByRef oForm As SAPbouiCOM.Form, ByVal sTABLA As String, ByRef dtDatos As System.Data.DataTable, ByRef dt As SAPbouiCOM.DataTable) As Boolean
        ComprobarDOCSEL = False

        Try
            For iCol As Integer = 0 To 12
                dtDatos.Columns.Add(dt.Columns.Item(iCol).Name)
            Next

            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sTABLA).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sTABLA).GetValue("Sel", i).ToString = "Y" Then
                    'Añadimos los registros
                    Dim oRow As DataRow = dtDatos.NewRow
                    For iCol As Integer = 0 To 12
                        oRow.Item(dt.Columns.Item(iCol).Name) = dt.Columns.Item(iCol).Cells.Item(i).Value
                    Next
                    dtDatos.Rows.Add(oRow)
                End If
            Next
            ComprobarDOCSEL = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ComprobarDOCSELLIB(ByRef oForm As SAPbouiCOM.Form, ByVal sTABLA As String, ByRef dtDatos As System.Data.DataTable, ByRef dt As SAPbouiCOM.DataTable) As System.Data.DataTable

        Try
            Dim dtDatosNoSelected As System.Data.DataTable = New System.Data.DataTable
            Dim cantSel = 0
            For iCol As Integer = 0 To dt.Columns.Count - 1
                dtDatos.Columns.Add(dt.Columns.Item(iCol).Name)
                dtDatosNoSelected.Columns.Add(dt.Columns.Item(iCol).Name)
            Next

            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sTABLA).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sTABLA).GetValue("Sel", i).ToString = "Y" Then
                    'Añadimos los registros
                    Dim oRow As DataRow = dtDatos.NewRow
                    For iCol As Integer = 0 To dt.Columns.Count - 1
                        oRow.Item(dt.Columns.Item(iCol).Name) = dt.Columns.Item(iCol).Cells.Item(i).Value
                    Next
                    dtDatos.Rows.Add(oRow)
                    cantSel += cantSel + 1
                Else
                    Dim oRow As DataRow = dtDatosNoSelected.NewRow
                    For iCol As Integer = 0 To dt.Columns.Count - 1
                        oRow.Item(dt.Columns.Item(iCol).Name) = dt.Columns.Item(iCol).Cells.Item(i).Value
                    Next
                    dtDatosNoSelected.Rows.Add(oRow)
                End If
            Next

            If (cantSel = 0) Then
                ComprobarDOCSELLIB = dtDatosNoSelected
            Else
                ComprobarDOCSELLIB = dtDatos
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ComprobarDOCENT(ByRef oForm As SAPbouiCOM.Form, ByVal sTABLA As String, ByRef dtDatos As System.Data.DataTable, ByRef dt As SAPbouiCOM.DataTable, ByRef sTipo As String) As System.Data.DataTable

        Try
            Dim dtDatosNoSelected As System.Data.DataTable = New System.Data.DataTable
            Dim cantSel = 0
            For iCol As Integer = 0 To dt.Columns.Count - 1
                dtDatos.Columns.Add(dt.Columns.Item(iCol).Name)
                dtDatosNoSelected.Columns.Add(dt.Columns.Item(iCol).Name)
            Next

            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sTABLA).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sTABLA).GetValue("Sel", i).ToString = "Y" And oForm.DataSources.DataTables.Item(sTABLA).GetValue("T. ENTRADA", i).ToString = sTipo Then
                    'Añadimos los registros
                    Dim oRow As DataRow = dtDatos.NewRow
                    For iCol As Integer = 0 To dt.Columns.Count - 1
                        oRow.Item(dt.Columns.Item(iCol).Name) = dt.Columns.Item(iCol).Cells.Item(i).Value
                    Next
                    dtDatos.Rows.Add(oRow)
                    cantSel += cantSel + 1
                Else
                    If oForm.DataSources.DataTables.Item(sTABLA).GetValue("T. ENTRADA", i).ToString = sTipo Then
                        Dim oRow As DataRow = dtDatosNoSelected.NewRow
                        For iCol As Integer = 0 To dt.Columns.Count - 1
                            oRow.Item(dt.Columns.Item(iCol).Name) = dt.Columns.Item(iCol).Cells.Item(i).Value
                        Next
                        dtDatosNoSelected.Rows.Add(oRow)
                        cantSel += cantSel + 1
                    End If
                End If
            Next

            If (cantSel = 0) Then
                ComprobarDOCENT = dtDatosNoSelected
            Else
                ComprobarDOCENT = dtDatos
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub Ini_Grid(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String
        Try
            oForm.Freeze(True)
#Region "Pdte"
            sSQL = "SELECT CAST('' as nVARCHAR(50)) ""T. SALIDA"", CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", CAST('' as nVARCHAR(50)) ""Nº INTERNO"", CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
            sSQL &= " 'N' ""AUTORIZADO"", CAST('' as nVARCHAR(50)) ""CÓDIGO"",  CAST('' as nVARCHAR(150))	""EMPRESA"", CAST('' as nVARCHAR(50)) ""CLASE EXP."", 'N' ""ROT. STOCK"", "
            sSQL &= " 'N' ""A"", CAST('' as nVARCHAR(50)) ""UBICACIÓN"", CAST('' as nVARCHAR(50)) ""ZONA TRANSPORTE"", 'N' ""Sel"" "
            sSQL &= "FROM DUMMY "
            oForm.DataSources.DataTables.Item("DTSPTE").ExecuteQuery(sSQL)
            FormateaGrid_SPTE(oForm)
            'oForm.DataSources.DataTables.Item("DTSPTE").Rows.Clear()
#End Region
#Region "LIB"
            sSQL = "SELECT CAST('' as nVARCHAR(50)) ""T. SALIDA"", CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", CAST('' as nVARCHAR(50)) ""Nº INTERNO"", CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
            sSQL &= " CAST('' as nVARCHAR(50)) ""CÓDIGO"",  CAST('' as nVARCHAR(150))	""EMPRESA"", CAST('' as nVARCHAR(50)) ""CLASE EXP."", 'N' ""ROT. STOCK"", "
            sSQL &= " 'N' ""A"", CAST('' as nVARCHAR(50)) ""UBICACIÓN"", CAST('' as nVARCHAR(50)) ""ZONA TRANSPORTE"", 0 ""Cant."", 0 ""Cant. Pdte."", 'N' ""Sel"" "
            sSQL &= "FROM DUMMY "
            oForm.DataSources.DataTables.Item("DTSLIB").ExecuteQuery(sSQL)
            FormateaGrid_SLIB(oForm)
            'oForm.DataSources.DataTables.Item("DTSLIB").Rows.Clear()
#End Region
#Region "COM"
            sSQL = "SELECT CAST('' as nVARCHAR(50)) ""T. SALIDA"", CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", CAST('' as nVARCHAR(50)) ""Nº INTERNO"", CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
            sSQL &= " CAST('' as nVARCHAR(50)) ""CÓDIGO"",  CAST('' as nVARCHAR(150))	""EMPRESA"", CAST('' as nVARCHAR(50)) ""CLASE EXP."", CAST('' as nVARCHAR(50)) ""AG. TRANSPORTE"",  "
            sSQL &= " 'PE' ""ESTADO"", 'N' ""Sel"" "
            sSQL &= "FROM DUMMY "
            oForm.DataSources.DataTables.Item("DTSCOM").ExecuteQuery(sSQL)
            FormateaGrid_SCOM(oForm)
            'oForm.DataSources.DataTables.Item("DTSCOM").Rows.Clear
#End Region
#Region "ENT"
            sSQL = "SELECT CAST('' as nVARCHAR(50)) ""T. ENTRADA"", CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", CAST('' as nVARCHAR(50)) ""Nº INTERNO"", CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
            sSQL &= " CAST('' as nVARCHAR(50)) ""CÓDIGO"",  CAST('' as nVARCHAR(150))	""EMPRESA"", CAST('' as nVARCHAR(50)) ""ESTADO"", CAST('' as nVARCHAR(50)) ""DOC. ENTRADA"", "
            sSQL &= " CAST('' as nVARCHAR(50)) ""ID DOC. ENTRADA"""
            sSQL &= "FROM DUMMY "
            oForm.DataSources.DataTables.Item("DTE").ExecuteQuery(sSQL)
            FormateaGrid_E(oForm)
            ' oForm.DataSources.DataTables.Item("DTE").Rows.Clear()
#End Region

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)

        End Try
    End Sub
    Private Function SetComboExp(ByRef oForm As SAPbouiCOM.Form)
        Dim whs = oForm.DataSources.UserDataSources.Item("UDALM").Value
        Dim sSQL = " SELECT CAST(""TrnspCode"" as NVARCHAR(50)) ,""TrnspName"" "
        sSQL &= " From OWHS  JOIN OSHP ON OWHS.""U_EXO_SUCURSAL"" = OSHP.""U_EXO_DELEGA"""
        sSQL &= " WHERE OWHS.""WhsCode"" = '" & whs & "'"
        sSQL &= " ORDER By  ""TrnspName"" "

        Try
            Dim oColumnCb = CType(oForm.Items.Item("cbEXPCB").Specific, SAPbouiCOM.ComboBox)
            objGlobal.funcionesUI.cargaCombo(oColumnCb.ValidValues, sSQL)
            oColumnCb.ValidValues.Add("-1", " ")
            oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
            oColumnCb.Select("-1", BoSearchKey.psk_ByValue)
        Catch ex As Exception

        End Try

        Try
            Dim oColumnCb = CType(oForm.Items.Item("cbEXPCBL").Specific, SAPbouiCOM.ComboBox)
            objGlobal.funcionesUI.cargaCombo(oColumnCb.ValidValues, sSQL)
            oColumnCb.ValidValues.Add("-1", " ")
            oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
            oColumnCb.Select("-1", BoSearchKey.psk_ByValue)
        Catch ex As Exception

        End Try
    End Function
    Private Sub FiltrarPDTE(ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim sSalidas As String = ""
        Dim sICD As String = "" : Dim sICH As String = ""
        Dim sEXPE As String = "" : Dim sTerri As String = ""
        Dim sSQL As String = ""
#End Region
        Try
            sSalidas = oForm.DataSources.UserDataSources.Item("UDSAL").Value.ToString
            sICD = oForm.DataSources.UserDataSources.Item("UDICD").Value.ToString
            sICH = oForm.DataSources.UserDataSources.Item("UDICH").Value.ToString
            sEXPE = oForm.DataSources.UserDataSources.Item("UDEXPE").Value.ToString
            sTerri = oForm.DataSources.UserDataSources.Item("UDTERRI").Value.ToString
            oForm.Freeze(True)
            Select Case sSalidas
                Case "-"
                    sSQL = "SELECT 
                            CAST('' as nVARCHAR(50)) ""T. SALIDA"", 
                            CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", 
                            CAST('' as nVARCHAR(50)) ""FECHA CREACION"", 
                            CAST('' as nVARCHAR(50)) ""FECHA ENTREGA"", 
                            CAST('' as nVARCHAR(50)) ""Nº INTERNO"", 
                            CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " 
                            ""COM"", 
                            CAST('' as nVARCHAR(50)) ""CÓDIGO"",  
                            CAST('' as nVARCHAR(150))	""EMPRESA"", 
                            CAST('' as nVARCHAR(50)) ""CLASE EXP."", 
                            'N' ""R.ST."", "
                    sSQL &= " 
                            'N' ""A"", 
                            CAST('' as nVARCHAR(50)) ""UBICACIÓN"", 
                            CAST('' as nVARCHAR(50)) ""ZONA TRANSPORTE"", 
                            'N' ""Sel"",
                            ""-1"" AS ""ObjectType"""
                    sSQL &= "FROM DUMMY "
                Case "TODOS"
#Region "Todos"
                    sSQL = "Select * FROM ("
                    sSQL &= "Select ""T. SALIDA"", ""DELEGACIÓN"", ""FECHA CREACION"", ""FECHA ENTREGA"", ""Nº INTERNO"", ""Nº DOCUMENTO"", ""COM"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", "
                    sSQL &= " ""ROT. STOCK"" As ""R.ST."", (CASE WHEN ""A"" = 'Y' THEN 'Sí' ELSE 'No' END) AS ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel"", '17' AS ""ObjectType""  FROM ""EXO_PEDIDOS_VENTA"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " And ""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""CLASE EXP.""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT ""T. SALIDA"", ""DELEGACIÓN"", ""FECHA CREACION"", ""FECHA ENTREGA"", ""Nº INTERNO"", ""Nº DOCUMENTO"", ""COM"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", ""ROT. STOCK"" AS ""R.ST."", "
                    sSQL &= " (CASE WHEN ""A"" = 'Y' THEN 'Sí' ELSE 'No' END) AS ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel"", '1250000001' AS ""ObjectType"" FROM ""EXO_SOL_TRASLADO"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and ""FromWhsCod""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT ""T. SALIDA"", ""DELEGACIÓN"", ""FECHA CREACION"", ""FECHA ENTREGA"", ""Nº INTERNO"", ""Nº DOCUMENTO"", ""COM"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", ""ROT. STOCK"" AS ""R.ST."", "
                    sSQL &= " (CASE WHEN ""A"" = 'Y' THEN 'Sí' ELSE 'No' END) AS ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel"", '234000032' AS ""ObjectType"" FROM ""EXO_SOL_DEVOLUCION"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and ""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""CLASE EXP.""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= ") Z ORDER BY ""T. SALIDA"", ""Nº DOCUMENTO"" "
#End Region
                Case "PEDVTA"
#Region "Pedidos de Ventas"
                    sSQL = "SELECT ""T. SALIDA"", ""DELEGACIÓN"", ""FECHA CREACION"", ""FECHA ENTREGA"", ""Nº INTERNO"", ""Nº DOCUMENTO"", ""COM"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", "
                    sSQL &= " ""ROT. STOCK"" AS ""R.ST."", (CASE WHEN ""A"" = 'Y' THEN 'Sí' ELSE 'No' END) AS ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel"", '17' AS ""ObjectType""  FROM ""EXO_PEDIDOS_VENTA"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and ""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""CLASE EXP.""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " ORDER BY ""T. SALIDA"", ""Nº DOCUMENTO"" "
#End Region
                Case "SOLTRA"
#Region "Sol de traslado"
                    sSQL = "SELECT ""T. SALIDA"", ""DELEGACIÓN"", ""FECHA CREACION"", ""FECHA ENTREGA"", ""Nº INTERNO"", ""Nº DOCUMENTO"", ""COM"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", ""ROT. STOCK"" AS ""R.ST."", "
                    sSQL &= " (CASE WHEN ""A"" = 'Y' THEN 'Sí' ELSE 'No' END) AS ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel"", '1250000001' AS ""ObjectType"" FROM ""EXO_SOL_TRASLADO"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and ""FromWhsCod""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""CLASE EXP.""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " ORDER BY ""T. SALIDA"", ""Nº DOCUMENTO"" "
#End Region
                Case "SDPROV"
#Region "Sol de Devolución"
                    sSQL = "SELECT ""T. SALIDA"", ""DELEGACIÓN"", ""FECHA CREACION"", ""FECHA ENTREGA"", ""Nº INTERNO"", ""Nº DOCUMENTO"",  ""COM"", ""CÓDIGO"",  ""EMPRESA"", ""CLASE EXP."", ""ROT. STOCK"" AS ""R.ST."", "
                    sSQL &= " (CASE WHEN ""A"" = 'Y' THEN 'Sí' ELSE 'No' END) AS ""A"", ""UBICACIÓN"", ""ZONA TRANSPORTE"", ""Sel"", '234000032' AS ""ObjectType"" FROM ""EXO_SOL_DEVOLUCION"" "
                    sSQL &= " WHERE 1=1 "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and ""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' and ""CÓDIGO""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (""CÓDIGO"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (""CÓDIGO""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (""CLASE EXP.""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " ORDER BY ""T. SALIDA"", ""Nº DOCUMENTO"" "
#End Region
            End Select
            oForm.DataSources.DataTables.Item("DTSPTE").ExecuteQuery(sSQL)
            FormateaGrid_SPTE(oForm)

            objGlobal.SBOApp.StatusBar.SetText("Datos de salidas Pdtes. Cargados con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub FiltrarLIB(ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim sSalidas As String = ""
        Dim sICD As String = "" : Dim sICH As String = ""
        Dim sEXPE As String = "" : Dim sTerri As String = ""
        Dim sSQL As String = ""
#End Region
        Try
            sSalidas = oForm.DataSources.UserDataSources.Item("UDSAL").Value.ToString
            sICD = oForm.DataSources.UserDataSources.Item("UDICD").Value.ToString
            sICH = oForm.DataSources.UserDataSources.Item("UDICH").Value.ToString
            sEXPE = oForm.DataSources.UserDataSources.Item("UDEXPE").Value.ToString
            sTerri = oForm.DataSources.UserDataSources.Item("UDTERRI").Value.ToString
            oForm.Freeze(True)
            Select Case sSalidas
                Case "-"
                    sSQL = "SELECT CAST('' as nVARCHAR(50)) ""T. SALIDA"", CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", CAST('' as nVARCHAR(50)) ""Nº INTERNO"", CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST('' as nVARCHAR(50)) ""CÓDIGO"",  CAST('' as nVARCHAR(150))	""EMPRESA"", CAST('' as nVARCHAR(50)) ""CLASE EXP."", 'N' ""ROT. STOCK"", "
                    sSQL &= " 'N' ""A"", CAST('' as nVARCHAR(50)) ""UBICACIÓN"", CAST('' as nVARCHAR(50)) ""ZONA TRANSPORTE"", 0 ""Cant."", 0 ""Cant. Pdte."", CAST('' as nVARCHAR(50)) ""Picking"", CAST('' as NVARCHAR(50)) ""Usuario Picking"", 'N' ""Sel"" "
                    sSQL &= "FROM DUMMY "
                Case "TODOS"
#Region "Todos"
                    sSQL = "SELECT * FROM ( "
                    sSQL &= "SELECT DISTINCT CAST('PEDVTA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(TL.""TrnsCode"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"",  "
                    sSQL &= " IFNULL(PK.""Cant."",0) ""Cant."", IFNULL(PK.""Cant."" - TR.""Cant. T"",0) ""Cant. Pdte."", PK.""Picking"", 
                              (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') FROM (SELECT X1.""USER_CODE"" from PKL1 X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" Where X0.""AbsEntry"" = TL.""PickIdNo"" Group by X1.""USER_CODE"") Y0)  AS ""Usuario Picking"", 
                            'N' ""Sel"" "
                    sSQL &= "FROM ORDR T0 "
                    sSQL &= " LEFT JOIN RDR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN ""VEXO_PICKING"" PK ON PK.""BaseObject""= T0.""ObjType"" and PK.""OrderEntry""= TL.""DocEntry""  "
                    sSQL &= " LEFT JOIN ""VEXO_TRASLADOS"" TR ON TR.""BaseObject""= T0.""ObjType"" and TR.""OrderEntry""= TL.""DocEntry"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and T0.""Confirmed""='Y' and TL.""PickStatus""<>'N' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT DISTINCT CAST('SOLTRA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", "
                    sSQL &= " IFNULL(PK.""Cant."",0) ""Cant."", IFNULL(PK.""Cant."" - TR.""Cant. T"",0) ""Cant. Pdte."",  PK.""Picking"", 
                              (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') FROM (SELECT X1.""USER_CODE"" from PKL1 X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" Where X0.""AbsEntry"" = TL.""PickIdNo"" Group by X1.""USER_CODE"") Y0) AS ""Usuario Picking"",
                            'N' ""Sel""  "
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN ""VEXO_PICKING"" PK ON PK.""BaseObject""= T0.""ObjType"" and PK.""OrderEntry""= TL.""DocEntry""  "
                    sSQL &= " LEFT JOIN ""VEXO_TRASLADOS"" TR ON TR.""BaseObject""= T0.""ObjType"" and TR.""OrderEntry""= TL.""DocEntry"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and T0.""U_EXO_STATUSP""='L' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""FromWhsCod""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT DISTINCT CAST('SDPROV' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", "
                    sSQL &= " IFNULL(TL.""Quantity"",0) ""Cant."", 0 ""Cant. Pdte."",0 ""Picking"", 
                            (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') FROM (SELECT X1.""USER_CODE"" from PKL1 X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" Where X0.""AbsEntry"" = TL.""PickIdNo"" Group by X1.""USER_CODE"") Y0) AS ""Usuario Picking"",
                            'N' ""Sel"" "
                    sSQL &= "FROM OPRR T0 "
                    sSQL &= " LEFT JOIN PRR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and T0.""Confirmed""='Y' and T0.""U_EXO_STATUSP""='L' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= ") Z ORDER BY ""T. SALIDA"", ""Nº DOCUMENTO"" "
#End Region
                Case "PEDVTA"
#Region "Pedidos de Ventas"
                    sSQL = "SELECT DISTINCT CAST('PEDVTA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(TL.""TrnsCode"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", "
                    sSQL &= " IFNULL(PK.""Cant."",0) ""Cant."", IFNULL(PK.""Cant."" - TR.""Cant. T"",0) ""Cant. Pdte."",  PK.""Picking"", 
                              (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') FROM (SELECT X1.""USER_CODE"" from PKL1 X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" Where X0.""AbsEntry"" = TL.""PickIdNo"" Group by X1.""USER_CODE"") Y0)  AS ""Usuario Picking"",
                              'N' ""Sel"" "
                    sSQL &= "FROM ORDR T0 "
                    sSQL &= " LEFT JOIN RDR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN ""VEXO_PICKING"" PK ON PK.""BaseObject""= T0.""ObjType"" and PK.""OrderEntry""= TL.""DocEntry"" "
                    sSQL &= " LEFT JOIN ""VEXO_TRASLADOS"" TR ON TR.""BaseObject""= T0.""ObjType"" and TR.""OrderEntry""= TL.""DocEntry"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and T0.""Confirmed""='Y' and TL.""PickStatus""<>'N' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " ORDER BY ""T. SALIDA"", ""Nº DOCUMENTO"" "
#End Region
                Case "SOLTRA"
#Region "Sol de traslado"
                    sSQL = "SELECT DISTINCT CAST('SOLTRA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", "
                    sSQL &= " IFNULL(PK.""Cant."",0) ""Cant."", IFNULL(PK.""Cant."" - TR.""Cant. T"",0) ""Cant. Pdte."",  PK.""Picking"", 
                            (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') FROM (SELECT X1.""USER_CODE"" from PKL1 X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" Where X0.""AbsEntry"" = TL.""PickIdNo"" Group by X1.""USER_CODE"") Y0) AS ""Usuario Picking"",
                            'N' ""Sel"" "
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN ""VEXO_PICKING"" PK ON PK.""BaseObject""= T0.""ObjType"" and PK.""OrderEntry""= TL.""DocEntry""  "
                    sSQL &= " LEFT JOIN ""VEXO_TRASLADOS"" TR ON TR.""BaseObject""= T0.""ObjType"" and TR.""OrderEntry""= TL.""DocEntry"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and  T0.""U_EXO_TIPO"" = 'ITC' AND T0.""U_EXO_STATUSP""='L' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""FromWhsCod""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " ORDER BY ""T. SALIDA"", ""Nº DOCUMENTO"" "
#End Region
                Case "SDPROV"
#Region "Sol de Devolución"
                    sSQL = "SELECT DISTINCT CAST('SDPROV' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", "
                    sSQL &= " ifnull(R.""ROTURA"",'N') ""ROT. STOCK"", "
                    sSQL &= " IFNULL(A.""A"",'N') ""A"", CAST(IFNULL(S.""Sit"",'SIN SITUACIÓN') as nVARCHAR(50)) ""UBICACIÓN"", CAST(TT.""descript"" as nVARCHAR(50)) ""ZONA TRANSPORTE"", "
                    sSQL &= " IFNULL(TL.""Quantity"",0) ""Cant."", 0 ""Cant. Pdte."", 0 ""Picking"", (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') FROM (SELECT X1.""USER_CODE"" from PKL1 X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" Where X0.""AbsEntry"" = TL.""PickIdNo"" Group by X1.""USER_CODE"") Y0) AS ""Usuario Picking"", 'N' ""Sel""  "
                    sSQL &= "FROM OPRR T0 "
                    sSQL &= " LEFT JOIN PRR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " WHERE TL.""LineStatus""='O' and T0.""Confirmed""='Y' and T0.""U_EXO_STATUSP""='L' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " ORDER BY ""T. SALIDA"", ""Nº DOCUMENTO"" "
#End Region
            End Select
            oForm.DataSources.DataTables.Item("DTSLIB").ExecuteQuery(sSQL)
            FormateaGrid_SLIB(oForm)

            objGlobal.SBOApp.StatusBar.SetText("Datos de salida liberados Cargados con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub FiltrarCOM(ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim sSalidas As String = ""
        Dim sICD As String = "" : Dim sICH As String = ""
        Dim sEXPE As String = "" : Dim sTerri As String = ""
        Dim sSQL As String = ""
#End Region
        Try
            sSalidas = oForm.DataSources.UserDataSources.Item("UDSAL").Value.ToString
            sICD = oForm.DataSources.UserDataSources.Item("UDICD").Value.ToString
            sICH = oForm.DataSources.UserDataSources.Item("UDICH").Value.ToString
            sEXPE = oForm.DataSources.UserDataSources.Item("UDEXPE").Value.ToString
            sTerri = oForm.DataSources.UserDataSources.Item("UDTERRI").Value.ToString
            oForm.Freeze(True)
            Select Case sSalidas
                Case "-"
                    sSQL = "SELECT CAST('' as nVARCHAR(50)) ""T. SALIDA"", CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", CAST('' as nVARCHAR(50)) ""Nº INTERNO"", CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST('' as nVARCHAR(50)) ""CÓDIGO"",  CAST('' as nVARCHAR(150))	""EMPRESA"", CAST('' as nVARCHAR(50)) ""CLASE EXP."", CAST('' as nVARCHAR(50)) ""AG. TRANSPORTE"",  "
                    sSQL &= " 'PP' ""ESTADO"", CAST('' AS NVARCHAR(50)) AS ""CODUSR"", 'N' ""Sel"" "
                    sSQL &= "FROM DUMMY "
                Case "TODOS"
#Region "Todos"
                    sSQL = "SELECT * FROM ("
                    sSQL &= " SELECT DISTINCT CAST('ALBVTA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", 
                              CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(TL.""TrnsCode"" as nVARCHAR(50)) ""CLASE EXP."", 
                                IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"",  "
                    sSQL &= " EST.""Estado"" ""ESTADO"", T0.""U_EXO_LSTEMB"" ""List. Embalaje"", 
                                (SELECT STRING_AGG(IFNULL(""USR"", ''),'-') 
                                    FROM (SELECT Distinct IFNULL(OUSR.""USER_CODE"", '') AS ""USR"" 
                                            from ""@EXO_LSTEMBL"" X0 LEFT JOIN OUSR ON X0.""U_EXO_CODUSU"" = OUSR.""USERID"" 
                                           Where X0.""DocEntry"" = E.""DocEntry"" Group by IFNULL(OUSR.""USER_CODE"", '') ) Y0) AS ""CODUSR"", 'N' ""Sel"" "
                    sSQL &= " FROM ODLN T0 "
                    sSQL &= " LEFT JOIN DLN1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=TL.""TrnsCode"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON  E.""DocEntry"" = T0.""U_EXO_LSTEMB"" "
                    sSQL &= " LEFT JOIN ""VEXO_PARRILLA_ESTADO_SALCOMP"" EST ON EST.""ORIGEN""='ALBVTA' and EST.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN ( Select ""DocEntry"", ""Status"" FROM ""@EXO_LSTEMB"" WHERE ""Canceled""='N' ) ""HeadPack"" ON E.""DocEntry"" = ""HeadPack"".""DocEntry"" "
                    sSQL &= " WHERE  T0.""CANCELED"" = 'N' AND T0.""U_EXO_STATUSP""='C' and (T0.""U_EXO_ESTPAC""='Pendiente' or T0.""U_EXO_ESTPAC""='En curso') and (IFNULL(""HeadPack"".""Status"",'') = 'O' OR IFNULL(""HeadPack"".""Status"",'') = '') "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (CAST(TL.""TrnsCode"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= " SELECT DISTINCT CAST('SOLTRA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50)) ""CLASE EXP."", 
                                IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"", "
                    sSQL &= " EST.""Estado"" ""ESTADO"", E.""DocEntry"" ""List. Embalaje"", (SELECT STRING_AGG(IFNULL(""USR"", ''),'-') FROM (SELECT Distinct IFNULL(OUSR.""USER_CODE"", '') AS ""USR"" from ""@EXO_LSTEMBL"" X0 LEFT JOIN OUSR ON X0.""U_EXO_CODUSU"" = OUSR.""USERID"" Where X0.""DocEntry"" = E.""DocEntry"" Group by IFNULL(OUSR.""USER_CODE"", '') ) Y0) AS ""CODUSR"", 'N' ""Sel"" "
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=T0.""U_EXO_CLASEE"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON  E.""DocEntry"" = T0.""U_EXO_LSTEMB"" "
                    sSQL &= " LEFT JOIN ""VEXO_PARRILLA_ESTADO_SALCOMP"" EST ON EST.""ORIGEN""='SOLTRA' and EST.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN ( SELECT ""DocEntry"", ""Status"" FROM ""@EXO_LSTEMB"" WHERE ""Canceled""='N' ) ""HeadPack"" ON E.""DocEntry"" = ""HeadPack"".""DocEntry"" "
                    sSQL &= " LEFT JOIN OTER TT On T1.""Territory""=TT.""territryID"" "
                    sSQL &= " WHERE T0.""CANCELED"" = 'N' AND T0.""U_EXO_TIPO"" = 'ITC' AND T0.""U_EXO_STATUSP""='C' and (IFNULL(""HeadPack"".""Status"",'') = 'O' OR IFNULL(""HeadPack"".""Status"",'') = '') "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""FromWhsCod""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= " SELECT DISTINCT CAST('DPROV' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", 
                                IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"",  "
                    sSQL &= " EST.""Estado"" ""ESTADO"", E.""DocEntry"" ""List. Embalaje"", (SELECT STRING_AGG(IFNULL(""USR"", ''),'-') FROM (SELECT Distinct IFNULL(OUSR.""USER_CODE"", '') AS ""USR"" from ""@EXO_LSTEMBL"" X0 LEFT JOIN OUSR ON X0.""U_EXO_CODUSU"" = OUSR.""USERID"" Where X0.""DocEntry"" = E.""DocEntry"" Group by IFNULL(OUSR.""USER_CODE"", '') ) Y0) AS ""CODUSR"",  'N' ""Sel"" "
                    sSQL &= " FROM ORPD T0 "
                    sSQL &= " LEFT JOIN RPD1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=TL.""TrnsCode"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON  E.""DocEntry"" = T0.""U_EXO_LSTEMB"" "
                    sSQL &= " LEFT JOIN ""VEXO_PARRILLA_ESTADO_SALCOMP"" EST ON EST.""ORIGEN""='DPROV' and EST.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN ( SELECT ""DocEntry"", ""Status"" FROM ""@EXO_LSTEMB"" WHERE ""Canceled""='N' ) ""HeadPack"" ON E.""DocEntry"" = ""HeadPack"".""DocEntry"" "
                    sSQL &= " WHERE T0.""CANCELED"" = 'N' AND T0.""U_EXO_STATUSP""='C' and (T0.""U_EXO_ESTPAC""='Pendiente' or T0.""U_EXO_ESTPAC""='En curso') and (IFNULL(""HeadPack"".""Status"",'') = 'O' OR IFNULL(""HeadPack"".""Status"",'') = '') "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= ") Z ORDER BY ""T. SALIDA"", ""Nº DOCUMENTO"" "
#End Region
                Case "PEDVTA"
#Region "Entregas de Ventas"
                    sSQL = " SELECT DISTINCT CAST('ALBVTA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(TL.""TrnsCode"" as nVARCHAR(50)) ""CLASE EXP."", IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"",  "
                    sSQL &= " EST.""Estado"" ""ESTADO"", T0.""U_EXO_LSTEMB"" ""List. Embalaje"", (SELECT STRING_AGG(IFNULL(""USR"", ''),'-') FROM (SELECT Distinct IFNULL(OUSR.""USER_CODE"", '') AS ""USR"" from ""@EXO_LSTEMBL"" X0 LEFT JOIN OUSR ON X0.""U_EXO_CODUSU"" = OUSR.""USERID"" Where X0.""DocEntry"" = E.""DocEntry"" Group by IFNULL(OUSR.""USER_CODE"", '') ) Y0) AS ""CODUSR"", 'N' ""Sel"" "
                    sSQL &= "FROM ODLN T0 "
                    sSQL &= " LEFT JOIN DLN1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=TL.""TrnsCode"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON  E.""DocEntry"" = T0.""U_EXO_LSTEMB"" "
                    sSQL &= " LEFT JOIN ""VEXO_PARRILLA_ESTADO_SALCOMP"" EST ON EST.""ORIGEN""='ALBVTA' and EST.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN ( SELECT ""DocEntry"", ""Status"" FROM ""@EXO_LSTEMB"" WHERE ""Canceled""='N' ) ""HeadPack"" ON E.""DocEntry"" = ""HeadPack"".""DocEntry"" "
                    sSQL &= " WHERE  T0.""CANCELED"" = 'N' AND T0.""U_EXO_STATUSP""='C' and (T0.""U_EXO_ESTPAC""='Pendiente' or T0.""U_EXO_ESTPAC""='En curso') and (IFNULL(""HeadPack"".""Status"",'') = 'O' OR IFNULL(""HeadPack"".""Status"",'') = '') "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (CAST(TL.""TrnsCode"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If


#End Region
                Case "SOLTRA"
#Region "Sol de traslado"
                    sSQL = "SELECT DISTINCT CAST('SOLTRA' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50)) ""CLASE EXP."", IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"", "
                    sSQL &= " EST.""Estado""  ""ESTADO"", E.""DocEntry"" ""List. Embalaje"", (SELECT STRING_AGG(IFNULL(""USR"", ''),'-') FROM (SELECT Distinct IFNULL(OUSR.""USER_CODE"", '') AS ""USR"" from ""@EXO_LSTEMBL"" X0 LEFT JOIN OUSR ON X0.""U_EXO_CODUSU"" = OUSR.""USERID"" Where X0.""DocEntry"" = E.""DocEntry"" Group by IFNULL(OUSR.""USER_CODE"", '') ) Y0) AS ""CODUSR"", 'N' ""Sel"" "
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=T0.""U_EXO_CLASEE"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON  E.""DocEntry"" = T0.""U_EXO_LSTEMB"" "
                    sSQL &= " LEFT JOIN ""VEXO_PARRILLA_ESTADO_SALCOMP"" EST ON EST.""ORIGEN""='SOLTRA' and EST.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN ( SELECT ""DocEntry"", ""Status"" FROM ""@EXO_LSTEMB"" WHERE ""Canceled""='N' ) ""HeadPack"" ON E.""DocEntry"" = ""HeadPack"".""DocEntry"" "
                    sSQL &= " LEFT JOIN OTER TT On T1.""Territory""=TT.""territryID"" "
                    sSQL &= " WHERE T0.""CANCELED"" = 'N' AND T0.""U_EXO_TIPO"" = 'ITC' AND T0.""U_EXO_STATUSP""='C' and (IFNULL(""HeadPack"".""Status"",'') = 'O' OR IFNULL(""HeadPack"".""Status"",'') = '') "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""FromWhsCod""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " ORDER BY ""T. SALIDA"", ""Nº DOCUMENTO"" "
#End Region
                Case "SDPROV"
#Region "Devolución"
                    sSQL = "SELECT DISTINCT CAST('DPROV' as nVARCHAR(50)) ""T. SALIDA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", CAST(T0.""TrnspCode"" as nVARCHAR(50)) ""CLASE EXP."", IFNULL(CAST(AG.""U_EXO_AGE"" as nVARCHAR(50)),'-1') ""AG. TRANSPORTE"",  "
                    sSQL &= " EST.""Estado"" ""ESTADO"", E.""DocEntry"" ""List. Embalaje"", (SELECT STRING_AGG(IFNULL(""USR"", ''),'-') FROM (SELECT Distinct IFNULL(OUSR.""USER_CODE"", '') AS ""USR"" from ""@EXO_LSTEMBL"" X0 LEFT JOIN OUSR ON X0.""U_EXO_CODUSU"" = OUSR.""USERID"" Where X0.""DocEntry"" = E.""DocEntry"" Group by IFNULL(OUSR.""USER_CODE"", '') ) Y0) AS ""CODUSR"", 'N' ""Sel"" "
                    sSQL &= " FROM ORPD T0 "
                    sSQL &= " LEFT JOIN RPD1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN OSHP  AG ON AG.""TrnspCode""=TL.""TrnsCode"" "
                    sSQL &= " LEFT JOIN ""@EXO_LSTEMBL"" E ON  E.""DocEntry"" = T0.""U_EXO_LSTEMB"" "
                    sSQL &= " LEFT JOIN ""VEXO_PARRILLA_ESTADO_SALCOMP"" EST ON EST.""ORIGEN""='DPROV' and EST.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN ( SELECT ""DocEntry"", ""Status"" FROM ""@EXO_LSTEMB"" WHERE ""Canceled""='N' ) ""HeadPack"" ON E.""DocEntry"" = ""HeadPack"".""DocEntry"" "
                    sSQL &= " WHERE T0.""CANCELED"" = 'N' AND T0.""U_EXO_STATUSP""='C' and (T0.""U_EXO_ESTPAC""='Pendiente' or T0.""U_EXO_ESTPAC""='En curso') and (IFNULL(""HeadPack"".""Status"",'') = 'O' OR IFNULL(""HeadPack"".""Status"",'') = '') "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And ( CAST(T0.""U_EXO_CLASEE"" as nVARCHAR(50))='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " ORDER BY ""T. SALIDA"", ""Nº DOCUMENTO"" "
#End Region
            End Select
            oForm.DataSources.DataTables.Item("DTSCOM").ExecuteQuery(sSQL)
            FormateaGrid_SCOM(oForm)

            objGlobal.SBOApp.StatusBar.SetText("Datos de salida completadas Cargados con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub FiltrarENT(ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim sEntradas As String = ""
        Dim sICD As String = "" : Dim sICH As String = ""
        Dim sEXPE As String = "" : Dim sTerri As String = ""
        Dim sSQL As String = ""
        Dim dateBack = Now.AddDays(7)

#End Region
        Try
            sEntradas = oForm.DataSources.UserDataSources.Item("UDENT").Value.ToString
            sICD = oForm.DataSources.UserDataSources.Item("UDICD").Value.ToString
            sICH = oForm.DataSources.UserDataSources.Item("UDICH").Value.ToString
            sEXPE = oForm.DataSources.UserDataSources.Item("UDEXPE").Value.ToString
            sTerri = oForm.DataSources.UserDataSources.Item("UDTERRI").Value.ToString
            oForm.Freeze(True)
            Select Case sEntradas
                Case "-"
                    sSQL = "SELECT CAST('' as nVARCHAR(50)) ""T. ENTRADA"", CAST('' as nVARCHAR(50)) ""DELEGACIÓN"", CAST('' as nVARCHAR(50)) ""FECHA CREACION"", CAST('' as nVARCHAR(50)) ""FECHA ENTREGA"", CAST('' as nVARCHAR(50)) ""Nº INTERNO"", CAST('' as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST('' as nVARCHAR(50)) ""CÓDIGO"",  CAST('' as nVARCHAR(150))	""EMPRESA"", CAST('' as nVARCHAR(50)) ""ESTADO"", CAST('' as nVARCHAR(50)) ""DOC. ENTRADA"", "
                    sSQL &= " CAST('' as nVARCHAR(50)) ""ID DOC. ENTRADA"", CAST('' AS NVARCHAR(50)) ""USR. REUBICACION"" , 'N' ""Sel"" "
                    sSQL &= "FROM DUMMY "
                Case "TODOS"
#Region "Todos"
                    sSQL = "SELECT * FROM ( "
                    sSQL &= " SELECT DISTINCT CAST('PED' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", T0.""DocDate"" AS ""FECHA CREACION"", T0.""DocDueDate"" AS ""FECHA ENTREGA"", 
                              CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    'sSQL &= " CAST((CASE WHEN T4.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST((CASE WHEN T4.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"","
                    sSQL &= " T4.""U_EXO_PACKING"" AS ""Packing"", (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') 
                                                                        FROM (SELECT X1.""USER_CODE"" from ""@EXO_PACKINGL"" X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" 
                                                                                    Where X0.""Code"" = T4.""U_EXO_PACKING"" Group by X1.""USER_CODE"") Y0) AS ""Usuario Reubicacion"",  "
                    sSQL &= " 'N' ""Sel"" "
                    sSQL &= " FROM OPOR T0 "
                    sSQL &= " LEFT JOIN POR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN PDN1 T3 ON TL.""DocEntry""=T3.""BaseEntry"" and T0.""ObjType""=T3.""BaseType"" and TL.""LineNum""=T3.""BaseLine"" "
                    sSQL &= " Left JOIN OPDN T4 On T3.""DocEntry""=T4.""DocEntry"" AND T4.""CANCELED"" = 'N' "
                    sSQL &= " LEFT JOIN WTQ1 T5 ON T5.""BaseEntry""=T3.""DocEntry"" and T5.""BaseType""=T4.""ObjType"" "
                    sSQL &= " LEFT JOIN OWTQ T6 ON T5.""DocEntry""=T6.""DocEntry"" AND T6.""CANCELED"" = 'N' "
                    sSQL &= " WHERE (((TL.""LineStatus""='O' or (TL.""LineStatus""='C' and IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') <> '')) and T0.""DocDueDate""<='" & dateBack.Year.ToString("0000") & dateBack.Month.ToString("00") & dateBack.Day.ToString("00") & "') "
                    sSQL &= " ) "
                    sSQL &= " and IFNULL(T4.""DocStatus"",'O')='O' and IFNULL(T6.""DocStatus"",'O')='O' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT DISTINCT CAST('STR' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", T0.""DocDate"" AS ""FECHA CREACION"", T0.""DocDueDate"" AS ""FECHA ENTREGA"", 
                              CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    sSQL &= " CAST((CASE WHEN T0.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"", 
                                T4.""U_EXO_PACKING"" AS ""Packing"", (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') 
                                                                        FROM (SELECT X1.""USER_CODE"" from ""@EXO_PACKINGL"" X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" 
                                                                        Where X0.""Code"" = T4.""U_EXO_PACKING"" Group by X1.""USER_CODE"") Y0) AS ""Usuario Reubicacion"", "
                    sSQL &= " 'N' ""Sel"" "
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL On TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 On T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 On T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R On R.""DocEntry""=T0.""DocEntry"" And R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S On S.""DocEntry""=T0.""DocEntry"" And S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A On A.""CardCode""=T0.""CardCode"" And A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT On T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN WTR1 T3 ON TL.""DocEntry""=T3.""BaseEntry"" And T0.""ObjType""=T3.""BaseType""  And TL.""LineNum""=T3.""BaseLine"" "
                    sSQL &= " Left JOIN OWTR T4 ON T3.""DocEntry""=T4.""DocEntry"" And T4.""CANCELED"" = 'N' "
                    sSQL &= " WHERE T0.""U_EXO_TIPO"" = 'ITC' and (T0.""DocDueDate""<='" & dateBack.Year.ToString("0000") & dateBack.Month.ToString("00") & dateBack.Day.ToString("00") & "' )"
                    sSQL &= " And IFNULL(T4.""DocStatus"",'O')<>'C' "
                    sSQL &= " And (((TL.""LineStatus""='O' and T0.""U_EXO_TIPO""='ITC') OR (TL.""LineStatus""='C' and IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'')<>'')) ) "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " UNION ALL "
                    sSQL &= "SELECT DISTINCT CAST('SDE' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", T0.""DocDate"" AS ""FECHA CREACION"", T0.""DocDueDate"" AS ""FECHA ENTREGA"", 
                              CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    sSQL &= " CAST((CASE WHEN T0.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"", 
                                T4.""U_EXO_PACKING"" AS ""Packing"", (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') 
                                                                        FROM (SELECT X1.""USER_CODE"" from ""@EXO_PACKINGL"" X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" 
                                                                        Where X0.""Code"" = T4.""U_EXO_PACKING"" Group by X1.""USER_CODE"") Y0) AS ""Usuario Reubicacion"",  "
                    sSQL &= " 'N' ""Sel"" "
                    sSQL &= "FROM ORRR T0 "
                    sSQL &= " LEFT JOIN RRR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN RDN1 T3 ON T0.""DocEntry""=T3.""BaseEntry"" and T0.""ObjType""=T3.""BaseType"" "
                    sSQL &= " Left JOIN ORDN T4 ON T3.""DocEntry""=T4.""DocEntry"" AND T4.""CANCELED"" = 'N' "
                    sSQL &= " WHERE ((TL.""LineStatus""='O' and T0.""DocDueDate""<='" & dateBack.Year.ToString("0000") & dateBack.Month.ToString("00") & dateBack.Day.ToString("00") & "') "
                    sSQL &= " ) "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= ") Z ORDER BY ""T. ENTRADA"", ""Nº DOCUMENTO"" "
#End Region
                Case "PED"
#Region "Pedidos de compra"
                    sSQL = "SELECT DISTINCT CAST('PED' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", T0.""DocDate"" AS ""FECHA CREACION"", T0.""DocDueDate"" AS ""FECHA ENTREGA"", 
                              CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    'sSQL &= " CAST((CASE WHEN T4.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST((CASE WHEN T4.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"", 
                                T4.""U_EXO_PACKING"" AS ""Packing"", 
                                (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') 
                                        FROM (SELECT X1.""USER_CODE"" 
                                                    from ""@EXO_PACKINGL"" X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" 
                                                    Where X0.""Code"" = T4.""U_EXO_PACKING"" 
                                                    Group by X1.""USER_CODE""
                                            ) Y0) AS ""Usuario Reubicacion"",  "
                    sSQL &= " 'N' ""Sel"" "
                    sSQL &= " FROM OPOR T0 "
                    sSQL &= " LEFT JOIN POR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " INNER JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN PDN1 T3 ON TL.""DocEntry""=T3.""BaseEntry"" and T0.""ObjType""=T3.""BaseType"" and TL.""LineNum""=T3.""BaseLine"" "
                    sSQL &= " Left JOIN OPDN T4 On T3.""DocEntry""=T4.""DocEntry"" AND T4.""CANCELED"" = 'N' "
                    sSQL &= " LEFT JOIN WTQ1 T5 ON T5.""BaseEntry""=T3.""DocEntry"" and T5.""BaseType""=T4.""ObjType"" "
                    sSQL &= " LEFT JOIN OWTQ T6 ON T5.""DocEntry""=T6.""DocEntry"" AND T6.""CANCELED"" = 'N' "
                    sSQL &= " WHERE (((TL.""LineStatus""='O' or (TL.""LineStatus""='C' and IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') <> '')) and T0.""DocDueDate""<='" & dateBack.Year.ToString("0000") & dateBack.Month.ToString("00") & dateBack.Day.ToString("00") & "') "
                    sSQL &= " ) "
                    sSQL &= " and IFNULL(T4.""DocStatus"",'O')='O' and IFNULL(T6.""DocStatus"",'O')='O' "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " ORDER BY ""T. ENTRADA"", ""Nº DOCUMENTO"" "
#End Region
                Case "STR"
#Region "Sol de traslado en destino"
                    sSQL = "SELECT DISTINCT CAST('STR' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", T0.""DocDate"" AS ""FECHA CREACION"", 
                                T0.""DocDueDate"" AS ""FECHA ENTREGA"", CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    sSQL &= " CAST((CASE WHEN T0.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"", 
                                T4.""U_EXO_PACKING"" AS ""Packing"", (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') 
                                                                        FROM (SELECT X1.""USER_CODE"" from ""@EXO_PACKINGL"" X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" 
                                                                        Where X0.""Code"" = T4.""U_EXO_PACKING"" Group by X1.""USER_CODE"") Y0) AS ""Usuario Reubicacion"",  "
                    sSQL &= " 'N' ""Sel"" "
                    sSQL &= "FROM OWTQ T0 "
                    sSQL &= " LEFT JOIN WTQ1 TL On TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 On T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 On T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R On R.""DocEntry""=T0.""DocEntry"" And R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S On S.""DocEntry""=T0.""DocEntry"" And S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A On A.""CardCode""=T0.""CardCode"" And A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT On T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN WTR1 T3 ON TL.""DocEntry""=T3.""BaseEntry"" and T0.""ObjType""=T3.""BaseType"" And TL.""LineNum""=T3.""BaseLine"" "
                    sSQL &= " Left JOIN OWTR T4 ON T3.""DocEntry""=T4.""DocEntry"" AND T4.""CANCELED"" = 'N' "
                    sSQL &= " WHERE T0.""U_EXO_TIPO"" = 'ITC' and (T0.""DocDueDate""<='" & dateBack.Year.ToString("0000") & dateBack.Month.ToString("00") & dateBack.Day.ToString("00") & "' )"
                    sSQL &= " And IFNULL(T4.""DocStatus"",'O')<>'C' "
                    sSQL &= " And (((TL.""LineStatus""='O' and T0.""U_EXO_TIPO""='ITC') OR (TL.""LineStatus""='C' and IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'')<>'')) ) "

                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " ORDER BY ""T. ENTRADA"", ""Nº DOCUMENTO"" "
#End Region
                Case "SDE"
#Region "Sol de Devolución de cliente"
                    sSQL = "SELECT DISTINCT CAST('SDE' as nVARCHAR(50)) ""T. ENTRADA"", CAST(IFNULL(T2.""Name"",' ') as nVARCHAR(50)) ""DELEGACIÓN"", T0.""DocDate"" AS ""FECHA CREACION"", T0.""DocDueDate"" AS ""FECHA ENTREGA"", 
                                CAST(T0.""DocEntry"" as nVARCHAR(50)) ""Nº INTERNO"", CAST(T0.""DocNum"" as nVARCHAR(50)) ""Nº DOCUMENTO"", "
                    sSQL &= " CAST(T0.""CardCode"" as nVARCHAR(50)) ""CÓDIGO"",  CAST(T0.""CardName"" as nVARCHAR(150))	""EMPRESA"", "
                    sSQL &= " CAST((CASE WHEN T0.""U_EXO_ESTPAC""='Completado' THEN 'Completado' WHEN T0.""DocStatus""='O' THEN 'Pendiente' WHEN T0.""DocStatus""='C' THEN 'Recibido' ELSE 'En curso' END ) as nVARCHAR(50)) ""ESTADO"", "
                    sSQL &= " CAST(IFNULL(CAST(T4.""DocNum"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""DOC. ENTRADA"",  CAST(IFNULL(CAST(T4.""DocEntry"" as NVARCHAR(50)),'') as nVARCHAR(50)) ""ID DOC. ENTRADA"", 
                                T4.""U_EXO_PACKING"" AS ""Packing"", 
                                (SELECT Distinct STRING_AGG(IFNULL(""USER_CODE"",''),'-') 
                                        FROM (SELECT X1.""USER_CODE"" 
                                                    from ""@EXO_PACKINGL"" X0 Left join OUSR X1 ON X1.""USERID"" = X0.""U_EXO_CODUSU"" 
                                                    Where X0.""Code"" = T4.""U_EXO_PACKING"" Group by X1.""USER_CODE"") Y0) AS ""Usuario Reubicacion"",  "
                    sSQL &= " 'N' ""Sel"" "
                    sSQL &= "FROM ORRR T0 "
                    sSQL &= " LEFT JOIN RRR1 TL ON TL.""DocEntry""=T0.""DocEntry"" "
                    sSQL &= " LEFT JOIN OCRD T1 ON T0.""CardCode""=T1.""CardCode"" "
                    sSQL &= " LEFT JOIN OUBR T2 ON T1.""U_EXO_DELE""=T2.""Code"" "
                    sSQL &= " LEFT JOIN ""EXO_ROTURA"" R ON R.""DocEntry""=T0.""DocEntry"" and R.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_SITUACION"" S ON S.""DocEntry""=T0.""DocEntry"" and S.""ObjType""=T0.""ObjType"" "
                    sSQL &= " LEFT JOIN ""EXO_A"" A ON A.""CardCode""=T0.""CardCode"" and A.""WhsCode""=TL.""WhsCode"" "
                    sSQL &= " LEFT JOIN OTER TT ON T1.""Territory""=TT.""territryID"" "
                    sSQL &= " LEFT JOIN RDN1 T3 ON T0.""DocEntry""=T3.""BaseEntry"" and T0.""ObjType""=T3.""BaseType"" "
                    sSQL &= " Left JOIN ORDN T4 ON T3.""DocEntry""=T4.""DocEntry"" AND T4.""CANCELED"" = 'N' "
                    sSQL &= " WHERE ((TL.""LineStatus""='O' and T0.""DocDueDate""<='" & dateBack.Year.ToString("0000") & dateBack.Month.ToString("00") & dateBack.Day.ToString("00") & "') "
                    sSQL &= " ) "
                    If CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sSQL &= " and TL.""WhsCode""='" & CType(oForm.Items.Item("cbALM").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
                    End If
                    If sICD <> "" And sICH <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' and T0.""CardCode""<='" & sICH & "' )"
                    ElseIf sICD <> "" Then
                        sSQL &= " And (T0.""CardCode"">='" & sICD & "' )"
                    ElseIf sICH <> "" Then
                        sSQL &= " And (T0.""CardCode""<='" & sICH & "' )"
                    End If
                    If sEXPE <> "-" Then
                        sSQL &= " And (T0.""TrnspCode""='" & sEXPE & "' )"
                    End If
                    If sTerri <> "-" Then
                        sSQL &= " And (T1.""Territory""='" & sTerri & "' )"
                    End If
                    sSQL &= " ORDER BY ""T. ENTRADA"", ""Nº DOCUMENTO"" "
#End Region
            End Select
            oForm.DataSources.DataTables.Item("DTE").ExecuteQuery(sSQL)
            FormateaGrid_E(oForm)

            objGlobal.SBOApp.StatusBar.SetText("Datos de Entrada Cargados con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_E(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)

            Dim grid = CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid)
            For i = 0 To grid.Columns.Count - 1
                grid.Columns.Item(i).TitleObject.Sortable = True
            Next

            For i = 0 To 13
                Select Case i
                    Case 0
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("PED", "Pedido de compra")
                        oColumnCb.ValidValues.Add("STR", "Sol. de traslado Destino")
                        oColumnCb.ValidValues.Add("SDE", "Solicitud de devolución de Clientes")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 4
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "22"
                        oColumnTxt.Editable = False
                    Case 8
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                    Case 9
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                    Case 10
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                        oColumnTxt.LinkedObjectType = "20"
                    Case 11
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                        oColumnTxt.LinkedObjectType = "EXO_PACKING"
                    Case 12
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                    Case 13
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                        oColumnChk = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.CheckBoxColumn)
                        oColumnChk.Editable = True
                        oColumnChk.Width = 30
                    Case Else
                        CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next

            CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item("Nº INTERNO").Visible = False
            CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item("ID DOC. ENTRADA").Visible = False
            CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item("DOC. ENTRADA"), SAPbouiCOM.EditTextColumn).LinkedObjectType = "20"
            CType(CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item("Nº DOCUMENTO"), SAPbouiCOM.EditTextColumn).LinkedObjectType = "17"
            CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).Columns.Item("Packing").TitleObject.Caption = "Packing list"
            CType(oform.Items.Item("grdE").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_SCOM(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            Dim grid = CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid)
            CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(11).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(11), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 0 To grid.Columns.Count - 1
                grid.Columns.Item(i).TitleObject.Sortable = True
            Next

            For i = 0 To 9
                Select Case i
                    Case 0
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("ALBVTA", "Entrega de clientes")
                        oColumnCb.ValidValues.Add("SOLTRA", "Sol. de traslado Origen")
                        oColumnCb.ValidValues.Add("DPROV", "Dev. de Proveedor")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 2
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "17"
                        oColumnTxt.Editable = False
                    Case 4
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "2"
                        oColumnTxt.Editable = False
                    Case 6
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        Try
                            sSQL = " SELECT CAST(""TrnspCode"" as NVARCHAR(50)) ,""TrnspName"" "
                            sSQL &= " From OSHP  "
                            sSQL &= " ORDER By  ""TrnspName"" "
                            objGlobal.funcionesUI.cargaCombo(oColumnCb.ValidValues, sSQL)
                            oColumnCb.ValidValues.Add("-1", " ")

                        Catch ex As Exception

                        End Try
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description

                        oColumnCb.Editable = False
                    Case 7
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)

                        sSQL = "SELECT '-1', ' ' FROM DUMMY "
                        sSQL &= " UNION ALL "
                        sSQL = " SELECT ""CardCode"" ,""CardFName"" WHERE ""QryGroup1""='Y' "
                        sSQL &= " From OCRD  "
                        Try
                            objGlobal.funcionesUI.cargaCombo(oColumnCb.ValidValues, sSQL)
                            'oColumnCb.ValidValues.Add("-1", " ")

                        Catch ex As Exception

                        End Try
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 9
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "EXO_LSTEMB"
                        oColumnTxt.Editable = False
                    Case 8
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("PP", "Packing Pendiente")
                        oColumnCb.ValidValues.Add("PC", "Packing En Curso")
                        oColumnCb.ValidValues.Add("PT", "Packing Completado")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case Else
                        CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next

            grid.Columns.Item("CODUSR").Editable = False
            grid.Columns.Item("CODUSR").TitleObject.Caption = "Usuario de Packing"
            grid.Columns.Item("Nº INTERNO").Visible = False
            CType(grid.Columns.Item("Nº DOCUMENTO"), SAPbouiCOM.EditTextColumn).LinkedObjectType = 17

            grid.SelectionMode = BoMatrixSelect.ms_Single
            CType(oform.Items.Item("grdSCOM").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
            oform.Settings.MatrixUID = grid.Item.UniqueID
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_SLIB(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            Dim grid = CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid)
            CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(15).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(15), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 0 To grid.Columns.Count - 1
                grid.Columns.Item(i).TitleObject.Sortable = True
            Next

            For i = 0 To 13
                Select Case i
                    Case 0
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("PEDVTA", "Pedido de clientes")
                        oColumnCb.ValidValues.Add("SOLTRA", "Sol. de traslado Origen")
                        oColumnCb.ValidValues.Add("SDPROV", "Sol. de dev. de Proveedor")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 3
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "17"
                        oColumnTxt.Editable = False
                    Case 4
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "2"
                        oColumnTxt.Editable = False
                    Case 7
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)

                        oColumnCb.ValidValues.Add("Y", "Sí")
                        oColumnCb.ValidValues.Add("N", "No")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                        oColumnCb.Visible = False
                    Case 8
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)

                        oColumnCb.ValidValues.Add("Y", "Sí")
                        oColumnCb.ValidValues.Add("N", "No")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 9, 10
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Visible = False
                    Case 6
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        Try
                            Dim whs = oform.DataSources.UserDataSources.Item("UDALM").Value
                            sSQL = " SELECT CAST(""TrnspCode"" as NVARCHAR(50)) ,""TrnspName"" "
                            sSQL &= " From OWHS  JOIN OSHP ON OWHS.""U_EXO_SUCURSAL"" = OSHP.""U_EXO_DELEGA"""
                            sSQL &= " WHERE OWHS.""WhsCode"" = '" & whs & "'"
                            sSQL &= " ORDER By  ""TrnspName"" "
                            objGlobal.funcionesUI.cargaCombo(oColumnCb.ValidValues, sSQL)
                            oColumnCb.ValidValues.Add("-1", " ")
                        Catch ex As Exception

                        End Try
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = True
                    Case 11, 12
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.RightJustified = True
                    Case 13
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "156"
                        oColumnTxt.Editable = False
                    Case Else
                        CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next

            CType(oform.Items.Item("grdSLIB").Specific, Grid).Columns.Item("Usuario Picking").Editable = False
            CType(oform.Items.Item("grdSLIB").Specific, Grid).Columns.Item("Nº INTERNO").Visible = False
            CType(oform.Items.Item("grdSLIB").Specific, Grid).Columns.Item("CLASE EXP.").Width = 150

            grid.SelectionMode = BoMatrixSelect.ms_Single
            'CType(oform.Items.Item("grdSLIB").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_SPTE(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            Dim grid = CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid)
            CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(14).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(14), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 0 To grid.Columns.Count - 1
                grid.Columns.Item(i).TitleObject.Sortable = True
            Next

            For i = 0 To 11
                Select Case i
                    Case 0
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("PEDVTA", "Pedido de clientes")
                        oColumnCb.ValidValues.Add("SOLTRA", "Sol. de traslado Origen")
                        oColumnCb.ValidValues.Add("SDPROV", "Sol. de dev. de Proveedor")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 4
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "17"
                        oColumnTxt.Editable = False
                    Case 5
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "2"
                        oColumnTxt.Editable = False
                    Case 7
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "2"
                        oColumnTxt.Editable = False
                    Case 10
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)

                        oColumnCb.ValidValues.Add("Y", "Sí")
                        oColumnCb.ValidValues.Add("N", "No")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 9
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        Try
                            Dim whs = oform.DataSources.UserDataSources.Item("UDALM").Value
                            sSQL = " SELECT CAST(""TrnspCode"" as NVARCHAR(50)) ,""TrnspName"" "
                            sSQL &= " From OWHS  JOIN OSHP ON OWHS.""U_EXO_SUCURSAL"" = OSHP.""U_EXO_DELEGA"""
                            sSQL &= " WHERE OWHS.""WhsCode"" = '" & whs & "'"
                            sSQL &= " ORDER By  ""TrnspName"" "
                            objGlobal.funcionesUI.cargaCombo(oColumnCb.ValidValues, sSQL)
                            oColumnCb.ValidValues.Add("-1", " ")

                        Catch ex As Exception

                        End Try

                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description

                        oColumnCb.Editable = True
                    Case Else
                        CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdSPTE").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next

            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("T. SALIDA").Width = 150
            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("DELEGACIÓN").Width = 90
            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("Nº DOCUMENTO").Width = 90
            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("FECHA CREACION").Width = 90
            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("FECHA ENTREGA").Width = 90
            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("COM").Width = 30
            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("EMPRESA").Width = 150
            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("R.ST.").Width = 35
            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("CLASE EXP.").Width = 150
            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("Nº INTERNO").Visible = False
            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("ObjectType").Visible = False
            CType(oform.Items.Item("grdSPTE").Specific, Grid).Columns.Item("UBICACIÓN").Editable = False
            grid.SelectionMode = BoMatrixSelect.ms_Single

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_RSTOCK(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            Dim ogrd = CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid)

            For i = 0 To ogrd.Columns.Count - 1
                Select Case i
                    Case 0
                        CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                        oColumnCb = CType(CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.ComboBoxColumn)
                        oColumnCb.ValidValues.Add("17", "Pedido de clientes")
                        oColumnCb.ValidValues.Add("1250000001", "Sol. de traslado Origen")
                        oColumnCb.ValidValues.Add("234000032", "Sol. de dev. de Proveedor")
                        oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
                        oColumnCb.Editable = False
                    Case 1
                        CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "17"
                        oColumnTxt.Editable = False
                    Case 4
                        CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "4"
                        oColumnTxt.Editable = False
                    Case 6
                        CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.RightJustified = True
                        oColumnTxt.Editable = False
                    Case 7
                        ogrd.Columns.Item(i).Editable = False
                        ogrd.Columns.Item(i).Visible = True
                        ogrd.Columns.Item(i).TitleObject.Caption = "Stock"
                    Case 8
                        ogrd.Columns.Item(i).Editable = False
                        ogrd.Columns.Item(i).Visible = True
                        ogrd.Columns.Item(i).RightJustified = True
                    Case 9
                        ogrd.Columns.Item(i).Editable = False
                        ogrd.Columns.Item(i).Visible = True
                        ogrd.Columns.Item(i).RightJustified = True
                    Case Else
                        CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next
            CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_PSTOCK(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            Dim grid = CType(oform.Items.Item("grdPSTOCK").Specific, SAPbouiCOM.Grid)
            For i = 0 To grid.Columns.Count - 1
                CType(oform.Items.Item("grdPSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Editable = False
            Next
            CType(oform.Items.Item("grdPSTOCK").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_PSTOCKCOMP(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            Dim grid = CType(oform.Items.Item("grdPSTOCK").Specific, SAPbouiCOM.Grid)
            For i = 0 To grid.Columns.Count - 1
                CType(oform.Items.Item("grdPSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Editable = False
            Next
            CType(oform.Items.Item("grdPSTOCK").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid_PSTOCKBULTOS(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            Dim grid = CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid)
            For i = 0 To grid.Columns.Count - 1
                CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(i).Editable = False
            Next
            If grid.Columns.Count > 1 Then
                grid.Columns.Item("Nº Interno").Visible = False

                CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(2).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.EditTextColumn)
                oColumnTxt.LinkedObjectType = "22"
                oColumnTxt.Editable = False

                grid.Columns.Item("Interno Entrada").Visible = False
                grid.Columns.Item("Interno Emb").Visible = False
                grid.Columns.Item("Pdte. Recibir").RightJustified = True
                grid.Columns.Item("Recibido").RightJustified = True
                grid.Columns.Item("Pdte. Reubicar").RightJustified = True
            End If

            CType(oform.Items.Item("grdRSTOCK").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
End Class
