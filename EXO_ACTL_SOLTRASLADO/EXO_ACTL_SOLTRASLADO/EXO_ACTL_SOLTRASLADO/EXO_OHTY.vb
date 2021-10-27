Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OHTY
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        If actualizar Then
            CargaRol()
        End If
    End Sub
    Private Sub CargaRol()
        Dim oRoleSrv As SAPbobsCOM.EmployeeRolesSetupService = Nothing
        Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing
        Dim addLine As SAPbobsCOM.EmployeeRoleSetup = Nothing
        Dim sExiste As String = "" : Dim sSQl As String = ""
        Try
            If objGlobal.refDi.comunes.esAdministrador Then
                sSQl = "SELECT ""name"" FROM ""OHTY"" WHERE ""name""='Almacén'"
                sExiste = objGlobal.refDi.SQL.sqlStringB1(sSQl)
                If sExiste <> "" Then
                    objGlobal.SBOApp.StatusBar.SetText("Ya existe el Rol ""Almacén"". No se creará de nuevo.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Else
                    oCmpSrv = objGlobal.compañia.GetCompanyService
                    oRoleSrv = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.EmployeeRolesSetupService), SAPbobsCOM.EmployeeRolesSetupService)

                    addLine = CType(oRoleSrv.GetDataInterface(SAPbobsCOM.EmployeeRolesSetupServiceDataInterfaces.erssEmployeeRoleSetup), SAPbobsCOM.EmployeeRoleSetup)

                    addLine.Name = "Almacén"
                    addLine.Description = "Almacén"
                    oRoleSrv.AddEmployeeRoleSetup(addLine)
                    objGlobal.SBOApp.StatusBar.SetText("Se ha creado el Rol ""Almacén"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If

            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            If addLine IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(addLine)
            If oRoleSrv IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRoleSrv)
            If oCmpSrv IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrv)
        End Try
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
End Class
