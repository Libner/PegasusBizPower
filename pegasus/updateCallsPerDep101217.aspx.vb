Imports System.Data.SqlClient
Public Class updateCallsPerDep101217
    Inherits System.Web.UI.Page
#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Public func As New bizpower.cfunc
    Public dr_m As SqlClient.SqlDataReader
    Protected dateM As String
    Protected dep As Integer
    Public pYear As String
    Protected call_Service, call_Sales As String



    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        pYear = Request.QueryString("y")
        'Put user code to initialize the page here
        Dim sql As String
        dep = 1

        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        If IsNumeric(pYear) Then
            sql = "SET DATEFORMAT dmy;SELECT  DateKey ,call_Service,call_Sales,Id from DimDate  where year(DateKey)=" & pYear
        Else
            Response.Write("error")
            Response.End()
        End If
        Dim cmdSelect As New SqlClient.SqlCommand(sql, con)
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        dr_m = cmdSelect.ExecuteReader()

        While dr_m.Read
            dateM = Day(dr_m("DateKey")) & "/" & Month(dr_m("DateKey")) & "/" & Year(dr_m("DateKey"))
            ID = dr_m("ID")
            If Not dr_m("call_Service") Is DBNull.Value Then
                call_Service = dr_m("call_Service")
            Else
                call_Service = 0
            End If

            If Not dr_m("call_Sales") Is DBNull.Value Then
                call_Sales = dr_m("call_Sales")
            Else
                call_Sales = 0
            End If

            Dim cmdUpdate As New SqlClient.SqlCommand("update_Departments_Calls", conU)
            cmdUpdate.CommandType = CommandType.StoredProcedure
            cmdUpdate.Parameters.Add("@pcall_Service", SqlDbType.Int).Value = call_Service
            cmdUpdate.Parameters.Add("@pcall_Sales", SqlDbType.Int).Value = call_Sales
            cmdUpdate.Parameters.Add("@dep", SqlDbType.Int).Value = dep
            cmdUpdate.Parameters.Add("@Id_DimDate", SqlDbType.Int).Value = ID
            cmdUpdate.Parameters.Add("@pKeY", SqlDbType.DateTime).Value = dateM
             conU.Open()
            cmdUpdate.ExecuteNonQuery()
            cmdUpdate.Dispose()
            conU.Close()

        End While
        dr_m.Close()
        cmdSelect.Dispose()
        con.Close()
        'Response.Write("end")



    End Sub

End Class
