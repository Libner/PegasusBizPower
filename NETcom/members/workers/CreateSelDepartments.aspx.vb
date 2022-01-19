Imports System.Data.SqlClient
Public Class CreateSelDepartments
    Inherits System.Web.UI.Page
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dtsDepartments As New DataTable
    Public sDepartments As String

    Protected UId As Integer = 1229



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

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here

        CreateArray()

    End Sub
    Private Sub CreateArray()
         Dim cmdSelect As New SqlClient.SqlCommand("SELECT departmentId, departmentName from Departments  order by PriorityLevel", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtsDepartments)
        con.Close()

        ' Dim list1 As New ListItem("בחר מחלקה", "0")
        sDepartments = "<option value=0>בחר מחלקה</option>"
        For i As Integer = 0 To dtsDepartments.Rows.Count - 1
            sDepartments = sDepartments & "<option value=" & dtsDepartments.Rows(i)("departmentId") & ">" & dtsDepartments.Rows(i)("departmentName") & "</option>"
        Next
        Response.Write(sDepartments)

    End Sub
End Class
