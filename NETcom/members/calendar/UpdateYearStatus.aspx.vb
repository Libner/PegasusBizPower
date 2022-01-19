Imports System.Data.SqlClient
Public Class UpdateYearStatus
    Inherits System.Web.UI.Page
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmd As New SqlClient.SqlCommand
    Public year, selV As String
   
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
        year = Request("year")
        selV = Request("selV")
        con.Open()

        cmd = New System.Data.SqlClient.SqlCommand("Update YearsStatus set Year_Status=@selV where Year=@year", con)
        cmd.Parameters.Add("@selV", SqlDbType.Char, 1).Value = selV
        cmd.Parameters.Add("@year", SqlDbType.VarChar, 10).Value = year
   
        cmd.CommandType = CommandType.Text
        ' cmd.Parameters.Add("@chkValue", SqlDbType.Int).Value = CInt(chk.Value())
        'cmd.Parameters.Add("@ProfId", SqlDbType.Int).Value = CInt(chk.Attributes("value"))
        cmd.ExecuteNonQuery()
        con.Close()
        If selV = "1" Then
            Response.Write(" עמוד תקין ")
        Else
            Response.Write(" עמוד לא תקין ")
        End If

    End Sub

End Class
