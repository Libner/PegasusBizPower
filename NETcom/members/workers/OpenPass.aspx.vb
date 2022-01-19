

Public Class OpenPass
    Inherits System.Web.UI.Page
    Protected pass As String
 
#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm

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
        Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT top 1 Pass_OpenFile from PassOpenFile", myConnection)
        cmdSelect.CommandType = CommandType.Text
        myConnection.Open()
        pass = cmdSelect.ExecuteScalar()
        myConnection.close()
        '  Response.Write("pass=" & pass)
        If Page.IsPostBack Then
            '   Response.Write(Request.Form("pass"))
            myConnection.Open()
            Dim cmd = New System.Data.SqlClient.SqlCommand("Update PassOpenFile set Pass_OpenFile=@PassOpenFile", myConnection)
            cmd.Parameters.Add("@PassOpenFile", SqlDbType.VarChar, 50).Value = Request.Form("pass")
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()
            myConnection.Close()
            Dim cScript As String


            cScript = "<script language='javascript'>self.close(); </script>"
            RegisterStartupScript("ReloadScrpt", cScript)
        End If


    End Sub

End Class
