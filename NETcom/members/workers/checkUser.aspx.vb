Public Class checkUser
    Inherits System.Web.UI.Page
    Protected fname, pass, result As String
    Dim myReader As System.Data.SqlClient.SqlDataReader
    Protected Divresult As PlaceHolder





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
        fname = Request.QueryString("fname")
        pass = Request.Form("pass")
        result = ""
        ' Response.Write(fname)
        If Page.IsPostBack Then
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            Dim cmdSelect As New SqlClient.SqlCommand("SELECT top 1 Pass_OpenFile from PassOpenFile where ltrim(rtrim(Pass_OpenFile))=@pass", myConnection)
            cmdSelect.Parameters.Add("@pass", SqlDbType.VarChar, 50).Value = Trim(Request.Form("pass"))
            cmdSelect.CommandType = CommandType.Text
            myConnection.Open()
            myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)


            ' Response.Write("result=" & result)

            If myReader.HasRows Then
                Dim cScript As String


                cScript = "<script language='javascript'>window.open('../../../Download/UsersFiles/" + fname + "','_blank');self.close(); </script>"
                RegisterStartupScript("ReloadScrpt", cScript)


            Else
                Divresult.Visible = True


            End If

        End If
    End Sub

End Class
