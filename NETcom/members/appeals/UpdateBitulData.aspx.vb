Imports System.Data.SqlClient

Public Class UpdateBitulData
    Inherits System.Web.UI.Page
    Protected appId As Integer
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

    Dim cmd As New SqlClient.SqlCommand

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
        appId = Request.Form("appId")

        con.Open()

        cmd = New System.Data.SqlClient.SqlCommand("UPDATE APPEALS  SET APPEAL_DATE_BITUL = GetDate()  WHERE(appeal_id=@appId and  APPEAL_DATE_BITUL IS NULL)", con)
        cmd.Parameters.Add("@appId", SqlDbType.Int).Value = appId
    
        cmd.CommandType = CommandType.Text
        ' cmd.Parameters.Add("@chkValue", SqlDbType.Int).Value = CInt(chk.Value())
        'cmd.Parameters.Add("@ProfId", SqlDbType.Int).Value = CInt(chk.Attributes("value"))
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

End Class
