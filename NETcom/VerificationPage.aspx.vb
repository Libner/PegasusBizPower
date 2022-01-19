Public Class VerificationPage
    Inherits System.Web.UI.Page
    Protected UserId As Integer
    Protected VCode, fValue As String
    Protected con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected plhError As PlaceHolder




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
        UserId = Trim(Request.Cookies("bizpegasus")("UserId"))
          If Request.Form.Count > 0 Then
            VCode = Request.Form("VCode")
            Dim cmdSelect = New SqlClient.SqlCommand("SELECT VerificationCode FROM Users " & _
        " WHERE (ltrim(rtrim(VerificationCode)) = @VCode and User_id=@UserId)", con)
            cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = CInt(UserId)
            cmdSelect.Parameters.Add("@VCode", SqlDbType.VarChar, 20).Value = Trim(CInt(VCode))
            con.Open()
            Try
                fValue = cmdSelect.ExecuteScalar
            Catch ex As Exception
            End Try
            con.Close()

            If fValue = "" Then
                plhError.Visible = True
                '     Response.Write("Error")
            Else
                cmdSelect = New SqlClient.SqlCommand("update Users set VerificationCode_Confirmed=1 ,DateUpdate_ConfirmedCode=GetDate()" & _
                " WHERE (User_id=@UserId)", con)
                cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = CInt(UserId)
                con.Open()
                cmdSelect.ExecuteNonQuery()
                con.Close()
                Response.Redirect("/netCom/default.asp")

            End If

        End If

    End Sub

End Class
