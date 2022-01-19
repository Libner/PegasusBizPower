Imports System.Data.SqlClient

Public Class UploadUserFiles
    Inherits System.Web.UI.Page
    Protected UserId, deleteId As Integer
    Protected UploadUserFiles, UploadUserFilesView As Boolean


    Protected rptData As Repeater
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Request.QueryString("UserId") <> "" Then
            UserId = Request.QueryString("UserId")
        Else
            UserId = 0
        End If
        'Response.Write("UserId=" & UserId)
        UploadUserFiles = Trim(Request.Cookies("bizpegasus")("UploadUserFiles"))
        UploadUserFilesView = Trim(Request.Cookies("bizpegasus")("UploadUserFilesView"))

        '  Response.End()
        If IsNumeric(Request.QueryString("deleteId")) Then
            deleteId = CInt(Request.QueryString("deleteId"))
            Dim cmdDelete As New SqlClient.SqlCommand("DELETE FROM UsersFiles WHERE Id = @deleteId", con)
            cmdDelete.CommandType = CommandType.Text
            cmdDelete.Parameters.Add("@deleteId", SqlDbType.Int).Value = deleteId
            con.Open()
            '  Try
            cmdDelete.ExecuteNonQuery()
            con.Close()
            Dim cScript As String

            Response.Redirect("UploadUserFiles.aspx?UserId=" & UserId)


            '  Catch ex As Exception
            'Response.Write(Err.Description)
            '   Finally

            '   End Try
        End If
        If Not Page.IsPostBack Then
            bindList()
        End If
    End Sub
    Public Sub bindList()
        Dim cmdSelect As New System.Data.SqlClient.SqlCommand

        cmdSelect.Connection = con
        cmdSelect.CommandText = "GetUsersFiles"
        cmdSelect.CommandType = CommandType.StoredProcedure
        cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = UserId
        con.Open()

        Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        rptData.DataSource = dr
        rptData.DataBind()


        dr.Close()
        cmdSelect.Dispose()
        con.Close()
    End Sub
End Class
