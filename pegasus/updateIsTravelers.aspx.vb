Public Class updateIsTravelers
    Inherits System.Web.UI.Page
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dr_m As SqlClient.SqlDataReader
    Public Docket_pfileNum As String
    Public appealid As String





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
        Dim sql As String
        sql = "select  distinct Docket_pfileNum,appeal_id from Tour_Travelers left join FORM_VALUE on  FIELD_VALUE=cast(Docket_pfileNum as varchar(max))"
        Dim cmdSelect As New SqlClient.SqlCommand(sql, con)
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        dr_m = cmdSelect.ExecuteReader()

        While dr_m.Read
            If Not dr_m("appeal_id") Is DBNull.Value Then
                appealid = dr_m("appeal_id")
                Dim Upd As New SqlClient.SqlCommand("UPDATE Appeals SET  APPEAL_TravelersCount=1 where appeal_id=" & appealid, conU)
                ' Response.Write("UPDATE Appeals SET  APPEAL_TravelersCount=1 where appeal_id=" & appealid & "<BR>")
                '  Response.End()


                '  Upd.Parameters.Add("@appealid", SqlDbType.Int).Value = appealid
                conU.Open()
                Upd.CommandType = CommandType.Text
                Upd.ExecuteNonQuery()
                conU.Close()

            End If

        End While
        dr_m.Close()
        cmdSelect.Dispose()
        con.Close()


    End Sub

End Class
