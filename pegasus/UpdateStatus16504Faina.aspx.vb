Public Class UpdateStatus16504Faina
    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Public dr_m As SqlClient.SqlDataReader
    Protected dateM As String
    Protected P16504 As String
    Public pYear, pMonth, APPEALDATE As String
    Protected RelationId As Integer

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
        pYear = 2017 'Request.QueryString("y")
        pMonth = 12 ' Request.QueryString("p")
        Dim sql As String

        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        sql = "SET DATEFORMAT dmy;SELECT APPEAL_ID,APPEAL_DATE,RelationId from Appeals where  month(APPEAL_DATE)=12 and year(APPEAL_DATE)=2017 and QUESTIONS_ID=16735 and RelationId>0"
        Dim cmdSelect As New SqlClient.SqlCommand(sql, con)

        con.Open()
        dr_m = cmdSelect.ExecuteReader()
        While dr_m.Read
            If Not dr_m("RelationId") Is DBNull.Value Then
                RelationId = dr_m("RelationId")
                '    Response.Write("RelationId=" & RelationId & "<BR>")
            End If
            If Not dr_m("APPEAL_DATE") Is DBNull.Value Then
                APPEALDATE = dr_m("APPEAL_DATE")
            End If
            Dim s = "SET DATEFORMAT dmy;UPDATE Appeals SET  APPEAL_STATUS=5  where APPEAL_ID=" & RelationId & ";UPDATE APPEALS SET APPEAL_Update_DATE='" & APPEALDATE & "' where appeal_id=" & RelationId
            'Response.Write(s & "<BR>")
            Dim Upd As New SqlClient.SqlCommand("SET DATEFORMAT dmy;UPDATE Appeals SET  APPEAL_STATUS=5  where APPEAL_ID=" & RelationId & ";UPDATE APPEALS SET APPEAL_Update_DATE='" & APPEALDATE & "' where appeal_id=" & RelationId, conU)
            conU.Open()
            Upd.CommandType = CommandType.Text
            Upd.ExecuteNonQuery()
            conU.Close()

        End While
        dr_m.Close()
        cmdSelect.Dispose()
        con.Close()



    End Sub

End Class
