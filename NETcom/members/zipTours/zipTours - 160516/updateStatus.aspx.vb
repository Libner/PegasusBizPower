Imports System.Data.SqlClient

Public Class updateStatus
    Inherits System.Web.UI.Page
    Protected depId, statusId, st As String
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
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
        st = Request("st")

        If st <> "" Then
            Dim stArr As Array
            stArr = Split(st, "-")
            statusId = stArr(1)
            depId = stArr(0)



            conPegasus.Open()

            cmd = New System.Data.SqlClient.SqlCommand("Update Tours_Departures set Status_ForBizpower=@statusId where Departure_Id=@depId", conPegasus)
            cmd.CommandType = CommandType.Text
            cmd.Parameters.Add("@statusId", SqlDbType.Int).Value = CInt(statusId)
            cmd.Parameters.Add("@depId", SqlDbType.Int).Value = CInt(depId)
            'cmd.Parameters.Add("@ProfId", SqlDbType.Int).Value = CInt(chk.Attributes("value"))
            cmd.ExecuteNonQuery()
            conPegasus.Close()
        End If
    End Sub

End Class
