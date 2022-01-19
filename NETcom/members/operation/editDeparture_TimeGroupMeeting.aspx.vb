Public Class editDeparture_TimeGroupMeeting
    Inherits System.Web.UI.Page
    Protected SiteId, DepartureId As Integer
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected func As New include.funcs
    Public Departure_TimeGroupMeeting, DepartureCode As String
    Protected WithEvents btnSubmit As UI.WebControls.Button

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        'If IsNothing(Request.Cookies(ConfigurationSettings.AppSettings("AdminCookieName"))) OrElse _
        'Not IsNumeric(Request.Cookies(ConfigurationSettings.AppSettings("AdminCookieName"))("workerid")) Then
        '    Response.Redirect("../default.aspx")
        'End If


        If IsNumeric(Request.QueryString("DepartureId")) Then
            DepartureId = CInt(Request.QueryString("DepartureId"))
        End If

        DepartureCode = Request.QueryString("DepartureCode")

        If Not Page.IsPostBack Then
            If DepartureId > 0 Then
                cmdSelect = New SqlClient.SqlCommand("SELECT Departure_TimeGroupMeeting FROM Tours_Departures " & _
                " WHERE (Departure_Id = @Departure_Id)", conPegasus)
                cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
                conPegasus.Open()
                Try
                    Departure_TimeGroupMeeting = cmdSelect.ExecuteScalar
                Catch ex As Exception

                End Try
                conPegasus.Close()
            End If
        Else
            cmdSelect = New SqlClient.SqlCommand("update Tours_Departures set Departure_TimeGroupMeeting=@Departure_TimeGroupMeeting" & _
          " WHERE (Departure_Id = @Departure_Id)", conPegasus)
            cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
            cmdSelect.Parameters.Add("@Departure_TimeGroupMeeting", SqlDbType.Text).Value = Request.Form("Departure_TimeGroupMeeting")
            conPegasus.Open()
            cmdSelect.ExecuteNonQuery()
            conPegasus.Close()
            Dim cScript As String
            cScript = "<script language='javascript'>opener.document.getElementById('Departure_TimeGroupMeeting_" & DepartureId & "').innerHTML = '" & Replace(Request.Form("Departure_TimeGroupMeeting"), "'", "&#" & AscW(Chr(39)) & ";") & "'; self.close(); </script>"
            RegisterStartupScript("ReloadScrpt", cScript)
        End If

    End Sub

End Class