Public Class EditCalendar
    Inherits System.Web.UI.Page
    Protected SiteId, DepartureId As Integer
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected func As New include.funcs
    Public fName, DepartureCode, fTitle, fValue As String
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents btnSubmit As UI.WebControls.Button

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
        If IsNumeric(Request.QueryString("DepartureId")) Then
            DepartureId = CInt(Request.QueryString("DepartureId"))
        End If

        fName = Request.QueryString("fName")
        DepartureCode = Request.QueryString("DepartureCode")
        ' Response.Write("fName=" & fName)
        Select Case fName
            Case "DateMeetingAfterTrip"
                fTitle = "תאריך פגישה לאחר טיול"
            Case "Departure_Itinerary"
                fTitle = "תאריך קבלת Itinerary"
            Case "Departure_DateBrief"
                fTitle = "תאריך בריף"
            Case "Departure_DateGroupMeeting"
                fTitle = "תאריך מפגש קבוצה"
        End Select
        If Not Page.IsPostBack Then
            If DepartureId > 0 Then
                cmdSelect = New SqlClient.SqlCommand("SET DATEFORMAT DMY;SELECT " & fName & " FROM Tours_Departures " & _
                " WHERE (Departure_Id = @Departure_Id)", conPegasus)
                cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
                conPegasus.Open()
                Try
                    fValue = cmdSelect.ExecuteScalar
                Catch ex As Exception

                End Try
                ' Response.Write("fValue=" & fValue)
                If fValue = "" Then
                    fValue = Now()
                End If
                conPegasus.Close()
            End If
        Else

            ''''   update(Tours_Departures) DateMeetingAfterTrip = NULL   where(Departure_Id = 6296)
            ' Response.Write("1=" & Request.Form("datepickerN"))
            '  Response.End()
            If IsDate(Request.Form("datepickerN")) Then
                cmdSelect = New SqlClient.SqlCommand("SET DATEFORMAT DMY;update Tours_Departures set " & fName & " = '" & Request.Form("datepickerN") & "' WHERE (Departure_Id = @Departure_Id)", conPegasus)
            Else
                cmdSelect = New SqlClient.SqlCommand("SET DATEFORMAT DMY;update Tours_Departures set " & fName & " = NULL WHERE (Departure_Id = @Departure_Id)", conPegasus)

            End If


            cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
            conPegasus.Open()
            cmdSelect.ExecuteNonQuery()
            conPegasus.Close()
            '  Response.Write(fName & ":" & Request.Form("fselect"))
            ' Response.End()
            '  Dim valueD As Date = Request.Form("datepickerN")

            ' Dim valueD As Date = Request.Form("datepickerN")
            Dim cScript As String
            cScript = "<script language='javascript'>opener.document.getElementById('" & fName & "_" & DepartureId & "').innerHTML = '" & Request.Form("datepickerN") & "'; self.close(); </script>"
            RegisterStartupScript("ReloadScrpt", cScript)
        End If
    End Sub

End Class
