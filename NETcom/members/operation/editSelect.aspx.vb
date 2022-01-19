Public Class editSelect
    Inherits System.Web.UI.Page
    Protected SiteId, DepartureId As Integer
    Protected conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected func As New include.funcs
    Public Departure_GuideTelphone As String
    Protected WithEvents btnSubmit As UI.WebControls.Button
    Protected fName, fValue, DepartureCode, fTitle As String
    Protected WithEvents fselect As System.Web.UI.HtmlControls.HtmlSelect


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
        fName = Request.QueryString("fName")
        DepartureCode = Request.QueryString("DepartureCode")
        Select Case fName
            Case "Departure_Costing"
                fTitle = "תמחיר"
            Case "Voucher_Simultaneous"
                fTitle = "שובר לסימולטני"
            Case "Voucher_Group"
                fTitle = "שוברי הוצאות קבוצה ומדריך"
            Case "GilboaHotel"
                fTitle = "הוקלדו מלונות בגלבוע "
            Case "Vouchers_Provider"
                fTitle = "שוברי ספק קרקע"
        End Select



        If Not Page.IsPostBack Then
            If DepartureId > 0 Then
                cmdSelect = New SqlClient.SqlCommand("SELECT " & fName & " FROM Tours_Departures " & _
                " WHERE (Departure_Id = @Departure_Id)", conPegasus)
                cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
                conPegasus.Open()
                Try
                    fValue = cmdSelect.ExecuteScalar
                Catch ex As Exception

                End Try
                conPegasus.Close()
            End If
        Else
            ' Response.Write("1=" & Request.Form("fselect"))
            'Response.End()
            cmdSelect = New SqlClient.SqlCommand("update Tours_Departures set " & fName & " = '" & Request.Form("fselect") & "' WHERE (Departure_Id = @Departure_Id)", conPegasus)
            cmdSelect.Parameters.Add("@Departure_Id", SqlDbType.Int).Value = CInt(DepartureId)
            conPegasus.Open()
            cmdSelect.ExecuteNonQuery()
            conPegasus.Close()
            '  Response.Write(fName & ":" & Request.Form("fselect"))
            ' Response.End()
            Dim cScript As String
            cScript = "<script language='javascript'>opener.document.getElementById('" & fName & "_" & DepartureId & "').innerHTML = '" & Replace(Request.Form("fselect"), "'", "&#" & AscW(Chr(39)) & ";") & "'; self.close(); </script>"
            RegisterStartupScript("ReloadScrpt", cScript)
        End If

    End Sub
   

End Class