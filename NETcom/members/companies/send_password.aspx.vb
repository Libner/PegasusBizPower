Public Class send_password
    Inherits System.Web.UI.Page
    Protected MemberId, CategoryId, TourId, DepartureId, ContactId, UserId, lang_id As Integer
    Protected FullName, Email, Phone, Cellular, Address, City, LoginName, Password As String
    Protected con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Dim dr As SqlClient.SqlDataReader
    Protected func As New bizpower.cfunc

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
        Server.ScriptTimeout = 10000
        Response.Buffer = False

        If IsNumeric(Request.QueryString("ContactId")) Then
            ContactId = CInt(Request.QueryString("ContactId"))
        Else
            ContactId = 0
        End If

        If IsNumeric(Request.QueryString("MemberId")) Then
            MemberId = CInt(Request.QueryString("MemberId"))
        Else
            MemberId = 0
        End If

        If MemberId > 0 Then
            cmdSelect = New SqlClient.SqlCommand("SELECT registerDate, updateDate, FullName, Phone, Cellular, Email, " & _
            " Address, City, PostIndex, LoginName, Password, Category_Id, Tour_Id, Departure_Id FROM dbo.Members " & _
            " WHERE (Member_Id = @MemberId)", con)
            cmdSelect.Parameters.Add("@MemberId", SqlDbType.Int).Value = CInt(MemberId)
            con.Open()
            Dim dr_m As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
            If dr_m.Read() Then
                If Not dr_m("FullName") Is DBNull.Value Then
                    FullName = Trim(dr_m("FullName"))
                End If
                If Not dr_m("Email") Is DBNull.Value Then
                    Email = Trim(dr_m("Email"))
                End If
                If Not dr_m("Phone") Is DBNull.Value Then
                    Phone = Trim(dr_m("Phone"))
                End If
                If Not dr_m("Cellular") Is DBNull.Value Then
                    Cellular = Trim(dr_m("Cellular"))
                End If
                If Not dr_m("Address") Is DBNull.Value Then
                    Address = Trim(dr_m("Address"))
                End If
                If Not dr_m("City") Is DBNull.Value Then
                    City = Trim(dr_m("City"))
                End If
                If Not dr_m("LoginName") Is DBNull.Value Then
                    LoginName = Trim(dr_m("LoginName"))
                End If
                If Not dr_m("Password") Is DBNull.Value Then
                    Password = Trim(dr_m("Password"))
                End If
                If Not dr_m("Category_Id") Is DBNull.Value Then
                    CategoryId = CInt(dr_m("Category_Id"))
                End If
                If Not dr_m("Tour_Id") Is DBNull.Value Then
                    TourId = CInt(dr_m("Tour_Id"))
                End If
                If Not dr_m("Departure_Id") Is DBNull.Value Then
                    DepartureId = CInt(dr_m("Departure_Id"))
                End If
            End If
            dr_m.Close()
            con.Close()

            Dim strBody As New System.Text.StringBuilder
            strBody.Append("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">" & vbCrLf)
            strBody.Append("<html><head><title></title>" & vbCrLf)
            strBody.Append("<style type=""text/css""> .page_title{" & vbCrLf)
            strBody.Append(" font-family:Arial;" & vbCrLf)
            strBody.Append(" FONT-WEIGHT: bolder;" & vbCrLf)
            strBody.Append(" font-size:12pt;" & vbCrLf)
            strBody.Append(" color:#E60009; " & vbCrLf)
            strBody.Append(" background-color:#E0E0E0;" & vbCrLf)
            strBody.Append(" padding-left:10px;" & vbCrLf)
            strBody.Append(" padding-right:10px;" & vbCrLf)
            strBody.Append(" line-height: 18px;" & vbCrLf)
            strBody.Append(" height: 26px;}" & vbCrLf)
            strBody.Append(" BODY" & vbCrLf)
            strBody.Append(" {" & vbCrLf)
            strBody.Append("  margin: 0px;" & vbCrLf)
            strBody.Append("  font-family:Arial;" & vbCrLf)
            strBody.Append("  COLOR: #666666;" & vbCrLf)
            strBody.Append(" }" & vbCrLf)
            strBody.Append(" TD" & vbCrLf)
            strBody.Append(" {" & vbCrLf)
            strBody.Append("  text-align: right;" & vbCrLf)
            ' strBody.Append("  direction: rtl;" & vbCrLf)
            strBody.Append(" }" & vbCrLf)
            strBody.Append("</style></head><body>" & vbCrLf)
            strBody.Append("<table cellpadding=3 cellspacing=1 border=0 width=550 align=center>" & vbCrLf)
            strBody.Append("<tr><td colspan=""2"">&nbsp;</td></tr>" & vbCrLf)
            strBody.Append("<tr><td colspan=""2"" align=""right"" dir=""rtl"">ברוך הבא לאתר של חברת פגסוס</td></tr>" & vbCrLf)
            strBody.Append("<tr><td colspan=""2"" align=""right"">אנא שמור פרטים אלו לצורך כניסה למערכת.</td></tr>" & vbCrLf)
            strBody.Append("<tr><td colspan=""2"" align=""right"">&nbsp;</td></tr>" & vbCrLf)
            strBody.Append("<tr><td colspan=""2"" class=""page_title"">פרטי כניסה</td></tr>" & vbCrLf)
            strBody.Append("<tr align=""right""><td bgcolor=""#F5F5F5""><b>" & LoginName & "</b></td><td nowrap bgcolor=""#F2F2F2""><b>שם משתמש</b></td></tr>" & vbCrLf)
            strBody.Append("<tr align=""right""><td bgcolor=""#F5F5F5"" dir=ltr><b>" & Password & "</b></td><td nowrap bgcolor=""#F2F2F2""><b>סיסמא</b></td></tr>" & vbCrLf)
            strBody.Append("<tr align=""right""><td bgcolor=""#F5F5F5"" dir=rtl>" & FullName & "</td><td nowrap bgcolor=""#F2F2F2"">שם מלא</td></tr>" & vbCrLf)
            strBody.Append("<tr align=""right""><td bgcolor=""#F5F5F5"">" & Email & "</td><td nowrap bgcolor=""#F2F2F2"">אימייל</td></tr>" & vbCrLf)
            strBody.Append("<tr align=""right""><td bgcolor=""#F5F5F5"">" & Phone & "</td><td nowrap bgcolor=""#F2F2F2"">טלפון</td></tr>" & vbCrLf)
            strBody.Append("<tr><td colspan=2 align=center height=10 valign=top></td></tr>")
            strBody.Append("<tr><td colspan=2 align=right height=10 valign=top><a href=""https://www.pegasusisrael.co.il/default.aspx"">לחץ כאן לכניסה לאתר</a></td></tr>")
            strBody.Append("<tr><td colspan=2 align=right height=10 valign=top>בברכה, צוות האתר</td></tr>")
            strBody.Append("</table></body></html>")
            'Response.Write strBody
            'Response.End		

            If Trim(Email) <> "" Then
                Dim Msg As New System.Web.Mail.MailMessage
                Msg.BodyFormat = System.Web.Mail.MailFormat.Html
                Msg.BodyEncoding = System.Text.Encoding.UTF8
                Msg.From = "info@pegasusisrael.co.il"
                Msg.Subject = "פרטי כניסה לאתר פגסוס"
                Msg.To = Email
                Msg.Body = strBody.ToString()
                System.Web.Mail.SmtpMail.SmtpServer = ConfigurationSettings.AppSettings.Item("smtp_server")
                Try

                    System.Web.Mail.SmtpMail.Send(Msg)

                Catch ex As Exception
                    Response.Write(Err.Description)
                    Response.End()
                End Try
                Msg = Nothing
            End If


        End If

        Response.Redirect("contact.asp?ContactId=" & ContactId)

    End Sub

End Class