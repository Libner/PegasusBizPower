Public Class send_password1
    Inherits System.Web.UI.Page
    Public supplierId As Integer
    Protected Passw, Email As String
    Protected func As New bizpower.cfunc
    Protected con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmdSelect As New SqlClient.SqlCommand



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
        Server.ScriptTimeout = 10000
        Response.Buffer = False

        If IsNumeric(Request.QueryString("supplierId")) Then
            supplierId = CInt(Request.QueryString("supplierId"))
        Else
            supplierId = 0
        End If
        If supplierId > 0 Then
            Passw = func.RandomPassword(8, 2, 3, 3)

            cmdSelect = New SqlClient.SqlCommand("update Suppliers set password=@Passw,PasswordDate=GetDate(),PasswordChangeDate=NULL where supplier_Id = @supplierId; SELECT supplier_Email1,supplier_Email2,supplier_Email3, supplier_Email4 FROM Suppliers " & _
                      " WHERE (supplier_Id = @supplierId)", con)
            cmdSelect.Parameters.Add("@supplierId", SqlDbType.Int).Value = CInt(supplierId)
            cmdSelect.Parameters.Add("@Passw", SqlDbType.VarChar, 10).Value = Passw
            con.Open()
            Dim dr_m As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
            If dr_m.Read() Then

                If Not dr_m("supplier_Email1") Is DBNull.Value Then
                    If Email = "" Then
                        Email = Trim(dr_m("supplier_Email1"))
                    Else
                        Email = Email & "," & Trim(dr_m("supplier_Email1"))
                    End If
                End If
                If Not dr_m("supplier_Email2") Is DBNull.Value Then

                    If Email = "" Then
                        Email = Trim(dr_m("supplier_Email2"))
                    Else
                        Email = Email & "," & Trim(dr_m("supplier_Email2"))
                    End If
                End If


                If Not dr_m("supplier_Email3") Is DBNull.Value Then

                    If Email = "" Then
                        Email = Trim(dr_m("supplier_Email3"))
                    Else
                        Email = Email & "," & Trim(dr_m("supplier_Email3"))
                    End If
                End If

                If Not dr_m("supplier_Email4") Is DBNull.Value Then

                    If Email = "" Then
                        Email = Trim(dr_m("supplier_Email4"))
                    Else
                        Email = Email & "," & Trim(dr_m("supplier_Email4"))
                    End If

                End If
            End If
            '   Response.Write(Email)

            '   Email = "furfaina@gmail.com.co.il"
            If Trim(Email) <> "" Then
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
                strBody.Append("  text-align: left;" & vbCrLf)
                ' strBody.Append("  direction: rtl;" & vbCrLf)
                strBody.Append(" }" & vbCrLf)
                strBody.Append("</style></head><body>" & vbCrLf)
                strBody.Append("<table cellpadding=3 cellspacing=1 border=0 width=550 align=center>" & vbCrLf)
                strBody.Append("<tr><td colspan=""2"">&nbsp;</td></tr>" & vbCrLf)
                strBody.Append("<tr><td colspan=""2"" align=""left"" dir=""ltr"">Welcome to the Pegasus website</td></tr>" & vbCrLf)
                strBody.Append("<tr><td colspan=""2"" align=""left"">&nbsp;Please note that passwords are case-sensitive</td></tr>" & vbCrLf)
                strBody.Append("<tr align=""left""><td nowrap bgcolor=""#F2F2F2""><b>password: </b></td><td bgcolor=""#F5F5F5"" dir=ltr><b>" & Passw & "</b></td></tr>" & vbCrLf)
                strBody.Append("<tr align=""left""><td bgcolor=""#F2F2F2"" colspan=2>The user's name will be forwarded by telephone</td></tr>" & vbCrLf)
                strBody.Append("<tr><td colspan=2 align=center height=10 valign=top></td></tr>")
                strBody.Append("<tr><td colspan=2 align=left height=10 valign=top><a href=""https://www.pegasusisrael.co.il/suppliers/login.aspx"">Click here to login </a></td></tr>")
                strBody.Append("<tr><td colspan=2 align=left height=10 valign=top>Best Regards,<BR>Pegasus operation team</td></tr>")
                strBody.Append("</table></body></html>")
                '     Response.Write(strBody)
                '     Response.End()
                '   Response.Write("--" & Email)
                '     Response.End()

                Dim Msg As New System.Web.Mail.MailMessage
                Msg.BodyFormat = System.Web.Mail.MailFormat.Html
                Msg.BodyEncoding = System.Text.Encoding.UTF8
                Msg.From = "info@pegasusisrael.co.il"
                Msg.Subject = "Your Pegasus personal account"
                Msg.To = Email
                Msg.Body = strBody.ToString()
                System.Web.Mail.SmtpMail.Send(Msg)
                Msg = Nothing
            End If

            Response.Redirect("default.asp")



        End If

    End Sub

End Class
