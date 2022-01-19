'Imports EASendMail
Imports System.Text.RegularExpressions

Public Class sendMailQ
    Inherits System.Web.UI.Page
    Protected hdnemailLimit, hdnfirst_slice, hdnsecond_slice, hdnpagesource, hdnUseBizLogo, hdnPageLang As HtmlControls.HtmlInputHidden
    Protected WithEvents pnlStartSend, pnlEndSend As WebControls.Panel
    Protected WithEvents form1 As HtmlControls.HtmlForm
    Protected UserID, OrgID, query_groupID, prodId, pageId, sqlstr, UseBizLogo, lang_id, _
    PRODUCT_TYPE, fromEmail, fromName, quest_id, questName, pr_language, _
    subjectEmail, FILE_ATTACHMENT, PageTitle, formLink As String
    Dim pagesource, first_slice, second_slice As New System.Text.StringBuilder
    Protected notSendCount, currEmailCount As Long
    Protected ind, ind_start, ind_end As Integer
    Protected func As New bizpower.cfunc
    Protected blocked_flag As Boolean
    Protected emailCount As Long = 0
    Protected emailLimit As Long = 0
    Protected emailCountLeft As Long = 0
    Protected sendCount As Long = 0
    Protected badCount As Integer = 0
    Protected strLocal As String = ""
    Const count_to_send As Long = 200
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
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
        UserID = Trim(Trim(Request.Cookies("bizpegasus")("UserID")))
        OrgID = Trim(Trim(Request.Cookies("bizpegasus")("OrgID")))
        query_groupID = Trim(Request("send_groupId"))
        prodId = Trim(Request("prodId"))
        pageId = Trim(Request("pageId"))
        blocked_flag = False

        Server.ScriptTimeout = 6000

        strLocal = "http://" & Trim(Request.ServerVariables("SERVER_NAME"))
        If Len(Trim(Request.ServerVariables("SERVER_PORT"))) > 0 Then
            strLocal = strLocal & ":" & Trim(Request.ServerVariables("SERVER_PORT"))
        End If
        strLocal = strLocal & Application("VirDir") & "/"

        If IsNumeric(prodId) = True Then
            sqlstr = "Select From_Mail, FROM_NAME, EMAIL_SUBJECT, Page_ID, " & _
            " QUESTIONS_ID, FILE_ATTACHMENT From PRODUCTS Where product_id=" & prodId & _
            " AND ORGANIZATION_ID=" & OrgID
            cmdSelect = New SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            Dim rs_product As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
            If rs_product.Read() Then
                fromEmail = IIf(IsDBNull(rs_product("From_Name")), "", rs_product("From_Mail"))
                fromName = IIf(IsDBNull(rs_product("From_Name")), "", rs_product("From_Name"))
                pageId = IIf(IsDBNull(rs_product("Page_Id")), "", rs_product("Page_Id"))
                quest_id = IIf(IsDBNull(rs_product("QUESTIONS_ID")), "", rs_product("QUESTIONS_ID"))
                subjectEmail = IIf(IsDBNull(rs_product("EMAIL_SUBJECT")), "", rs_product("EMAIL_SUBJECT"))
                FILE_ATTACHMENT = IIf(IsDBNull(rs_product("FILE_ATTACHMENT")), "", rs_product("FILE_ATTACHMENT"))
            End If
            con.Close()

            If Len(Trim(fromEmail)) = 0 Then
                sqlstr = "Select EMAIL From USERS  WHERE USER_ID = " & UserID
                cmdSelect = New SqlClient.SqlCommand(sqlstr, con)
                con.Open()
                fromEmail = cmdSelect.ExecuteScalar()
                con.Close()
            End If

            If IsNumeric(quest_id) Then  ' יש טופס בדף	
                cmdSelect = New SqlClient.SqlCommand("Select Product_Name,FILE_ATTACHMENT,ATTACHMENT_TITLE  From Products Where product_id=" & quest_id & " and ORGANIZATION_ID=" & OrgID, con)
                con.Open()
                Dim dr_prod As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
                If dr_prod.Read() Then
                    questName = IIf(IsDBNull(dr_prod("Product_Name")), "", dr_prod("Product_Name"))
                End If
                con.Close()
            End If

            If IsNumeric(pageId) Then
                sqlstr = "SELECT Page_Title, FORM_LINK_IMAGE FROM Pages WHERE Page_Id=" & pageId
                cmdSelect = New SqlClient.SqlCommand(sqlstr, con)
                con.Open()
                Dim dr_page As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
                If dr_page.Read() Then
                    PageTitle = IIf(IsDBNull(dr_page("Page_Title")), "", dr_page("Page_Title"))
                    formLink = IIf(IsDBNull(dr_page("FORM_LINK_IMAGE")), "", dr_page("FORM_LINK_IMAGE"))
                End If
                con.Close()
            End If

        End If

        If Request.Form.Count = 0 Then

            sqlstr = "Select Email_Limit From Organizations WHERE ORGANIZATION_ID = " & OrgID
            cmdSelect = New SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            emailLimit = cmdSelect.ExecuteScalar()
            con.Close()
            hdnemailLimit.Value = Server.HtmlEncode(emailLimit)

            sqlstr = "Select UseBizLogo From Organizations WHERE ORGANIZATION_ID = " & OrgID
            cmdSelect = New SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            UseBizLogo = cmdSelect.ExecuteScalar()
            con.Close()
            hdnUseBizLogo.Value = Server.HtmlEncode(UseBizLogo)

            sqlstr = "SELECT Page_Lang FROM Pages WHERE Page_Id=" & pageId
            cmdSelect = New SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            Dim tmp = cmdSelect.ExecuteScalar()
            If Not IsDBNull(tmp) Then
                lang_id = Trim(tmp)
            Else
                lang_id = ""
            End If
            con.Close()

            hdnPageLang.Value = lang_id

            pagesource = New System.Text.StringBuilder

            sqlstr = "SELECT Page_Source FROM Pages WHERE Page_Id=" & pageId
            cmdSelect = New SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            tmp = cmdSelect.ExecuteScalar()
            If Not IsDBNull(tmp) Then
                pagesource.Append(tmp)
            End If
            con.Close()

            first_slice = New System.Text.StringBuilder : second_slice = New System.Text.StringBuilder

            If IsNumeric(quest_id) Then  ' יש טופס בדף	
                Dim image_name As String = "id=""tofes"""
                If pagesource.Length > 0 Then
                    ind = InStr(pagesource.ToString(), image_name)
                    If ind = 0 Then
                        image_name = "id=tofes"
                        ind = InStr(pagesource.ToString(), image_name)
                    End If
                    If ind > 0 Then
                        ind_start = InStrRev(pagesource.ToString(), "<", ind)
                        ind_end = InStr(ind, pagesource.ToString(), ">")
                        first_slice.Append(Mid(pagesource.ToString(), 1, ind_start - 1))
                        second_slice.Append(Mid(pagesource.ToString(), ind_end + 1))
                    End If
                End If
            End If

            If pagesource.Length > 0 Then
                hdnfirst_slice.Value = Server.HtmlEncode(first_slice.ToString())
                hdnsecond_slice.Value = Server.HtmlEncode(second_slice.ToString())
                hdnpagesource.Value = Server.HtmlEncode(pagesource.ToString())
            End If

            Dim cmdCount As New SqlClient.SqlCommand("get_send_details", con)
            cmdCount.CommandType = CommandType.StoredProcedure
            cmdCount.Parameters.Add("@prodId", SqlDbType.Int).Value = CLng(prodId)
            cmdCount.Parameters.Add("@countSend", 0).Direction = ParameterDirection.Output
            cmdCount.Parameters.Add("@countnotSend", 0).Direction = ParameterDirection.Output
            cmdCount.Parameters.Add("@countBad", 0).Direction = ParameterDirection.Output
            con.Open()
            cmdCount.ExecuteNonQuery()
            con.Close()

            If IsNumeric(cmdCount.Parameters("@countSend").Value) Then
                ViewState("sendCount") = CLng(cmdCount.Parameters("@countSend").Value)
            Else
                ViewState("sendCount") = 0
            End If
            If IsNumeric(cmdCount.Parameters("@countnotSend").Value) Then
                ViewState("notSendCount") = CLng(cmdCount.Parameters("@countnotSend").Value)
            Else
                ViewState("notSendCount") = 0
            End If
            If IsNumeric(cmdCount.Parameters("@countBad").Value) Then
                ViewState("badCount") = cmdCount.Parameters("@countBad").Value
            Else
                ViewState("badCount") = 0
            End If
        End If

        If IsNumeric(prodId) = True Then
            sendMail()
        End If

    End Sub

    Sub sendMail()

        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmdSelect As New SqlClient.SqlCommand
        Dim emailRemainder, currBadCount As Integer
        Dim strLink As String = ""
        Dim url, PEOPLE_ID, GROUPID, peopleEMAIL, peopleName As String
        Dim form_content As New System.Text.StringBuilder

        emailLimit = Server.HtmlDecode(hdnemailLimit.Value)
        UseBizLogo = Server.HtmlDecode(hdnUseBizLogo.Value)
        lang_id = Server.HtmlDecode(hdnPageLang.Value)

        first_slice = New System.Text.StringBuilder : first_slice.Append(Server.HtmlDecode(hdnfirst_slice.Value))
        second_slice = New System.Text.StringBuilder : second_slice.Append(Server.HtmlDecode(hdnsecond_slice.Value))
        pagesource = New System.Text.StringBuilder : pagesource.Append(Server.HtmlDecode(hdnpagesource.Value))
        pagesource.Replace(".", "&#46;")

        If IsNumeric(ViewState("sendCount")) Then
            sendCount = CLng(ViewState("sendCount"))
        Else
            sendCount = 0
        End If
        If IsNumeric(ViewState("notSendCount")) Then
            notSendCount = CLng(ViewState("notSendCount"))
        Else
            notSendCount = 0
        End If
        If IsNumeric(ViewState("badCount")) Then
            badCount = CLng(ViewState("badCount"))
        Else
            badCount = 0
        End If

        currEmailCount = notSendCount
        emailCountLeft = sendCount - notSendCount
        emailCount = 0

        If IsNumeric(emailLimit) Then
            emailLimit = CLng(emailLimit)
        End If

        Dim charset As String = "windows-1255"

        If currEmailCount > 0 And emailCount <= emailLimit Then

            currBadCount = 0
            Dim intCount As Integer = 0

            cmdSelect = New SqlClient.SqlCommand("getTopPeoplesToSend", con)
            cmdSelect.CommandType = CommandType.StoredProcedure
            cmdSelect.Parameters.Add("@count_to_send", SqlDbType.Int).Value = CLng(count_to_send)
            cmdSelect.Parameters.Add("@prodId", SqlDbType.Int).Value = CLng(prodId)
            con.Open()
            Dim dr_emails As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()
            While dr_emails.Read()
                intCount = intCount + 1

                PEOPLE_ID = Trim(dr_emails("PEOPLE_ID"))
                GROUPID = Trim(dr_emails("GROUPID"))
                peopleEMAIL = Trim(dr_emails("PEOPLE_EMAIL"))
                peopleName = Trim(dr_emails("PEOPLE_NAME"))

                '//start of sending Email			
                url = strLocal & "/sale.asp?" & func.Encode("prodId=" & prodId & "&pageId=" & pageId & "&C=" & PEOPLE_ID & "&O=" & OrgID)
                Dim strBody As New System.Text.StringBuilder("")
                strBody.Append("<HTML><HEAD><META http-equiv=Content-Type content=""text/html; charset=windows-1255"">") : strBody.Append(vbCrLf)
                strBody.Append("<title>") : strBody.Append(PageTitle) : strBody.Append("</title>") : strBody.Append(vbCrLf)
                strBody.Append("<style>BODY,P,TD {Font-Family:Arial; Font-Size:10pt}</style></head>") : strBody.Append(vbCrLf)
                strBody.Append("<body topmargin=5 leftmargin=0 rightmargin=0>") : strBody.Append(vbCrLf)
                strBody.Append("<center><P align=""center"" width=""100%"" style=""")
                strBody.Append(" font-size:10pt;font-weight:600;font-family:Arial"">If you can't see this letter properly, please ")
                strBody.Append("<A style=""color:blue;font-weight:600;text-decoration:underline;font-size=10pt"" href=""" & url)
                strBody.Append(""" target=""_blank"">click here</A> to see it on the web</P></center>") : strBody.Append(vbCrLf)
                strBody.Append("<table cellpadding=0 cellspacing=0 width=620 align=center border=0><tr><td>")
                form_content = New System.Text.StringBuilder
                If Trim(quest_id) <> "" Or IsDBNull(quest_id) = False Then
                    form_content = createForm(quest_id, prodId, pageId, PEOPLE_ID, OrgID)
                End If

                If first_slice.Length > 0 Or form_content.Length > 0 Or second_slice.Length > 0 Then
                    pagesource = New System.Text.StringBuilder
                    pagesource.Append(first_slice) : pagesource.Append(form_content) : pagesource.Append(second_slice)
                End If

                Dim OrgName As String
                OrgName = HttpUtility.UrlDecode(Trim(Request.Cookies("bizpegasus")("ORGNAME")), System.Text.Encoding.GetEncoding(1255))

                strBody.Append(pagesource)
                strBody.Append("</td></tr><tr><td align=""center"" bgcolor=""#FFFFFF""><table cellpadding=1 cellspacing=0 width=620 border=0 align=center>")
                strBody.Append("<tr><td height=5 nowrap></td></tr>") : strBody.Append(vbCrLf)
                strBody.Append("<tr><td align=right style=""font-family:Arial; font-size:8pt; font-weight:500; color:#6A6A6A; vertical-align:top;"" width=100% dir=ltr>")
                If Trim(UseBizLogo) = "1" Then
                    strBody.Append("<img src=""") : strBody.Append(strLocal) : strBody.Append("/images/")
                    strBody.Append(prodId) : strBody.Append("_") : strBody.Append(PEOPLE_ID) : strBody.Append(".aspx"" border=0 hspace=5 width=""1"" height=""18"">")
                Else
                    strBody.Append("<img src=""") : strBody.Append(strLocal) : strBody.Append("/images/") : strBody.Append(prodId)
                    strBody.Append("_") : strBody.Append(PEOPLE_ID) : strBody.Append(".aspx"" border=0 hspace=5 width=""1"" height=""18"">")
                End If
                strBody.Append("</td>")

                strBody.Append("<td align=right valign=""top"" nowrap style=""font-family:Arial; font-size:10pt; font-weight:500; color:#6A6A6A; vertical-align:top;"" dir=rtl>")
                If Trim(lang_id) = "1" Then 'Hebrew
                    strBody.Append("דיוור זה נשלח אליך מ") : strBody.Append(OrgName) : strBody.Append(" באמצעות מערכת ה")
                    strBody.Append("<a href=""http://pegasus.bizpower.co.il"">ניוזלטר</a>")
                    strBody.Append("&nbsp;<a href=""http://pegasus.bizpower.co.il"">BIZPOWER</a>")
                    strBody.Append(",<br>במידה ואינך מעוניין לקבל דיוור מ") : strBody.Append(OrgName)
                    strBody.Append("&nbsp;<a style=""font-family:Arial; font-size:10pt; font-weight:500; color:#6A6A6A; vertical-align:top;"" href=""") : strBody.Append(strLocal) : strBody.Append("/remove_email.asp?")
                    strBody.Append(func.Encode("OrgID=" & OrgID & "&PEOPLE_ID=" & PEOPLE_ID)) : strBody.Append(""" target=""_new"">לחץ כאן</a>")
                Else
                    strBody.Append("This mail was sent to you from&nbsp;") : strBody.Append(OrgName) : strBody.Append(",&nbsp;")
                    strBody.Append("<a style=""font-family:Arial; font-size:10pt; font-weight:500; color:#6A6A6A; vertical-align:top;"" href=""") : strBody.Append(strLocal) : strBody.Append("/remove_email.asp?")
                    strBody.Append(func.Encode("OrgID=" & OrgID & "&PEOPLE_ID=" & PEOPLE_ID)) : strBody.Append(""" target=""_new"">Click here</a>&nbsp;if you wish to receive no further emails from&nbsp;") : strBody.Append(OrgName)
                End If
                strBody.Append("&nbsp;</td></tr>") : strBody.Append(vbCrLf)
                strBody.Append("<tr><td height=10 nowrap></td></tr></table>") : strBody.Append(vbCrLf)
                strBody.Append("</td></tr></table></body></html>")

                'peopleEMAIL = "olga@eltam.com"
                'Response.Write(strBody)
                'Response.End()

                'Dim oMail As SmtpMail = New SmtpMail("ES-AA1141023508-00535-D2D646C730E540AC24BDDC671C4BB432")
                'Dim oSmtp As SmtpClient = New SmtpClient
                'Dim oServer As New SmtpServer(ConfigurationSettings.AppSettings.Item("smtp_server"))
                'oMail.Charset = charset

                'Dim errStr As String = ""
                'Dim IsServerValid As Boolean = False

                'Try
                '    Dim servers() As SmtpServer = DnsQueryEx.QueryServers(peopleEMAIL)
                '    If servers.Length >= 0 Then
                '        IsServerValid = True
                '    End If
                'Catch ex As Exception
                '    errStr = peopleEMAIL & " - not exists"
                '    IsServerValid = False
                'End Try

                'If IsServerValid Then

                '    If Len(fromName) > 0 Then
                '        oMail.From = New MailAddress(fromName, fromEmail)
                '    Else
                '        oMail.From = New MailAddress(fromEmail)
                '    End If

                '    'Please separate multiple addresses by comma(,)
                '    oMail.To = New AddressCollection(peopleEMAIL)

                '    'To avoid too many email addresses appear in To header, using the following code only
                '    'display the current recipient
                '    Dim strRecipient As String = ""
                '    If Len(peopleName) > 0 Then
                '        strRecipient = peopleName & " "
                '    End If
                '    strRecipient &= "<" & peopleEMAIL & ">"
                '    oMail.Headers.ReplaceHeader("To", """{$var_rcptname}"" <{$var_rcptaddr}>")
                '    oMail.Headers.ReplaceHeader("X-Rcpt-To", New AddressCollection(strRecipient).ToEncodedString(HeaderEncodingType.EncodingAuto, charset))

                '    oMail.Subject = subjectEmail
                '    oMail.HtmlBody = strBody.ToString()

                '    'if you want EASendMail service to send the email in the specified date, use the following code.
                '    oMail.Date = Now
                ' Response.Write(peopleEMAIL)
                '  Response.End()
                Dim objMail As New System.Web.Mail.MailMessage
                objMail.From = fromEmail
                objMail.To = peopleEMAIL

                objMail.Subject = subjectEmail
                objMail.Body = strBody.ToString()
                objMail.BodyEncoding = System.Text.Encoding.GetEncoding(1255)
                objMail.BodyFormat = Mail.MailFormat.Html

                If Len(Trim(FILE_ATTACHMENT)) > 0 Then
                    Dim fs As IO.File
                    If fs.Exists(Server.MapPath("../../../download/products/") & "/" & FILE_ATTACHMENT) Then
                        objMail.Attachments.Add(Server.MapPath("../../../download/products/") & "/" & FILE_ATTACHMENT)
                    End If
                    fs = Nothing
                End If



                Try
                    'System.Web.Mail.SmtpMail.SmtpServer = "jack"
                    System.Web.Mail.SmtpMail.SmtpServer = ConfigurationSettings.AppSettings.Item("smtp_server")
                    System.Web.Mail.SmtpMail.Send(objMail)
                Catch ex As Exception
                    ' Response.Write(peopleEMAIL & "<BR>")
                End Try
                objMail = Nothing

                'If Len(Trim(FILE_ATTACHMENT)) > 0 Then
                '    Dim fs As IO.File
                '    If fs.Exists(Server.MapPath("../../../download/products/") & "/" & FILE_ATTACHMENT) Then
                '        oMail.AddAttachment(Server.MapPath("../../../download/products/") & "/" & FILE_ATTACHMENT)
                '    End If
                '    fs = Nothing
                'End If

                '     oSmtp.SendMailToQueue(oServer, oMail)

                emailCount = emailCount + 1
                emailRemainder = emailLimit - emailCount

                Dim con_pr As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
                sqlstr = "Update PRODUCT_CLIENT Set DATE_SEND = getDate() Where PRODUCT_ID = @prodId" & _
                " And PEOPLE_ID = @peopleId And GROUPID = @groupId And ORGANIZATION_ID = @OrgID"
                Dim cmdUpdate As New SqlClient.SqlCommand(sqlstr, con_pr)
                cmdUpdate.CommandType = CommandType.Text
                cmdUpdate.Parameters.Add("@prodId", SqlDbType.Int).Value = CLng(prodId)
                cmdUpdate.Parameters.Add("@groupId", SqlDbType.Int).Value = CLng(GROUPID)
                cmdUpdate.Parameters.Add("@peopleId", SqlDbType.Int).Value = CLng(PEOPLE_ID)
                cmdUpdate.Parameters.Add("@OrgID", SqlDbType.Int).Value = CLng(OrgID)
                con_pr.Open()
                cmdUpdate.ExecuteNonQuery()
                con_pr.Close()

                '''Else 'Mark email as not valid email

                '''Dim con_pr As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
                '''sqlstr = "Update PRODUCT_CLIENT SET IsEmailValid = 0 WHERE (PRODUCT_ID = @prodId)" & _
                '''" AND (PEOPLE_ID = @peopleId) AND (GROUPID = @groupId) AND (ORGANIZATION_ID = @OrgID);" & _
                '''" UPDATE PEOPLES Set IsEmailValid = 0 WHERE (PEOPLE_ID = @peopleId) AND (GROUPID = @groupId) " & _
                '''" AND (ORGANIZATION_ID = @OrgID)"
                '''Dim cmdUpdate As New SqlClient.SqlCommand(sqlstr, con_pr)
                '''cmdUpdate.CommandType = CommandType.Text
                '''cmdUpdate.Parameters.Add("@prodId", SqlDbType.Int).Value = CLng(prodId)
                '''cmdUpdate.Parameters.Add("@groupId", SqlDbType.Int).Value = CLng(GROUPID)
                '''cmdUpdate.Parameters.Add("@peopleId", SqlDbType.Int).Value = CLng(PEOPLE_ID)
                '''cmdUpdate.Parameters.Add("@OrgID", SqlDbType.Int).Value = CLng(OrgID)
                '''con_pr.Open()
                '''cmdUpdate.ExecuteNonQuery()
                '''con_pr.Close()

                '''currBadCount = currBadCount + 1

                '''End If

                If IsNumeric(emailLimit) Then
                    If emailCount >= emailLimit Then 'End of organization email limit
                        Exit While
                    End If
                End If
            End While

            con.Close()


        End If

        'Update organization email limit
        cmdSelect = New SqlClient.SqlCommand("Update ORGANIZATIONS SET Email_Limit = Email_Limit - @emailCount WHERE ORGANIZATION_ID = @OrgID", con)
        cmdSelect.CommandType = CommandType.Text
        cmdSelect.Parameters.Add("@emailCount", SqlDbType.Int).Value = CLng(emailCount)
        cmdSelect.Parameters.Add("@OrgID", SqlDbType.Int).Value = CLng(OrgID)
        con.Open()
        cmdSelect.ExecuteNonQuery()
        con.Close()

        ViewState("badCount") = badCount + currBadCount
        ViewState("sendCount") = sendCount + (emailCount)
        ViewState("notSendCount") = notSendCount - (emailCount)

        If (emailCount = 0 And currBadCount = 0) Or (emailCount > emailLimit) Then  'End Of sending

            pnlStartSend.Visible = False
            pnlEndSend.Visible = True

            'Update send date
            Dim cmdUpd As New SqlClient.SqlCommand("updProductSendDate", con)
            cmdUpd.CommandType = CommandType.StoredProcedure
            cmdUpd.Parameters.Add("@prodId", SqlDbType.Int).Value = CLng(prodId)
            cmdUpd.Parameters.Add("@OrgID", SqlDbType.Int).Value = CLng(OrgID)
            con.Open()
            cmdUpd.ExecuteNonQuery()
            con.Close()

            Dim cScript As String = "<script language='javascript'> opener.location.href = 'products.asp'; </script>"
            RegisterStartupScript("parentwindow", cScript)
        Else 'Submit another time

            Dim cScript As String = "<script>" & Chr(10) & Chr(13) & _
            "document.forms(0).submit();" & Chr(10) & Chr(13) & _
            "</script>"
            RegisterStartupScript("submitscript", cScript)
        End If

    End Sub

    Function createForm(ByVal quest_id, ByVal prodId, ByVal pageId, ByVal PEOPLE_ID, ByVal OrgID)

        Dim strLink As String = strLocal & "/netcom/seker/seker.asp?" & func.Encode("P=" & prodId & "&C=" & PEOPLE_ID & "&O=" & OrgID)
        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmdSelect As New SqlClient.SqlCommand
        Dim pertext, perTextImgAlign, perbgcolor, pertype, persize, percolorname, percolor
        Dim form_content As New System.Text.StringBuilder

        If Trim(quest_id) <> "" And Trim(pageId) <> "" Then ' יש טופס בדף

            If Trim(formLink) = "image" Then 'image                 

                form_content.Append("<a ")
                If Trim(PEOPLE_ID) = "" Then
                    form_content.Append(" href=""#"" OnClick=""return false;""> ")
                Else
                    form_content.Append(" href=""" & func.vFix(strLink) & """ target=""_blank"">")
                End If
                form_content.Append(" <img src=""" & strLocal & "/netcom/GetImage.asp?DB=Page&FIELD=LINK_IMAGE&ID=" & pageId & """ border=""0""></a>")

            ElseIf Trim(formLink) = "link" Then 'text
                'Response.Write image_name
                'Response.End
                sqlstr = "select LINK_TEXT,LINK_TEXT_ALIGN,LINK_BGCOLOR,LINK_FONT_TYPE,LINK_FONT_SIZE,LINK_FONT_COLOR from Pages where Page_Id=" & pageId
                con.Open()
                cmdSelect = New SqlClient.SqlCommand(sqlstr, con)
                Dim dr_link As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
                If dr_link.Read() Then
                    pertext = Trim(dr_link("LINK_TEXT"))
                    perTextImgAlign = Trim(dr_link("LINK_TEXT_ALIGN"))
                    perbgcolor = Trim(dr_link("LINK_BGCOLOR"))
                    pertype = Trim(dr_link("LINK_FONT_TYPE"))
                    persize = Trim(dr_link("LINK_FONT_SIZE"))
                    percolorname = Trim(dr_link("LINK_FONT_COLOR"))
                    percolor = Trim(dr_link("LINK_FONT_COLOR"))
                End If
                dr_link.Close()
                con.Close()

                If Len(pertext) = 0 Or IsDBNull(pertext) Then
                    pertext = questName
                End If
                If Len(percolor) = 0 Or IsDBNull(percolor) Then
                    percolor = "#000000"
                End If
                If Len(pertype) = 0 Or IsDBNull(pertype) Then
                    pertype = "STRONG"
                End If
                If Len(persize) = 0 Or IsDBNull(persize) Then
                    persize = "2"
                End If
                If Len(perbgcolor) = 0 Or IsDBNull(perbgcolor) Then
                    perbgcolor = "transparent"
                End If
                If Len(perTextImgAlign) = 0 Or IsDBNull(perTextImgAlign) Then
                    perTextImgAlign = "center"
                End If

                form_content = New System.Text.StringBuilder

                form_content.Append("<p align=""" & perTextImgAlign & """ dir=rtl width=100% bgcolor=""" & perbgcolor & """>" & _
                "&nbsp;<a bgcolor=""" & perbgcolor & """ target=""_blank"" style=""letter-spacing:1px""")
                If Trim(PEOPLE_ID) = "" Then
                    form_content.Append(" href=""#"" OnClick=""return false;"" ")
                Else
                    form_content.Append(" href=" & func.vFix(strLink))
                End If
                form_content.Append("><font color=""" & percolor & """ size=""" & persize & """>")
                If Trim(pertype) = "I" Then
                    form_content.Append("<i>")
                ElseIf Trim(pertype) = "STRONG" Then
                    form_content.Append("<b>")
                ElseIf Trim(pertype) = "SUB" Then
                    form_content.Append("<SUB>")
                ElseIf Trim(pertype) = "SUP" Then
                    form_content.Append("<SUP>")
                ElseIf Trim(pertype) = "STRIKE" Then
                    form_content.Append("<STRIKE>")
                End If

                form_content.Append(pertext)

                If Trim(pertype) = "I" Then
                    form_content.Append("</i>")
                ElseIf Trim(pertype) = "STRONG" Then
                    form_content.Append("</b>")
                ElseIf Trim(pertype) = "SUB" Then
                    form_content.Append("</SUB>")
                ElseIf Trim(pertype) = "SUP" Then
                    form_content.Append("</SUP>")
                ElseIf Trim(pertype) = "STRIKE" Then
                    form_content.Append("</STRIKE>")
                End If

                form_content.Append("</font></a>&nbsp;</p>")
            End If
        End If

        createForm = form_content
    End Function

End Class
