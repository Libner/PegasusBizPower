Imports System.Text.RegularExpressions
Public Class sendMail
    Inherits System.Web.UI.Page
    Protected hdnemailLimit, hdnfirst_slice, hdnsecond_slice, hdnpagesource, hdnUseBizLogo, hdnPageLang As HtmlControls.HtmlInputHidden
    Protected WithEvents pnlStartSend, pnlEndSend As WebControls.Panel
    Protected WithEvents form1 As HtmlControls.HtmlForm
    Protected UserID, OrgID, query_groupID, prodId, pageId, sqlstr, UseBizLogo, PageLang, _
    PRODUCT_TYPE, fromEmail, fromName, quest_id, questName, pr_language, _
    subjectEmail, FILE_ATTACHMENT, PageTitle, pagesource, formLink, first_slice, second_slice As String
    Protected totalCount, currEmailCount, ind, ind_start, ind_end As Integer
    Protected func As New bizpower.cfunc
    Protected blocked_flag As Boolean
    Protected emailCount As Integer = 0
    Protected emailLimit As Integer = 0
    Protected emailCountLeft As Integer = 0
    Protected sendCount As Integer = 0
    Protected strLocal As String = ""
    Const count_to_send As Integer = 100
    Dim con As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmdSelect As New System.Data.SqlClient.SqlCommand

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
            cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
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
                cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
                con.Open()
                fromEmail = cmdSelect.ExecuteScalar()
                con.Close()
            End If

            If IsNumeric(quest_id) Then  ' יש טופס בדף	
                cmdSelect = New System.Data.SqlClient.SqlCommand("Select Product_Name,FILE_ATTACHMENT,ATTACHMENT_TITLE  From Products Where product_id=" & quest_id & " and ORGANIZATION_ID=" & OrgID, con)
                con.Open()
                Dim dr_prod As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
                If dr_prod.Read() Then
                    questName = IIf(IsDBNull(dr_prod("Product_Name")), "", dr_prod("Product_Name"))
                End If
                con.Close()
            End If

            If IsNumeric(pageId) Then
                sqlstr = "SELECT Page_Title, FORM_LINK_IMAGE FROM Pages WHERE Page_Id=" & pageId
                cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
                con.Open()
                Dim dr_page As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
                If dr_page.Read() Then
                    PageTitle = IIf(IsDBNull(dr_page("Page_Title")), "", dr_page("Page_Title"))
                    formLink = IIf(IsDBNull(dr_page("FORM_LINK_IMAGE")), "", dr_page("FORM_LINK_IMAGE"))
                End If
                con.Close()
            End If

        End If

        If Not Page.IsPostBack Then

            sqlstr = "Select Email_Limit From Organizations WHERE ORGANIZATION_ID = " & OrgID
            cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            emailLimit = cmdSelect.ExecuteScalar()
            con.Close()
            hdnemailLimit.Value = Server.HtmlEncode(emailLimit)

            sqlstr = "Select UseBizLogo From Organizations WHERE ORGANIZATION_ID = " & OrgID
            cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            UseBizLogo = cmdSelect.ExecuteScalar()
            con.Close()
            hdnUseBizLogo.Value = Server.HtmlEncode(UseBizLogo)

            sqlstr = "SELECT Page_Lang FROM Pages WHERE Page_Id=" & pageId
            cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            Dim tmp = cmdSelect.ExecuteScalar()
            If Not IsDBNull(tmp) Then
                PageLang = Trim(tmp)
            Else
                PageLang = ""
            End If
            con.Close()

            hdnPageLang.Value = PageLang

            sqlstr = "SELECT Page_Source FROM Pages WHERE Page_Id=" & pageId
            cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            tmp = cmdSelect.ExecuteScalar()
            If Not IsDBNull(tmp) Then
                pagesource = CStr(tmp)
            Else
                pagesource = ""
            End If
            con.Close()

            Dim image_name As String = "id=""tofes"""
            If Len(pagesource) > 0 Then
                ind = InStr(pagesource, image_name)
                If ind = 0 Then
                    image_name = "id=tofes"
                    ind = InStr(pagesource, image_name)
                End If
                If ind > 0 Then
                    ind_start = InStrRev(pagesource, "<", ind)
                    ind_end = InStr(ind, pagesource, ">")
                    first_slice = Mid(pagesource, 1, ind_start - 1)
                    second_slice = Mid(pagesource, ind_end + 1)
                End If

                hdnfirst_slice.Value = Server.HtmlEncode(first_slice)
                hdnsecond_slice.Value = Server.HtmlEncode(second_slice)
                hdnpagesource.Value = Server.HtmlEncode(pagesource)
            End If
        End If

        If IsNumeric(prodId) = True Then
            sendMail()
        End If

    End Sub

    Sub sendMail()

        Dim con As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmdSelect As New System.Data.SqlClient.SqlCommand
        Dim emailRemainder As Integer
        Dim strLink As String = ""
        Dim url, form_content, PEOPLE_ID, GROUPID, peopleEMAIL As String
        Dim dt_emails As DataTable

        sqlstr = "Select TOP " & CStr(count_to_send) & " PEOPLE_ID, PEOPLE_EMAIL, GROUPID FROM PRODUCT_CLIENT WHERE " & _
               " PRODUCT_ID = " & prodId & " And DATE_SEND IS NULL Order BY GROUPID, PEOPLE_ID"
        cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
        con.Open()
        Dim da As New SqlClient.SqlDataAdapter(cmdSelect)
        dt_emails = New DataTable
        dt_emails.BeginLoadData()
        da.Fill(dt_emails)
        dt_emails.EndLoadData()
        da.Dispose()
        con.Close()

        emailLimit = Server.HtmlDecode(hdnemailLimit.Value)
        UseBizLogo = Server.HtmlDecode(hdnUseBizLogo.Value)
        PageLang = Server.HtmlDecode(hdnPageLang.Value)
        first_slice = Server.HtmlDecode(hdnfirst_slice.Value)
        second_slice = Server.HtmlDecode(hdnsecond_slice.Value)
        pagesource = Replace(Server.HtmlDecode(hdnpagesource.Value), ".", "&#46;")

        sqlstr = "Select Count(PEOPLE_ID) FROM PRODUCT_CLIENT WHERE " & _
               " PRODUCT_ID = " & prodId & " And DATE_SEND IS NOT NULL"
        Dim cmdCount As New SqlClient.SqlCommand(sqlstr, con)
        con.Open()
        sendCount = cmdCount.ExecuteScalar()
        con.Close()

        sqlstr = "Select Count(PEOPLE_ID) FROM PRODUCT_CLIENT WHERE " & _
               " PRODUCT_ID = " & prodId & " And DATE_SEND IS NULL"
        cmdCount = New SqlClient.SqlCommand(sqlstr, con)
        con.Open()
        totalCount = cmdCount.ExecuteScalar()
        con.Close()

        currEmailCount = dt_emails.Rows.Count
        emailCountLeft = totalCount - dt_emails.Rows.Count
        emailCount = 0

        If IsNumeric(emailLimit) Then
            emailLimit = CLng(emailLimit)
        End If

        If currEmailCount > 0 And emailCount <= emailLimit Then
            For pp As Integer = 0 To dt_emails.Rows.Count - 1
                PEOPLE_ID = Trim(dt_emails.Rows(pp)("PEOPLE_ID"))
                GROUPID = Trim(dt_emails.Rows(pp)("GROUPID"))
                peopleEMAIL = Trim(dt_emails.Rows(pp)("PEOPLE_EMAIL"))

                'Response.Write(ex.Message & vbCrLf & "currEmailCount=" & currEmailCount & " emailCount= " & emailCount & _
                '" emailLimit= " & emailLimit & " totalCount=" & totalCount & " count_to_send = " & _
                'count_to_send & " dt_emails.Rows.Count = " & dt_emails.Rows.Count)
                'Response.End()

                '//start of sending Email			
                url = strLocal & "/sale.asp?" & func.Encode("prodId=" & prodId & "&pageId=" & pageId & "&C=" & PEOPLE_ID & "&O=" & OrgID)
                Dim Body_message As New System.Text.StringBuilder("")
                Body_message.Append("<HTML><HEAD><META http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & Chr(13) & Chr(10))
                Body_message.Append("<title>" & PageTitle & "</title>" & Chr(13) & Chr(10))
                Body_message.Append("<style>Body,P,TD {Font-Family:Arial; Font-Size:10pt}</style></head>" & Chr(13) & Chr(10))
                Body_message.Append("<body topmargin=5 leftmargin=0 rightmargin=0>" & Chr(13) & Chr(10))
                Body_message.Append("<center><P align=""center"" width=""100%"" style=""")
                Body_message.Append(" font-size:10pt;font-weight:600;font-family:Arial"">If you can't see this letter properly, please ")
                Body_message.Append("<A style=""color:blue;font-weight:600;text-decoration:underline;font-size=10pt"" href=""" & url)
                Body_message.Append(""" target=""_blank"">click here</A> to see it on the web</P></center>" & Chr(13) & Chr(10))
                Body_message.Append("<table cellpadding=0 cellspacing=0 width=620 align=center border=0><tr><td>")
                form_content = ""
                If Trim(quest_id) <> "" Or IsDBNull(quest_id) = False Then
                    form_content = createForm(quest_id, prodId, pageId, PEOPLE_ID, OrgID)
                End If

                If Len(first_slice) > 0 Or Len(form_content) > 0 Or Len(second_slice) > 0 Then
                    pagesource = first_slice & form_content & second_slice
                End If

                Dim OrgName As String
                OrgName = HttpUtility.UrlDecode(Trim(Request.Cookies("bizpegasus")("ORGNAME")), System.Text.Encoding.GetEncoding(1255))

                Body_message.Append(pagesource)
                Body_message.Append("</td></tr><tr><td align=""center"" bgcolor=""#FFFFFF""><table cellpadding=1 cellspacing=0 width=620 border=0 align=center>")
                Body_message.Append("<tr><td height=5 nowrap></td></tr>" & Chr(13) & Chr(10))
                Body_message.Append("<tr><td align=right style=""font-family:Arial (Hebrew); font-size:8pt; font-weight:500; color:#6A6A6A; vertical-align:top;"" width=100% dir=ltr>")
                If Trim(UseBizLogo) = "1" Then
                    Body_message.Append("<A style=""font-family:Arial (Hebrew); font-size:8pt; font-weight:500; color:#6A6A6A; vertical-align:top;"" href=""")
                    Body_message.Append("http://pegasus.bizpower.co.il"" target=""_blank"" dir=ltr><img src=""" & strLocal & "/images/")
                    Body_message.Append(prodId & "_" & PEOPLE_ID & ".aspx"" border=0 hspace=5></A>")
                Else
                    Body_message.Append("<img src=""" & strLocal & "/images/" & prodId & "_" & PEOPLE_ID & ".aspx"" border=0 hspace=5>")
                End If
                Body_message.Append("</td>")

                Dim strTmp As String
                If Trim(PageLang) = "1" Then 'Hebrew
                    strTmp = "דואר זה נשלח אליך מ" & OrgName & ",&nbsp;במידה ואינך מעוניין לקבל דיוור מ" & OrgName & _
                    "&nbsp;<a style=""font-family:Arial (Hebrew); font-size:8pt; font-weight:500; color:#6A6A6A; vertical-align:top;"" href=""" & strLocal & "/remove_email.asp?" & _
                    func.Encode("OrgID=" & OrgID & "&PEOPLE_ID=" & PEOPLE_ID) & """ target=""_new"">לחץ כאן</a>"
                Else
                    strTmp = "This mail was sent to you from&nbsp;" & OrgName & ",&nbsp;" & _
                    "<a style=""font-family:Arial (Hebrew); font-size:8pt; font-weight:500; color:#6A6A6A; vertical-align:top;"" href=""" & strLocal & "/remove_email.asp?" & _
                    func.Encode("OrgID=" & OrgID & "&PEOPLE_ID=" & PEOPLE_ID) & """ target=""_new"">Click here</a>&nbsp;to unsubscribe"
                End If

                Body_message.Append("<td align=right valign=""top"" nowrap style=""font-family:Arial (Hebrew); font-size:8pt; font-weight:500; color:#6A6A6A; vertical-align:top;"" dir=rtl>")
                Body_message.Append(strTmp)
                Body_message.Append("&nbsp;</td></tr>" & Chr(13) & Chr(10))
                Body_message.Append("<tr><td height=10 nowrap></td></tr></table>" & Chr(13) & Chr(10))
                Body_message.Append("</td></tr></table></body></html>")

                'peopleEMAIL = "olga@eltam.com"
                'Response.Write(Body_message)
                'Response.End()

                If RegExpTest(peopleEMAIL) Then
                    Dim objMail As System.Web.Mail.MailMessage
                    objMail = New System.Web.Mail.MailMessage

                    If Len(fromName) > 0 Then
                        objMail.From = "(" + fromName + ") " + fromEmail
                    Else
                        objMail.From = fromEmail
                    End If
                    objMail.To = peopleEMAIL
                    objMail.Subject = subjectEmail
                    objMail.Body = Body_message.ToString()
                    objMail.BodyEncoding = System.Text.Encoding.GetEncoding(1255)
                    objMail.BodyFormat = Mail.MailFormat.Html
                    If Len(Trim(FILE_ATTACHMENT)) > 0 Then
                        Dim fs As IO.File
                        If fs.Exists(Server.MapPath("../../../download/products/") & "/" & FILE_ATTACHMENT) Then
                            objMail.Attachments.Add(New Mail.MailAttachment(Server.MapPath("../../../download/products/") & "/" & FILE_ATTACHMENT))
                        End If
                        fs = Nothing
                    End If
                    System.Web.Mail.SmtpMail.Send(objMail)

                    sqlstr = "Update PRODUCT_CLIENT Set DATE_SEND = getDate() Where PRODUCT_ID = " & prodId & _
                    " And PEOPLE_ID = " & PEOPLE_ID & " And GROUPID = " & GROUPID & " And ORGANIZATION_ID = " & OrgID
                    cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
                    con.Open()
                    cmdSelect.ExecuteNonQuery()
                    con.Close()

                    emailCount = emailCount + 1

                    emailRemainder = emailLimit - emailCount
                    sqlstr = "Update ORGANIZATIONS SET Email_Limit = Email_Limit - 1 WHERE ORGANIZATION_ID = " & OrgID
                    cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
                    con.Open()
                    cmdSelect.ExecuteNonQuery()
                    con.Close()

                End If

                If IsNumeric(emailLimit) Then
                    If emailCount >= emailLimit Then
                        Exit For
                    End If
                End If
            Next
        End If

        If IsNumeric(GROUPID) Then
            sqlstr = "Select groupID From PRODUCT_GROUPS WHERE PRODUCT_ID = " & prodId & _
            " And groupID = " & GROUPID & " And ORGANIZATION_ID = " & OrgID & " And DATE_SEND is NULL"
            cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            Dim rs_group As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleResult)
            If Not rs_group.HasRows Then
                sqlstr = "Update PRODUCT_GROUPS Set DATE_SEND = getDate() Where PRODUCT_ID = " & prodId & _
                " And groupID = " & GROUPID
                Dim con1 As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
                Dim cmdInsert As New System.Data.SqlClient.SqlCommand(sqlstr, con1)
                con1.Open()
                cmdInsert.ExecuteNonQuery()
                con1.Close()
            End If
            con.Close()
        End If

        sqlstr = "Update Products Set DATE_SEND_END = getDate() Where PRODUCT_ID = " & prodId & _
        " And ORGANIZATION_ID = " & OrgID
        Dim con_pr As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmdUpd As New System.Data.SqlClient.SqlCommand(sqlstr, con_pr)
        con_pr.Open()
        cmdUpd.ExecuteNonQuery()
        con_pr.Close()

        If (emailCount = 0) Or (emailCount > emailLimit) Then
            pnlStartSend.Visible = False
            pnlEndSend.Visible = True

            Dim cScript As String = "<script language='javascript'> opener.location.href = 'products.asp'; </script>"
            RegisterStartupScript("parentwindow", cScript)
        Else
            Dim cScript As String = "<script>" & Chr(10) & Chr(13) & _
            "document.forms(0).submit();" & Chr(10) & Chr(13) & _
            "</script>"
            RegisterStartupScript("submitscript", cScript)
        End If

    End Sub

    Function RegExpTest(ByVal sEmail) As Boolean
        Dim regEx As Regex
        Dim retVal As Boolean
        regEx = New Regex("^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$")

        retVal = regEx.IsMatch(sEmail)

        RegExpTest = retVal
    End Function

    Function createForm(ByVal quest_id, ByVal prodId, ByVal pageId, ByVal PEOPLE_ID, ByVal OrgID)

        Dim strLink As String = strLocal & "/netcom/seker/seker.asp?" & func.Encode("P=" & prodId & "&C=" & PEOPLE_ID & "&O=" & OrgID)
        Dim con As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmdSelect As New System.Data.SqlClient.SqlCommand
        Dim form_content, pertext, perTextImgAlign, perbgcolor, pertype, persize, _
        percolorname, percolor

        If Trim(quest_id) <> "" And Trim(pageId) <> "" Then ' יש טופס בדף

            If Trim(formLink) = "image" Then 'image                 

                form_content = "<a "
                If Trim(PEOPLE_ID) = "" Then
                    form_content = form_content & " href=""#"" OnClick=""return false;""> "
                Else
                    form_content = form_content & " href=""" & func.vFix(strLink) & """ target=""_blank"">"
                End If
                form_content = form_content & " <img src=""" & strLocal & "/netcom/GetImage.asp?DB=Page&FIELD=LINK_IMAGE&ID=" & pageId & """ border=""0""></a>"

            ElseIf Trim(formLink) = "link" Then 'text
                'Response.Write image_name
                'Response.End
                sqlstr = "select LINK_TEXT,LINK_TEXT_ALIGN,LINK_BGCOLOR,LINK_FONT_TYPE,LINK_FONT_SIZE,LINK_FONT_COLOR from Pages where Page_Id=" & pageId
                con.Open()
                cmdSelect = New System.Data.SqlClient.SqlCommand(sqlstr, con)
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

                form_content = "<p align=""" & perTextImgAlign & """ dir=rtl width=100% bgcolor=""" & perbgcolor & """>" & _
                "&nbsp;<a bgcolor=""" & perbgcolor & """ target=""_blank"" style=""letter-spacing:1px"""
                If Trim(PEOPLE_ID) = "" Then
                    form_content = form_content & " href=""#"" OnClick=""return false;"" "
                Else
                    form_content = form_content & " href=" & func.vFix(strLink)
                End If
                form_content = form_content & "><font color=""" & percolor & """ size=""" & persize & """>"
                If Trim(pertype) = "I" Then
                    form_content = form_content & "<i>"
                ElseIf Trim(pertype) = "STRONG" Then
                    form_content = form_content & "<b>"
                ElseIf Trim(pertype) = "SUB" Then
                    form_content = form_content & "<SUB>"
                ElseIf Trim(pertype) = "SUP" Then
                    form_content = form_content & "<SUP>"
                ElseIf Trim(pertype) = "STRIKE" Then
                    form_content = form_content & "<STRIKE>"
                End If

                form_content = form_content & pertext

                If Trim(pertype) = "I" Then
                    form_content = form_content & "</i>"
                ElseIf Trim(pertype) = "STRONG" Then
                    form_content = form_content & "</b>"
                ElseIf Trim(pertype) = "SUB" Then
                    form_content = form_content & "</SUB>"
                ElseIf Trim(pertype) = "SUP" Then
                    form_content = form_content & "</SUP>"
                ElseIf Trim(pertype) = "STRIKE" Then
                    form_content = form_content & "</STRIKE>"
                End If

                form_content = form_content & "</font></a>&nbsp;</p>"
            End If
        End If

        createForm = form_content
    End Function

End Class
