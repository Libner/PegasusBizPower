Public Class editText
    Inherits System.Web.UI.Page
    Public copypageId, pageSource, pageId, templateId, pwidth, pageTitle, _
    TemplateTitle, TemplateSource, userId, orgID, strLocal, num, sqlstr, _
    count_pages, procent, user_name, org_name, str_confirm, PageLang, _
    str_alert, page_title_, perSize, temp_text As String
    Protected WithEvents SqlConn, conn1 As System.Data.SqlClient.SqlConnection
    Protected WithEvents catCMD As System.Data.SqlClient.SqlCommand
    Public myReader, rstitle, rsbuttons As System.Data.SqlClient.SqlDataReader
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents Button7 As System.Web.UI.WebControls.Button
    Protected WithEvents Button6 As System.Web.UI.WebControls.Button
    Protected WithEvents title1 As System.Web.UI.HtmlControls.HtmlTextArea
    Protected WithEvents pageId1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents template_id As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents pageTitle1 As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents template_title As System.Web.UI.WebControls.Label
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents PHAlert As System.Web.UI.WebControls.PlaceHolder
    Public dir_var, align_var, dir_obj_var, lang_id As String
    Public arrTitles(1), arrButtons(1)
    Public title_count, button_count As Integer

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
        strLocal = "http://" & Trim(Request.ServerVariables("SERVER_NAME"))
        If Len(Trim(Request.ServerVariables("SERVER_PORT"))) > 0 Then
            strLocal = strLocal & ":" & Trim(Request.ServerVariables("SERVER_PORT"))
        End If
        strLocal = strLocal & Application("VirDir") & "/"
        userId = Trim(Request.Cookies("bizpegasus")("UserId"))
        orgID = Trim(Request.Cookies("bizpegasus")("orgId"))
        perSize = Trim(Request.Cookies("bizpegasus")("perSize"))
        lang_id = Trim(Request.Cookies("bizpegasus")("LANGID"))

        If IsNumeric(lang_id) = False Or lang_id = "" Then
            lang_id = "1"
        End If
        If lang_id = "2" Then
            dir_var = "rtl" : align_var = "left" : dir_obj_var = "ltr" : temp_text = "No template"
        Else
            dir_var = "ltr" : align_var = "right" : dir_obj_var = "rtl" : temp_text = "ללא תבנית"
        End If
        pageTitle1.Style.Add("direction", dir_obj_var)
        If Trim(pageId) = "" Or Trim(templateId) = "" Then
            template_title.Text = temp_text
        Else
            template_title.Text = ""
        End If

        If IsNumeric(Request.Form("wpageID")) = True Then
            pageId = Request.Form("wpageID")
        ElseIf IsNumeric(Request("pageID")) = True Then
            pageId = Request("pageID")
        Else
            pageId = pageId1.Value
            'Response.Write(pageId)
            'Response.End()
        End If

        templateId = Request("template_id")
        org_name = Request.Cookies("bizpegasus")("ORGNAME")
        org_name = HttpContext.Current.Server.UrlDecode(org_name)
        user_name = Trim(Request.Cookies("bizpegasus")("UserName"))
        user_name = HttpContext.Current.Server.UrlDecode(user_name)

        If Trim(pageId) <> "" Then
            pageId1.Value = pageId
        End If

        If Trim(templateId) <> "" Then
            template_id.Value = templateId
        End If

        conn1 = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        sqlstr = "Select word_id, word From translate Where lang_id = " & lang_id & " And page_id = 26 Order By word_id"
        catCMD = New System.Data.SqlClient.SqlCommand(sqlstr, conn1)
        conn1.Open()
        rstitle = catCMD.ExecuteReader()
        ReDim arrTitles(rstitle.RecordsAffected)
        title_count = 0
        While rstitle.Read()
            title_count = title_count + 1
            ReDim Preserve arrTitles(title_count)
            arrTitles(Trim(rstitle("word_id"))) = Trim(rstitle("word"))
        End While
        rstitle.Close()
        conn1.Close()

        sqlstr = "select button_id, value From buttons Where lang_id = " & lang_id & " Order By button_id"
        catCMD = New System.Data.SqlClient.SqlCommand(sqlstr, conn1)
        conn1.Open()
        rsbuttons = catCMD.ExecuteReader()
        button_count = 0
        While rsbuttons.Read()
            button_count = button_count + 1
            ReDim Preserve arrButtons(button_count)
            arrButtons(Trim(rsbuttons("button_id"))) = Trim(rsbuttons("value"))
        End While
        rstitle.Close()
        conn1.Close()
        Button2.Text = arrButtons(2) : Button1.Text = arrButtons(1) : Button6.Text = arrButtons(6) : Button7.Text = arrButtons(7)


        SqlConn = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        If Not IsPostBack Then
            pwidth = "610" : PageLang = "1"
            If IsNumeric(Trim(pageId)) = True Then
                'Response.Write(pageId)
                'Response.End()
                catCMD = New System.Data.SqlClient.SqlCommand("Select Page_Title, Page_Source, Page_Lang From Pages Where Page_Id=" & pageId & " And Organization_id = " & orgID, SqlConn)
                SqlConn.Open()
                myReader = catCMD.ExecuteReader(CommandBehavior.SingleRow)

                If myReader.Read() Then

                    If Not myReader("Page_Source") Is DBNull.Value Then
                        pageSource = Trim(myReader("Page_Source"))
                    Else
                        pageSource = "<p dir=rtl>&nbsp;</p>"
                    End If

                    If Not myReader("Page_Title") Is DBNull.Value Then
                        pageTitle = Trim(myReader("Page_Title"))
                    Else
                        pageTitle = ""
                    End If

                    If Not myReader("Page_Lang") Is DBNull.Value Then
                        PageLang = Trim(myReader("Page_Lang"))
                    Else
                        PageLang = ""
                    End If
                Else
                    Response.Redirect("editPage.aspx")
                End If

                myReader.Close()
                SqlConn.Close()
                catCMD.Dispose()

                title1.Style.Item("WIDTH") = "645px"
                'title1.Style.Item("HEIGHT") = "600px"
                title1.Value = pageSource

                pageTitle1.Value = pageTitle

            End If
            If Trim(copypageId) <> "" Then
                catCMD = New System.Data.SqlClient.SqlCommand("Select Page_Source From Pages Where Page_Id=" & copypageId & " And Organization_id = " & orgID, SqlConn)
                SqlConn.Open()
                myReader = catCMD.ExecuteReader(CommandBehavior.SingleRow)

                If myReader.Read() Then

                    If Not myReader("Page_Source") Is DBNull.Value Then
                        pageSource = Trim(myReader("Page_Source"))
                    Else
                        pageSource = "<p dir=rtl>&nbsp;</p>"
                    End If

                End If

                myReader.Close()
                SqlConn.Close()
                catCMD.Dispose()

                title1.Style.Item("WIDTH") = "645px"
                'title1.Style.Item("HEIGHT") = "600px"
                title1.Value = pageSource

                'pageTitle1.Value = pageTitle
            End If
        End If

        If Trim(templateId) <> "" Then

            'Response.Write(templateId)
            'Response.End()

            SqlConn = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            catCMD = New System.Data.SqlClient.SqlCommand("Select Template_Title,Template_Source From Templates WHERE Template_ID = " & templateId, SqlConn)
            SqlConn.Open()
            myReader = catCMD.ExecuteReader(CommandBehavior.SingleRow)

            If myReader.Read() Then

                If Not myReader("Template_Source") Is DBNull.Value Then
                    TemplateSource = Trim(myReader("Template_Source"))
                Else
                    TemplateSource = ""
                End If

                If Not myReader("Template_Title") Is DBNull.Value Then
                    TemplateTitle = Trim(myReader("Template_Title"))
                Else
                    TemplateTitle = ""
                End If

            End If

            myReader.Close()
            SqlConn.Close()
            catCMD.Dispose()

            title1.Style.Item("WIDTH") = "645px"
            'title1.Style.Item("HEIGHT") = "600px"
            title1.Value = TemplateSource
            template_id.Value = "" 'להוריד העתקת תבנית 
            template_title.Text = TemplateTitle
        End If

        If IsNumeric(pageId) Then
            page_title_ = "עדכון דף מעוצב"
        Else
            page_title_ = "בניית דף מעוצב"
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Trim(pageTitle1.Value) = "" Then
            Dim stralert As String
            If Trim(lang_id) = "1" Then
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('חובה למלא שם של דף מעוצב');" & vbCrLf
                stralert += "</script>" & vbCrLf
            Else
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('Please insert the page name');" & vbCrLf
                stralert += "</script>" & vbCrLf
            End If

            PHAlert.Controls.Add(New LiteralControl(stralert))
        ElseIf Trim(title1.Value) = "" Then
            Dim stralert As String
            If Trim(lang_id) = "1" Then
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('חובה למלא את השדה טקסט');" & vbCrLf
                stralert += "</script>" & vbCrLf
            Else
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('Please insert the page content');" & vbCrLf
                stralert += "</script>" & vbCrLf
            End If
            PHAlert.Controls.Add(New LiteralControl(stralert))
        Else
            catCMD = New System.Data.SqlClient.SqlCommand("updatePageContent", SqlConn)
            catCMD.CommandType = CommandType.StoredProcedure
            If Trim(pageId) = "" Or IsNumeric(pageId) = False Then
                pageId = 0
            End If

            catCMD.Parameters.Add("@pageId", pageId)
            catCMD.Parameters.Add("@orgId", orgID)
            catCMD.Parameters.Add("@userId", userId)
            catCMD.Parameters.Add("@PageLang", Trim(Request.Form("PageLang")))
            catCMD.Parameters.Add("@pageTitle", pageTitle1.Value)
            catCMD.Parameters.Add("@pageSource", title1.InnerText)

            SqlConn.Open()
            catCMD.ExecuteNonQuery()
            SqlConn.Close()
            catCMD.Dispose()

            'אם נמחק טופס לנקות שדות
            If InStr(title1.InnerText, "tofes") = 0 And IsNumeric(pageId) Then
                sqlstr = "UPDATE Pages SET Product_Id = NULL, FORM_LINK_IMAGE=NULL,FORM_LINK=NULL," & _
                " LINK_IMAGE=NULL, LINK_IMAGE_ALIGN=NULL,LINK_TEXT=NULL,LINK_TEXT_ALIGN=NULL," & _
                " LINK_BGCOLOR=NULL, LINK_FONT_TYPE=NULL, LINK_FONT_SIZE=NULL, LINK_FONT_COLOR=NULL," & _
                " FORM_BGCOLOR=NULL, FORM_FONT_TYPE=NULL, FORM_FONT_SIZE=NULL, FORM_FONT_COLOR=NULL" & _
                " WHERE PAGE_ID = " & pageId
                catCMD = New System.Data.SqlClient.SqlCommand(sqlstr, SqlConn)
                SqlConn.Open()
                catCMD.ExecuteReader(CommandBehavior.Default)
                SqlConn.Close()
                catCMD.Dispose()
            End If

            If pageId = 0 Then
                catCMD = New System.Data.SqlClient.SqlCommand("Select Max(Page_ID) From Pages", SqlConn)
                SqlConn.Open()
                myReader = catCMD.ExecuteReader(CommandBehavior.SingleRow)

                If myReader.Read() Then

                    If Not myReader(0) Is DBNull.Value Then
                        pageId = Trim(myReader(0))
                    Else
                        pageId = "0"
                    End If

                End If

                myReader.Close()
                SqlConn.Close()
                catCMD.Dispose()
            End If

            Response.Redirect("default.asp")

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        
        Response.Redirect("default.asp")

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If Trim(pageTitle1.Value) = "" Then
            Dim stralert As String
            If Trim(lang_id) = "1" Then
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('חובה למלא שם של דף מעוצב');" & vbCrLf
                stralert += "</script>" & vbCrLf
            Else
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('Please insert the page name');" & vbCrLf
                stralert += "</script>" & vbCrLf
            End If
            PHAlert.Controls.Add(New LiteralControl(stralert))
        ElseIf Trim(title1.Value) = "" Then
            Dim stralert As String
            If Trim(lang_id) = "1" Then
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('חובה למלא את השדה טקסט');" & vbCrLf
                stralert += "</script>" & vbCrLf
            Else
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('Please insert the page content');" & vbCrLf
                stralert += "</script>" & vbCrLf
            End If
            PHAlert.Controls.Add(New LiteralControl(stralert))
        Else
            catCMD = New System.Data.SqlClient.SqlCommand("updatePageContent", SqlConn)
            catCMD.CommandType = CommandType.StoredProcedure
            If Trim(pageId) = "" Or IsNumeric(pageId) = False Then
                pageId = 0
            End If

            catCMD.Parameters.Add("@pageId", pageId)
            catCMD.Parameters.Add("@orgId", orgID)
            catCMD.Parameters.Add("@userId", userId)
            catCMD.Parameters.Add("@PageLang", Trim(Request.Form("PageLang")))
            catCMD.Parameters.Add("@pageTitle", pageTitle1.Value)
            catCMD.Parameters.Add("@pageSource", title1.InnerText)

            SqlConn.Open()
            catCMD.ExecuteNonQuery()
            SqlConn.Close()
            catCMD.Dispose()

            'אם נמחק טופס לנקות שדות
            If InStr(title1.InnerText, "tofes") = 0 And IsNumeric(pageId) Then
                sqlstr = "UPDATE Pages SET PRODUCT_ID = NULL, FORM_LINK_IMAGE=NULL,FORM_LINK=NULL," & _
                " LINK_IMAGE=NULL, LINK_IMAGE_ALIGN=NULL,LINK_TEXT=NULL,LINK_TEXT_ALIGN=NULL," & _
                " LINK_BGCOLOR=NULL, LINK_FONT_TYPE=NULL, LINK_FONT_SIZE=NULL, LINK_FONT_COLOR=NULL," & _
                " FORM_BGCOLOR=NULL, FORM_FONT_TYPE=NULL, FORM_FONT_SIZE=NULL, FORM_FONT_COLOR=NULL" & _
                " WHERE PAGE_ID = " & pageId
                catCMD = New System.Data.SqlClient.SqlCommand(sqlstr, SqlConn)
                SqlConn.Open()
                catCMD.ExecuteReader(CommandBehavior.Default)
                SqlConn.Close()
                catCMD.Dispose()
            End If

            If pageId = 0 Then
                catCMD = New System.Data.SqlClient.SqlCommand("Select Max(Page_ID) From Pages", SqlConn)
                SqlConn.Open()
                myReader = catCMD.ExecuteReader(CommandBehavior.SingleRow)

                If myReader.Read() Then

                    If Not myReader(0) Is DBNull.Value Then
                        pageId = Trim(myReader(0))
                    Else
                        pageId = "0"
                    End If

                End If

                myReader.Close()
                SqlConn.Close()
                catCMD.Dispose()
            End If
            Dim stralert As String
            stralert = "<script language='javascript'>" & vbCrLf
            stralert += "page = window.open('result.asp?pageId=" + pageId + "','Page','toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=780, height=540, left=2, top=2');"
            stralert += "</script>" & vbCrLf

            PHAlert.Controls.Add(New LiteralControl(stralert))
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If Trim(pageTitle1.Value) = "" Then
            Dim stralert As String
            If Trim(lang_id) = "1" Then
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('חובה למלא שם של דף מעוצב');" & vbCrLf
                stralert += "</script>" & vbCrLf
            Else
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('Please insert the page name');" & vbCrLf
                stralert += "</script>" & vbCrLf
            End If
            PHAlert.Controls.Add(New LiteralControl(stralert))
        ElseIf Trim(title1.Value) = "" Then
            Dim stralert As String
            If Trim(lang_id) = "1" Then
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('חובה למלא את השדה טקסט');" & vbCrLf
                stralert += "</script>" & vbCrLf
            Else
                stralert = "<script language='javascript'>" & vbCrLf
                stralert += "alert('Please insert the page content');" & vbCrLf
                stralert += "</script>" & vbCrLf
            End If
            PHAlert.Controls.Add(New LiteralControl(stralert))
        Else
            catCMD = New System.Data.SqlClient.SqlCommand("updatePageContent", SqlConn)
            catCMD.CommandType = CommandType.StoredProcedure
            If Trim(pageId) = "" Or IsNumeric(pageId) = False Then
                pageId = 0
            End If

            catCMD.Parameters.Add("@pageId", pageId)
            catCMD.Parameters.Add("@orgId", orgID)
            catCMD.Parameters.Add("@userId", userId)
            catCMD.Parameters.Add("@PageLang", Trim(Request.Form("PageLang")))
            catCMD.Parameters.Add("@pageTitle", pageTitle1.Value)
            catCMD.Parameters.Add("@pageSource", title1.InnerText)
            SqlConn.Open()
            catCMD.ExecuteNonQuery()
            SqlConn.Close()
            catCMD.Dispose()

            'אם נמחק טופס לנקות שדות
            If InStr(title1.InnerText, "tofes") = 0 And IsNumeric(pageId) Then
                sqlstr = "UPDATE Pages SET Product_Id = NULL, FORM_LINK_IMAGE=NULL,FORM_LINK=NULL," & _
                " LINK_IMAGE=NULL, LINK_IMAGE_ALIGN=NULL,LINK_TEXT=NULL,LINK_TEXT_ALIGN=NULL," & _
                " LINK_BGCOLOR=NULL, LINK_FONT_TYPE=NULL, LINK_FONT_SIZE=NULL, LINK_FONT_COLOR=NULL," & _
                " FORM_BGCOLOR=NULL, FORM_FONT_TYPE=NULL, FORM_FONT_SIZE=NULL, FORM_FONT_COLOR=NULL" & _
                " WHERE PAGE_ID = " & pageId
                catCMD = New System.Data.SqlClient.SqlCommand(sqlstr, SqlConn)
                SqlConn.Open()
                catCMD.ExecuteReader(CommandBehavior.Default)
                SqlConn.Close()
                catCMD.Dispose()
            End If

            If pageId = 0 Then
                catCMD = New System.Data.SqlClient.SqlCommand("Select Max(Page_ID) From Pages", SqlConn)
                SqlConn.Open()
                myReader = catCMD.ExecuteReader(CommandBehavior.SingleRow)

                If myReader.Read() Then

                    If Not myReader(0) Is DBNull.Value Then
                        pageId = Trim(myReader(0))
                    Else
                        pageId = "0"
                    End If

                End If

                myReader.Close()
                SqlConn.Close()
                catCMD.Dispose()
            End If

            Response.Redirect("editPage.aspx?pageID=" & pageId)

        End If
    End Sub

End Class
