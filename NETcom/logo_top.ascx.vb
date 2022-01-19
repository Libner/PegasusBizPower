Public Class logo_top_webflowCSS
    Inherits System.Web.UI.UserControl
    Public ShowLogo As String = "1"
    Public strLocal, userId, orgID, lang_id, dir_var, align_var, dir_obj_var, img_, bgr_pos, perSize, org_name, user_name As String
    Protected WithEvents conn As System.Data.SqlClient.SqlConnection
    Dim myReader, rs_groups As System.Data.SqlClient.SqlDataReader
    Protected WithEvents catCMD As System.Data.SqlClient.SqlCommand
    Dim sqlStr, newID, newTitle, newDate, is_groups As String
    Protected WithEvents NewsPH As System.Web.UI.WebControls.PlaceHolder
    Dim strPH, loginname, password, wizard_id, formid, form_type, task_inout, task_date As String
    Dim func As bizpower.cfunc

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
        'Session.LCID = 2057
        strLocal = "http://" & Trim(Request.ServerVariables("SERVER_NAME"))
        If Len(Trim(Request.ServerVariables("SERVER_PORT"))) > 0 Then
            strLocal = strLocal & ":" & Trim(Request.ServerVariables("SERVER_PORT"))
        End If
        strLocal = strLocal & Application("VirDir") & "/"

        conn = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        func = New bizpower.cfunc
        If Request.Form("username") <> "" Then
            loginname = func.sFix(Request.Form("username"))
            password = func.sFix(Request.Form("password"))
            wizard_id = func.sFix(Request.Form("wizard_id"))

            sqlStr = "SELECT USER_ID, USERS.ORGANIZATION_ID, USERS.LOGINNAME, USERS.PASSWORD," & _
            " CHIEF ,FIRSTNAME,LASTNAME,ORGANIZATIONS.ORGANIZATION_NAME,RowsInList,ORGANIZATION_LOGO, LANG_ID, " & _
            " USERS.WORK_PRICING,USERS.SURVEYS,USERS.EMAILS,USERS.COMPANIES,USERS.TASKS,USERS.CASH_FLOW, " & _
            " IsNULL(ORGANIZATIONS.FILE_UPLOAD, 0) as FILE_UPLOAD, IsNULL(USERS.EDIT_APPEAL, 0) as EDIT_APPEAL " & _
            " FROM USERS INNER JOIN ORGANIZATIONS ON USERS.ORGANIZATION_ID=ORGANIZATIONS.ORGANIZATION_ID " & _
            " WHERE ORGANIZATIONS.ACTIVE = 1 and USERS.ACTIVE = 1 and USERS.loginName='" & func.sFix(loginname) & _
            "' AND USERS.Password='" & func.sFix(password) & "'"
            'Response.Write sqlStr
            'Response.End
            catCMD = New System.Data.SqlClient.SqlCommand(sqlStr, conn)
            conn.Open()
            myReader = catCMD.ExecuteReader(CommandBehavior.SingleRow)

            If myReader.Read() Then
                Response.Cookies("bizpegasus")("UserId") = Trim(myReader("USER_ID"))
                Response.Cookies("bizpegasus")("wizardId") = wizard_id
                Response.Cookies("bizpegasus")("OrgID") = Trim(myReader("ORGANIZATION_ID"))
                Response.Cookies("bizpegasus")("LANGID") = myReader("LANG_ID")
                Response.Cookies("bizpegasus")("ORGNAME") = HttpUtility.UrlEncode(Trim(myReader("ORGANIZATION_NAME")), System.Text.Encoding.GetEncoding("windows-1255"))
                Response.Cookies("bizpegasus")("UserName") = HttpUtility.UrlEncode(Trim(myReader("FIRSTNAME")) & " " & Trim(myReader("LASTNAME")), System.Text.Encoding.GetEncoding("windows-1255"))
                Response.Cookies("bizpegasus")("RowsInList") = Trim(myReader("RowsInList"))
                Response.Cookies("bizpegasus")("perSize") = myReader("ORGANIZATION_LOGO").ActualSize
                Response.Cookies("bizpegasus")("Chief") = Trim(myReader("Chief"))
                Response.Cookies("bizpegasus")("WORKPRICING") = Trim(myReader("WORK_PRICING"))
                Response.Cookies("bizpegasus")("SURVEYS") = Trim(myReader("SURVEYS"))
                Response.Cookies("bizpegasus")("EMAILS") = Trim(myReader("EMAILS"))
                Response.Cookies("bizpegasus")("COMPANIES") = Trim(myReader("COMPANIES"))
                Response.Cookies("bizpegasus")("TASKS") = Trim(myReader("TASKS"))
                Response.Cookies("bizpegasus")("CASHFLOW") = Trim(myReader("CASH_FLOW"))
                Response.Cookies("bizpegasus")("CASHFLOW") = Trim(myReader("CASH_FLOW"))
                Response.Cookies("bizpegasus")("FILEUP") = Trim(myReader("FILE_UPLOAD"))
                Response.Cookies("bizpegasus")("EDITAPPEAL") = Trim(myReader("EDIT_APPEAL"))

                sqlStr = "Select Top 1 group_Id from Users_Groups WHERE ORGANIZATION_ID= " & Trim(myReader("ORGANIZATION_ID"))
                'Response.Write sqlStr
                catCMD = New System.Data.SqlClient.SqlCommand(sqlStr, conn)
                conn.Open()
                rs_groups = catCMD.ExecuteReader(CommandBehavior.SingleRow)
                If rs_groups.Read() Then
                    is_groups = "1"
                Else
                    is_groups = "0"
                End If
                rs_groups = Nothing

                Response.Cookies("bizpegasus")("ISGROUPS") = CStr(is_groups)

                Response.Cookies("bizpegasus").Expires = DateTime.Now().AddYears(1) 'DateAdd("yyyy", 1, Now())
            Else
                myReader.Close()
                conn.Close()
                Response.Redirect(strLocal)
                Response.End()
            End If
            myReader.Close()
            conn.Close()

            If Trim(Request.Cookies("bizpegasus")("OrgID")) <> "" Then
                sqlStr = "Select * from titles WHERE ORGANIZATION_ID=" & Trim(Request.Cookies("bizpegasus")("OrgID")) & " order by object_id"
                'Response.Write sql_obj
                'Response.End
                catCMD = New System.Data.SqlClient.SqlCommand(sqlStr, conn)
                conn.Open()
                myReader = catCMD.ExecuteReader(CommandBehavior.CloseConnection)
                Do While myReader.Read()
                    Select Case myReader("object_id")
                        Case "1"
                            Response.Cookies("bizpegasus")("Projectone") = Trim(myReader("title_organization_one"))
                            Response.Cookies("bizpegasus")("ProjectMulti") = Trim(myReader("title_organization_multi"))
                        Case "2"
                            Response.Cookies("bizpegasus")("CompaniesOne") = Trim(myReader("title_organization_one"))
                            Response.Cookies("bizpegasus")("CompaniesMulti") = Trim(myReader("title_organization_multi"))
                        Case "3"
                            Response.Cookies("bizpegasus")("ContactsOne") = Trim(myReader("title_organization_one"))
                            Response.Cookies("bizpegasus")("ContactsMulti") = Trim(myReader("title_organization_multi"))
                        Case "4"
                            Response.Cookies("bizpegasus")("TasksOne") = Trim(myReader("title_organization_one"))
                            Response.Cookies("bizpegasus")("TasksMulti") = Trim(myReader("title_organization_multi"))
                        Case "5"
                            Response.Cookies("bizpegasus")("ActivitiesOne") = Trim(myReader("title_organization_one"))
                            Response.Cookies("bizpegasus")("ActivitiesMulti") = Trim(myReader("title_organization_multi"))
                            'לינקים עליונים	
                        Case "6"
                            Response.Cookies("bizpegasus")("title1") = Trim(myReader("title_organization_one"))
                        Case "7"
                            Response.Cookies("bizpegasus")("title2") = Trim(myReader("title_organization_one"))
                        Case "8"
                            Response.Cookies("bizpegasus")("title3") = Trim(myReader("title_organization_one"))
                        Case "9"
                            Response.Cookies("bizpegasus")("title4") = Trim(myReader("title_organization_one"))
                        Case "10"
                            Response.Cookies("bizpegasus")("title5") = Trim(myReader("title_organization_one"))
                        Case "11"
                            Response.Cookies("bizpegasus")("title6") = Trim(myReader("title_organization_one"))
                    End Select
                Loop
                myReader.Close()
            End If
        Else
            wizard_id = Trim(Request.Cookies("bizpegasus")("wizardId"))
        End If

        userId = Trim(Request.Cookies("bizpegasus")("UserId"))
        orgID = Trim(Request.Cookies("bizpegasus")("OrgID"))
        If orgID = "" Or userId = "" Then
            Response.Redirect(strLocal)
        End If
        lang_id = Trim(Request.Cookies("bizpegasus")("LANGID"))
        perSize = Trim(Request.Cookies("bizpegasus")("perSize"))

        org_name = Request.Cookies("bizpegasus")("ORGNAME")
        org_name = HttpContext.Current.Server.UrlDecode(org_name)
        user_name = Trim(Request.Cookies("bizpegasus")("UserName"))
        user_name = HttpContext.Current.Server.UrlDecode(user_name)

        If IsNumeric(lang_id) = False Or Trim(lang_id) = "" Then
            lang_id = "1"
        End If
        If lang_id = "2" Then
            dir_var = "rtl"
            align_var = "left"
            dir_obj_var = "ltr"
            img_ = "_eng"
            bgr_pos = "'top right'"
        Else
            dir_var = "ltr"
            align_var = "right"
            dir_obj_var = "rtl"
            img_ = ""
            bgr_pos = "'top left'"
        End If

        ' כניסה מהתוכנה Desktop
        formid = Trim(Request.Form("formid"))
        form_type = Trim(Request.Form("form_type"))
        newID = Trim(Request.Form("newid"))
        task_inout = Trim(Request.Form("task_inout"))
        task_date = Trim(Request.Form("date_task"))
        If formid <> "" Then
            If form_type = "1" Then
                Response.Redirect("members/appeals/feedback.asp?appid=" & formid)
            ElseIf form_type = "2" Then
                Response.Redirect("members/appeals/appeal_card.asp?appid=" & formid)
            End If
        ElseIf task_inout <> "" Then
            Response.Redirect("members/tasks/default.asp?T=" & task_inout & "&start_date=" & task_date & "&end_date=" & task_date)
        ElseIf newID <> "" Then
            Response.Redirect("news.asp?ID=" & newID)
        End If

        'create news line
        sqlStr = "SELECT new_ID,New_Title,New_Date FROM News"
        sqlStr += " WHERE New_Site_Visible=1 "
        If Trim(lang_id) = "1" Then
            sqlStr = sqlStr + " AND Category_Id=1 "
        Else
            sqlStr = sqlStr + " AND Category_Id=2 "
        End If
        sqlStr += " ORDER BY New_Date DESC,New_ID DESC"

        catCMD = New System.Data.SqlClient.SqlCommand(sqlStr, conn)
        conn.Open()
        myReader = catCMD.ExecuteReader(CommandBehavior.CloseConnection)
        Do While myReader.Read()
            newID = myReader("New_ID")
            newTitle = myReader("New_Title")
            newDate = myReader("New_Date")

            strPH += "&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;"
            strPH += "<a class='homeNews' href='" & strLocal & "netcom/news.asp?ID=" & newID & "'>>> " & newTitle & "</a>"

        Loop
        myReader.Close()

        If strPH <> "" Then
            NewsPH.Controls.Add(New LiteralControl(strPH))
        End If


    End Sub

End Class
