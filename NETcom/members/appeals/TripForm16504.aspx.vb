Imports System.Data.SqlClient
Public Class TripForm16504
    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim BizpowerDBName As String = ConfigurationSettings.AppSettings("BizpowerDBName")
    
    Public contactID As String
    Public companyID, UserID, OrgID As String
    Public dtCountries, dtSerias, dtCompetitor As New DataTable
    Public TitleCompetitor As String
    Protected appealId As Integer
    Protected pTitle As Integer

    Dim cmdSelect As New SqlClient.SqlCommand
    Dim primKeydtCountries(0), primKeydtSerias(0), primKeydtCompetitor(0) As Data.DataColumn
    Protected dateBack As Date
    Protected v40167 As String
    Protected WithEvents CRMCountry, selSerias, selCompetitor As System.Web.UI.HtmlControls.HtmlSelect
    Protected contactName, contactPhone, contactCellular, contactEmail As String



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
        UserID = Trim(Request.Cookies("bizpegasus")("UserID"))

        OrgID = Trim(Trim(Request.Cookies("bizpegasus")("ORGID")))
        contactID = Request.QueryString("contactID")

        contactID = ""
        dateBack = DateAdd("d", +2, Now())
        If IsNumeric(Request("contactID")) Then
            contactID = Request("contactID")

            Dim sqlSelectC = New SqlClient.SqlCommand("SELECT contacts.contact_id,contacts.company_id,contacts.contact_name," & _
                        " contacts.phone,contacts.cellular,contacts.email FROM contacts  WHERE  contacts.contact_id =@contactID", con)
            sqlSelectC.Parameters.Add("@contactID", SqlDbType.Int).Value = contactID
            sqlSelectC.CommandType = CommandType.Text
            con.Open()
            Dim dr As SqlClient.SqlDataReader = sqlSelectC.ExecuteReader()
            If dr.Read() Then
                If Not IsDBNull(dr("company_id")) Then
                    companyID = Trim(dr("company_id"))
                Else
                    companyID = 0
                End If
                If Not IsDBNull(dr("contact_name")) Then
                    contactName = Trim(dr("contact_name"))
                Else
                    contactName = ""
                End If
                If Not IsDBNull(dr("phone")) Then
                    contactPhone = Trim(dr("phone"))
                Else
                    contactPhone = ""
                End If
                If Not IsDBNull(dr("cellular")) Then
                    contactCellular = Trim(dr("cellular"))
                Else
                    contactCellular = ""
                End If
                If Not IsDBNull(dr("email")) Then
                    contactEmail = Trim(dr("email"))
                Else
                    contactEmail = ""
                End If

            End If
            con.Close()

            Dim cmdSelect As New SqlClient.SqlCommand("select count(*) as r  from appeals where contact_id=@contactID and  questions_id=16735", con)
            cmdSelect.Parameters.Add("@contactID", SqlDbType.Int).Value = contactID
            cmdSelect.CommandType = CommandType.Text
            con.Open()
            Dim tmp = cmdSelect.ExecuteScalar()
            cmdSelect.Dispose()
            con.Close()
            If tmp > 0 Then
                v40167 = "��"
            Else
                v40167 = "��"
            End If

            cmdSelect = New SqlClient.SqlCommand("SELECT  COUNT(*) AS Expr1 FROM  APPEALS WHERE (CONTACT_ID = @ContactId ) AND (QUESTIONS_ID = 16735) AND (YEAR(GETDATE()) - YEAR(APPEAL_DATE) <= 1)" & _
            " and not exists (select *  FROM FORM_VALUE where FORM_VALUE.Appeal_Id=APPEALS.Appeal_Id  and (FIELD_ID = 40661)  And  (FIELD_VALUE = 'on'))", con)
            cmdSelect.Parameters.Add("@contactID", SqlDbType.Int).Value = contactID
            cmdSelect.CommandType = CommandType.Text
            con.Open()
            Dim tmpA = cmdSelect.ExecuteScalar()
            cmdSelect.Dispose()
            con.Close()
            If tmpA > 0 Then
                TitleCompetitor = "��� �� ��� �����?"
                pTitle = 40847
            Else
                TitleCompetitor = "�� �� ���� ��� �����?"
                pTitle = 40848
            End If
        Else
            TitleCompetitor = "�� �� ���� ��� �����?"
            pTitle = 40848
        End If
        If Not Page.IsPostBack Then
            'Response.Write("not Page.IsPostBack")
            GetCRMCountries()
            '==get country by serias
            GetSerias()
            GetCompetitor()
        Else
            SaveData()
        End If
    End Sub
    Private Sub SaveData()
        Dim sqlIns As SqlClient.SqlCommand
        '===============added by Mila (from TrimForm16504NewContact.aspx)
        '=====================add/update contact
        Dim contactName, contactPhone, contactCellular, contactEmail As String
        contactName = Request("contact_name")
        contactPhone = Request("phone")
        contactCellular = Request("cellular")
        contactEmail = Request("email")


        If func.fixNumeric(contactID) > 0 Then
            If Len(Request("contact_name")) > 0 And Request.QueryString("isSelectedContact") = "1" Then
                Dim sqlUpd = New SqlClient.SqlCommand("UPDATE contacts SET CONTACT_NAME=@contactName, email=@contactEmail , phone=@contactPhone " & _
                          ", cellular=@contactCellular,date_update=getDate() WHERE contact_ID=@contactID ", con)
                sqlUpd.Parameters.Add("@contactName", SqlDbType.VarChar).Value = contactName
                sqlUpd.Parameters.Add("@contactEmail", SqlDbType.VarChar).Value = contactEmail
                sqlUpd.Parameters.Add("@contactPhone", SqlDbType.VarChar).Value = contactPhone
                sqlUpd.Parameters.Add("@contactCellular", SqlDbType.VarChar).Value = contactCellular
                sqlUpd.Parameters.Add("@contactID", SqlDbType.Int).Value = contactID
                con.Open()
                sqlUpd.ExecuteNonQuery()
                con.Close()
                '--insert into changes table
                Dim sqlstrIns = New SqlClient.SqlCommand("INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
                " SELECT '��� ���', ' ��:'  + IsNULL(CONTACT_NAME, '') + ' ����: ' + IsNULL(cellular, ''), CONTACT_ID, '�����', getDate()," & UserID & _
                " FROM dbo.CONTACTS WHERE (CONTACT_ID = " & contactID & ")", con)

                con.Open()
                sqlstrIns.ExecuteNonQuery()
                con.Close()
            End If
        Else
            'chk contact
            If Len(Request("contact_name")) > 0 Then

                Dim found_contact_id = 0
                Dim sqlCheck As New SqlClient.SqlCommand("EXEC dbo.contacts_chk_phone	@OrgId='" & OrgID & "', @EditContactId='" & contactID & "', " & _
                " @cp='" & contactCellular & "', @pn='" & contactPhone & "'", con)
                sqlCheck.CommandType = CommandType.Text
                con.Open()
                found_contact_id = sqlCheck.ExecuteScalar()
                cmdSelect.Dispose()
                con.Close()
                If found_contact_id > 0 Then
                    contactID = CLng(found_contact_id)

                Else
                    'Dim s = "SET DATEFORMAT MDY; SET NOCOUNT ON; " & _
                    '" INSERT INTO CONTACTS(12681,contact_name,email,phone,cellular,date_update,Organization_ID)" & _
                    '" VALUES (" & companyID & ",'" & contactName & "','" & contactEmail & "','" & _
                    'contactPhone & "','" & contactCellular & "'" & ", getDate(), " & OrgID & "); SELECT @@IDENTITY AS NewID"
                    'Response.Write(s)
                    'Response.End()

                    Dim sqlstrIns = New SqlClient.SqlCommand("SET DATEFORMAT MDY; SET NOCOUNT ON; " & _
                    " INSERT INTO CONTACTS(company_id,contact_name,email,phone,cellular,date_update,Organization_ID)" & _
                    " VALUES (12681,'" & contactName & "','" & contactEmail & "','" & _
                    contactPhone & "','" & contactCellular & "'" & ", getDate(), " & OrgID & "); SELECT @@IDENTITY AS NewID", con)

                    con.Open()
                    contactID = sqlstrIns.ExecuteScalar()
                    con.Close()
                    sqlstrIns = New SqlClient.SqlCommand("INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
                    " SELECT '��� ���', ' ��:'  + IsNULL(CONTACT_NAME, '') + ' ����: ' + IsNULL(cellular, ''), CONTACT_ID, '�����', getDate()," & UserID & _
                    " FROM dbo.CONTACTS WHERE (CONTACT_ID = " & contactID & ")", con)
                    con.Open()
                    sqlstrIns.ExecuteNonQuery()
                    con.Close()
                End If
            End If
        End If

        '===============================================================


        Dim CRMCountryId, selSeriasId, selCompetitorId, pTitle As Integer

        Dim CRMCountryName, selSeriasName, field40109, field40167, field40846, field40170, field40103, field40847, field40848 As String
        'Response.Write("savedata")
        'Dim x

        'For Each x In Request.Form

        '    Response.Write(x)
        '    Response.Write(" - [" & Request.Form(x) & "]" & "<br>")
        'Next
        Dim CompetitorName As String

        pTitle = Request.Form("pTitle")
        CRMCountryId = Request.Form("CRMCountry")
        ' CRMCountryName = Request.Form("input_CRMCountry")


        selSeriasId = Request.Form("selSerias")
        selSeriasName = Request.Form("input_selSerias")
        field40109 = Request.Form("field40109") '����� ���� ���� �����
        field40167 = Request.Form("field40167")
        field40846 = Request.Form("field40846") '��� ������� ���� 
        selCompetitorId = func.fixNumeric(Request.Form("selCompetitor"))
        field40170 = Request.Form("field40170") ' ---��� ��� ����� 
        field40103 = Request.Form("field40103")  '���� ������
        If selCompetitorId > 0 Then
            cmdSelect = New SqlClient.SqlCommand("SELECT  Competitor_Name FROM  Compare_Competitors WHERE (Competitor_Id = @selCompetitorId ) ", con)
            cmdSelect.Parameters.Add("@selCompetitorId", SqlDbType.Int).Value = selCompetitorId
            cmdSelect.CommandType = CommandType.Text
            con.Open()
            CompetitorName = cmdSelect.ExecuteScalar()
            cmdSelect.Dispose()
            con.Close()
        Else
            CompetitorName = ""
        End If

        If pTitle = 40847 Then
            field40847 = CompetitorName  '��� �� ��� �����
            field40848 = ""
        Else

            field40848 = CompetitorName '�� �� ���� ��� �����?
            field40847 = ""
        End If


        '    '!!!!!!==============create new appeal
        '    '=============================���� �������==================

        Dim sqlstring = "Set NOCOUNT On;Set DATEFORMAT dmy;" & _
        "INSERT INTO appeals (appeal_date,appeal_SeriasId,questions_id,USER_ID,User_Id_Order_Owner,Department_Id,appeal_status,appeal_deleted,type_id,organization_id,company_ID,contact_ID,appeal_CountryId,Reason_Id) " & _
        "VALUES (getDate(),@selSeriasId,@quest_id,@USER_ID,@UserIdOrderOwner,@DepartmentId,'1',0,2,@OrgID,@companyID,@contactID,@appeal_CountryId,1);" & _
        "SELECT @@IDENTITY AS NewID"
        cmdSelect = New SqlClient.SqlCommand(sqlstring, con)
        cmdSelect.Parameters.Add("@USER_ID", SqlDbType.Int).Value = UserID
        cmdSelect.Parameters.Add("@DepartmentId", SqlDbType.Int).Value = DBNull.Value
        cmdSelect.Parameters.Add("@quest_id", SqlDbType.Int).Value = 16504 '- ���� �������
        cmdSelect.Parameters.Add("@OrgID", SqlDbType.Int).Value = 264 'CInt(OrgID)
        cmdSelect.Parameters.Add("@selSeriasId", SqlDbType.Int).Value = selSeriasId
        cmdSelect.Parameters.Add("@UserIdOrderOwner", SqlDbType.Int).Value = 0
        If func.fixNumeric(companyID) <> 0 Then
            cmdSelect.Parameters.Add("@companyID", SqlDbType.Int).Value = CInt(companyID)
        Else
            cmdSelect.Parameters.Add("@companyID", SqlDbType.Int).Value = DBNull.Value
        End If
        cmdSelect.Parameters.Add("@contactID", SqlDbType.Int).Value = CInt(contactID)
        cmdSelect.Parameters.Add("@appeal_CountryId", SqlDbType.Int).Value = IIf(CInt(CRMCountryId) > 0, CRMCountryId, DBNull.Value)

        con.Open()
        appealId = cmdSelect.ExecuteScalar
        con.Close()

        'insert field

        '------40847   '��� �� ��� �����
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40847)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40847
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()
        '------40848   ''�� �� ���� ��� �����?

        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40848)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40848
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()
        ' ----40167 ---��� ���� �� ����� ����? ------
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40167)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40167
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()

        '------40109 ����� ���� ���� ����� 
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40109)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40109
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()
        '----40103   ---���� ������--------------------
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40103)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40103
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()
        '--40170 ---��� ��� ����� 
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40170)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40170
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()
        '--- 40846 ��� ������� ����
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40846)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40846
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()


        'insert changes ����� ���� 
        sqlIns = New SqlClient.SqlCommand("INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
        "SELECT '����', '��: ' + P.Product_Name + ', ��� ���:'  + IsNULL(CONTACT_NAME, ''), Appeal_Id, '�����', getDate(), " & UserID & _
        " FROM dbo.Appeals A LEFT OUTER JOIN dbo.Products P ON P.Product_Id = A.Questions_id " & _
        " LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = A.Contact_Id WHERE (Appeal_Id = " & appealId & ")", con)
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()

        If Request.Form("send_task") = "1" Then
            'add task
            Response.Redirect("../tasks/addtask.asp?appealID=" & appealId) ''& "&task_reciver_id=" & UserID)
            'addTask(appealId)
        Else
            'close popup
            Dim cScript As String


            cScript = "<script language='javascript'>alert('����� ���� ������');self.close() ;window.opener.document.location.reload(true); </script>"
            '' Response.Write(cScript)
            ''  Response.End()

            RegisterStartupScript("ReloadScrpt", cScript)


            ' RegisterStartupScript("ReloadScrpt", cScript)

            'Page.ClientScript.RegisterStartupScript(Me.GetType(), "ReloadScrpt", cScript)
            '  RegisterStartupScript("ReloadScrpt", cScript)
            'check RESPONSIBLE and send mail
            'sqlstr = "Select Langu, Product_Name, PRODUCT_DESCRIPTION, RESPONSIBLE From Products WHERE PRODUCT_ID = " & quest_id
            ''' message str_alert = "����� ���� ������"
        End If

     
    End Sub
    ''don't automatically
    'Sub addTask(ByVal appealId As Integer)
    '    Dim taskId As Integer = 0
    '    Dim fromEmail As String = ""

    '    Dim task_date = "getdate()"
    '    Dim task_open_date = "getdate()"
    '    Dim task_content = "��� ����� �������"
    '    Dim task_project_id = "NULL"
    '    Dim OrgID = "264"
    '    Dim task_appeal_id = CStr(appealId)
    '    Dim parentID = "NULL"
    '    Dim task_types = "2005" '���� �����
    '    Dim task_reciver_id = UserID
    '    Dim File_Name = ""

    '    Dim sqlstr As String = "SET DATEFORMAT DMY; SET NOCOUNT ON; Insert Into tasks (company_id,contact_id,project_id,User_ID,ORGANIZATION_ID,appeal_ID,parent_ID,task_date,task_open_date,task_content,task_types,task_status,reciver_id,attachment) " & _
    '      " values (" & companyId & "," & contactID & "," & task_project_id & "," & UserID & "," & OrgID & "," & task_appeal_id & "," & parentID & "," & _
    '      task_date & "','" & task_open_date & "','" & task_content & "','" & task_types & "','1'," & task_reciver_id & ",''); SELECT @@IDENTITY AS NewID"
    '    'Response.Write sqlstr
    '    'Response.End

    '    cmdSelect = New SqlClient.SqlCommand(sqlstr, con)
    '    con.Open()
    '    Try
    '        taskId = cmdSelect.ExecuteScalar
    '        con.Close()
    '    Catch ex As Exception
    '    Finally
    '        con.Close()
    '    End Try

    '    If taskId > 0 Then
    '        '--insert into changes table
    '        sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
    '        " SELECT '�����', '��: ' + U.FIRSTNAME + ' ��� ���:'  + IsNULL(CONTACT_NAME, ''), task_id, '�����', getDate(), " & UserID & _
    '        " FROM dbo.tasks T LEFT OUTER JOIN dbo.USERS U ON U.User_Id = T.reciver_id " & _
    '        " LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = T.Contact_Id WHERE (Task_Id = " & taskId & ")"
    '        con.Open()
    '        Try
    '            cmdSelect.ExecuteNonQuery()
    '            con.Close()
    '        Catch ex As Exception
    '        Finally
    '            con.Close()
    '        End Try


    '        cmdSelect = New SqlClient.SqlCommand("Select EMAIL From USERS Where USER_ID =@UserID", con)
    '        con.Open()
    '        cmdSelect.Parameters.Add("@UserID", SqlDbType.Int).Value = UserID
    '        Try
    '            fromEmail = cmdSelect.ExecuteScalar
    '            con.Close()
    '        Catch ex As Exception
    '        Finally
    '            con.Close()
    '        End Try



    '    End If
    'End Sub

    Sub GetCRMCountries()
        'Dim cmdSelect As New SqlClient.SqlCommand("SELECT Country_Id, Country_Name from Countries  order by Country_Name", conPegasus)
        'SELECT FROM BIZPOWER YAADIM
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT distinct Country_Id, Country_Name from tours " & _
" inner join " & BizpowerDBName & ".dbo.Countries on tours.Country_CRM=Countries.Country_Id " & _
" inner join " & BizpowerDBName & ".dbo.Series on tours.SeriasId=Series.Series_Id " & _
" order by Country_Name", conPegasus)
        conPegasus.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtCountries)
        conPegasus.Close()
        primKeydtCountries(0) = dtCountries.Columns("Country_ID")
        dtCountries.PrimaryKey = primKeydtCountries
        CRMCountry.Items.Clear()
        Dim list1 As New ListItem("", "0")
        CRMCountry.Items.Add(list1)
        For i As Integer = 0 To dtCountries.Rows.Count - 1
            Dim list As New ListItem(dtCountries.Rows(i)("Country_Name"), dtCountries.Rows(i)("Country_Id"))
            If Request.QueryString("CRMCountry") > 0 And Request.QueryString("CRMCountry") = dtCountries.Rows(i)("Country_Id") Then
                list.Selected = True
            End If
            CRMCountry.Items.Add(list)
        Next
    End Sub
    Sub GetSerias()
        'Dim cmdSelect As New SqlClient.SqlCommand("SELECT Country_Id, Country_Name from Countries  order by Country_Name", conPegasus)
        'SELECT FROM BIZPOWER YAADIM
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT  distinct SeriasId,Series_Name FROM Tours T inner join " & BizpowerDBName & ".dbo.Series S on T.SeriasId=S.Series_Id  ORDER BY  Series_Name", conPegasus)
        conPegasus.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtSerias)
        conPegasus.Close()
        primKeydtSerias(0) = dtSerias.Columns("SeriasId")
        dtSerias.PrimaryKey = primKeydtSerias
        selSerias.Items.Clear()
        Dim list1 As New ListItem("", "0")
        selSerias.Items.Add(list1)
        For i As Integer = 0 To dtSerias.Rows.Count - 1
            Dim list As New ListItem(dtSerias.Rows(i)("Series_Name"), dtSerias.Rows(i)("SeriasId"))
            If Request.QueryString("selSerias") > 0 And Request.QueryString("selSerias") = dtSerias.Rows(i)("SeriasId") Then
                list.Selected = True
            End If
            selSerias.Items.Add(list)
        Next
    End Sub
    Sub GetCompetitor()
        '==added by Mila - get all competitors without relation to active screens
        '       Dim cmdSelect As New SqlClient.SqlCommand("SELECT distinct m.Competitor_Id,m.Competitor_Name, min(m.ord) from ( select row_number() OVER (ORDER BY Competitor_Name,End_Date desc, Start_Date desc) as ord, " & _
        '" Compare_Screens.Competitor_Id,Competitor_Name from Compare_Screens  inner join Compare_Competitors on Compare_Competitors.Competitor_Id=Compare_Screens.Competitor_Id " & _
        '" where   datediff(dd,Compare_Screens.[Start_Date],getdate())>=0 and datediff(dd,Compare_Screens.[End_Date],getdate())<=0 ) as m group by  m.Competitor_Id,m.Competitor_Name" & _
        '" order by  min(m.ord)", con)

        Dim cmdSelect As New SqlClient.SqlCommand("SELECT Competitor_Id,Competitor_Name from Compare_Competitors  order by  Competitor_Name", con)

        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtCompetitor)
        con.Close()
        primKeydtCompetitor(0) = dtCompetitor.Columns("Competitor_Id")
        dtCompetitor.PrimaryKey = primKeydtCompetitor
        selCompetitor.Items.Clear()
        Dim list1 As New ListItem("", "")
        selCompetitor.Items.Add(list1)
        For i As Integer = 0 To dtCompetitor.Rows.Count - 1
            Dim list As New ListItem(dtCompetitor.Rows(i)("Competitor_Name"), dtCompetitor.Rows(i)("Competitor_Id"))
            If Request.QueryString("selCompetitor") > 0 And Request.QueryString("selCompetitor") = dtCompetitor.Rows(i)("Competitor_Id") Then
                list.Selected = True
            End If
            selCompetitor.Items.Add(list)
        Next

    End Sub


End Class
