Imports System.Data.SqlClient
Public Class TripForm16504
    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Public dtCountries, dtCompetitor As New DataTable
    Public contactID, companyId, TitleCompetitor As String
    Protected appealId As Integer
    Protected pTitle As Integer

    Dim cmdSelect As New SqlClient.SqlCommand
    Dim primKeydtCountries(0), primKeydtCompetitor(0) As Data.DataColumn
    Protected dateBack As Date
    Protected v40167 As String
    Protected WithEvents CRMCountry, selSerias, selCompetitor As System.Web.UI.HtmlControls.HtmlSelect



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
        contactID = ""
        dateBack = DateAdd("d", +2, Now())
        If IsNumeric(Request("contactID")) Then
            contactID = Request("contactID")
            Dim cmdSelect As New SqlClient.SqlCommand("select count(*) as r  from appeals where contact_id=@contactID and  questions_id=16735", con)
            cmdSelect.Parameters.Add("@contactID", SqlDbType.Int).Value = contactID
            cmdSelect.CommandType = CommandType.Text
            con.Open()
            Dim tmp = cmdSelect.ExecuteScalar()
            cmdSelect.Dispose()
            con.Close()
            If tmp > 0 Then
                v40167 = "כן"
            End If

        cmdSelect = New SqlClient.SqlCommand("SELECT  COUNT(*) AS Expr1 FROM  APPEALS WHERE (CONTACT_ID = @ContactId ) AND (QUESTIONS_ID = 16504) AND (YEAR(GETDATE()) - YEAR(APPEAL_DATE) <= 2)", con)
        cmdSelect.Parameters.Add("@contactID", SqlDbType.Int).Value = contactID
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        Dim tmpA = cmdSelect.ExecuteScalar()
        cmdSelect.Dispose()
        con.Close()
        If tmpA > 0 Then
                TitleCompetitor = "מול מי אתה משווה?"
                pTitle = 40847
        Else
                TitleCompetitor = "עם מי נסעת שנה שעברה?"
                pTitle = 40848
        End If
        End If
        If Not Page.IsPostBack Then
            'Response.Write("not Page.IsPostBack")
            GetCRMCountries()
            GetCompetitor()
        Else
            SaveData()
        End If
    End Sub
    Private Sub SaveData()
        Dim sqlIns As SqlClient.SqlCommand
        Dim companyID, UserID, OrgID, contactID As String
        OrgID = Trim(Trim(Request.Cookies("bizpegasus")("ORGID")))
        contactID = Request.QueryString("contactID")
        UserID = Trim(Request.Cookies("bizpegasus")("UserID"))
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
        field40109 = Request.Form("field40109") 'באיזה חודש תרצה לטייל
        field40167 = Request.Form("field40167")
        field40846 = Request.Form("field40846") 'כמה מטיילים תהיו 
        selCompetitorId = Request.Form("selCompetitor")
        field40170 = Request.Form("field40170") ' ---מתי כדי לחזור 
        field40103 = Request.Form("field40103")  'תוכן הפנייה
        If IsNumeric(selCompetitorId) Then
            cmdSelect = New SqlClient.SqlCommand("SELECT  Competitor_Name FROM  Compare_Competitors WHERE (Competitor_Id = @selCompetitorId ) ", con)
            cmdSelect.Parameters.Add("@selCompetitorId", SqlDbType.Int).Value = selCompetitorId
            cmdSelect.CommandType = CommandType.Text
            con.Open()
            CompetitorName = cmdSelect.ExecuteScalar()
            cmdSelect.Dispose()
            con.Close()

        End If

        If pTitle = 40847 Then
            field40847 = CompetitorName  'מול מי אתה משווה
            field40848 = ""
        Else

            field40848 = CompetitorName 'עם מי נסעת שנה שעברה?
            field40847 = ""
        End If


        '    '!!!!!!==============create new appeal
        '    '=============================טופס מתעניין==================

        Dim sqlstring = "Set NOCOUNT On;Set DATEFORMAT dmy;" & _
        "INSERT INTO appeals (appeal_date,appeal_SeriasId,questions_id,USER_ID,User_Id_Order_Owner,Department_Id,appeal_status,appeal_deleted,type_id,organization_id,company_ID,contact_ID,appeal_CountryId,Reason_Id) " & _
        "VALUES (getDate(),@selSeriasId,@quest_id,@USER_ID,@UserIdOrderOwner,@DepartmentId,'1',0,2,@OrgID,@companyID,@contactID,@appeal_CountryId,1);" & _
        "SELECT @@IDENTITY AS NewID"
        cmdSelect = New SqlClient.SqlCommand(sqlstring, con)
        cmdSelect.Parameters.Add("@USER_ID", SqlDbType.Int).Value = UserID
        cmdSelect.Parameters.Add("@DepartmentId", SqlDbType.Int).Value = DBNull.Value
        cmdSelect.Parameters.Add("@UserIdOrderOwner", SqlDbType.Int).Value = 0
        cmdSelect.Parameters.Add("@quest_id", SqlDbType.Int).Value = 16504 '- טופס מתעניין
        cmdSelect.Parameters.Add("@OrgID", SqlDbType.Int).Value = 264 'CInt(OrgID)
        cmdSelect.Parameters.Add("@selSeriasId", SqlDbType.Int).Value = selSeriasId
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

        '------40847   'מול מי אתה משווה
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40847)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40847
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()
        '------40848   ''עם מי נסעת שנה שעברה?

        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40848)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40848
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()
        ' ----40167 ---האם נסעת עם פגסוס בעבר? ------
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40167)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40167
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()

        '------40109 באיזה חודש תרצה לטייל 
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40109)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40109
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()
        '----40103   ---תוכן הפנייה--------------------
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40103)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40103
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()
        '--40170 ---מתי כדי לחזור 
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40170)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40170
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()
        '--- 40846 כמה מטיילים תהיו
        sqlIns = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID ,@Field_Id,@Field_Value)", con)
        sqlIns.Parameters.Add("@appID", SqlDbType.Int).Value = func.fixNumeric(appealId)
        sqlIns.Parameters.Add("@Field_Id", SqlDbType.Int).Value = func.fixNumeric(40846)
        sqlIns.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = field40846
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()


        'insert changes הוספת תופס 
        sqlIns = New SqlClient.SqlCommand("INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
        "SELECT 'טופס', 'שם: ' + P.Product_Name + ', איש קשר:'  + IsNULL(CONTACT_NAME, ''), Appeal_Id, 'הוספה', getDate(), " & UserID & _
        " FROM dbo.Appeals A LEFT OUTER JOIN dbo.Products P ON P.Product_Id = A.Questions_id " & _
        " LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = A.Contact_Id WHERE (Appeal_Id = " & appealId & ")", con)
        con.Open()
        sqlIns.ExecuteNonQuery()
        con.Close()
        Dim cScript As String


        cScript = "<script language='javascript'>alert('הטופס נקלט במערכת');self.close() ;window.opener.document.location.reload(true); </script>"
        Response.Write(cScript)
        Response.End()

             RegisterStartupScript("ReloadScrpt", cScript)


        RegisterStartupScript("ReloadScrpt", cScript)

        'Page.ClientScript.RegisterStartupScript(Me.GetType(), "ReloadScrpt", cScript)
        RegisterStartupScript("ReloadScrpt", cScript)
        'check RESPONSIBLE and send mail
        'sqlstr = "Select Langu, Product_Name, PRODUCT_DESCRIPTION, RESPONSIBLE From Products WHERE PRODUCT_ID = " & quest_id
        ''' message str_alert = "הטופס נקלט במערכת"
    End Sub
    Sub GetCRMCountries()
        'Dim cmdSelect As New SqlClient.SqlCommand("SELECT Country_Id, Country_Name from Countries  order by Country_Name", conPegasus)
        'SELECT FROM BIZPOWER YAADIM
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT Country_Id, Country_Name from Countries  order by Country_Name", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtCountries)
        con.Close()
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
    Sub GetCompetitor()
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT distinct m.Competitor_Id,m.Competitor_Name, min(m.ord) from ( select row_number() OVER (ORDER BY Competitor_Name,End_Date desc, Start_Date desc) as ord, " & _
 " Compare_Screens.Competitor_Id,Competitor_Name from Compare_Screens  inner join Compare_Competitors on Compare_Competitors.Competitor_Id=Compare_Screens.Competitor_Id " & _
 " where   datediff(dd,Compare_Screens.[Start_Date],getdate())>=0 and datediff(dd,Compare_Screens.[End_Date],getdate())<=0 ) as m group by  m.Competitor_Id,m.Competitor_Name" & _
 " order by  min(m.ord)", con)

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
