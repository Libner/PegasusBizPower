Imports System.Data.SqlClient
Imports System.Web.Mail
Public Class SendReservationForm

    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Dim BizpowerDBName As String = ConfigurationSettings.AppSettings("BizpowerDBName")
    Dim SiteDBName As String = ConfigurationSettings.AppSettings("PegasusSiteDBName")
    Dim isSuccess As Boolean = False
    Dim userRequestFromMailId As Integer = 1422 'avir!!
    Dim userFromToursLogId As Integer = 1469 'avir!!

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

        Dim userId, orgID, lang_id, DepartmentId, companyID As String

        userId = Trim(Request.Cookies("bizpegasus")("UserId"))
        orgID = Trim(Request.Cookies("bizpegasus")("OrgID"))
        lang_id = Trim(Request.Cookies("bizpegasus")("LANGID"))

        Dim LogId As Integer
        Dim Logs As String
        Dim LogsList() As String
        Dim resList As New ArrayList
        Dim ContactId, CountryId, TourId, SeriasId, countEntersLastYear As Integer
        Dim isExistsAppleal16735, isExistsAppleal16504, isExistsDeparture As Boolean
        Dim Series_Name, Country_Name As String

        Dim appID As Integer = 0
        Dim status As String

        If func.fixNumeric(userId) > 0 Then
            If IsNumeric(lang_id) = False Or Trim(lang_id) = "" Then
                lang_id = "1"
            End If

            Logs = func.dbNullFix(Request("Logs"))
            LogId = func.fixNumeric(Request("logid"))
            LogsList = Logs.Split(",")

            'Response.Write("<br>Logs=" & Logs)
            For i As Integer = 0 To LogsList.Length - 1
                LogId = func.fixNumeric(LogsList(i))
                'Response.Write("<br>LogId=" & LogId)
                If LogId > 0 Then
                    'seng new appeal
                    cmdSelect = New SqlClient.SqlCommand("SELECT  Contact_Id, Log_ContactsToursInteresting.Tour_Id, Tours.Country_CRM " & _
                    " FROM Log_ContactsToursInteresting inner join Tours on Log_ContactsToursInteresting.Tour_Id=Tours.Tour_Id where Log_Id=@LogId", conPegasus)
                    cmdSelect.CommandType = CommandType.Text
                    cmdSelect.Parameters.Add("@LogId", SqlDbType.Int).Value = LogId
                    conPegasus.Open()
                    Dim rd As SqlDataReader
                    rd = cmdSelect.ExecuteReader
                    If rd.Read Then
                        ContactId = func.fixNumeric(rd("Contact_Id"))
                        CountryId = func.fixNumeric(rd("Country_CRM"))
                        TourId = func.fixNumeric(rd("Tour_Id"))
                    End If
                    conPegasus.Close()

                    ' 'test=================
                    ' '	תוכן הפניה
                    ' 'get enters count for all serias in yad 
                    ' Dim seriasCountList1 As New ArrayList
                    ' Dim seriasNameList1 As New ArrayList

                    ' 'count enters to Serias
                    ' cmdSelect = New SqlClient.SqlCommand("select count(Log_Id) as countEntersLastYear,tours.SeriasId," & BizpowerDBName & ".dbo.Series.Series_Name " & _
                    ' " from Log_ContactsToursInterestingEnters inner join tours on Log_ContactsToursInterestingEnters.Tour_Id=tours.Tour_Id " & _
                    ' " inner join " & BizpowerDBName & ".dbo.Series on tours.SeriasId=" & BizpowerDBName & ".dbo.Series.Series_Id " & _
                    '" where Contact_Id=@contactId and tours.Country_CRM=@CountryId and datediff(dd,[Enter_Date],getdate())<=366" & _
                    '" group by tours.SeriasId," & BizpowerDBName & ".dbo.Series.Series_Name", conPegasus)
                    ' cmdSelect.CommandType = CommandType.Text
                    ' cmdSelect.Parameters.Add("@contactId", SqlDbType.Int).Value = ContactId
                    ' cmdSelect.Parameters.Add("@CountryId", SqlDbType.Int).Value = CountryId
                    ' conPegasus.Open()
                    ' Dim drs1 As SqlDataReader
                    ' drs1 = cmdSelect.ExecuteReader
                    ' Do While drs1.Read
                    '     seriasCountList1.Add(drs1("countEntersLastYear"))
                    '     seriasNameList1.Add(func.dbNullFix(drs1("Series_Name")))
                    ' Loop
                    ' conPegasus.Close()

                    ' cmdSelect = New SqlClient.SqlCommand("select Country_Name from countries where Country_Id=@CountryId", con)
                    ' cmdSelect.CommandType = CommandType.Text
                    ' cmdSelect.Parameters.Add("@CountryId", SqlDbType.Int).Value = CountryId
                    ' con.Open()
                    ' Try
                    '     Country_Name = func.dbNullFix(cmdSelect.ExecuteScalar)
                    '     con.Close()
                    ' Catch ex As Exception

                    ' Finally
                    '     con.Close()
                    ' End Try

                    ' Dim strContent1 As String
                    ' strContent1 = strContent1 & " הלקוח נכנס לאתר האינטרנט ליעד " & Country_Name & " בשנה האחרונה"
                    ' For s1 As Integer = 0 To seriasCountList1.Count - 1
                    '     strContent1 = strContent1 & "<br>" & seriasCountList1(s1) & " " & " פעמים " & seriasNameList1(s1) & " לסדרה "
                    ' Next
                    ' Response.Write(strContent1)
                    ' Response.End()
                    ' '=====================






                    If ContactId > 0 And TourId > 0 Then
                        'check if exists tofes mit'anen to yad in last 3 monthes

                        cmdSelect = New SqlClient.SqlCommand("select dbo.chk_Appeal16504_LogMembersToursInteresting(@contactID,@CountryId) as ExistsAppleal16504", con)
                        cmdSelect.CommandType = CommandType.Text
                        cmdSelect.Parameters.Add("@contactID", SqlDbType.Int).Value = CInt(ContactId)
                        cmdSelect.Parameters.Add("@CountryId", SqlDbType.Int).Value = CountryId
                        con.Open()
                        Try
                            isExistsAppleal16504 = func.dbNullBool(cmdSelect.ExecuteScalar)
                            con.Close()
                        Catch ex As Exception

                        Finally
                            con.Close()
                        End Try
                        ' Response.Write("<br>isExistsAppleal16504=" & isExistsAppleal16504)


                        'if  exisits return logid as row with appeal
                        If isExistsAppleal16504 Then
                            resList.Add(LogId)
                        Else
                            'if not exisit add new

                            'it's were trip in the past - to any country (@countryId=0)
                            cmdSelect = New SqlClient.SqlCommand("select dbo.chk_Appeal16735_LogMembersToursInteresting(@contactID,0) as ExistsAppleal16735", con)
                            cmdSelect.CommandType = CommandType.Text
                            cmdSelect.Parameters.Add("@contactID", SqlDbType.Int).Value = CInt(ContactId)
                            cmdSelect.Parameters.Add("@CountryId", SqlDbType.Int).Value = CountryId
                            con.Open()
                            Try
                                isExistsAppleal16735 = func.dbNullBool(cmdSelect.ExecuteScalar)
                                con.Close()
                            Catch ex As Exception

                            Finally
                                con.Close()
                            End Try

                            cmdSelect = New SqlClient.SqlCommand("select Country_Name from countries where Country_Id=@CountryId", con)
                            cmdSelect.CommandType = CommandType.Text
                            cmdSelect.Parameters.Add("@CountryId", SqlDbType.Int).Value = CountryId
                            con.Open()
                            Try
                                Country_Name = func.dbNullFix(cmdSelect.ExecuteScalar)
                                con.Close()
                            Catch ex As Exception

                            Finally
                                con.Close()
                            End Try

                            'there is tour to yad in the closest year
                            cmdSelect = New SqlClient.SqlCommand("select top 1 Departure_Id from tours_Departurs inner join Tours on tours_Departurs.Tour_Id = Tours.Tour_Id where Tour_CRM=@CountryId and datediff(dd,getdate(),Departure_Date)<=366", conPegasus)
                            cmdSelect.CommandType = CommandType.Text
                            cmdSelect.Parameters.Add("@CountryId", SqlDbType.Int).Value = CountryId
                            conPegasus.Open()
                            Try
                                If func.fixNumeric(cmdSelect.ExecuteScalar) > 0 Then
                                    isExistsDeparture = True
                                End If
                                conPegasus.Close()
                            Catch ex As Exception

                            Finally
                                conPegasus.Close()
                            End Try




                            cmdSelect = New SqlClient.SqlCommand("select Department_Id from Users where User_Id=@UserIdOrderOwner", con)
                            cmdSelect.CommandType = CommandType.Text
                            cmdSelect.Parameters.Add("@UserIdOrderOwner", SqlDbType.Int).Value = CInt(userId)
                            con.Open()
                            Try
                                DepartmentId = cmdSelect.ExecuteScalar
                                con.Close()
                            Catch ex As Exception

                                'Response.Write("<br>DepartmentId err=" & Err.Description)
                            Finally
                                con.Close()
                                DepartmentId = 0
                            End Try

                            cmdSelect = New SqlClient.SqlCommand("Select company_id From companies WHERE Organization_ID = @OrgID AND private_flag = 1", con)
                            cmdSelect.CommandType = CommandType.Text
                            cmdSelect.Parameters.Add("@OrgID", SqlDbType.Int).Value = CInt(orgID)
                            con.Open()
                            Try
                                companyID = cmdSelect.ExecuteScalar
                                con.Close()
                            Catch ex As Exception
                                companyID = ""
                                '' Response.Write("<br>get companyID: " & Err.Description)
                            Finally
                                con.Close()
                            End Try




                            Dim sqlstring As String = "SET NOCOUNT ON;SET DATEFORMAT dmy;" & _
                                           "INSERT INTO appeals (appeal_date,questions_id,USER_ID,User_Id_Order_Owner,Department_Id,appeal_status,appeal_deleted,type_id,organization_id,company_ID,contact_ID,appeal_CountryId,Tour_ID,Reason_Id,IsMarketingEmail) " & _
                                           "VALUES (getDate(),@quest_id,@USER_ID,0,@DepartmentId,'1',0,2,@OrgID,@companyID,@contactID,@appeal_CountryId,@TourId,@Reason_Id,1);" & _
                                           "SELECT @@IDENTITY AS NewID"
                            cmdSelect = New SqlClient.SqlCommand(sqlstring, con)
                            cmdSelect.Parameters.Add("@USER_ID", SqlDbType.Int).Value = userFromToursLogId 'userId
                            cmdSelect.Parameters.Add("@DepartmentId", SqlDbType.Int).Value = func.fixNumeric(DepartmentId)
                            cmdSelect.Parameters.Add("@quest_id", SqlDbType.Int).Value = 16504 'quest_id = "16504" - טופס מתעניין
                            cmdSelect.Parameters.Add("@OrgID", SqlDbType.Int).Value = CInt(orgID)
                            If func.fixNumeric(companyID) <> 0 Then
                                cmdSelect.Parameters.Add("@companyID", SqlDbType.Int).Value = CInt(companyID)
                            Else
                                cmdSelect.Parameters.Add("@companyID", SqlDbType.Int).Value = DBNull.Value
                            End If
                            cmdSelect.Parameters.Add("@contactID", SqlDbType.Int).Value = CInt(ContactId)
                            cmdSelect.Parameters.Add("@appeal_CountryId", SqlDbType.Int).Value = IIf(CInt(CountryId) > 0, CountryId, DBNull.Value)
                            cmdSelect.Parameters.Add("@TourId", SqlDbType.Int).Value = CInt(TourId)
                            'cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = DBNull.Value
                            cmdSelect.Parameters.Add("@Reason_Id", SqlDbType.Int).Value = 3 'בעקבות התעניינות באתר החברה"


                            con.Open()
                            Try
                                appID = cmdSelect.ExecuteScalar
                                con.Close()
                            Catch ex As Exception
                                appID = 0
                                status = "ERROR"
                                'Response.Write("<br>INSERT appeal: " & Err.Description)
                            Finally
                                con.Close()
                            End Try

                            'Response.Write("<br>INSERT appID: " & appID)
                            If appID > 0 Then


                                cmdSelect = New SqlClient.SqlCommand("INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
                                "SELECT 'טופס', 'שם: ' + P.Product_Name + ', איש קשר:'  + IsNULL(CONTACT_NAME, ''), Appeal_Id, 'הוספה', getDate(), " & userId & _
                                " FROM dbo.Appeals A LEFT OUTER JOIN dbo.Products P ON P.Product_Id = A.Questions_id " & _
                                " LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = A.Contact_Id WHERE (Appeal_Id = @appID)", con)
                                cmdSelect.Parameters.Add("@appID", SqlDbType.Int).Value = CInt(appID)
                                con.Open()
                                Try
                                    cmdSelect.ExecuteNonQuery()
                                    con.Close()
                                Catch ex As Exception
                                    status = "ERROR"
                                Finally
                                    con.Close()
                                End Try


                                'FIELDS======================

                                ''''		הגורם ליצירת הטופס -- hidden by Mila 23/10/2019 - field is moved from dymnamical fields and it's added to application - Reason_Id
                                '''cmdSelect = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID,@Field_Id,@Field_Value)", con)

                                '''cmdSelect.Parameters.Add("@appID", SqlDbType.Int).Value = CInt(appID)
                                '''cmdSelect.Parameters.Add("@Field_Id", SqlDbType.Int).Value = 40842
                                '''cmdSelect.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = "בעקבות התעניינות באתר החברה" '3 'בעקבות התעניינות באתר החברה
                                '''con.Open()
                                '''Try
                                '''    cmdSelect.ExecuteNonQuery()
                                '''    con.Close()
                                '''Catch ex As Exception
                                '''    status = "ERROR"
                                '''Finally
                                '''    con.Close()
                                '''End Try

                                '=======month -- field40109
                                cmdSelect = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID,@Field_Id,@Field_Value)", con)
                                cmdSelect.Parameters.Add("@appID", SqlDbType.Int).Value = CInt(appID)
                                cmdSelect.Parameters.Add("@Field_Id", SqlDbType.Int).Value = 40109
                                cmdSelect.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = "" 'empty
                                con.Open()
                                    Try
                                    cmdSelect.ExecuteNonQuery()
                                    con.Close()
                                    Catch ex As Exception
                                        status = "ERROR"
                                        Response.Write("<br>INSERT Field 40109: - " & Err.Description)
                                    Finally
                                    con.Close()
                                    End Try

                                '	האם נסעת עם פגסוס בעבר 
                                cmdSelect = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID,@Field_Id,@Field_Value)", con)

                                cmdSelect.Parameters.Add("@appID", SqlDbType.Int).Value = CInt(appID)
                                cmdSelect.Parameters.Add("@Field_Id", SqlDbType.Int).Value = 40167
                                cmdSelect.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = IIf(isExistsAppleal16735, "כן", "לא") 'האם נסעת עם פגסוס בעבר 
                                con.Open()
                                Try
                                    cmdSelect.ExecuteNonQuery()
                                    con.Close()
                                Catch ex As Exception
                                    status = "ERROR"
                                Finally
                                    con.Close()
                                End Try

                                '	כיצד הגעת אלינו
                                cmdSelect = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID,@Field_Id,@Field_Value)", con)

                                cmdSelect.Parameters.Add("@appID", SqlDbType.Int).Value = CInt(appID)
                                cmdSelect.Parameters.Add("@Field_Id", SqlDbType.Int).Value = 40108
                                cmdSelect.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = "אתר האינטרנט" 'כיצד הגעת אלינו 
                                con.Open()
                                Try
                                    cmdSelect.ExecuteNonQuery()
                                    con.Close()
                                Catch ex As Exception
                                    status = "ERROR"
                                Finally
                                    con.Close()
                                End Try


                                '	תוכן הפניה
                                'get enters count for all serias in yad 

                                Dim seriasCountList As New ArrayList
                                Dim seriasNameList As New ArrayList
                                seriasCountList.Clear()
                                seriasNameList.Clear()
                                'count enters to Serias
                                'cmdSelect = New SqlClient.SqlCommand("select count(Log_Id) as countEntersLastYear,tours.SeriasId," & BizpowerDBName & ".dbo.Series.Series_Name " & _
                                cmdSelect = New SqlClient.SqlCommand("select sum(countDayEnters) as countEntersLastYear,tours.SeriasId," & BizpowerDBName & ".dbo.Series.Series_Name " & _
                                 " from Log_ContactsToursInterestingEnters inner join tours on Log_ContactsToursInterestingEnters.Tour_Id=tours.Tour_Id " & _
                                 " inner join " & BizpowerDBName & ".dbo.Series on tours.SeriasId=" & BizpowerDBName & ".dbo.Series.Series_Id " & _
                                " where Contact_Id=@contactId and tours.Country_CRM=@CountryId and datediff(dd,[Enter_Date],getdate())<=366" & _
                                " group by tours.SeriasId," & BizpowerDBName & ".dbo.Series.Series_Name", conPegasus)
                                cmdSelect.CommandType = CommandType.Text
                                cmdSelect.Parameters.Add("@contactId", SqlDbType.Int).Value = ContactId
                                cmdSelect.Parameters.Add("@CountryId", SqlDbType.Int).Value = CountryId
                                conPegasus.Open()
                                Dim drs As SqlDataReader
                                drs = cmdSelect.ExecuteReader
                                Do While drs.Read
                                    seriasCountList.Add(drs("countEntersLastYear"))
                                    seriasNameList.Add(func.dbNullFix(drs("Series_Name")))
                                Loop
                                conPegasus.Close()

                               
                                'cmdSelect = New SqlClient.SqlCommand("select Series_Name from Series where Series_Id =@SeriasId", con)
                                'cmdSelect.CommandType = CommandType.Text
                                'cmdSelect.Parameters.Add("@SeriasId", SqlDbType.Int).Value = SeriasId
                                'conPegasus.Open()
                                'Try
                                '    Series_Name = func.dbNullFix(cmdSelect.ExecuteScalar)
                                '    con.Close()
                                'Catch ex As Exception

                                'Finally
                                '    conPegasus.Close()
                                'End Try


                                'תוכן הפניה 
                                cmdSelect = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID,@Field_Id,@Field_Value)", con)

                                cmdSelect.Parameters.Add("@appID", SqlDbType.Int).Value = CInt(appID)
                                cmdSelect.Parameters.Add("@Field_Id", SqlDbType.Int).Value = 40103
                                Dim strContent As String
                                strContent = "בשנה האחרונה" & " הלקוח נכנס לאתר האינטרנט ליעד " & Country_Name
                                For s As Integer = 0 To seriasCountList.Count - 1
                                    strContent = strContent & Chr(10) & seriasCountList(s)
                                    strContent = strContent & "  " & "פעמים"
                                    strContent = strContent & " לסדרה "
                                    strContent = strContent & "  " & seriasNameList(s)
                                Next
                                cmdSelect.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = strContent
                                con.Open()
                                Try
                                    cmdSelect.ExecuteNonQuery()
                                    con.Close()
                                Catch ex As Exception
                                    status = "ERROR"
                                Finally
                                    con.Close()
                                End Try

                                '	מה הסיכוי שירשם לטיול של פגסוס 40169
                                cmdSelect = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID,@Field_Id,@Field_Value)", con)
                                cmdSelect.Parameters.Add("@appID", SqlDbType.Int).Value = CInt(appID)
                                cmdSelect.Parameters.Add("@Field_Id", SqlDbType.Int).Value = 40169
                                cmdSelect.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = "" 'empty
                                con.Open()
                                Try
                                    cmdSelect.ExecuteNonQuery()
                                    con.Close()
                                Catch ex As Exception
                                    status = "ERROR"
                                    Response.Write("<br>INSERT Field 40109: - " & Err.Description)
                                Finally
                                    con.Close()
                                End Try

                                ''''		האם מתאים לקבלת מייל שיווקי עם לינק רישום -- hidden by Mila 23/10/2019 - field is moved from dymnamical fields and it's added to application - IsMarketingEmail
                                '''cmdSelect = New SqlClient.SqlCommand("INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (@appID,@Field_Id,@Field_Value)", con)

                                '''cmdSelect.Parameters.Add("@appID", SqlDbType.Int).Value = CInt(appID)
                                '''cmdSelect.Parameters.Add("@Field_Id", SqlDbType.Int).Value = 40843
                                '''cmdSelect.Parameters.Add("@Field_Value", SqlDbType.VarChar).Value = IIf(isExistsDeparture, "כן", "לא")
                                '''con.Open()
                                '''Try
                                '''    cmdSelect.ExecuteNonQuery()
                                '''    con.Close()
                                '''Catch ex As Exception
                                '''    status = "ERROR"
                                '''Finally
                                '''    con.Close()
                                '''End Try


                                If status = "ERROR" Then
                                    'roll back - operation is not finished correctly - delete appeal and fielda join to appeal
                                    cmdSelect = New SqlClient.SqlCommand("delete FORM_VALUE where Appeal_Id=@appID;delete from appeals where Appeal_Id=@appID", con)
                                    cmdSelect.Parameters.Add("@appID", SqlDbType.Int).Value = CInt(appID)
                                    con.Open()
                                    cmdSelect.ExecuteNonQuery()
                                    con.Close()
                                    appID = 0
                                Else
                                    resList.Add(LogId)
                                    'Response.Write("<br>resList.len=" & resList.Count & " add log :" & LogId)
                                End If

                            Else
                                '' Response.Write("<br>app_id=0")
                                status = "ERROR"
                                appID = 0
                            End If


                        End If


                    End If

                End If
            Next
        End If
        Try
            If resList.Count = 0 Then
                Response.Write("ERROR")
            Else
                Response.Write(Strings.Join(resList.ToArray, ","))
            End If
        Catch ex As Exception
            'Response.Write(Err.Description)
        End Try
    End Sub

End Class
