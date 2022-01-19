Imports System.Data.SqlClient
Public Class updateTimingAutoSending

    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Dim BizpowerDBName As String = ConfigurationSettings.AppSettings("BizpowerDBName")
    Dim SiteDBName As String = ConfigurationSettings.AppSettings("PegasusSiteDBName")
    Dim isSuccess As Boolean = False
    Public UpdateTimingPermitted As Boolean
    Public Frequancy, DateInterval, DayNumber As Integer
    Public countriesSequence, timingTitle As String
    Dim countriesList() As String
    Protected WithEvents sCountries, selDayNumber As System.Web.UI.HtmlControls.HtmlSelect
    Public dtCountries As New DataTable
    Dim primKeydtCountries(0) As Data.DataColumn
    Public TimingId As Integer
    Protected WithEvents btnSubmit As Web.UI.WebControls.LinkButton
    Dim userId, orgID, lang_id, DepartmentId, companyID As String

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

 
        userId = Trim(Request.Cookies("bizpegasus")("UserId"))
        orgID = Trim(Request.Cookies("bizpegasus")("OrgID"))
        lang_id = Trim(Request.Cookies("bizpegasus")("LANGID"))

        If Request.Cookies("bizpegasus")("UpdateTimingMailReservationForm") = "1" Then
            UpdateTimingPermitted = True
        End If

        TimingId = func.fixNumeric(Request.QueryString("TimingId"))

        '==========מועד השליחה========================================================
        'ללא שליחה אוטומטיט - value=0  - no job schedule - disable job
        'אחת ליום - value=1  - 1 day
        'אחת לשבוע - value=7 - 1 week
        'אחת לחודש - value=30 - 1 month

        '=============מספר יום=================
        'יום מסוים להרצה במקרה שנבחר פעם בשבוע או פעם בחודש
        '(למשל כל יום ראשון כל שבוע או כל יום 15 בכל חודש)

        '==========טווח תאריכי פתיחת טפסי הרישום ביחס למועד ההפצה=====================
        'יומיים אחרונים - value=2 - 2 last days by default!!!
        'השבוע אחרון - value=7 - 1 last week
        'חודש אחרון value=30 - 1 last month
        'חודשיים אחרונים - value=60 - 2 last monthes
        'שלושה חודשים אחרונים - value=90 - 3 last monthes
        'חצי שנה אחרונה - value=180 - last half of last year
        '==========יעדים נבחרים=======================================================
        'countries CRM ID list with comma delimeter / null("") - all countries
        If Not IsPostBack Then
            GetCountries()
            If TimingId > 0 Then
                GetData()
            End If
        End If

    End Sub
    Sub GetCountries()
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
        sCountries.Items.Clear()
        Dim list1 As New ListItem("----בחר יעדים----", "")
        sCountries.Items.Add(list1)
        Dim list2 As New ListItem("----כל היעדים----", "0")
        sCountries.Items.Add(list2)
        For i As Integer = 0 To dtCountries.Rows.Count - 1
            Dim list As New ListItem(dtCountries.Rows(i)("Country_Name"), dtCountries.Rows(i)("Country_Id"))
            sCountries.Items.Add(list)
        Next
    End Sub

    Sub GetData()
        Frequancy = 0
        DateInterval = 2
        countriesSequence = ""
        DayNumber = 0

        cmdSelect = New SqlClient.SqlCommand("select Timing_Title,Sending_Frequancy,Sending_Forms_DateInterval,Sending_countriesList,Sending_Forms_DayNumber from ReservationForm_SendMail_Job_Timing where Timing_Id=@Timing_Id", con)

        cmdSelect.Parameters.Add("@Timing_Id", func.fixNumeric(TimingId))
        con.Open()
        Dim dr As SqlDataReader
        dr = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
        If dr.Read Then
            timingTitle = func.dbNullFix(dr("Timing_Title"))
            Frequancy = func.fixNumeric(dr("Sending_Frequancy"))
            DateInterval = func.fixNumeric(dr("Sending_Forms_DateInterval"))
            countriesSequence = func.dbNullFix(dr("Sending_countriesList"))
            DayNumber = func.fixNumeric(dr("Sending_Forms_DayNumber"))
        End If
        con.Close()

        'mark items to selected from countriesSequence
        If countriesSequence <> "" Then
            countriesList = countriesSequence.Replace(" ", "").Split(",")
            If countriesList.Length = 0 Then
                sCountries.Items(1).Selected = True
            Else
                For i As Integer = 0 To countriesList.Length - 1
                    sCountries.Items.FindByValue(countriesList(i)).Selected = True
                Next
            End If
        End If
    End Sub


    Sub btnSubmit_onClick(ByVal s As Object, ByVal e As EventArgs) Handles btnSubmit.Click
        Dim checkFrequancy, checkInterval, countries, RunDayNumber, timingTitle As String

        timingTitle = func.dbNullFix(Request.Form("timingTitle"))
        Frequancy = func.dbNullFix(Request.Form("checkFrequancy"))
        checkInterval = func.dbNullFix(Request.Form("checkInterval"))
        countries = func.dbNullFix(Request.Form("sCountries")).Replace(" ", "")
        RunDayNumber = func.dbNullFix(Request.Form("selDayNumber"))
        If checkInterval = "" Then
            checkInterval = "2"
        End If


        Dim sqlStr As String

        If TimingId = 0 Then
            sqlStr = "insert into ReservationForm_SendMail_Job_Timing (Timing_Title,Sending_Frequancy,Sending_Forms_DateInterval,Sending_countriesList,Sending_Forms_DayNumber,Creation_Date,CreationUser_Id) " & _
            " values (@Timing_Title,@Frequancy,@DateInterval,@countriesList,@DayNumber,getdate(),@userId)"
        Else
            sqlStr = "update ReservationForm_SendMail_Job_Timing set Timing_Title=@Timing_Title,Sending_Frequancy=@Frequancy,Sending_Forms_DateInterval=@DateInterval,Sending_countriesList=@countriesList,Sending_Forms_DayNumber=@DayNumber,Modify_Date=getdate(),ModifyUser_Id=@userId" & _
        " where Timing_Id=@Timing_Id"
        End If
        con.Open()
        Dim cmdUpd As New SqlCommand(sqlStr, con)
        cmdUpd.Parameters.Add("@Timing_Title", func.dbNullFix(timingTitle))
        cmdUpd.Parameters.Add("@Frequancy", func.fixNumeric(Frequancy))
        If checkInterval = "" Then
            cmdUpd.Parameters.Add("@DateInterval", 2) 'by default
        Else
            cmdUpd.Parameters.Add("@DateInterval", func.fixNumeric(checkInterval))
        End If
        If RunDayNumber = "" Then
            cmdUpd.Parameters.Add("@DayNumber", 1) 'by default
        Else
            cmdUpd.Parameters.Add("@DayNumber", func.fixNumeric(RunDayNumber))
        End If
        If TimingId > 0 Then
            cmdUpd.Parameters.Add("@Timing_Id", func.fixNumeric(TimingId))
        End If

        cmdUpd.Parameters.Add("@userId", func.fixNumeric(userId))

        cmdUpd.Parameters.Add("@countriesList", countries)
        cmdUpd.ExecuteNonQuery()
        con.Close()

        GetData()


        '' Create the TaskService object.
        ''???domain, user, paswword
        '' Dim ts = New TaskService("RemoteServerName", "User", "Domain", "Password")

        'Dim Scheduler As New TaskScheduler.TaskSchedulerClass
        '''???? - user....
        'Scheduler.Connect(Nothing, Nothing, Nothing, Nothing)

        'Dim folder As TaskScheduler.ITaskFolder = Scheduler.GetFolder("\")
        'Dim Task As TaskScheduler.IRegisteredTask = folder.GetTask("pegasus_SendReservationForm")
        'Dim NewTask As TaskScheduler.ITaskDefinition = Task.Definition

        'With (NewTask)
        '    If func.fixNumeric(Frequancy) = 1 Then
        '        Dim Trigger As TaskScheduler.IDailyTrigger = .Triggers.Item(1)
        '        With Trigger
        '            Dim StartBoundary As String = "00:01"
        '            .StartBoundary = StartBoundary
        '        End With
        '    End If
        '    If func.fixNumeric(Frequancy) = 7 Then
        '        Dim Trigger As TaskScheduler.IWeeklyTrigger = .Triggers.Item(1)
        '        With Trigger
        '            Dim StartBoundary As String = "00:01"
        '            .StartBoundary = StartBoundary
        '            .DaysOfWeek = 1
        '        End With
        '    End If
        '    If func.fixNumeric(Frequancy) = 30 Then
        '        Dim Trigger As TaskScheduler.IMonthlyTrigger = .Triggers.Item(1)
        '        With Trigger
        '            Dim StartBoundary As String = "00:01"
        '            .StartBoundary = StartBoundary
        '            .DaysOfMonth = 1
        '        End With
        '    End If
        'End With
        'folder.RegisterTaskDefinition("pegasus_SendReservationForm", NewTask, TaskScheduler._TASK_CREATION.TASK_UPDATE, "", "", TaskScheduler._TASK_LOGON_TYPE.TASK_LOGON_PASSWORD, "")

    End Sub
End Class
