Imports System.Data.SqlClient
Public Class jobTimingList
    Inherits System.Web.UI.Page
    Protected WithEvents rptTimings As Repeater
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
  
    Dim cmdSelect As New SqlClient.SqlCommand
     Public UpdateTimingPermitted As Boolean

    Public userId As String
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
        'permitions to new appeal (mitan'en) creating 

        userId = Trim(Request.Cookies("bizpegasus")("UserId"))



        If Not Page.IsPostBack Then

            cmdSelect = New SqlClient.SqlCommand("SELECT  Timing_Id, Timing_Title, Sending_Frequancy, Sending_Forms_DateInterval, dbo.get_CountriesListById(Sending_countriesList) AS Sending_countriesList, Sending_Forms_DayNumber FROM ReservationForm_SendMail_Job_Timing", con)
            cmdSelect.CommandType = CommandType.Text
            con.Open()
            rptTimings.DataSource = cmdSelect.ExecuteReader
            rptTimings.DataBind()

        End If


       
    End Sub
   


    Private Sub rptTimings_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptTimings.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then


        End If
    End Sub

    Function getTimingVisiblity(ByVal Frequancy) As String
        If func.fixNumeric(Frequancy) > 0 Then
            getTimingVisiblity = "<span style='color:green'>פעיל</span>"
        Else
            getTimingVisiblity = "<span style='color:red'>לא פעיל</span>"
        End If
    End Function
    Function getTimingFrequancy(ByVal Frequancy) As String
        'ללא שליחה אוטומטיט - value=0  - no job schedule - disable job
        'אחת ליום - value=1  - 1 day
        'אחת לשבוע - value=7 - 1 week
        'אחת לחודש - value=30 - 1 month
        If func.fixNumeric(Frequancy) = 0 Then
            getTimingFrequancy = "לא פעיל"
        End If
        If func.fixNumeric(Frequancy) = 1 Then
            getTimingFrequancy = "אחת ליום"
        End If
        If func.fixNumeric(Frequancy) = 7 Then
            getTimingFrequancy = "אחת לשבוע"
        End If
        If func.fixNumeric(Frequancy) = 30 Then
            getTimingFrequancy = "אחת לחודש"
        End If
    End Function
    Function getTimingDay(ByVal DayNumber, ByVal Frequancy) As String
        '=============מספר יום=================
        'יום מסוים להרצה במקרה שנבחר פעם בשבוע או פעם בחודש
        '(למשל כל יום ראשון כל שבוע או כל יום 15 בכל חודש)
        Dim strDay As String = ""
        If func.fixNumeric(DayNumber) = -1 And func.fixNumeric(Frequancy) = 30 Then
            strDay = "יום אחרון בחודש"
        End If
        If func.fixNumeric(DayNumber) = 1 And func.fixNumeric(Frequancy) = 30 Then
            strDay = "יום ראשון בחודש"
        End If
        If func.fixNumeric(DayNumber) > 1 And func.fixNumeric(DayNumber) <= 30 And func.fixNumeric(Frequancy) = 30 Then
            strDay = "יום " & DayNumber & " בחודש"
        End If


        If func.fixNumeric(DayNumber) = -1 And func.fixNumeric(Frequancy) = 7 Then
            strDay = "יום אחרון שבוע"
        End If
        If func.fixNumeric(DayNumber) = 1 And func.fixNumeric(Frequancy) = 7 Then
            strDay = "יום ראשון שבוע"
        End If
        If func.fixNumeric(DayNumber) > 1 And func.fixNumeric(DayNumber) <= 7 And func.fixNumeric(Frequancy) = 7 Then
            strDay = "יום " & DayNumber & " שבוע"
        End If
        getTimingDay = strDay
    End Function


    Function getTimingFormsDateInterval(ByVal FormsDateInterval) As String
        '==========טווח תאריכי פתיחת טפסי הרישום ביחס למועד ההפצה=====================
        'יומיים אחרונים - value=2 - 2 last days by default!!!
        'השבוע אחרון - value=7 - 1 last week
        'חודש אחרון value=30 - 1 last month
        'חודשיים אחרונים - value=60 - 2 last monthes
        'שלושה חודשים אחרונים - value=90 - 3 last monthes
        'חצי שנה אחרונה - value=180 - last half of last year
        If func.fixNumeric(FormsDateInterval) = 2 Then
            getTimingFormsDateInterval = "יומיים אחרונים"
        End If
        If func.fixNumeric(FormsDateInterval) = 7 Then
            getTimingFormsDateInterval = "השבוע אחרון"
        End If
        If func.fixNumeric(FormsDateInterval) = 30 Then
            getTimingFormsDateInterval = "חודש אחרון"
        End If
        If func.fixNumeric(FormsDateInterval) = 60 Then
            getTimingFormsDateInterval = "חודשיים אחרונים"
        End If
        If func.fixNumeric(FormsDateInterval) = 90 Then
            getTimingFormsDateInterval = "שלושה חודשים אחרונים"
        End If
        If func.fixNumeric(FormsDateInterval) = 180 Then
            getTimingFormsDateInterval = "חצי שנה אחרונה"
        End If
    End Function
End Class

