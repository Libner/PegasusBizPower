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
            getTimingVisiblity = "<span style='color:green'>����</span>"
        Else
            getTimingVisiblity = "<span style='color:red'>�� ����</span>"
        End If
    End Function
    Function getTimingFrequancy(ByVal Frequancy) As String
        '��� ����� �������� - value=0  - no job schedule - disable job
        '��� ���� - value=1  - 1 day
        '��� ����� - value=7 - 1 week
        '��� ����� - value=30 - 1 month
        If func.fixNumeric(Frequancy) = 0 Then
            getTimingFrequancy = "�� ����"
        End If
        If func.fixNumeric(Frequancy) = 1 Then
            getTimingFrequancy = "��� ����"
        End If
        If func.fixNumeric(Frequancy) = 7 Then
            getTimingFrequancy = "��� �����"
        End If
        If func.fixNumeric(Frequancy) = 30 Then
            getTimingFrequancy = "��� �����"
        End If
    End Function
    Function getTimingDay(ByVal DayNumber, ByVal Frequancy) As String
        '=============���� ���=================
        '��� ����� ����� ����� ����� ��� ����� �� ��� �����
        '(���� �� ��� ����� �� ���� �� �� ��� 15 ��� ����)
        Dim strDay As String = ""
        If func.fixNumeric(DayNumber) = -1 And func.fixNumeric(Frequancy) = 30 Then
            strDay = "��� ����� �����"
        End If
        If func.fixNumeric(DayNumber) = 1 And func.fixNumeric(Frequancy) = 30 Then
            strDay = "��� ����� �����"
        End If
        If func.fixNumeric(DayNumber) > 1 And func.fixNumeric(DayNumber) <= 30 And func.fixNumeric(Frequancy) = 30 Then
            strDay = "��� " & DayNumber & " �����"
        End If


        If func.fixNumeric(DayNumber) = -1 And func.fixNumeric(Frequancy) = 7 Then
            strDay = "��� ����� ����"
        End If
        If func.fixNumeric(DayNumber) = 1 And func.fixNumeric(Frequancy) = 7 Then
            strDay = "��� ����� ����"
        End If
        If func.fixNumeric(DayNumber) > 1 And func.fixNumeric(DayNumber) <= 7 And func.fixNumeric(Frequancy) = 7 Then
            strDay = "��� " & DayNumber & " ����"
        End If
        getTimingDay = strDay
    End Function


    Function getTimingFormsDateInterval(ByVal FormsDateInterval) As String
        '==========���� ������ ����� ���� ������ ���� ����� �����=====================
        '������ ������� - value=2 - 2 last days by default!!!
        '����� ����� - value=7 - 1 last week
        '���� ����� value=30 - 1 last month
        '������� ������� - value=60 - 2 last monthes
        '����� ������ ������� - value=90 - 3 last monthes
        '��� ��� ������ - value=180 - last half of last year
        If func.fixNumeric(FormsDateInterval) = 2 Then
            getTimingFormsDateInterval = "������ �������"
        End If
        If func.fixNumeric(FormsDateInterval) = 7 Then
            getTimingFormsDateInterval = "����� �����"
        End If
        If func.fixNumeric(FormsDateInterval) = 30 Then
            getTimingFormsDateInterval = "���� �����"
        End If
        If func.fixNumeric(FormsDateInterval) = 60 Then
            getTimingFormsDateInterval = "������� �������"
        End If
        If func.fixNumeric(FormsDateInterval) = 90 Then
            getTimingFormsDateInterval = "����� ������ �������"
        End If
        If func.fixNumeric(FormsDateInterval) = 180 Then
            getTimingFormsDateInterval = "��� ��� ������"
        End If
    End Function
End Class

