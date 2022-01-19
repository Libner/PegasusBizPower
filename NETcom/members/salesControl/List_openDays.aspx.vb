Imports System.Data.SqlClient
Public Class List_openDays
    Inherits System.Web.UI.Page
    Protected func As New bizpower.cfunc
    Protected dateStart, dateEnd As String
    Protected uid As Integer
    Protected pdateStart, pdateEnd As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected rptData As Repeater


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
        If Request("dStart") = "" Then
            dateStart = "1/" & Month(Now()) & "/" & Year(Now())
        Else
            dateStart = Request("dStart")
        End If
        If Request("dEnd") = "" Then
            dateEnd = Now().ToString("dd/MM/yyyy")
        Else
            dateEnd = Request("dEnd")
        End If
        If IsNumeric(Request.QueryString("uid")) Then
            uid = Request.QueryString("uid")
        End If

        pdateStart = Year(dateStart) & "/" & Month(dateStart) & "/" & Day(dateStart)
        pdateEnd = Year(dateEnd) & "/" & Month(dateEnd) & "/" & Day(dateEnd)
        If Not IsPostBack Then
            GetData()
        End If
    End Sub
    Sub GetData()
        Dim dr As SqlDataReader
        ' Dim cmdSelect = New SqlCommand("select DateKey,WeekdayNameHeb from DimDate where   NOT EXISTS( SELECT LogInDate FROM Users_LogIn " & _
        '"  WHERE DimDate.DateKey =Users_LogIn.LogInDate and User_Id=" & uid & " ) and  DateDiff(d,DateKey, convert(smalldatetime,'" & pdateStart & "',101)) <= 0 " & _
        '"  AND DateDiff(d,DateKey, convert(smalldatetime,'" & pdateEnd & "',101)) >= 0 " & _
        '"   and DayTypeId<>'0'", con)
        Dim cmdSelect = New SqlCommand("select DateKey,WeekdayNameHeb from DimDate where   NOT EXISTS( SELECT LogInDate FROM Users_LogIn " & _
   "  WHERE  DateDiff(d,DimDate.DateKey,Users_LogIn.LogInDate)=0  and User_Id=" & uid & " ) and  DateDiff(d,DateKey, convert(smalldatetime,'" & pdateStart & "',101)) <= 0 " & _
   "  AND DateDiff(d,DateKey, convert(smalldatetime,'" & pdateEnd & "',101)) >= 0 " & _
   "   and DayTypeId<>'0'", con)
        'and  IsHoliday=0
        con.Open()
        dr = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        '  Response.End()
        rptData.DataSource = dr
        rptData.DataBind()
        con.Close()





    End Sub

End Class
