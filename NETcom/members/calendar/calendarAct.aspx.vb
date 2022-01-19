Imports System.Data.SqlClient
Public Class calendarAct
    Inherits System.Web.UI.Page
    Protected i, j, yy As Integer
    Public dt As String
    Protected func As New bizpower.cfunc

    Protected currYear, YearStatus, YearStatusName As String
    Public dayT As String
    Protected WithEvents rptDays, rptMonth As Repeater

    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))




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
        Session.LCID = 1037
        dt = Year(Now())
        If Not IsNumeric(Request("y")) Then
            currYear = dt
        Else
            currYear = Request("y")

        End If
        ' Response.Write(currYear)
        ' Response.End()
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT Year_Status from YearsStatus where Year='" & currYear & "'", con)
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        Dim myReader As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()
        If myReader.Read Then
            ' Response.Write("yes")
            If Not IsDBNull(myReader("Year_Status")) Then
                YearStatus = myReader("Year_Status")
            End If
        End If
        con.Close()

        If YearStatus = "1" Then
            YearStatusName = "עמוד תקין"
        Else
            YearStatusName = "עמוד לא תקין"
        End If


        Dim Months As New ArrayList
        For m As Integer = 12 To 1 Step -1
            Months.Add(m)
        Next

        rptMonth.DataSource = Months
        rptMonth.DataBind()


    End Sub


    Private Sub rptMonth_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptMonth.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim rptDays As Repeater
            Dim month As Integer = e.Item.DataItem
            rptDays = e.Item.FindControl("rptDays")

            Dim cmdSelect = New SqlCommand("SELECT DateKey, CalendarYear, CalendarMonth, CalendarMonthName, CalendarDay,HolidayName,DayFontColor, WeekdayNumber, WeekdayNameHeb,DayTypeName,DayTypeColor,IsHoliday " & _
                 "  FROM DimDate  WHERE CalendarYear = @currYear and CalendarMonth=@month ORDER BY CalendarDay, CalendarMonth", con)
            cmdSelect.Parameters.Add("@currYear", SqlDbType.Int).Value = CInt(currYear)
            cmdSelect.Parameters.Add("@month", SqlDbType.Int).Value = CInt(month)

            con.Open()
            Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

            If dr.HasRows Then
                '  Response.Write(PageSize.SelectedValue)
                '  Response.End()
                rptDays.DataSource = dr
                rptDays.DataBind()

            End If
            con.Close()

        End If




    End Sub
End Class