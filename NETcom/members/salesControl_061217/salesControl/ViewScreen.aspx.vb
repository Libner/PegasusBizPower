Imports System.Data.SqlClient

Public Class ViewScreen
    Inherits System.Web.UI.Page
    Protected i, j, yy, mm As Integer
    Public dt, mt As String
    Protected func As New bizpower.cfunc
    Protected Current_MonthDays, pMonthDays As Integer
    Protected MonthDays As Decimal
    Protected currYear, currMonth As String
    Public dayT As String
    Protected rptDep As Repeater
    Protected WithEvents rptDays As Repeater
    Protected PEdit As String
    Public dep As String
    Protected lbl16735, lbl16735_bitul As Label
    Protected WithEvents seldep As System.Web.UI.HtmlControls.HtmlSelect
    Public dtDep As New DataTable
    Dim primKeydtDep(0) As Data.DataColumn
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dr_m As SqlClient.SqlDataReader
    Protected WithEvents rptTableService, rptTableSales As Repeater
    Protected sumDays, sumP16719Kishurit_Sales, sumP16724ContactUs_Sales, sumcall_Sales, sumP16735, sumP16735_Bitulim, sumP16724ContactUs, sumP16719Kishurit As Decimal
    Protected sumPotencialSales, sumPotencialSevice, sumP16504, sum16735Days, sumP16719Kishurit_Service, sumP17012ContactUs_Sales As Decimal
    Protected sumContactUs_Service, sumCallService, sumProcAzlaha, sumProcKishurit_Service, sumProcKishurit_Sales As Decimal

    Protected resSP16719Kishurit_Service As Decimal

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
        PEdit = Request.QueryString("PEdit")
        ' Response.Write("PEdit=" & PEdit)
        If Request("seldep") <> "" Then
            dep = Request("seldep")
        Else
            dep = "0"
        End If
        dt = Year(Now())
        mt = Month(Now())
        If Not IsNumeric(Request("y")) Then
            currYear = dt
        Else
            currYear = Request("y")

        End If

        If Not IsNumeric(Request("m")) Then
            currMonth = mt
        Else
            currMonth = Request("m")

        End If
        pMonthDays = DateTime.DaysInMonth(currYear, currMonth)
        MonthDays = CalcMonthDays(currMonth, currYear)
         ' Response.Write(MonthDays)
        '   Response.Write(currYear & ":" & currMonth)
        If Not IsPostBack Then
            GetDep()

        End If

        MonthData()
        TableSales()
        TableService()


    End Sub

    Function CalcMonthDays(ByVal pMonth As String, ByVal pYear As String) As String
        Dim result As String = ""
        Dim query As String

        Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        If DateDiff("d", "01/" & pMonth & "/" & pYear, Now()) < pMonthDays Then
            query = " where(Month(DateKey) = " & pMonth & " And Year(DateKey) = " & pYear & " And day(DateKey) < " & Day(Now()) & ")"

        Else
            query = " where(Month(DateKey) = " & pMonth & " And Year(DateKey) = " & pYear & ")"

        End If
        '  Response.Write(query)


        Dim cmdSelect As New SqlClient.SqlCommand("SELECT Sum(convert(decimal(6,2),DayTypeId)) as MonthDays from DimDate " & query, myConnection)
        cmdSelect.CommandType = CommandType.Text
        myConnection.Open()

        Dim myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        While myReader.Read()
            If Not IsDBNull(myReader("MonthDays")) Then

                result = myReader("MonthDays")

            Else
                result = "0"
            End If




        End While
        'If result = "" Then
        '    result = "הכל"
        'End If
        myConnection.Close()


        Return result
    End Function
    Sub TableSales()
        Dim cmdSelect As New SqlClient.SqlCommand("SET DATEFORMAT dmy;SELECT sum(call_Service)as Scall_Service,sum(call_Sales) as Scall_Sales,sum(IsNull(P16719Kishurit_Service,0)) as SP16719Kishurit_Service," & _
        " MONTH(DateKey) as SMonth,year(DateKey) as SYear from DimDate where " & _
        " month(DateKey)='" & currMonth & "' and (datediff(yyyy,DateKey,'01/" & currMonth & "/" & currYear & "')>=0  and datediff(yyyy,DateKey,'01/" & currMonth & "/" & currYear & "')<3)" & _
        " group by MONTH(DateKey),year(DateKey) order by year(DateKey) desc", con)
        cmdSelect.CommandType = CommandType.Text
        con.Open()

        Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rptTableSales.DataSource = dr
            rptTableSales.DataBind()
            rptTableSales.Visible = True
        End If

        dr.Close()
        cmdSelect.Dispose()
        con.Close()

    End Sub
    Sub TableService()
        Dim cmdSelect As New SqlClient.SqlCommand("SET DATEFORMAT dmy;SELECT sum(call_Service)as Scall_Service,sum(call_Sales) as Scall_Sales,sum(IsNull(P16719Kishurit_Service,0)) as SP16719Kishurit_Service," & _
        " MONTH(DateKey) as SMonth,year(DateKey) as SYear from DimDate where " & _
        " month(DateKey)='" & currMonth & "' and (datediff(yyyy,DateKey,'01/" & currMonth & "/" & currYear & "')>=0  and datediff(yyyy,DateKey,'01/" & currMonth & "/" & currYear & "')<3)" & _
        " group by MONTH(DateKey),year(DateKey) order by year(DateKey) desc", con)
        cmdSelect.CommandType = CommandType.Text
        con.Open()

        Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rptTableService.DataSource = dr
            rptTableService.DataBind()
            rptTableService.Visible = True
        End If

        dr.Close()
        cmdSelect.Dispose()
        con.Close()

    End Sub
    Sub GetDep()
        Dim cmdSelect = New SqlCommand("Select departmentId, departmentName,PriorityLevel from Departments order by PriorityLevel", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtDep)
        con.Close()
        primKeydtDep(0) = dtDep.Columns("departmentId")
        dtDep.PrimaryKey = primKeydtDep
        seldep.Items.Clear()
        '  Dim list1 As New ListItem("הכל", "0")
        '  seldep.Items.Add(list1)
        For i As Integer = 0 To dtDep.Rows.Count - 1
            Dim list As New ListItem(dtDep.Rows(i)("departmentName"), dtDep.Rows(i)("departmentId"))
            If Request("seldep") > 0 And Request("seldep") = dtDep.Rows(i)("departmentId") Then
                list.Selected = True
            ElseIf i = 0 Then
                list.Selected = True
                dep = dtDep.Rows(i)("departmentId")
            End If
            seldep.Items.Add(list)
        Next
        'Response.Write("dep=" & dep)
        ' Response.End()


    End Sub

    Sub MonthData()
        ' Response.Write("<BR>" & currYear & ":" & currMonth)
        'Dim sql = "SET DATEFORMAT DMY;SELECT A.DateKey,DayTypeId ,IsNull(P16719Kishurit_Sales,0) as P16719Kishurit_Sales,IsNull(P16719Kishurit_Service,0) as P16719Kishurit_Service," & _
        '" isNull(P16724ContactUs_Sales,0) as P16724ContactUs_Sales,IsNull(P16724ContactUs_Service,0) as P16724ContactUs_Service,IsNull(P17012ContactUs_Sales,0) as  P17012ContactUs_Sales,IsNull(P16735,0) as P16735," & _
        '" IsNull(P16735_Bitulim,0) as P16735_Bitulim, IsNull(P17012ContactUs_Service,0) as  P17012ContactUs_Service,IsNull(P16504,0) as P16504, call_Service,call_Sales, CalendarYear, " & _
        '" CalendarDay , HolidayName, DayFontColor, WeekdayNumber, WeekdayNameHeb, DayTypeName, DayTypeColor, IsHoliday) FROM DimDate A  left join Departments_Data DD on DD.DateKey=A.DateKey " & _
        '"  WHERE CalendarYear = " & currYear & " and CalendarMonth=" & currMonth & " and DD.departmentId=" & dep & " ORDER BY CalendarDay, CalendarMonth"
        'Response.Write(sql)
        'Response.End()
        ' Dim cmdSelect = New SqlCommand("SET DATEFORMAT DMY;SELECT DateKey,DayTypeId ,IsNull(P16719Kishurit_Sales,0) as P16719Kishurit_Sales,IsNull(P16719Kishurit_Service,0) as P16719Kishurit_Service,isNull(P16724ContactUs_Sales,0) as P16724ContactUs_Sales,IsNull(P16724ContactUs_Service,0) as P16724ContactUs_Service,IsNull(P17012ContactUs_Sales,0) as  P17012ContactUs_Sales, IsNull(P17012ContactUs_Service,0) as  P17012ContactUs_Service,IsNull(P16504,0) as P16504, call_Service,call_Sales, CalendarYear, CalendarMonth, CalendarMonthName, CalendarDay,HolidayName,DayFontColor, WeekdayNumber, WeekdayNameHeb,DayTypeName,DayTypeColor,IsHoliday " & _
        '               "  FROM DimDate  WHERE CalendarYear = @currYear and CalendarMonth=@month ORDER BY CalendarDay, CalendarMonth", con)

        Dim cmdSelect = New SqlCommand("SET DATEFORMAT DMY;SELECT A.DateKey,DayTypeId ,IsNull(P16719Kishurit_Sales,0) as P16719Kishurit_Sales,IsNull(P16719Kishurit_Service,0) as P16719Kishurit_Service," & _
        " isNull(P16724ContactUs_Sales,0) as P16724ContactUs_Sales,IsNull(P16724ContactUs_Service,0) as P16724ContactUs_Service,IsNull(P17012ContactUs_Sales,0) as  P17012ContactUs_Sales,IsNull(P16735,0) as P16735," & _
        " IsNull(P16735_Bitulim,0) as P16735_Bitulim, IsNull(P17012ContactUs_Service,0) as  P17012ContactUs_Service,IsNull(P16504,0) as P16504, call_Service,call_Sales, CalendarYear, " & _
        " CalendarDay , HolidayName, DayFontColor, WeekdayNumber, WeekdayNameHeb, DayTypeName, DayTypeColor, IsHoliday FROM DimDate A  left join Departments_Data DD on DD.DateKey=A.DateKey " & _
        "  WHERE CalendarYear = @currYear and CalendarMonth=@month  and DD.departmentId=@dep ORDER BY CalendarDay, CalendarMonth ", con)
        '
        cmdSelect.Parameters.Add("@currYear", SqlDbType.Int).Value = CInt(currYear)
        cmdSelect.Parameters.Add("@month", SqlDbType.Int).Value = CInt(currMonth)
        cmdSelect.Parameters.Add("@dep", SqlDbType.Int).Value = CInt(dep)
        con.Open()
        Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rptDays.DataSource = dr
            rptDays.DataBind()
        Else
            rptDays.Visible = False

        End If
        con.Close()

    End Sub

    Private Sub rptDays_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptDays.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            sumDays = sumDays + CDbl(e.Item.DataItem("DayTypeId"))
            sumP16719Kishurit_Sales = sumP16719Kishurit_Sales + CDbl(e.Item.DataItem("P16719Kishurit_Sales"))
            sumP16724ContactUs_Sales = sumP16724ContactUs_Sales + CDbl(e.Item.DataItem("P16724ContactUs_Sales")) + CDbl(e.Item.DataItem("P17012ContactUs_Sales"))
            sumcall_Sales = sumcall_Sales + CDbl(e.Item.DataItem("call_Sales"))
            sumP16735 = sumP16735 + CDbl(e.Item.DataItem("P16735"))
            sumP16735_Bitulim = sumP16735_Bitulim + CDbl(e.Item.DataItem("P16735_Bitulim"))
            sumP16724ContactUs = sumP16724ContactUs + CDbl(e.Item.DataItem("P16724ContactUs_Service")) + CDbl(e.Item.DataItem("P16724ContactUs_Sales")) + CDbl(e.Item.DataItem("P17012ContactUs_Service")) + CDbl(e.Item.DataItem("P17012ContactUs_Sales"))
            sumP16719Kishurit = sumP16719Kishurit + (CDbl(e.Item.DataItem("P16719Kishurit_Service") + CDbl(e.Item.DataItem("P16719Kishurit_Sales"))))
            sumPotencialSales = sumPotencialSales + CDbl(e.Item.DataItem("P16719Kishurit_Sales")) + CDbl(e.Item.DataItem("P16724ContactUs_Sales")) + CDbl(e.Item.DataItem("P17012ContactUs_Sales")) + CDbl(e.Item.DataItem("call_Sales"))
            sumP16504 = sumP16504 + CDbl(e.Item.DataItem("P16504"))
            'netto rishum sum16735Days
            sum16735Days = sum16735Days + (CDbl(e.Item.DataItem("P16735")) - CDbl(e.Item.DataItem("P16735_Bitulim")))
            sumP16719Kishurit_Service = sumP16719Kishurit_Service + CDbl(e.Item.DataItem("P16719Kishurit_Service"))
            sumContactUs_Service = sumContactUs_Service + CDbl(e.Item.DataItem("P16724ContactUs_Service")) + CDbl(e.Item.DataItem("P17012ContactUs_Service"))

            sumCallService = sumCallService + CDbl(e.Item.DataItem("call_Service"))
            sumPotencialSevice = sumPotencialSevice + (CDbl(e.Item.DataItem("P16719Kishurit_Service")) + CDbl(e.Item.DataItem("P16724ContactUs_Service")) + CDbl(e.Item.DataItem("call_Service")))
            'If e.Item.DataItem("P16504") <> 0 Then
            'sumProcAzlaha = sumProcAzlaha + (CDbl(e.Item.DataItem("P16735")) / (CDbl(e.Item.DataItem("P16504")) * 1.8)) * 100

            'End If
            If e.Item.DataItem("call_Sales") <> 0 Then
                sumProcKishurit_Sales = sumProcKishurit_Sales + e.Item.DataItem("P16719Kishurit_Sales") / e.Item.DataItem("call_Sales") * 100
            End If
            'If sumCallService > 0 Then
            '  sumProcKishurit_Service = sumProcKishurit_Service + CDbl(e.Item.DataItem("P16719Kishurit_Service")) / CDbl(e.Item.DataItem("call_Sales")) * 100
            ' sumProcKishurit_Service = (sumP16719Kishurit_Service / sumCallService) * 100
            ' End  If
            'sumProcAzlaha=  String.Format("{0:0.00}", 100 * CInt(sum16735Days) / (CInt(sumP16504) * 1.8))



            '  Response.Write(sumDays)
            '''    Dim dd As Date
            '''    Dim count, countBitul As String
            '''    lbl16735 = e.Item.FindControl("lbl16735")
            '''    dd = e.Item.DataItem("DateKey")

            '''    count = func.GetCountSaleByDate(dep, 40660, dd)
            '''    lbl16735.Text = count
            '''    lbl16735_bitul = e.Item.FindControl("lbl16735_bitul")


            '''    countBitul = func.GetCountSaleByDate(dep, 40661, dd)
            '''    lbl16735_bitul.Text = countBitul



        End If

    End Sub

    Private Sub rptTableService_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptTableService.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            'If DateDiff("d", "01/" & e.Item.DataItem("SMonth") & "/" & e.Item.DataItem("SYear"), Now()) < MonthDays Then
            Dim LblKishurit As Label
            LblKishurit = e.Item.FindControl("LblKishurit")
            ' Response.Write("f=" & e.Item.DataItem("SP16719Kishurit_Service"))
            'End If

            ' resSP16719Kishurit_Service = CDbl(e.Item.DataItem("SP16719Kishurit_Service")) / 15
            resSP16719Kishurit_Service = Math.Round(CDbl(e.Item.DataItem("SP16719Kishurit_Service")) / MonthDays, 0)
            LblKishurit.Text = resSP16719Kishurit_Service
            '    Response.Write("s=" & resSP16719Kishurit_Service & "<BR>")
        End If


    End Sub
End Class
