Imports System.Data.SqlClient
Public Class targetScreen
    Inherits System.Web.UI.Page
    Protected i, j, yy, mm As Integer
    Public dt, mt As String
    Protected func As New bizpower.cfunc
    Protected MonthDays, Current_MonthDays As Integer
    Protected currYear, currMonth As String
    Public dayT As String
    Protected rptDep As Repeater
    Protected WithEvents rptData As Repeater
    Protected PEdit As String
    Public dep As String
    Protected lbl16735, lbl16735_bitul As Label
    Protected WithEvents seldep As System.Web.UI.HtmlControls.HtmlSelect
    Public dtDep As New DataTable
    Dim primKeydtDep(0) As Data.DataColumn
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dr_m As SqlClient.SqlDataReader
    Protected rptTableService As Repeater
    Public dtData As New DataTable
    Protected WithEvents rptMonth As Repeater
    Public SumSales, SumSalesYear, SumRow3, SumKishuriotYear, SumKishuriot, SumSales_Point, SumActiveTime, SumActiveTimeYear, SumActiveTime_Point, SumKishuritService_Point As String
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
        '  Response.Write("edit=" & Request.Cookies("bizpegasus")("salesControlPointsEdit"))

        PEdit = Request.Cookies("bizpegasus")("salesControlPointsEdit") ' Request.QueryString("PEdit")

        'Response.Write("PEdit=" & PEdit)
        If Not IsPostBack Then
            GetDep()

        End If
        If Request("seldep") <> "" Then
            dep = Request("seldep")
            'Else
            '    dep = "2"
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
        MonthDays = DateTime.DaysInMonth(currYear, currMonth)

        Dim Months As New ArrayList
        For m As Integer = 1 To 12 Step 1
            Months.Add(m)
        Next

        rptMonth.DataSource = Months
        rptMonth.DataBind()

        'Response.Write(MonthDays)
        '   Response.Write(currYear & ":" & currMonth)

        'dtData.Columns.Add("ID", GetType(Integer))
        'dtData.Columns.Add("departmentId", GetType(Integer))
        'dtData.Columns.Add("Dim_Month", GetType(Integer))
        'dtData.Columns.Add("DimYear", GetType(Integer))
        'dtData.Columns.Add("Sales", GetType(Double))
        'dtData.Columns.Add("Sales_Point", GetType(Double))
        'dtData.Columns.Add("KishuritService_Point", GetType(Double))
        'dtData.Columns.Add("Kishuriot", GetType(Double))
        'dtData.Columns.Add("ActiveTime_Point", GetType(Double))
        'dtData.Columns.Add("ActiveTime", GetType(Double))
        'Dim cmdSelect = New SqlCommand("Select * from Departments_Data_Monthly  where departmentId=" & dep & "and month=1", con)
        'con.Open()
        'Dim ad As New SqlClient.SqlDataAdapter
        'ad.SelectCommand = cmdSelect
        'ad.Fill(dtDep)


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
        ''  Dim list1 As New ListItem("כל מחלקות החברה", "0")
        ''   seldep.Items.Add(list1)
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
    Private Sub rptMonth_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptMonth.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then

            Dim month As Integer = e.Item.DataItem
            rptData = e.Item.FindControl("rptData")
            ' Response.Write(currYear & "-" & month & "-" & dep & "<BR>")
            ' Response.Write(month & "<BR>")
            'Response.Write(currYear & "<BR>")
            'Response.Write(dep & "<BR>")
            Dim cmdSelect = New SqlCommand("SELECT Dim_Month,IsNull(dbo.GetKishuriot(@currYear-1,Dim_Month,@dep),0) as Kishuriot,IsNull(Kishuriot,0) as KishuriotYear, " & _
            " IsNull(dbo.GetSales(@currYear-1,Dim_Month,@dep),0) as Sales,IsNull(Sales,0) as SalesYear," & _
            " IsNull(dbo.GetActiveTime(@currYear-1,Dim_Month,@dep),0) as ActiveTime,isNull(ActiveTime,0) as ActiveTimeYear, " & _
            " isNull(Sales_Point,0) as Sales_Point,IsNull(KishuritService_Point,0) as KishuritService_Point,IsNull(ActiveTime_Point,0) as ActiveTime_Point,IsNull(dbo.GetSActiveTime(@currYear-1,Dim_Month,@dep),0) as ActiveTimeBefore from Departments_Data_Monthly" & _
            "   WHERE DimYear = @currYear and Dim_Month=@month and  departmentId=@dep", con)
            cmdSelect.Parameters.Add("@currYear", SqlDbType.Int).Value = CInt(currYear)
            cmdSelect.Parameters.Add("@month", SqlDbType.Int).Value = CInt(month)
            cmdSelect.Parameters.Add("@dep", SqlDbType.Int).Value = CInt(dep)
            con.Open()
            Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

            If dr.HasRows Then
                '    '  Response.Write(PageSize.SelectedValue)
                '    '  Response.End()
                rptData.DataSource = dr
                rptData.DataBind()
                con.Close()
            Else
                con.Close()
                Dim cmdSelectBlank = New SqlCommand("SELECT top 1 @month as Dim_Month,0 as Sales,0 as SalesYear,0 as Sales_Point,0 as KishuritService_Point,0 as Kishuriot,0 as KishuriotYear,0 as ActiveTime_Point,0 as ActiveTime,0 as ActiveTimeYear,0 as ActiveTimeBefore from Departments_Data_Monthly" & _
                              " ", con)
                '   cmdSelect.Parameters.Add("@currYear", SqlDbType.Int).Value = CInt(currYear)
                cmdSelectBlank.Parameters.Add("@month", SqlDbType.Int).Value = CInt(month)
                ' cmdSelect.Parameters.Add("@dep", SqlDbType.Int).Value = CInt(dep)
                con.Open()
                Dim drBlank As SqlDataReader = cmdSelectBlank.ExecuteReader(CommandBehavior.CloseConnection)

                If drBlank.HasRows Then
                    '    '  Response.Write(PageSize.SelectedValue)
                    '    '  Response.End()
                    rptData.DataSource = drBlank
                    rptData.DataBind()
                End If
                con.Close()
            End If


        End If
        

        'Response.Write(rptData.DataSource.DataItem("Sales"))
        'SumSales = SumSales + e.Item.DataItem("Sales")
        'SumSales = e.Item.DataItem("Sales").GetType.ToString



    End Sub



    Private Sub rptData_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptData.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            SumSalesYear = SumSalesYear + e.Item.DataItem("SalesYear")
            SumSales = SumSales + e.Item.DataItem("Sales")
            SumKishuriotYear = SumKishuriotYear + e.Item.DataItem("KishuriotYear")
            SumKishuriot = SumKishuriot + e.Item.DataItem("Kishuriot")

            SumSales_Point = SumSales_Point + e.Item.DataItem("Sales_Point")
            SumActiveTime = SumActiveTime + e.Item.DataItem("ActiveTime")
            SumActiveTimeYear = SumActiveTimeYear + e.Item.DataItem("ActiveTimeYear")
            SumActiveTime_Point = SumActiveTime_Point + e.Item.DataItem("ActiveTime_Point")
            SumKishuritService_Point = SumKishuritService_Point + e.Item.DataItem("KishuritService_Point")
            If e.Item.DataItem("Sales_Point") > 0 Then
                SumRow3 = SumRow3 + Math.Round((e.Item.DataItem("Sales") / e.Item.DataItem("Sales_Point") * 100), 2)
            End If

        End If
    End Sub
End Class
