Public Class info
    Inherits System.Web.UI.Page
    Protected UserId, OrgId, currentMonthRCount, pastMonthRCount, userTarget, maxCurrentMonthRCount, _
    sumCurrentMonthRCount, currDays, currDaysInMonth, futureOrderMoney As Integer
    Protected OrderPrice As Double = 0.0
    Protected func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))
    Dim cmdSelect As New SqlClient.SqlCommand

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
        UserId = func.fixNumeric(Request.Cookies("bizpegasus")("UserId"))
        OrgId = func.fixNumeric(Request.Cookies("bizpegasus")("orgId"))

        If (UserId = 0 Or OrgId = 0) Then
            Response.End()
        End If

        cmdSelect = New SqlClient.SqlCommand("SELECT " & _
        " SUM(CASE WHEN (Month(Appeal_Date) = @CurrMonth AND [User_Id] = @UserId) THEN CountPeople ELSE 0 END) as [currentMonthRCount], " & _
        " SUM(CASE WHEN (Month(Appeal_Date) = (@CurrMonth - 1) AND [User_Id] = @UserId) THEN CountPeople ELSE 0 END) as [pastMonthRCount], " & _
        " SUM(CASE WHEN (Month(Appeal_Date) = @CurrMonth) THEN CountPeople ELSE 0 END) as [sumCurrentMonthRCount] " & _
        " from dbo.tbl_reservations() R  WHERE (Year(Appeal_Date) = @CurrYear)", con)
        cmdSelect.CommandType = CommandType.Text
        cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = CInt(UserId)
        cmdSelect.Parameters.Add("@CurrMonth", SqlDbType.Int).Value = Now.Month
        cmdSelect.Parameters.Add("@CurrYear", SqlDbType.Int).Value = Now.Year
        con.Open()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
        If dr.Read() Then
            currentMonthRCount = func.fixNumeric(dr("currentMonthRCount"))
            pastMonthRCount = func.fixNumeric(dr("pastMonthRCount"))
            sumCurrentMonthRCount = func.fixNumeric(dr("sumCurrentMonthRCount"))
        End If
        cmdSelect.Dispose() : dr.Close() : con.Close()

        cmdSelect = New SqlClient.SqlCommand("WITH tmp AS (SELECT [User_Id], SUM(CountPeople) OVER " & _
       " (Partition BY [User_Id]) as [sumCurrentMonthRCount] from dbo.tbl_reservations() R WHERE " & _
      " (Month(Appeal_Date) = @CurrMonth) AND (Year(Appeal_Date) = @CurrYear)) " & _
     " SELECT MAX(sumCurrentMonthRCount) as [maxCurrentMonthRCount] FROM tmp", con)

        cmdSelect.Parameters.Add("@CurrMonth", SqlDbType.Int).Value = Now.Month
        cmdSelect.Parameters.Add("@CurrYear", SqlDbType.Int).Value = Now.Year
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        maxCurrentMonthRCount = func.fixNumeric(cmdSelect.ExecuteScalar())
        cmdSelect.Dispose() : con.Close()

        cmdSelect = New SqlClient.SqlCommand("SELECT Month_Min_Order, Order_Price " & _
        " FROM [USERS] WHERE ([User_Id] = @UserId)", con)
        cmdSelect.CommandType = CommandType.Text
        cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = CInt(UserId)
        con.Open()
        dr = cmdSelect.ExecuteReader(CommandBehavior.SingleRow)
        If dr.Read() Then
            userTarget = func.fixNumeric(dr("Month_Min_Order"))
            OrderPrice = func.fixNumeric(dr("Order_Price"))
        End If
        cmdSelect.Dispose() : dr.Close() : con.Close()

        currDaysInMonth = System.DateTime.DaysInMonth(Now.Year, Now.Month) - 4

        currDays = System.DateTime.Today.Day

        futureOrderMoney = ((currentMonthRCount / currDays * 1.0) * 1.1) * currDaysInMonth * OrderPrice
    End Sub

End Class
