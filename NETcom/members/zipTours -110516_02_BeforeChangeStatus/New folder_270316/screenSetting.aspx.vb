Public Class screenSetting
    Inherits System.Web.UI.Page
    Protected rptTitle, rptCustomers As Repeater
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected WithEvents sStatus, sSeries, sDepartments, sUsers As System.Web.UI.HtmlControls.HtmlSelect
    Public dtStatus, dtSeries, dtsDepartments, dtUsers As New DataTable

    Dim cmdSelect As New SqlClient.SqlCommand
    Dim primKeydtStatus(0), primKeydtSeries(0), primKeydtDepartments(0), primKeydtUsers(0) As Data.DataColumn
    Protected WithEvents btnSearch, btnSearchAll As Web.UI.WebControls.LinkButton




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
        If Not Page.IsPostBack Then
            cmdSelect = New SqlClient.SqlCommand("SELECT Column_Name  FROM ScreenSetting where Column_Visible=1 order by Column_Id desc", con)
            cmdSelect.CommandType = CommandType.Text

            con.Open()
            rptTitle.DataSource = cmdSelect.ExecuteReader()
            rptTitle.DataBind()
            con.Close()
            '   GetData()
            rptCustomers.DataSource = GetCustomersData(1)
            rptCustomers.DataBind()

            If ViewColumn(1) = True Then
                GetStatus()
            End If
            If ViewColumn(2) = True Then
                GetSeries()
            End If
            If ViewColumn(5) = True Then
                GettDepartments()
            End If
            GetUsers()
        End If
    End Sub
    Sub GetUsers()
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT USER_ID, FIRSTNAME + ' ' + LASTNAME as Username from Users  where [ACTIVE]=1 order by LASTNAME,FIRSTNAME", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtUsers)
        con.Close()
        primKeydtUsers(0) = dtUsers.Columns("USER_ID")
        dtUsers.PrimaryKey = primKeydtUsers
        sUsers.Items.Clear()
        Dim list1 As New ListItem("הכל", "0")
        sUsers.Items.Add(list1)
        For i As Integer = 0 To dtUsers.Rows.Count - 1
            Dim list As New ListItem(dtUsers.Rows(i)("Username"), dtUsers.Rows(i)("USER_ID"))
            sUsers.Items.Add(list)
        Next
    End Sub
    Sub GettDepartments()
        Dim cmdSelect As New SqlClient.SqlCommand("select  Dep_Id,Dep_Name from Departments ORDER BY Dep_Name", conPegasus)
        conPegasus.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtsDepartments)
        conPegasus.Close()
        primKeydtDepartments(0) = dtsDepartments.Columns("Dep_Id")
        dtsDepartments.PrimaryKey = primKeydtDepartments

        sDepartments.Items.Clear()
        Dim list1 As New ListItem("הכל", "0")
        sDepartments.Items.Add(list1)
        For i As Integer = 0 To dtsDepartments.Rows.Count - 1
            Dim list As New ListItem(dtsDepartments.Rows(i)("Dep_Name"), dtsDepartments.Rows(i)("Dep_Id"))
            sDepartments.Items.Add(list)
        Next


    End Sub
    Sub GetSeries()
        Dim cmdSelect As New SqlClient.SqlCommand("select  Series_Id,Series_Name from Series ORDER BY Series_Name", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtSeries)
        con.Close()
        primKeydtSeries(0) = dtStatus.Columns("Series_Id")
        dtSeries.PrimaryKey = primKeydtSeries
        sSeries.Items.Clear()
        Dim list1 As New ListItem("הכל", "0")
        sSeries.Items.Add(list1)
        For i As Integer = 0 To dtSeries.Rows.Count - 1
            Dim list As New ListItem(dtSeries.Rows(i)("Series_Name"), dtSeries.Rows(i)("Series_Id"))
            sSeries.Items.Add(list)
        Next


    End Sub
    Sub GetStatus()
        Dim cmdSelect As New SqlClient.SqlCommand("select  status_id,status_Name,status_Color,status_FntColor from Status_ForBizpower ORDER BY status_id", conPegasus)
        conPegasus.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtStatus)
        conPegasus.Close()
        primKeydtStatus(0) = dtStatus.Columns("status_id")
        dtStatus.PrimaryKey = primKeydtStatus

        sStatus.Items.Clear()
        Dim list1 As New ListItem("הכל", "0")
        sStatus.Items.Add(list1)
        For i As Integer = 0 To dtStatus.Rows.Count - 1
            Dim list As New ListItem(dtStatus.Rows(i)("status_Name"), dtStatus.Rows(i)("status_id"))
            list.Attributes.Add("style", "background-color:" & dtStatus.Rows(i)("status_Color") & ";color:" & dtStatus.Rows(i)("status_FntColor"))
            sStatus.Items.Add(list)
        Next


    End Sub
    'Sub GetData()
    '    Dim cmdSelectP = New SqlClient.SqlCommand("get_BizpowerDepartures", conPegasus)
    '    cmdSelectP.CommandType = CommandType.StoredProcedure

    '    conPegasus.Open()
    '    rptData.DataSource = cmdSelectP.ExecuteReader()
    '    rptData.DataBind()
    '    conPegasus.Close()

    'End Sub
    Public Function ViewColumn(ByVal ObjectId As Object) As Boolean

        Dim cmdSelect As New SqlClient.SqlCommand("SELECT Column_Visible FROM ScreenSetting WHERE Column_Id=" & ObjectId, con)
        con.Open()
        Dim tmp = cmdSelect.ExecuteScalar()
        cmdSelect.Dispose()
        con.Close()

        If IsDBNull(tmp) Then
            tmp = False
        Else
            tmp = Trim(tmp)
        End If
        Return tmp
    End Function
    Public Shared Function GetCustomers(ByVal pageIndex As Integer) As String
   
        Return GetCustomersData(pageIndex).GetXml
    End Function
    Public Shared Function GetCustomersData(ByVal pageIndex As Integer) As DataSet
        '  Response.write("pageIndex=" & pageIndex)
        '   Response.end()
        Dim query As String = "[GetCustomersPageWise]"
        Dim cmd As New SqlClient.SqlCommand(query)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add("@PageIndex", pageIndex)
        cmd.Parameters.Add("@PageSize", 100)
        cmd.Parameters.Add("@PageCount", SqlDbType.Int, 4).Direction = ParameterDirection.Output
        Return GetData(cmd)
    End Function
    Private Shared Function GetData(ByVal cmd As SqlClient.SqlCommand) As DataSet
        Dim conD As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
        Dim sda As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter
        cmd.Connection = conD
        sda.SelectCommand = cmd
        Dim ds As DataSet = New DataSet
        sda.Fill(ds, "Customers")
        Dim dt As DataTable = ds.Tables(0)
        dt.Columns.Add("PageCount")
        dt.NewRow()
        dt.Rows(0)(0) = cmd.Parameters("@PageCount").Value

        'Dim dt As DataTable = New DataTable("PageCount")
        'dt.Columns.Add("PageCount")
        'dt.Rows(0)(0) = cmd.Parameters("@PageCount").Value
        'dt.Rows.Add(dt.Rows(0)(0))

        'ds.Tables.Add(dt)


        Return ds
    End Function
End Class
