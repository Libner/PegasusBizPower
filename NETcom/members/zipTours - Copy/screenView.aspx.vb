Imports System.Data.SqlClient
Public Class screenView
    Inherits System.Web.UI.Page
    Protected rptTitle As Repeater
    Protected WithEvents rptCustomers As Repeater
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected totalRows As Integer = 0
    Protected TotalPages As Integer = 0
    Protected CurrentIndex As Integer = 0

    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected WithEvents sStatus, sSeries, sDepartments, sUsers As System.Web.UI.HtmlControls.HtmlSelect
    Public dtStatus, dtSeries, dtsDepartments, dtUsers As New DataTable
    Protected sPayFromDate, sPayToDate As System.Web.UI.HtmlControls.HtmlInputText

    Dim cmdSelect As New SqlClient.SqlCommand
    Dim primKeydtStatus(0), primKeydtSeries(0), primKeydtDepartments(0), primKeydtUsers(0) As Data.DataColumn
    Protected WithEvents btnSearchAll As Web.UI.WebControls.LinkButton
    Protected WithEvents pnlPages, pnlSearchMess As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents PageSize, pageList As Web.UI.WebControls.DropDownList
    Protected lblCount, lblTotalPages As Web.UI.WebControls.Label
    Protected WithEvents cmdPrev, cmdNext, btnIsPaid, btnSearch As Web.UI.WebControls.LinkButton
    Protected WithEvents btnSaveChecked As Web.UI.WebControls.Button
    Protected FromDate, ToDate As String





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
            'rptCustomers.DataSource = GetCustomersData(1)
            'rptCustomers.DataBind()

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
            bindList()
        End If

    End Sub
    Sub btnSearch_onClick(ByVal s As Object, ByVal e As EventArgs) Handles btnSearch.Click
        pageList.SelectedIndex = 0
        If pageList.Items.Count > 1 Then
            CurrentIndex = CInt(pageList.SelectedIndex)
        Else
            CurrentIndex = 0
        End If

        bindList()
    End Sub
    Sub btnSearchAll_onClick(ByVal s As Object, ByVal e As EventArgs) Handles btnSearchAll.Click
        pageList.SelectedIndex = 0
        CurrentIndex = CInt(pageList.SelectedIndex)
        sDepartments.Value = 0
        sUsers.Value = 0
        sSeries.Value = 0
        sStatus.Value = 0
        sPayFromDate.value = ""
        sPayToDate.Value = ""
        FromDate = ""
        ToDate = ""
        bindList()
    End Sub
    Sub GetUsers()
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT USER_ID, FIRSTNAME + ' ' + LASTNAME as Username from Users  where [ACTIVE]=1 order by FIRSTNAME asc,LASTNAME asc", con)
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
    Public Sub bindList()
        If IsDate(Request.Form("sPayFromDate")) Then
            FromDate = Request.Form("sPayFromDate")
        Else
            FromDate = ""
        End If
        If IsDate(Request.Form("sPayToDate")) Then
            ToDate = Request.Form("sPayToDate")
        Else
            ToDate = ""
        End If

        cmdSelect = New SqlCommand("GetScreenView", conPegasus)
        cmdSelect.CommandType = CommandType.StoredProcedure
        cmdSelect.Parameters.Add("@statusId", SqlDbType.Int).Value = sStatus.Value
        cmdSelect.Parameters.Add("@seriesId", SqlDbType.Int).Value = sSeries.Value
        cmdSelect.Parameters.Add("@depId", SqlDbType.Int).Value = sDepartments.Value
        cmdSelect.Parameters.Add("@userId", SqlDbType.Int).Value = sUsers.Value
        cmdSelect.Parameters.Add("@fromDate", SqlDbType.VarChar, 30).Value = FromDate
        cmdSelect.Parameters.Add("@toDate", SqlDbType.VarChar, 30).Value = ToDate

        'cmdSelect.Parameters.Add("@TourId", SqlDbType.Int).Value = drpTours.SelectedValue
        'cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = drpDepartures.SelectedValue
        'cmdSelect.Parameters.Add("@srchName", SqlDbType.VarChar, 100).Value = sName.Value
        'cmdSelect.Parameters.Add("@srchEmail", SqlDbType.VarChar, 100).Value = sEmail.Value
        cmdSelect.Parameters.Add("@PageSize", DbType.Int32).Value = Trim(PageSize.SelectedValue)
        cmdSelect.Parameters.Add("@PageNumber", DbType.Int32).Value = CInt(CurrentIndex) + 1
        'return value - total records count
        cmdSelect.Parameters.Add("@CountRecords", System.Data.SqlDbType.Int).Value = 0
        cmdSelect.Parameters("@CountRecords").Direction = ParameterDirection.Output
        conPegasus.Open()
        Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        'Response.Write("<br>PageSize.SelectedValue=" & PageSize.SelectedValue)
        'Response.Write("<br>@TourId=" & cmdSelect.Parameters("@TourId").Value)
        'Response.Write("<br>@DepartureId=" & cmdSelect.Parameters("@DepartureId").Value)
        'Response.Write("<br>@srchName=" & cmdSelect.Parameters("@srchName").Value)
        'Response.Write("<br>@srchEmail=" & cmdSelect.Parameters("@srchEmail").Value)
        'Response.Write("<br>@PageSize=" & cmdSelect.Parameters("@PageSize").Value)
        'Response.Write("<br>@PageNumber=" & cmdSelect.Parameters("@PageNumber").Value)
        'Response.Write("<br>@CountRecords=" & cmdSelect.Parameters("@CountRecords").Value)
        'Response.Write("<br>CurrentIndex=" & (CInt(CurrentIndex) + 1))
        'Response.Write("<br>PageSize=" & PageSize.SelectedValue)

        If dr.HasRows Then

            rptCustomers.DataSource = dr
            rptCustomers.DataBind()
            'pnlSearchMess.Visible = False
            rptCustomers.Visible = True

            totalRows = cmdSelect.Parameters("@CountRecords").Value()
            'Response.Write("totalRows=" & totalRows)
            If totalRows > PageSize.SelectedValue Then
                pnlPages.Visible = True
                lblCount.Text = totalRows

                TotalPages = Math.Ceiling(totalRows / PageSize.SelectedValue)
                BuildNumericPages()

            Else

                pnlPages.Visible = True
                lblCount.Text = totalRows

                TotalPages = Math.Ceiling(totalRows / PageSize.SelectedValue)
                BuildNumericPages()
            End If

        Else
            '  pnlSearchMess.Visible = True : 
            rptCustomers.Visible = False : pnlPages.Visible = False
        End If

        dr.Close()
        conPegasus.Close()

    End Sub
    Public Sub PageList_Click(ByVal s As System.Object, ByVal e As System.EventArgs) Handles pageList.SelectedIndexChanged
        CurrentIndex = CInt(pageList.SelectedItem.Value)
        bindList()
    End Sub

    Sub cmdPrev_onClick(ByVal s As Object, ByVal e As EventArgs) Handles cmdPrev.Click
        pageList.SelectedIndex = pageList.SelectedIndex - 1
        CurrentIndex = CInt(pageList.SelectedItem.Value)
        bindList()
    End Sub

    Sub cmdNext_onClick(ByVal s As Object, ByVal e As EventArgs) Handles cmdNext.Click
        pageList.SelectedIndex = pageList.SelectedIndex + 1
        CurrentIndex = CInt(pageList.SelectedIndex)
        bindList()
    End Sub

    Public Sub PageSize_Click(ByVal s As System.Object, ByVal e As System.EventArgs) Handles PageSize.SelectedIndexChanged
        CurrentIndex = 0
        bindList()
    End Sub
    Sub BuildNumericPages()
        Dim i As Integer
        Dim item As ListItem

        pageList.Items.Clear()
        For i = 1 To TotalPages
            item = New ListItem(i.ToString(), (i - 1).ToString())
            pageList.Items.Add(item)
        Next
        If (TotalPages <= 0) Then
            pageList.Enabled = False
            pageList.SelectedIndex = 0
            pageList.Visible = False
        Else 'Populate the list      
            pageList.Enabled = True
            pageList.Visible = True
            pageList.SelectedIndex = CurrentIndex
            lblTotalPages.Text = TotalPages
        End If
        If CurrentIndex = 0 Then
            cmdPrev.Enabled = False
        Else
            cmdPrev.Enabled = True
        End If
        If CurrentIndex = TotalPages - 1 Then
            cmdNext.Enabled = False
        Else
            cmdNext.Enabled = True
        End If

    End Sub



End Class

