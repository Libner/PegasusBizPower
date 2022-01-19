Imports System.Data.SqlClient
Public Class WorkScreen
    Inherits System.Web.UI.Page
    Protected rptData As Repeater
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected totalRows As Integer = 0
    Protected TotalPages As Integer = 0
    Protected CurrentIndex As Integer = 0
    Protected WithEvents pnlPages, pnlSearchMess As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents PageSize, pageList As Web.UI.WebControls.DropDownList
    Protected lblCount, lblTotalPages As Web.UI.WebControls.Label
    Protected WithEvents cmdPrev, cmdNext As Web.UI.WebControls.LinkButton


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
        If Not Page.IsPostBack Then
            bindList()
        End If


    End Sub

    Public Sub bindList()
        ' Response.Write(Request.Form("sPayFromDate"))
        ' Response.End()
        'If IsDate(Request.Form("sPayFromDate")) Then
        '    FromDate = Request.Form("sPayFromDate")
        'Else
        '    FromDate = ""
        'End If
        'If IsDate(Request.QueryString("sFromD")) Then
        '    FromDate = Request.QueryString("sFromD")
        'End If

        'If IsDate(Request.Form("sPayToDate")) Then
        '    ToDate = Request.Form("sPayToDate")
        'Else
        '    ToDate = ""
        'End If
        'If IsDate(Request.QueryString("sToD")) Then
        '    ToDate = Request.QueryString("sToD")
        'End If
        '      Response.Write(sCountries.Value)

        ''--------- cmdSelect = New SqlCommand("GetScreenAdminStatic", conPegasus)
        cmdSelect = New SqlCommand("GetWorkScreen", conPegasus)
        cmdSelect.CommandType = CommandType.StoredProcedure
        'cmdSelect.Parameters.Add("@TourId", SqlDbType.Int).Value = drpTours.SelectedValue
        'cmdSelect.Parameters.Add("@statusId", SqlDbType.Int).Value = sStatus.Value
        'cmdSelect.Parameters.Add("@seriesId", SqlDbType.Int).Value = sSeries.Value
        'cmdSelect.Parameters.Add("@depId", SqlDbType.Int).Value = sDepartments.Value
        'cmdSelect.Parameters.Add("@userId", SqlDbType.Int).Value = sUsers.Value
        'cmdSelect.Parameters.Add("@fromDate", SqlDbType.VarChar, 30).Value = FromDate
        'cmdSelect.Parameters.Add("@toDate", SqlDbType.VarChar, 30).Value = ToDate
        'cmdSelect.Parameters.Add("@countryId", SqlDbType.Int).Value = sCountries.Value
        'cmdSelect.Parameters.Add("@srchEmail", SqlDbType.VarChar, 100).Value = sEmail.Value
        cmdSelect.Parameters.Add("@PageSize", DbType.Int32).Value = Trim(PageSize.SelectedValue)
        cmdSelect.Parameters.Add("@PageNumber", DbType.Int32).Value = CInt(CurrentIndex) + 1
        'return value - total records count
        'cmdSelect.Parameters.Add("@sortQuery", SqlDbType.VarChar, 500).Value = sortQuery
        'Response.Write(sortQuery)
        'Response.End()
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
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rptData.DataSource = dr
            rptData.DataBind()
            'pnlSearchMess.Visible = False
            rptData.Visible = True

            totalRows = cmdSelect.Parameters("@CountRecords").Value()
            'Response.Write("totalRows=" & totalRows)
            If totalRows > PageSize.SelectedValue Then
                pnlPages.Visible = True
                lblCount.Text = totalRows

                TotalPages = Math.Ceiling(totalRows / PageSize.SelectedValue)
                '       Response.Write(TotalPages)
                '      Response.End()
                BuildNumericPages()

            Else

                pnlPages.Visible = True
                lblCount.Text = totalRows

                TotalPages = Math.Ceiling(totalRows / PageSize.SelectedValue)
                BuildNumericPages()
            End If

        Else
            '  pnlSearchMess.Visible = True : 
            rptData.Visible = False : pnlPages.Visible = False
        End If

        dr.Close()
        conPegasus.Close()

    End Sub
    Public Sub PageList_Click(ByVal s As System.Object, ByVal e As System.EventArgs) Handles pageList.SelectedIndexChanged
        CurrentIndex = CInt(pageList.SelectedItem.Value)
        '  Response.Write("CurrentIndex=" & CurrentIndex)
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
            If Request.QueryString("sPage") > 0 And Request.QueryString("sPage") = item.Value Then
                '   item.Selected = True
                CurrentIndex = Request.QueryString("sPage")
            End If
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
