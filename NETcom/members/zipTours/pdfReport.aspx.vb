Imports System.Data.SqlClient
Public Class pdfReport
    Inherits System.Web.UI.Page
    Protected rptTitle As Repeater
    Protected WithEvents rptCustomers As Repeater
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected totalRows As Integer = 0
    Protected TotalPages As Integer = 0
    Protected CurrentIndex As Integer = 0
    Protected sPayFromDate, sPayToDate As System.Web.UI.HtmlControls.HtmlInputText

    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected sStatus, sSeries, sDepartments, sUsers, sCountries As String
    Public dtStatus, dtSeries, dtsDepartments, dtUsers, dtCountries As New DataTable

    Dim cmdSelect As New SqlClient.SqlCommand
    Dim primKeydtStatus(0), primKeydtSeries(0), primKeydtDepartments(0), primKeydtUsers(0), primKeydtCountries(0) As Data.DataColumn
    Protected WithEvents btnSearchAll As Web.UI.WebControls.LinkButton
    Protected WithEvents pnlPages, pnlSearchMess As System.Web.UI.WebControls.PlaceHolder
     Protected WithEvents btnSaveChecked As Web.UI.WebControls.Button
    Protected FromDate, ToDate, sortQuery, sortQuery1 As String

    Public qrystring As String





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
        sortQuery = ""
        Dim r As Integer
        '''--  sortQuery = " Dep_Name ASC"
        qrystring = Request.ServerVariables("QUERY_STRING")
        r = qrystring.IndexOf("sort")
        If r > 0 Then
            '  Response.Write("<BR>gg=" & Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)) & "<BR>")
            sortQuery1 = Replace(Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)), "&", ",")
            sortQuery1 = Replace(sortQuery1, "=", " ")
            sortQuery1 = Replace(sortQuery1, "sort_10", "CountlastMonth")
            sortQuery1 = Replace(sortQuery1, "sort_11", "User_Name")
            sortQuery1 = Replace(sortQuery1, "sort_12", "Flying_Company")
            sortQuery1 = Replace(sortQuery1, "sort_13", "Price")
            sortQuery1 = Replace(sortQuery1, "sort_14", "CountryName")

            sortQuery1 = Replace(sortQuery1, "sort_1", "Status_Id")
            sortQuery1 = Replace(sortQuery1, "sort_2", "Series_Name")
            sortQuery1 = Replace(sortQuery1, "sort_3", "Departure_Date")
            sortQuery1 = Replace(sortQuery1, "sort_4", "CountMembers")
            sortQuery1 = Replace(sortQuery1, "sort_5", "Dep_Name")
            sortQuery1 = Replace(sortQuery1, "sort_6", "CountMembersBitulim")
            sortQuery1 = Replace(sortQuery1, "sort_7", "LastDateSale")
            sortQuery1 = Replace(sortQuery1, "sort_8", "CountlastWeek")
            sortQuery1 = Replace(sortQuery1, "sort_9", "Countlast2Week")

            sortQuery1 = sortQuery1.Substring(1)

            sortQuery = sortQuery1
            ' Response.Write("<BR>" & sortQuery1 & "<BR>")
            '         Response.End() '   sortQuery = sortQuery1
        End If


        '   Response.Write(qrystring & "<BR>")
        '  Response.End()
        If Not Page.IsPostBack Then
            cmdSelect = New SqlClient.SqlCommand("SELECT Column_Id, Column_Name  FROM ScreenSetting order by Column_Order desc", con)
            cmdSelect.CommandType = CommandType.Text

            con.Open()
            rptTitle.DataSource = cmdSelect.ExecuteReader()
            rptTitle.DataBind()
            con.Close()
            '   GetData()
            'rptCustomers.DataSource = GetCustomersData(1)
            'rptCustomers.DataBind()

       
            bindList()
        End If

    End Sub
   
    
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
        ' Response.Write(Request.Form("sPayFromDate"))
        ' Response.End()
        If IsDate(Request.Form("sPayFromDate")) Then
            FromDate = Request.Form("sPayFromDate")
        Else
            FromDate = ""
        End If
        If IsDate(Request.QueryString("sFromD")) Then
            FromDate = Request.QueryString("sFromD")
        End If

        If IsDate(Request.Form("sPayToDate")) Then
            ToDate = Request.Form("sPayToDate")
        Else
            ToDate = ""
        End If
        If IsDate(Request.QueryString("sToD")) Then
            ToDate = Request.QueryString("sToD")
        End If
        '      Response.Write(sCountries.Value)
        If Request("sStatus") <> "" Then
            sStatus = Request("sStatus")
        Else
            sStatus = 0
        End If

        cmdSelect = New SqlCommand("GetScreenAdminStatic_Report", conPegasus)
        cmdSelect.CommandType = CommandType.StoredProcedure
        'cmdSelect.Parameters.Add("@TourId", SqlDbType.Int).Value = drpTours.SelectedValue
        cmdSelect.Parameters.Add("@statusId", SqlDbType.Int).Value = 0 ' Request("sStatus")
        cmdSelect.Parameters.Add("@seriesId", SqlDbType.Int).Value = 0 ' Request("sSeries")
        cmdSelect.Parameters.Add("@depId", SqlDbType.Int).Value = 0 ' Request("sDepartments")
        cmdSelect.Parameters.Add("@userId", SqlDbType.Int).Value = 0 ' Request("sUsers")
        cmdSelect.Parameters.Add("@fromDate", SqlDbType.VarChar, 30).Value = FromDate
        cmdSelect.Parameters.Add("@toDate", SqlDbType.VarChar, 30).Value = ToDate
        cmdSelect.Parameters.Add("@countryId", SqlDbType.Int).Value = 0 ' Request("sCountries")
        'cmdSelect.Parameters.Add("@srchEmail", SqlDbType.VarChar, 100).Value = sEmail.Value
          'return value - total records count
        cmdSelect.Parameters.Add("@sortQuery", SqlDbType.VarChar, 500).Value = sortQuery
        'Response.Write(sortQuery)
        'Response.End()
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
            rptCustomers.DataSource = dr
            rptCustomers.DataBind()
            'pnlSearchMess.Visible = False
            rptCustomers.Visible = True

            '  totalRows = cmdSelect.Parameters("@CountRecords").Value()
            'Response.Write("totalRows=" & totalRows)


        Else
            '  pnlSearchMess.Visible = True : 
            rptCustomers.Visible = False
        End If

        dr.Close()
        conPegasus.Close()

    End Sub

    Private Sub rptCustomers_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptCustomers.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim chkTour As HtmlInputCheckBox
            chkTour = e.Item.FindControl("chkTour")
        
            If e.Item.DataItem("Checked_ForBizpower") = True Then
                chkTour.Checked = True
                chkTour.Attributes.Add("onclick", "checkDisable('" & e.Item.DataItem("Departure_id") & "');")


            Else
                chkTour.Checked = False
            End If
            chkTour.Value = e.Item.DataItem("Departure_id")

            'Dim slTour As HtmlSelect
            'slTour = e.Item.FindControl("slTour")
            'Dim cmdSelect As New SqlClient.SqlCommand("select  status_id,status_Name,status_Color,status_FntColor from Status_ForBizpower ORDER BY status_id", con)
            'con.Open()
            'Dim ad As New SqlClient.SqlDataAdapter
            'ad.SelectCommand = cmdSelect
            'ad.Fill(slTour)
            'con.Close()
            'primKeydtStatus(0) = slTour.Columns("status_id")
            'slTour.PrimaryKey = primKeydtStatus

            'slTour.Items.Clear()
            'For i As Integer = 0 To slTour.Rows.Count - 1
            '    Dim list As New ListItem(dtStslTouratus.Rows(i)("status_Name"), slTour.Rows(i)("status_id"))
            '    list.Attributes.Add("style", "background-color:" & slTour.Rows(i)("status_Color") & ";color:" & slTour.Rows(i)("status_FntColor"))
            '    slTour.Items.Add(list)
            'Next
        End If
    End Sub


End Class

