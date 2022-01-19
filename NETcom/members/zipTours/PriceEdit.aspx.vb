Imports System.Data.SqlClient
Public Class PriceEdit
    Inherits System.Web.UI.Page
    Protected rptTitle As Repeater
    Protected WithEvents rptCustomers As Repeater
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected totalRows As Integer = 0
    Protected AllCountMembers As Integer = 0
    Protected AllCountMembersBitulim As Integer = 0
    Protected AllCountlastWeek As Integer = 0
    Protected AllCountlast2Week As Integer = 0
    Protected AllCountlastMonth As Integer = 0
    Protected TotalPages As Integer = 0
    Protected CurrentIndex As Integer = 0
    Protected sPayFromDate, sPayToDate As System.Web.UI.HtmlControls.HtmlInputText

    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected WithEvents sStatus, sSeries, sDepartments, sUsers, sCountries As System.Web.UI.HtmlControls.HtmlSelect
    Public dtStatus, dtSeries, dtsDepartments, dtUsers, dtCountries As New DataTable

    Dim cmdSelect As New SqlClient.SqlCommand
    Dim primKeydtStatus(0), primKeydtSeries(0), primKeydtDepartments(0), primKeydtUsers(0), primKeydtCountries(0) As Data.DataColumn
    Protected WithEvents btnSearchAll As Web.UI.WebControls.LinkButton
    Protected WithEvents pnlPages, pnlSearchMess As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents PageSize, pageList As Web.UI.WebControls.DropDownList
    Protected lblCount, lblTotalPages As Web.UI.WebControls.Label
    Protected WithEvents cmdPrev, cmdNext, btnIsPaid, btnSearch As Web.UI.WebControls.LinkButton
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
            'cmdSelect = New SqlClient.SqlCommand("SELECT Column_Id, Column_Name  FROM ScreenSetting order by Column_Order desc", con)
            'cmdSelect.CommandType = CommandType.Text

            'con.Open()
            'RptTitleBottom.DataSource = cmdSelect.ExecuteReader()
            'RptTitleBottom.DataBind()

            con.Close()

          
          
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
        FromDate = ""
        ToDate = ""

       
        '      Response.Write(sCountries.Value)

        cmdSelect = New SqlCommand("GetScreenPriceEdit", conPegasus)
        cmdSelect.CommandType = CommandType.StoredProcedure
         'return value - total records count
        'Response.End()
        cmdSelect.Parameters.Add("@CountRecords", System.Data.SqlDbType.Int).Value = 0
        cmdSelect.Parameters("@CountRecords").Direction = ParameterDirection.Output
        cmdSelect.Parameters.Add("@AllCountMembers", System.Data.SqlDbType.Int).Value = 0
        cmdSelect.Parameters("@AllCountMembers").Direction = ParameterDirection.Output
        cmdSelect.Parameters.Add("@AllCountMembersBitulim", System.Data.SqlDbType.Int).Value = 0
        cmdSelect.Parameters("@AllCountMembersBitulim").Direction = ParameterDirection.Output
        cmdSelect.Parameters.Add("@AllCountlastWeek", System.Data.SqlDbType.Int).Value = 0
        cmdSelect.Parameters("@AllCountlastWeek").Direction = ParameterDirection.Output
        cmdSelect.Parameters.Add("@AllCountlast2Week", System.Data.SqlDbType.Int).Value = 0
        cmdSelect.Parameters("@AllCountlast2Week").Direction = ParameterDirection.Output
        cmdSelect.Parameters.Add("@AllCountlastMonth", System.Data.SqlDbType.Int).Value = 0
        cmdSelect.Parameters("@AllCountlastMonth").Direction = ParameterDirection.Output


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

            totalRows = cmdSelect.Parameters("@CountRecords").Value()
            AllCountMembers = cmdSelect.Parameters("@AllCountMembers").Value()
            AllCountMembersBitulim = cmdSelect.Parameters("@AllCountMembersBitulim").Value()
            AllCountlastWeek = cmdSelect.Parameters("@AllCountlastWeek").Value()
            AllCountlast2Week = cmdSelect.Parameters("@AllCountlast2Week").Value()
            AllCountlastMonth = cmdSelect.Parameters("@AllCountlastMonth").Value()
            'Response.Write("totalRows=" & totalRows)
        

        Else
            '  pnlSearchMess.Visible = True : 
            rptCustomers.Visible = False
        End If

        dr.Close()
        conPegasus.Close()

    End Sub
   


   

  
  
End Class


