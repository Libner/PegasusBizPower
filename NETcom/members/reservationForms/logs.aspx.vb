Imports System.Data.SqlClient
Public Class logsRegistrationSendings
    Inherits System.Web.UI.Page
    Protected WithEvents rptTitle As Repeater
    Protected WithEvents rptLogs As Repeater
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected totalRows As Integer = 0
    Protected TotalPages As Integer = 0
    Protected CurrentIndex As Integer = 0
    Protected sReservationFromDate, sReservationToDate, sCountSendings, sSendingFromDate, sSendingToDate, sContactName As System.Web.UI.HtmlControls.HtmlInputText

    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected WithEvents sDepartments, sCreateFormUsers, sDepartures, sSeries, sCountries, sGuides, sSendingFormUsers As System.Web.UI.HtmlControls.HtmlSelect
    Public dtDepartures, dtDepartments, dtSeries, dtCountries, dtGuides, dtUsers As New DataTable

    Dim cmdSelect As New SqlClient.SqlCommand
    Dim primKeydtDepartures(0), primKeydtDepartments(0), primKeydtSeries(0), primKeydtCountries(0), primKeydtGuides(0), primKeydtUsers(0) As Data.DataColumn
    Protected WithEvents btnSearchAll As Web.UI.WebControls.LinkButton
    Protected WithEvents pnlPages, pnlSearchMess As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents PageSize, pageList As Web.UI.WebControls.DropDownList
    Protected lblCount, lblTotalPages As Web.UI.WebControls.Label
    Protected WithEvents cmdPrev, cmdNext, btnIsPaid, btnSearch, btnSearchRight As Web.UI.WebControls.LinkButton
    Protected WithEvents btnSaveChecked As Web.UI.WebControls.Button
    Protected FromInsertDate, ToInsertDate, FromSendingDate, ToSendingDate As String
    Protected sortQuery, sortQuery1 As String

    Public qrystring As String
    Public dtScreenColumns As New DataTable
    Public UpdateTimingPermitted As Boolean
    Public AddAppealPermitted As Boolean

    'Dim arrColumnsNames() As String = {"ContactName", "Phone", "Email", "Departures", "Countries", "Series", "LastEnterDate", "CountEnters", "ExistsAppleal16735", "ExistsAppleal16504"}
    'Dim arrColumnsDesc() As String = {"שם הלקוח", "טלפון ", "אימייל", "שם הקטגוריה", "שם  היעד", "שם הסדרה", "תאריך אחרון", "מספר כניסות", "טייל ביעד?", "נוצר טופס מתעניין?"}
    Public dtTest As DataGrid

    Public userId As String
    Public queryParamSearch, queryParamSort As String
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
        'permitions to new appeal (mitan'en) creating 
        AddAppealPermitted = True '???for all users which are permitted to see this page

        userId = Trim(Request.Cookies("bizpegasus")("UserId"))


        setTitles()


        If Not Page.IsPostBack Then


            rptTitle.DataSource = dtScreenColumns
            rptTitle.DataBind()

            GetCountries()

            bindList()
        End If


        If Request.QueryString("sPage") = "" Then
            queryParamSearch = queryParamSearch & "&sPage=" & (CurrentIndex + 1).ToString
        Else
            Dim spage As String
            spage = Request.QueryString("sPage")
            queryParamSearch = queryParamSearch.Replace("&sPage=" & spage, "&sPage=" & (CurrentIndex + 1).ToString)
        End If

        If queryParamSearch.IndexOf("sort_") = -1 Then
            queryParamSearch = queryParamSearch & queryParamSort
        End If
        If queryParamSearch.StartsWith("&") Then
            queryParamSearch = Right(queryParamSearch, Len(queryParamSearch) - 1)
        End If
        queryParamSearch = queryParamSearch.Replace("&", "~")
        queryParamSearch = Server.UrlEncode(queryParamSearch)
    End Sub
    Sub setTitles()
        If Not IsNothing(viewstate("ScreenColumns")) Then
            dtScreenColumns = viewstate("ScreenColumns")
        End If
        If dtScreenColumns.Rows.Count = 0 Then

            cmdSelect = New SqlClient.SqlCommand("SELECT Column_Id, Column_Name,Column_Title,Column_DBName  FROM Screen_ReservationFormSendingLogs_Setting order by Column_Order", con)
            cmdSelect.CommandType = CommandType.Text
            con.Open()
            Dim ad As New SqlClient.SqlDataAdapter
            ad.SelectCommand = cmdSelect
            ad.Fill(dtScreenColumns)
            con.Close()
            viewstate("ScreenColumns") = dtScreenColumns
        End If

    End Sub

    Private Sub rptTitle_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptTitle.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim phSortArrows As PlaceHolder
            phSortArrows = e.Item.FindControl("phSortArrows")
            If e.Item.DataItem("Column_Name") = "Series" Then
                phSortArrows.Visible = True
            End If
        End If
    End Sub

    Sub setSortFields()
        sortQuery = ""
        Dim r As Integer
        '''--  sortQuery = " Dep_Name ASC"
        qrystring = Request.ServerVariables("QUERY_STRING")
        r = qrystring.IndexOf("sort")
        If r > 0 Then
            'Response.Write("<BR>gg=" & Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)) & "<BR>")
            sortQuery1 = Replace(Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)), "&", ",")
            sortQuery1 = Replace(sortQuery1, "=", " ")
            ' Response.Write("<BR>" & dtScreenColumns.Rows.Count & "<BR>")
            For i As Integer = 0 To dtScreenColumns.Rows.Count - 1
                sortQuery1 = Replace(sortQuery1, "sort_" & dtScreenColumns.Rows(i)("Column_Id"), dtScreenColumns.Rows(i)("Column_DBName"))
            Next

            sortQuery1 = sortQuery1.Substring(1)

            sortQuery = sortQuery1


            If Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)) <> "" Then
                queryParamSort = Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring))
            End If
            'Response.Write("<BR>" & queryParamSearch & "<BR>")
            '         Response.End() '   sortQuery = sortQuery1
        End If
    End Sub
    Sub btnSearch_onClick(ByVal s As Object, ByVal e As EventArgs) Handles btnSearch.Click, btnSearchRight.Click
        pageList.SelectedIndex = 0
        If pageList.Items.Count > 1 Then
            CurrentIndex = CInt(pageList.SelectedIndex)
        Else
            CurrentIndex = 0
        End If

        bindList()
    End Sub
    Sub GetCountries()
        'Dim cmdSelect As New SqlClient.SqlCommand("SELECT Country_Id, Country_Name from Countries  order by Country_Name", conPegasus)
        'SELECT FROM BIZPOWER YAADIM
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT Country_Id, Country_Name from Countries  order by Country_Name", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtCountries)
        con.Close()
        primKeydtCountries(0) = dtCountries.Columns("Country_ID")
        dtCountries.PrimaryKey = primKeydtCountries
        sCountries.Items.Clear()
        Dim list1 As New ListItem("", "")
         sCountries.Items.Add(list1)
        sCountries.Items.Add(New ListItem("כל היעדים", "כל היעדים"))
        For i As Integer = 0 To dtCountries.Rows.Count - 1
            Dim list As New ListItem(dtCountries.Rows(i)("Country_Name"), dtCountries.Rows(i)("Country_Id"))
            If Request.QueryString("sCountries") > 0 And Request.QueryString("sCountries") = dtCountries.Rows(i)("Country_Id") Then
                list.Selected = True

            End If
            sCountries.Items.Add(list)
        Next
    End Sub



    Sub btnSearchAll_onClick(ByVal s As Object, ByVal e As EventArgs) Handles btnSearchAll.Click
        pageList.SelectedIndex = 0
        CurrentIndex = CInt(pageList.SelectedIndex)
        sCountries.Value = 0
        sCountSendings.Value = ""
        sSendingFromDate.Value = ""
        sSendingToDate.Value = ""
        FromSendingDate = ""
        ToSendingDate = ""
        sortQuery = ""

        qrystring = ""
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
    Sub setSearchFields()




        If func.dbNullFix(Request.Form("sCountries")) <> "" Then
            Dim arrCountries() As String
            arrCountries = Request.Form("sCountries").Split(",")

            For i As Integer = 0 To arrCountries.Length - 1
                sCountries.Items.FindByValue(arrCountries(i)).Selected = True
            Next
        ElseIf func.dbNullFix(Request.QueryString("sCountries")) <> "" Then
            Dim arrCountries() As String
            arrCountries = Request.Form("sCountries").Split(",")
            For i As Integer = 0 To arrCountries.Length - 1
                sCountries.Items.FindByValue(arrCountries(i)).Selected = True
            Next
        Else
            sCountries.Items(0).Selected = True
        End If
        If sCountries.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sCountries=" & sCountries.Value
        End If


        If IsDate(Request.Form("sSendingFromDate")) Then
            FromSendingDate = Request.Form("sSendingFromDate")
        ElseIf IsDate(Request.QueryString("sSendingFromDate")) Then
            FromSendingDate = Request.QueryString("sSendingFromDate")
        Else
            FromSendingDate = ""
        End If
        sSendingFromDate.Value = FromSendingDate
        If sSendingFromDate.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sSendingFromDate=" & sSendingFromDate.Value
        End If

        If IsDate(Request.Form("sSendingToDate")) Then
            ToSendingDate = Request.Form("sSendingToDate")
        ElseIf IsDate(Request.QueryString("sSendingToDate")) Then
            ToSendingDate = Request.QueryString("sSendingToDate")
        Else
            ToSendingDate = ""
        End If
        sSendingToDate.Value = ToSendingDate
        If sSendingToDate.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sSendingToDate=" & sSendingToDate.Value
        End If

    End Sub
    Public Sub bindList()
        ' Response.Write(Request.Form("sPayFromDate"))
        ' Response.End()


        setSearchFields()

        setSortFields()

        cmdSelect = New SqlCommand("get_LogReservationFormsSendings", con)
        cmdSelect.CommandType = CommandType.StoredProcedure
        'cmdSelect.Parameters.Add("@TourId", SqlDbType.Int).Value = drpTours.SelectedValue

        cmdSelect.Parameters.Add("@siteDBName", SqlDbType.VarChar, 20).Value = ConfigurationSettings.AppSettings("PegasusSiteDBName")
        cmdSelect.Parameters.Add("@countSendings", SqlDbType.VarChar, 20).Value = sCountSendings.Value
        cmdSelect.Parameters.Add("@fromSendingDate", SqlDbType.VarChar, 30).Value = FromSendingDate
        cmdSelect.Parameters.Add("@toSendingDate", SqlDbType.VarChar, 30).Value = ToSendingDate
        cmdSelect.Parameters.Add("@countries", SqlDbType.VarChar).Value = func.dbNullFix(Request("sCountries"))

        cmdSelect.Parameters.Add("@PageSize", DbType.Int32).Value = Trim(PageSize.SelectedValue)
        cmdSelect.Parameters.Add("@PageNumber", DbType.Int32).Value = CInt(CurrentIndex) + 1
        'return value - total records count
        cmdSelect.Parameters.Add("@sortQuery", SqlDbType.VarChar, 500).Value = sortQuery
        'Response.Write(sortQuery)
        'Response.End()
        cmdSelect.Parameters.Add("@CountRecords", System.Data.SqlDbType.Int).Value = 0
        cmdSelect.Parameters("@CountRecords").Direction = ParameterDirection.Output

        con.Open()
        'For Each p As SqlClient.SqlParameter In cmdSelect.Parameters
        '    Response.Write("<br>" & p.ParameterName & "=" & p.Value)
        'Next
        'Response.End()
        Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rptLogs.DataSource = dr
            rptLogs.DataBind()
            'pnlSearchMess.Visible = False
            rptLogs.Visible = True

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
            rptLogs.Visible = False : pnlPages.Visible = False
        End If

        dr.Close()
        conPegasus.Close()

    End Sub
    Public Sub PageList_Click(ByVal s As System.Object, ByVal e As System.EventArgs) Handles pageList.SelectedIndexChanged
        CurrentIndex = CInt(pageList.SelectedItem.Value)
        bindList()

        Dim cScript = "<script language='javascript'>window.parent.scrollTo(0, 0);</script>"
        RegisterStartupScript("", cScript)
    End Sub

    Sub cmdPrev_onClick(ByVal s As Object, ByVal e As EventArgs) Handles cmdPrev.Click
        pageList.SelectedIndex = pageList.SelectedIndex - 1
        CurrentIndex = CInt(pageList.SelectedItem.Value)
        bindList()

        Dim cScript = "<script language='javascript'>window.parent.scrollTo(0, 0);</script>"
        RegisterStartupScript("", cScript)
    End Sub

    Sub cmdNext_onClick(ByVal s As Object, ByVal e As EventArgs) Handles cmdNext.Click

        pageList.SelectedIndex = pageList.SelectedIndex + 1
        CurrentIndex = CInt(pageList.SelectedIndex)

        bindList()

        Dim cScript = "<script language='javascript'>window.parent.scrollTo(0, 0);</script>"
        RegisterStartupScript("", cScript)



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


    Private Sub rptLogs_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptLogs.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then


        End If
    End Sub



End Class

