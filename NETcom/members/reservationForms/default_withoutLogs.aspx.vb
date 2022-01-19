Imports System.Data.SqlClient
Public Class adminRegistrationScreen
    Inherits System.Web.UI.Page
    Protected WithEvents rptTitle As Repeater
    Protected WithEvents rptLogs As Repeater
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected totalRows As Integer = 0
    Protected TotalPages As Integer = 0
    Protected CurrentIndex As Integer = 0
    Protected sReservationFromDate, sReservationToDate, sCountSendings, sLastSendingFromDate, sLastSendingToDate, sContactName As System.Web.UI.HtmlControls.HtmlInputText

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

        If Request.Cookies("bizpegasus")("UpdateTimingMailReservationForm") = "1" Then
            UpdateTimingPermitted = True
        End If
        'Response.Write(AddAppealPermitted)

        '  Response.End()
        setTitles()

      
        If Not Page.IsPostBack Then


            rptTitle.DataSource = dtScreenColumns
            rptTitle.DataBind()

            ''''rptTitle.DataSource = arrColumnsNames
            ''''rptTitle.DataBind()

            GetSeries()
            GetDepartures()
            GetGuides()
            GetDepartments()
            GetUsers()
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

            cmdSelect = New SqlClient.SqlCommand("SELECT Column_Id, Column_Name,Column_Title,Column_DBName  FROM Screen_ReservationFormSending_Setting order by Column_Order", con)
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
        Dim list1 As New ListItem("", "0")
        sCountries.Items.Add(list1)
        For i As Integer = 0 To dtCountries.Rows.Count - 1
            Dim list As New ListItem(dtCountries.Rows(i)("Country_Name"), dtCountries.Rows(i)("Country_Id"))
            If Request.QueryString("sCountries") > 0 And Request.QueryString("sCountries") = dtCountries.Rows(i)("Country_Id") Then
                list.Selected = True

            End If
            sCountries.Items.Add(list)
        Next
    End Sub

    Sub GetSeries()
        Dim cmdSelect As New SqlClient.SqlCommand("select  Series_Id,Series_Name from Series ORDER BY Series_Name", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtSeries)
        con.Close()
        primKeydtSeries(0) = dtSeries.Columns("Series_Id")
        dtSeries.PrimaryKey = primKeydtSeries
        sSeries.Items.Clear()
        Dim list1 As New ListItem("", "0")
        sSeries.Items.Add(list1)
        ' Dim list1 As New ListItem("הכל", "0")
        'sSeries.Items.Add(list1)
        For i As Integer = 0 To dtSeries.Rows.Count - 1
            Dim list As New ListItem(dtSeries.Rows(i)("Series_Name"), dtSeries.Rows(i)("Series_Id"))
            If Request.QueryString("sSer") > 0 And Request.QueryString("sSer") = dtSeries.Rows(i)("Series_Id") Then
                list.Selected = True
            End If
            sSeries.Items.Add(list)
        Next

    End Sub

    Sub GetDepartures()
        Dim strsql As String
        strsql = "SELECT distinct tt.* from (SELECT  Tour_Reservations.Departure_Id, (IsNULL(Departure_Code, '') + ' ' + IsNULL(Date_Begin, '') " & _
        " + ' - ' + IsNULL(Date_End, '')) as Departure_Name,Departure_Code,Departure_Date  FROM dbo.Tours_Departures  INNER JOIN " & _
         "     Tour_Reservations ON Tours_Departures.Departure_Id = Tour_Reservations.Departure_Id" & _
         ") as tt  ORDER BY Departure_Code,Departure_Date desc"

        Dim cmdSelect As New SqlClient.SqlCommand(strsql, conPegasus)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtDepartures)
        con.Close()
        primKeydtDepartures(0) = dtDepartures.Columns("Departure_id")
        dtDepartures.PrimaryKey = primKeydtDepartures
        sDepartures.Items.Clear()
        Dim list1 As New ListItem("", "0")
        sDepartures.Items.Add(list1)
        ' Dim list1 As New ListItem("הכל", "0")
        'sSeries.Items.Add(list1)
        For i As Integer = 0 To dtDepartures.Rows.Count - 1
            Dim list As New ListItem(dtDepartures.Rows(i)("Departure_Name"), dtDepartures.Rows(i)("Departure_Id"))
            If Request.QueryString("sDepartures") > 0 And Request.QueryString("sDepartures") = dtDepartures.Rows(i)("Departure_Id") Then
                list.Selected = True
            End If
            sDepartures.Items.Add(list)
        Next

    End Sub

    Sub GetGuides()
        Dim strsql As String
        strsql = "SELECT        Guide_Id, ISNULL(Guide_LName, '') + ' ' + ISNULL(Guide_FName, '') AS Guide_Name FROM  Guides" & _
         "  ORDER BY Guide_LName,Guide_FName desc"

        Dim cmdSelect As New SqlClient.SqlCommand(strsql, conPegasus)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtGuides)
        con.Close()
        primKeydtGuides(0) = dtGuides.Columns("Guide_id")
        dtGuides.PrimaryKey = primKeydtGuides
        sGuides.Items.Clear()
        Dim list1 As New ListItem("", "0")
        sGuides.Items.Add(list1)
        ' Dim list1 As New ListItem("הכל", "0")
        'sSeries.Items.Add(list1)
        For i As Integer = 0 To dtGuides.Rows.Count - 1
            Dim list As New ListItem(dtGuides.Rows(i)("Guide_Name"), dtGuides.Rows(i)("Guide_Id"))
            If Request.QueryString("sGuides") > 0 And Request.QueryString("sGuides") = dtGuides.Rows(i)("Guide_Id") Then
                list.Selected = True
            End If
            sGuides.Items.Add(list)
        Next

    End Sub

    Sub GetDepartments()
        Dim strsql As String
        strsql = "SELECT   departmentId, departmentName FROM Departments" & _
         "  ORDER BY PriorityLevel"

        Dim cmdSelect As New SqlClient.SqlCommand(strsql, con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtDepartments)
        con.Close()
        primKeydtDepartments(0) = dtDepartments.Columns("departmentId")
        dtDepartments.PrimaryKey = primKeydtDepartments
        sDepartments.Items.Clear()
        Dim list1 As New ListItem("", "0")
        sDepartments.Items.Add(list1)
        ' Dim list1 As New ListItem("הכל", "0")
        'sSeries.Items.Add(list1)
        For i As Integer = 0 To dtDepartments.Rows.Count - 1
            Dim list As New ListItem(dtDepartments.Rows(i)("departmentName"), dtDepartments.Rows(i)("departmentId"))
            If Request.QueryString("sDepartments") > 0 And Request.QueryString("sDepartments") = dtDepartments.Rows(i)("departmentId") Then
                list.Selected = True
            End If
            sDepartments.Items.Add(list)
        Next

    End Sub
    Sub GetUsers()
        Dim strsql As String
        strsql = "SELECT USER_ID,ISNULL(LASTNAME, '') + ' ' + ISNULL(FIRSTNAME, '') AS USER_Name FROM   USERS" & _
         "  ORDER BY LASTNAME,FIRSTNAME"

        Dim cmdSelect As New SqlClient.SqlCommand(strsql, con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtUsers)
        con.Close()
        primKeydtUsers(0) = dtUsers.Columns("USER_ID")
        dtUsers.PrimaryKey = primKeydtUsers

        sCreateFormUsers.Items.Clear()
        Dim list1 As New ListItem("", "0")
        sCreateFormUsers.Items.Add(list1)

        sSendingFormUsers.Items.Clear()
        Dim list2 As New ListItem("", "0")
        sSendingFormUsers.Items.Add(list2)

        For i As Integer = 0 To dtUsers.Rows.Count - 1
            Dim lists1 As New ListItem(dtUsers.Rows(i)("USER_Name"), dtUsers.Rows(i)("USER_ID"))
            If Request.QueryString("sCreateFormUsers") > 0 And Request.QueryString("sCreateFormUsers") = dtUsers.Rows(i)("USER_ID") Then
                lists1.Selected = True
            End If
            sCreateFormUsers.Items.Add(lists1)

            Dim lists2 As New ListItem(dtUsers.Rows(i)("USER_Name"), dtUsers.Rows(i)("USER_ID"))
            If Request.QueryString("sSendingFormUsers") > 0 And Request.QueryString("sSendingFormUsers") = dtUsers.Rows(i)("USER_ID") Then
                lists2.Selected = True
            End If
            sSendingFormUsers.Items.Add(lists2)
        Next

    End Sub

    Sub btnSearchAll_onClick(ByVal s As Object, ByVal e As EventArgs) Handles btnSearchAll.Click
        pageList.SelectedIndex = 0
        CurrentIndex = CInt(pageList.SelectedIndex)
        sDepartures.Value = 0
        sContactName.Value = ""
        sSeries.Value = 0
        sCountries.Value = 0
        sGuides.Value = 0
        sDepartments.Value = 0
        sCountries.Value = 0
        sCreateFormUsers.Value = ""
        sReservationFromDate.Value = ""
        sReservationToDate.Value = ""
        sCountSendings.Value = ""
        sLastSendingFromDate.Value = ""
        sLastSendingToDate.Value = ""
        sSendingFormUsers.Value = ""
        FromInsertDate = ""
        ToInsertDate = ""
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



        If func.dbNullFix(Request.Form("sContactName")) <> "" Then
            sContactName.Value = Request.Form("sContactName")
        ElseIf func.dbNullFix(Request.QueryString("sContactName")) <> "" Then
            sContactName.Value = Request.QueryString("sContactName")
        Else
            sContactName.Value = ""
        End If
        If sContactName.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sContactName=" & sContactName.Value
        End If

        If func.dbNullFix(Request.Form("sSeries")) <> "" Then
            sSeries.Items.FindByValue(Request.Form("sSeries")).Selected = True
        ElseIf func.dbNullFix(Request.QueryString("sSeries")) <> "" Then
            sSeries.Items.FindByValue(Request.QueryString("sSeries")).Selected = True
        Else
            sSeries.Items(0).Selected = True
        End If
        If sSeries.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sSeries=" & sSeries.Items(sSeries.SelectedIndex).Value
        End If

        If func.dbNullFix(Request.Form("sGuides")) <> "" Then
            sGuides.Items.FindByValue(Request.Form("sGuides")).Selected = True
        ElseIf func.dbNullFix(Request.QueryString("sGuides")) <> "" Then
            sGuides.Items.FindByValue(Request.QueryString("sGuides")).Selected = True
        Else
            sGuides.Items(0).Selected = True
        End If
        If sGuides.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sGuides=" & sGuides.Items(sGuides.SelectedIndex).Value
        End If

        If func.dbNullFix(Request.Form("sDepartures")) <> "" Then
            sDepartures.Items.FindByValue(Request.Form("sDepartures")).Selected = True
        ElseIf func.dbNullFix(Request.QueryString("sDepartures")) <> "" Then
            sDepartures.Items.FindByValue(Request.QueryString("sDepartures")).Selected = True
        Else
            sDepartures.Items(0).Selected = True
        End If
        If sDepartures.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sDepartures=" & sDepartures.Items(sDepartures.SelectedIndex).Value
        End If

        If func.dbNullFix(Request.Form("sCreateFormUsers")) <> "" Then
            sCreateFormUsers.Items.FindByValue(Request.Form("sCreateFormUsers")).Selected = True
        ElseIf func.dbNullFix(Request.QueryString("sCreateFormUsers")) <> "" Then
            sCreateFormUsers.Items.FindByValue(Request.QueryString("sCreateFormUsers")).Selected = True
        Else
            sCreateFormUsers.Items(0).Selected = True
        End If
        If sCreateFormUsers.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sCreateFormUsers=" & sCreateFormUsers.Items(sCreateFormUsers.SelectedIndex).Value
        End If

        If func.dbNullFix(Request.Form("sDepartments")) <> "" Then
            sDepartments.Items.FindByValue(Request.Form("sDepartments")).Selected = True
        ElseIf func.dbNullFix(Request.QueryString("sDepartments")) <> "" Then
            sDepartments.Items.FindByValue(Request.QueryString("sDepartments")).Selected = True
        Else
            sDepartments.Items(0).Selected = True
        End If
        If sDepartments.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sDepartments=" & sDepartments.Items(sDepartments.SelectedIndex).Value
        End If

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

        If IsDate(Request.Form("sReservationFromDate")) Then
            FromInsertDate = Request.Form("sReservationFromDate")
        ElseIf IsDate(Request.QueryString("sReservationFromDate")) Then
            FromInsertDate = Request.QueryString("sReservationFromDate")
        Else
            FromInsertDate = ""
        End If
        sReservationFromDate.Value = FromInsertDate
        If sReservationFromDate.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sReservationFromDate=" & sReservationFromDate.Value
        End If

        If IsDate(Request.Form("sReservationToDate")) Then
            ToInsertDate = Request.Form("sReservationToDate")
        ElseIf IsDate(Request.QueryString("sReservationToDate")) Then
            ToInsertDate = Request.QueryString("sReservationToDate")
        Else
            ToInsertDate = ""
        End If
        sReservationToDate.Value = ToInsertDate
        If sReservationToDate.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sReservationToDate=" & sReservationToDate.Value
        End If

        If IsDate(Request.Form("sLastSendingFromDate")) Then
            FromSendingDate = Request.Form("sLastSendingFromDate")
        ElseIf IsDate(Request.QueryString("sLastSendingFromDate")) Then
            FromSendingDate = Request.QueryString("sLastSendingFromDate")
        Else
            FromSendingDate = ""
        End If
        sLastSendingFromDate.Value = FromSendingDate
        If sLastSendingFromDate.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sLastSendingFromDate=" & sLastSendingFromDate.Value
        End If

        If IsDate(Request.Form("sLastSendingToDate")) Then
            ToSendingDate = Request.Form("sLastSendingToDate")
        ElseIf IsDate(Request.QueryString("sLastSendingToDate")) Then
            ToSendingDate = Request.QueryString("sLastSendingToDate")
        Else
            ToSendingDate = ""
        End If
        sLastSendingToDate.Value = ToSendingDate
        If sLastSendingToDate.Value <> "" Then
            queryParamSearch = queryParamSearch & "&sLastSendingToDate=" & sLastSendingToDate.Value
        End If

    End Sub
    Public Sub bindList()
        ' Response.Write(Request.Form("sPayFromDate"))
        ' Response.End()


        setSearchFields()

        setSortFields()

        cmdSelect = New SqlCommand("get_ReservationFormsSendings", con)
        cmdSelect.CommandType = CommandType.StoredProcedure
        'cmdSelect.Parameters.Add("@TourId", SqlDbType.Int).Value = drpTours.SelectedValue

        cmdSelect.Parameters.Add("@siteDBName", SqlDbType.VarChar, 20).Value = ConfigurationSettings.AppSettings("PegasusSiteDBName")

        cmdSelect.Parameters.Add("@seriesId", SqlDbType.VarChar).Value = sSeries.Value
        'cmdSelect.Parameters.Add("@contactId", SqlDbType.VarChar).Value = sContactName.Value
        cmdSelect.Parameters.Add("@contactName", SqlDbType.VarChar).Value = sContactName.Value
        cmdSelect.Parameters.Add("@guideId", SqlDbType.VarChar).Value = sGuides.Value
        cmdSelect.Parameters.Add("@departureId", SqlDbType.VarChar).Value = sDepartures.Value
        cmdSelect.Parameters.Add("@departmentId", SqlDbType.VarChar).Value = sDepartments.Value
        cmdSelect.Parameters.Add("@createFormUserId", SqlDbType.VarChar).Value = sCreateFormUsers.Value
        cmdSelect.Parameters.Add("@fromInsertDate", SqlDbType.VarChar, 30).Value = FromInsertDate
        cmdSelect.Parameters.Add("@toInsertDate", SqlDbType.VarChar, 30).Value = ToInsertDate
        cmdSelect.Parameters.Add("@countSendings", SqlDbType.VarChar, 20).Value = sCountSendings.Value
        cmdSelect.Parameters.Add("@fromSendingDate", SqlDbType.VarChar, 30).Value = FromSendingDate
        cmdSelect.Parameters.Add("@toSendingDate", SqlDbType.VarChar, 30).Value = ToSendingDate
        cmdSelect.Parameters.Add("@sendingFormUserId", SqlDbType.VarChar).Value = sSendingFormUsers.Value
        cmdSelect.Parameters.Add("@countryId", SqlDbType.VarChar).Value = Request("sCountries")

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

            'send to future and visibled, not canceled departures if 
            ' don't exists for this contact appel with type 16735 ( hatum rishum )  during closed 6 monthes
            ' and returned reservertion form has sent by system less from 3 times

            If DateDiff(DateInterval.Day, CDate(e.Item.DataItem("departure_Date")), Date.Today) < 0 _
And CBool(e.Item.DataItem("Departure_Vis")) _
And Not CBool(e.Item.DataItem("isCanceled")) _
And Not CBool(e.Item.DataItem("isExists16735For6monthes")) _
And func.fixNumeric(e.Item.DataItem("countSendings")) < 3 Then
                Dim phchkSendForm As PlaceHolder
                phchkSendForm = e.Item.FindControl("phchkSendForm")
                phchkSendForm.Visible = True
                '   Response.Write(e.Item.DataItem("Checked_ForBizpower"))

            End If
        End If
    End Sub



End Class

