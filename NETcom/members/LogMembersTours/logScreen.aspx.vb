Imports System.Data.SqlClient
Public Class LogMembersToursScreen
    Inherits System.Web.UI.Page
    Protected WithEvents rptTitle As Repeater
    Protected WithEvents rptLogs As Repeater
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected totalRows As Integer = 0
    Protected TotalPages As Integer = 0
    Protected CurrentIndex As Integer = 0
    Protected sContactName, sPhone, sEmail, sLastEnterFromDate, sLastEnterToDate, sCountEnters As System.Web.UI.HtmlControls.HtmlInputText

    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected WithEvents sSubCategoriesTours, sSeries, sCountries, sExistsAppleal16735, sExistsAppleal16504 As System.Web.UI.HtmlControls.HtmlSelect
    Public dtSubCategoriesTours, dtSeries, dtCountries As New DataTable

    Dim cmdSelect As New SqlClient.SqlCommand
    Dim primKeydtSubCategoriesTours(0), primKeydtSeries(0), primKeydtCountries(0) As Data.DataColumn
    Protected WithEvents btnSearchAll As Web.UI.WebControls.LinkButton
    Protected WithEvents pnlPages, pnlSearchMess As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents PageSize, pageList As Web.UI.WebControls.DropDownList
    Protected lblCount, lblTotalPages As Web.UI.WebControls.Label
    Protected WithEvents cmdPrev, cmdNext, btnIsPaid, btnSearch As Web.UI.WebControls.LinkButton
    Protected WithEvents btnSaveChecked As Web.UI.WebControls.Button
    Protected FromDate, ToDate, sortQuery, sortQuery1 As String

    Public qrystring, sortQrystring As String
    Public dtScreenColumns As New DataTable
    Public AddAppealPermitted As Boolean
    'Dim arrColumnsNames() As String = {"ContactName", "Phone", "Email", "SubCategoriesTours", "Countries", "Series", "LastEnterDate", "CountEnters", "ExistsAppleal16735", "ExistsAppleal16504"}
    'Dim arrColumnsDesc() As String = {"שם הלקוח", "טלפון ", "אימייל", "שם הקטגוריה", "שם  היעד", "שם הסדרה", "תאריך אחרון", "מספר כניסות", "טייל ביעד?", "נוצר טופס מתעניין?"}


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

        If Request.Cookies("bizpegasus")("AddAppealFromLogToursInteresting") = "1" Then
            AddAppealPermitted = True
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
            GetCountries()
            GetSubCategoriesTours()
            bindList()
        End If

    End Sub
    Sub setTitles()
        If Not IsNothing(viewstate("ScreenColumns")) Then
            dtScreenColumns = viewstate("ScreenColumns")
        End If
        If dtScreenColumns.Rows.Count = 0 Then

        cmdSelect = New SqlClient.SqlCommand("SELECT Column_Id, Column_Name,Column_Title,Column_DBName  FROM Screen_LogMembersTours_Setting order by Column_Order", con)
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
            If e.Item.DataItem("Column_Name") = "ContactName" Then
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
            sortQrystring = Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring))
            sortQuery1 = Replace(Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)), "&", ",")
            sortQuery1 = Replace(sortQuery1, "=", " ")
            ' Response.Write("<BR>" & dtScreenColumns.Rows.Count & "<BR>")
            For i As Integer = 0 To dtScreenColumns.Rows.Count - 1
                sortQuery1 = Replace(sortQuery1, "sort_" & dtScreenColumns.Rows(i)("Column_Id"), dtScreenColumns.Rows(i)("Column_DBName"))
            Next

            sortQuery1 = sortQuery1.Substring(1)

            sortQuery = sortQuery1
            ' Response.Write("<BR>" & sortQuery1 & "<BR>")
            '         Response.End() '   sortQuery = sortQuery1
        End If

        'Response.Write("<BR>sortQuery=" & sortQuery & "<BR>")
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

    Sub GetSubCategoriesTours()
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT  SubCategory_Id, Tours_Categories.Category_Name + ' - ' + SubCategory_Name as SubCategory_Name " & _
  "FROM Tours_SubCategories INNER JOIN Tours_Categories ON Tours_SubCategories.Category_Id = Tours_Categories.Category_Id " & _
 " ORDER BY Tours_Categories.Category_Order, Tours_SubCategories.SubCategory_Order", conPegasus)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtSubCategoriesTours)
        con.Close()
        primKeydtSubCategoriesTours(0) = dtSubCategoriesTours.Columns("Series_Id")
        dtSubCategoriesTours.PrimaryKey = primKeydtSubCategoriesTours
        sSubCategoriesTours.Items.Clear()
        Dim list1 As New ListItem("", "0")
        sSubCategoriesTours.Items.Add(list1)
        ' Dim list1 As New ListItem("הכל", "0")
        'sSeries.Items.Add(list1)
        For i As Integer = 0 To dtSubCategoriesTours.Rows.Count - 1
            Dim list As New ListItem(dtSubCategoriesTours.Rows(i)("SubCategory_Name"), dtSubCategoriesTours.Rows(i)("SubCategory_Id"))
            If Request.QueryString("sSubCategoriesTours") > 0 And Request.QueryString("sSubCategoriesTours") = dtSubCategoriesTours.Rows(i)("SubCategory_Id") Then
                list.Selected = True
            End If
            sSubCategoriesTours.Items.Add(list)
        Next


    End Sub

    Sub btnSearchAll_onClick(ByVal s As Object, ByVal e As EventArgs) Handles btnSearchAll.Click
        pageList.SelectedIndex = 0
        CurrentIndex = CInt(pageList.SelectedIndex)
        sSubCategoriesTours.Value = 0
        sSeries.Value = 0
        sCountries.Value = 0
        sExistsAppleal16735.Value = 0
        sExistsAppleal16504.Value = 0
        sContactName.Value = ""
        sPhone.Value = ""
        sEmail.Value = ""
        sLastEnterFromDate.Value = ""
        sLastEnterToDate.Value = ""
        sCountEnters.Value = ""
        FromDate = ""
        ToDate = ""
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
        If IsDate(Request.Form("sLastEnterFromDate")) Then
            FromDate = Request.Form("sLastEnterFromDate")
        ElseIf IsDate(Request.QueryString("sLastEnterFromDate")) Then
            FromDate = Request.QueryString("sLastEnterFromDate")
        Else
            FromDate = ""
        End If
        sLastEnterFromDate.Value = FromDate

        If IsDate(Request.Form("sLastEnterToDate")) Then
            ToDate = Request.Form("sLastEnterToDate")
        ElseIf IsDate(Request.QueryString("sLastEnterToDate")) Then
            ToDate = Request.QueryString("sLastEnterToDate")
        Else
            ToDate = ""
        End If
        sLastEnterToDate.Value = ToDate

        If func.dbNullFix(Request.Form("sContactName")) <> "" Then
            sContactName.Value = Request.Form("sContactName")
        ElseIf func.dbNullFix(Request.QueryString("sContactName")) <> "" Then
            sContactName.Value = Request.QueryString("sContactName")
        Else
            sContactName.Value = ""
        End If
    End Sub
    Public Sub bindList()
        ' Response.Write(Request.Form("sPayFromDate"))
        ' Response.End()


        setSearchFields()

        setSortFields()

        cmdSelect = New SqlCommand("get_LogContactsToursInteresting", con)
        cmdSelect.CommandType = CommandType.StoredProcedure
        'cmdSelect.Parameters.Add("@TourId", SqlDbType.Int).Value = drpTours.SelectedValue

        cmdSelect.Parameters.Add("@siteDBName", SqlDbType.VarChar, 20).Value = ConfigurationSettings.AppSettings("PegasusSiteDBName")
        cmdSelect.Parameters.Add("@ExistsAppleal16504", SqlDbType.VarChar, 20).Value = sExistsAppleal16504.Value
        cmdSelect.Parameters.Add("@ExistsAppleal16735", SqlDbType.VarChar, 20).Value = sExistsAppleal16735.Value
        cmdSelect.Parameters.Add("@CountEnters", SqlDbType.VarChar, 20).Value = sCountEnters.Value
        cmdSelect.Parameters.Add("@seriesId", SqlDbType.VarChar).Value = sSeries.Value
        cmdSelect.Parameters.Add("@countryId", SqlDbType.VarChar).Value = sCountries.Value
        cmdSelect.Parameters.Add("@subCatId", SqlDbType.VarChar).Value = sSubCategoriesTours.Value
        cmdSelect.Parameters.Add("@fromDate", SqlDbType.VarChar, 30).Value = FromDate
        cmdSelect.Parameters.Add("@toDate", SqlDbType.VarChar, 30).Value = ToDate
        cmdSelect.Parameters.Add("@Email", SqlDbType.VarChar, 100).Value = sEmail.Value
        cmdSelect.Parameters.Add("@Phone", SqlDbType.VarChar, 100).Value = sPhone.Value
        cmdSelect.Parameters.Add("@ContactName", SqlDbType.VarChar, 100).Value = sContactName.Value

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
        ' Response.End()
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
            Dim phAppleal16735 As PlaceHolder
            phAppleal16735 = e.Item.FindControl("phAppleal16735")
            Dim phAppleal16504 As PlaceHolder
            phAppleal16504 = e.Item.FindControl("phAppleal16504")
            '   Response.Write(e.Item.DataItem("Checked_ForBizpower"))

            Dim phChkSendAppleal As PlaceHolder
            phChkSendAppleal = e.Item.FindControl("phChkSendAppleal")
            '   Response.Write(e.Item.DataItem("Checked_ForBizpower"))

            If func.fixNumeric(e.Item.DataItem("Tour_Id")) > 0 Then
                phAppleal16504.Visible = True
                phAppleal16735.Visible = True

                'Response.Write("<br>ExistsAppleal16504=" & e.Item.DataItem("ExistsAppleal16504"))

                If Not func.dbNullBool(e.Item.DataItem("ExistsAppleal16504")) And Not func.dbNullBool(e.Item.DataItem("ExistsAppleal16735")) And AddAppealPermitted Then
                    phChkSendAppleal.Visible = True
                End If
            End If

        End If
    End Sub

    Private Sub btnSaveChecked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveChecked.Click
        saveChekedTours()
    End Sub
    Sub saveChekedTours()
        Dim chk As HtmlInputCheckBox
        Dim i = 0
        Dim j As Integer
        j = rptLogs.Items.Count()

        Dim cmd As System.Data.SqlClient.SqlCommand


        For i = 0 To j - 1
            chk = CType(rptLogs.Items(i).FindControl("chkTour"), HtmlInputCheckBox)
            If chk.Checked Then


                '   Response.Write("checked=" & chk.Value())  '//////code  from betipul net admin addprof
                conPegasus.Open()

                cmd = New System.Data.SqlClient.SqlCommand("Update Tours_Departures set Checked_ForBizpower=1 where Departure_Id=@chkValue ", conPegasus)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@chkValue", SqlDbType.Int).Value = CInt(chk.Value())
                'cmd.Parameters.Add("@ProfId", SqlDbType.Int).Value = CInt(chk.Attributes("value"))
                cmd.ExecuteNonQuery()
                conPegasus.Close()
            End If
        Next

    End Sub

End Class

