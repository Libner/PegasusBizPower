Imports System.Data.SqlClient
Public Class ScreenControl
    Inherits System.Web.UI.Page
    Protected WithEvents sSeries, sGuides, sTours As System.Web.UI.HtmlControls.HtmlSelect
    Protected sPayFromDate, sPayToDate, FromGrade, ToGrade As System.Web.UI.HtmlControls.HtmlInputText
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Public dtSeries, dtGuides, dtTours As New DataTable
    Dim primKeydtSeries(0), primKeydtGuides(0), primKeydtTours(0) As Data.DataColumn
    Protected totalRows As Integer = 0
    Protected TotalPages As Integer = 0
    Protected CurrentIndex As Integer = 0
    Protected FromDate, ToDate, sortQuery, sortQuery1, psSeries, psTours, psGuides As String
    Protected pFromGrade, pToGrade As Decimal
   
    Protected WithEvents pnlPages, pnlSearchMess As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents PageSize, pageList As Web.UI.WebControls.DropDownList
    Protected lblCount, lblTotalPages As Web.UI.WebControls.Label
    Protected WithEvents cmdPrev, cmdNext, btnIsPaid, btnSearch As Web.UI.WebControls.LinkButton
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected WithEvents rptCustomers As Repeater
    Protected SeriesSelect, GuidesSelect, ToursSelect As System.Web.UI.HtmlControls.HtmlInputHidden


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
        'Response.Write("CurrentIndex=" & CurrentIndex)

      
 
        Dim r As Integer
        '''--  sortQuery = " Dep_Name ASC"
        qrystring = Request.ServerVariables("QUERY_STRING")

        ' Response.Write("qrystring=" & qrystring)
        r = qrystring.IndexOf("sort")
        If r > 0 Then
            sortQuery1 = Replace(Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)), "&", ",")
            sortQuery1 = Replace(sortQuery1, "=", " ")
            sortQuery1 = Replace(sortQuery1, "sort_5", "Tour_Grade")
            sortQuery1 = Replace(sortQuery1, "sort_4", "Guide_LName")
            sortQuery1 = Replace(sortQuery1, "sort_3", "Departure_Date")
            sortQuery1 = Replace(sortQuery1, "sort_2", "Series_Name")
            sortQuery1 = Replace(sortQuery1, "sort_1", "Tour_Name")

            sortQuery1 = sortQuery1.Substring(1)

            sortQuery = sortQuery1
            ' Response.Write("sortQuery=" & sortQuery)
        End If
        ' GetSeries()
        'GetGuides()
        'GetTours()
        'bindList()
 
        If Not Page.IsPostBack Then   ''change 27/03/2018
            bindList()
        End If

    End Sub
    Public Sub bindList()
        ' Response.Write("bl")
        '      Response.Write(Request.Form("sPayFromDate"))
        '      Response.End()
        If IsDate(Request.Form("sPayFromDate")) Then
            FromDate = Request.Form("sPayFromDate")
        Else
            FromDate = "01/11/2016"
        End If
        If IsDate(Request.Form("sPayToDate")) Then
            ToDate = Request.Form("sPayToDate")
        Else
            Dim NextTime As Date = Now
            ToDate = NextTime.AddDays(2)
        End If

        If IsNumeric(Request.Form("FromGrade")) Then
            pFromGrade = Request.Form("FromGrade")
        Else
            pFromGrade = 0
            If IsNumeric(Request.QueryString("FromGrade")) Then
                pFromGrade = Request.QueryString("FromGrade")
            End If
        End If
        If IsNumeric(Request.Form("ToGrade")) Then
            pToGrade = Request.Form("ToGrade")
        Else
            pToGrade = 0
            If IsNumeric(Request.QueryString("ToGrade")) Then
                pToGrade = Request.QueryString("ToGrade")
            End If
        End If
        'If pToGrade > 0 And pFromGrade = 0 Then
        '    pFromGrade = 1
        'End If
     
        If pFromGrade > 0 And pToGrade = 0 Then
            pToGrade = 100
        End If
        'Response.Write("pFromGrade=" & pFromGrade & "--pToGrade=" & pToGrade)


        If Request.Form("SeriesSelect") <> "" Then
            psSeries = Request.Form("SeriesSelect")
        Else
            psSeries = 0
        End If

        If Request.Form("ToursSelect") <> "" Then
            psTours = Request.Form("ToursSelect")
        Else
            psTours = ""
        End If
        If Request.Form("GuidesSelect") <> "" Then
            psGuides = Request.Form("GuidesSelect")
        Else
            psGuides = 0
        End If
        If qrystring = "" Then
            qrystring = "FromGrade=" & pFromGrade & "&ToGrade=" & pToGrade
            '& "&sGuides=" & psGuides & "&sPayFromDate=" & FromDate & "&sPayToDate=" & ToDate & "&sSeries=" & psSeries & "&sTours=" & psTours
        End If

        'Response.Write("-" & pFromGrade & "-" & pToGrade)

        'Response.Write("@sortQuery=" & sortQuery)
        '     Response.End()

        cmdSelect = New SqlCommand("GetScreenFeedbackControl", conPegasus)
        cmdSelect.CommandType = CommandType.StoredProcedure
        cmdSelect.Parameters.Add("@seriesId", SqlDbType.VarChar, 300).Value = psSeries 'sSeries.Value
        cmdSelect.Parameters.Add("@GuideId", SqlDbType.VarChar, 300).Value = psGuides 'sGuides.Value
        cmdSelect.Parameters.Add("@TourId", SqlDbType.VarChar, 300).Value = psTours   'sTours.Value

        cmdSelect.Parameters.Add("@fromDate", SqlDbType.VarChar, 30).Value = FromDate
        cmdSelect.Parameters.Add("@toDate", SqlDbType.VarChar, 30).Value = ToDate
        cmdSelect.Parameters.Add("@FromGrade", SqlDbType.Decimal).Value = pFromGrade
        cmdSelect.Parameters.Add("@ToGrade", SqlDbType.Decimal).Value = pToGrade
        'cmdSelect.Parameters.Add("@srchEmail", SqlDbType.VarChar, 100).Value = sEmail.Value
        cmdSelect.Parameters.Add("@PageSize", DbType.Int32).Value = Trim(PageSize.SelectedValue)
        cmdSelect.Parameters.Add("@PageNumber", DbType.Int32).Value = CInt(CurrentIndex) + 1

        'return value - total records count
        cmdSelect.Parameters.Add("@sortQuery", SqlDbType.VarChar, 500).Value = sortQuery
        'Response.Write(sSeries.Value & ":" & ":" & sGuides.Value & ":" & sTours.Value & ":" & sortQuery)
        '  Response.End()
        cmdSelect.Parameters.Add("@CountRecords", System.Data.SqlDbType.Int).Value = 0
        cmdSelect.Parameters("@CountRecords").Direction = ParameterDirection.Output
        conPegasus.Open()
        Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
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
            rptCustomers.Visible = False : pnlPages.Visible = False
        End If

        dr.Close()
        conPegasus.Close()

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
    Public Sub PageList_Click(ByVal s As System.Object, ByVal e As System.EventArgs) Handles pageList.SelectedIndexChanged
        CurrentIndex = CInt(pageList.SelectedItem.Value)
        'Response.Write("PageList_Click=" & CurrentIndex)
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
        ' Response.Write("next")
        pageList.SelectedIndex = pageList.SelectedIndex + 1
        'Response.Write("pageList.SelectedIndex=" & pageList.SelectedIndex)
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
    'Sub GetSeries()
    '    Dim cmdSelect As New SqlClient.SqlCommand("select  Series_Id,Series_Name from Series ORDER BY Series_Name", con)
    '    con.Open()
    '    Dim ad As New SqlClient.SqlDataAdapter
    '    ad.SelectCommand = cmdSelect
    '    ad.Fill(dtSeries)
    '    con.Close()
    '    primKeydtSeries(0) = dtSeries.Columns("Series_Id")
    '    dtSeries.PrimaryKey = primKeydtSeries
    '    sSeries.Items.Clear()
    '    Dim list1 As New ListItem("הכל", "0")
    '    sSeries.Items.Add(list1)
    '    For i As Integer = 0 To dtSeries.Rows.Count - 1
    '        Dim list As New ListItem(dtSeries.Rows(i)("Series_Name"), dtSeries.Rows(i)("Series_Id"))
    '        If Request.Form("sSeries") > 0 And Request.Form("sSeries") = dtSeries.Rows(i)("Series_Id") Then
    '            list.Selected = True
    '        End If
    '        sSeries.Items.Add(list)
    '    Next



    'End Sub
    ''Sub GetGuides()
    ''    Dim cmdSelect As New SqlClient.SqlCommand("SELECT Guide_Id, (Guide_FName + '  ' + Guide_LName) as Guide_Name  FROM Guides  where Guide_Vis=1 ORDER BY Guide_LName, Guide_FName", conPegasus)
    ''    conPegasus.Open()
    ''    Dim ad As New SqlClient.SqlDataAdapter
    ''    ad.SelectCommand = cmdSelect
    ''    ad.Fill(dtGuides)
    ''    conPegasus.Close()
    ''    primKeydtGuides(0) = dtGuides.Columns("Guide_Id")
    ''    dtGuides.PrimaryKey = primKeydtGuides
    ''    sGuides.Items.Clear()
    ''    Dim list1 As New ListItem("הכל", "0")
    ''    sGuides.Items.Add(list1)
    ''    For i As Integer = 0 To dtGuides.Rows.Count - 1
    ''        Dim list As New ListItem(dtGuides.Rows(i)("Guide_Name"), dtGuides.Rows(i)("Guide_Id"))
    ''        If Request.Form("sGuides") > 0 And Request.Form("sGuides") = dtGuides.Rows(i)("Guide_Id") Then
    ''            list.Selected = True
    ''        End If

    ''        sGuides.Items.Add(list)
    ''    Next


    ''End Sub
    'Sub GetTours()
    '    Dim cmdSelect As New SqlClient.SqlCommand("set dateformat dmy;SELECT distinct  Tours.Tour_Id ,Tour_Name=case when isnull(Tour_PageTitle,'')='' then Tour_Name else Tour_PageTitle end " & _
    '     " FROM   Tours_Departures		INNER JOIN Tours ON Tours_Departures.Tour_Id = Tours.Tour_Id 		where	DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date_End) <= -1 " & _
    '     "order by Tours.Tour_Id desc", conPegasus)
    '    conPegasus.Open()
    '    Dim ad As New SqlClient.SqlDataAdapter
    '    ad.SelectCommand = cmdSelect
    '    ad.Fill(dtTours)
    '    conPegasus.Close()
    '    primKeydtTours(0) = dtTours.Columns("TourId")
    '    dtTours.PrimaryKey = primKeydtTours
    '    sTours.Items.Clear()
    '    Dim list1 As New ListItem("הכל", "0")
    '    sTours.Items.Add(list1)
    '    For i As Integer = 0 To dtTours.Rows.Count - 1
    '        Dim list As New ListItem(dtTours.Rows(i)("Tour_Name"), dtTours.Rows(i)("Tour_Id"))
    '        If Request.Form("sTours") > 0 And Request.Form("sTours") = dtTours.Rows(i)("Tour_Id") Then
    '            list.Selected = True
    '        End If

    '        sTours.Items.Add(list)
    '    Next


    'End Sub


    Private Sub rptCustomers_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptCustomers.ItemDataBound
        Dim tdRes, tdProc As System.Web.UI.HtmlControls.HtmlTableCell
        Dim conF As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionStringS"))

        Dim rptDeparture_Id As String
        Dim IsTravelersTour As String
        Dim rptCountMembers As String
        Dim CountFeedBack As Integer
        Dim CountPeopleFeedBack As Integer
        Dim procPeople As Decimal
        'Dim rptGrade As Integer
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then

            rptDeparture_Id = e.Item.DataItem("Departure_Id")
            IsTravelersTour = e.Item.DataItem("IsTravelersTour")
            tdRes = e.Item.FindControl("tdRes")
            tdProc = e.Item.FindControl("tdProc")
            If Not IsDBNull(e.Item.DataItem("CountMembers")) Then
                rptCountMembers = e.Item.DataItem("CountMembers")
                Dim cmdSelectF = New SqlCommand("select count(Id) from FeedBack_Form  WHERE FeedBack_Status=2 and  Departure_Id = @DepartureId", conF)
                cmdSelectF.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(rptDeparture_Id)
                conF.Open()
                If Not IsDBNull(cmdSelectF.ExecuteScalar()) Then
                    CountFeedBack = cmdSelectF.ExecuteScalar()
                Else
                    '        'CountFeedBack = 1
                    CountFeedBack = 0
                End If
                conF.Close()
                CountPeopleFeedBack = func.NumFeedBackPeople(CInt(rptDeparture_Id))
                If IsTravelersTour = "1" Then
                    tdRes.InnerText = CountFeedBack & "/" & rptCountMembers

                Else
                    tdRes.InnerText = CountFeedBack & "/" & CountPeopleFeedBack & "/" & rptCountMembers
                End If
                If rptCountMembers > 0 Then
                    If IsTravelersTour = "1" Then
                        procPeople = Math.Round((CountFeedBack / rptCountMembers) * 100, 2)
                    Else
                        procPeople = Math.Round((CountPeopleFeedBack / rptCountMembers) * 100, 2)

                    End If
                Else
                        procPeople = 0
                    End If
                    If procPeople > 0 Then
                        tdProc.InnerText = procPeople & "%"
                    Else
                        tdProc.InnerText = ""
                    End If


                End If

            End If

        ''rptCustomers()
    End Sub
End Class
