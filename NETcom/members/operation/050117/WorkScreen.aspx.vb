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
    Protected SeriasId, query As String
    Protected tab As Integer
    Protected WithEvents btnSearchAll, btnSearch As Web.UI.WebControls.LinkButton
    Protected sFromDateEnd, sToDateEnd, sFromDate, sToDate, sFromItineraryDate, sToItineraryDate, sFromBrifDate, sToBrifDate, sFromGroupMeetingDate, sToGroupMeetingDate, sFromMeetingAfterTripDate, sToMeetingAfterTripDate As System.Web.UI.HtmlControls.HtmlInputText
    Protected tabName As String
    Protected pSuppliers As String = ""

    Public dtSeries, dtUsers, dtGuides, dtSuppliers As New DataTable
    Dim primKeydtSeries(0), primKeydtUsers(0), primKeydtGuides(0), primKeydtSuppliers(0) As Data.DataColumn
    Protected WithEvents sSeries, sUsers, sGuides, sSuppliers As System.Web.UI.HtmlControls.HtmlSelect
    Protected FromDateEnd, ToDateEnd, FromDate, ToDate, pFromItineraryDate, pToItineraryDate, pFromBrifDate, pToBrifDate, pFromGroupMeetingDate, pToGroupMeetingDate, pFromMeetingAfterTripDate, pToMeetingAfterTripDate As String
    Protected pcodeTiul, pGilboaHotel, pVoucher_Simultaneous, pVoucher_Group, pDeparture_Costing, pVouchers_Provider, pStatus, sortQuery, sortQuery1, qrystring As String




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

        sortQuery = ""
        Dim r As Integer
        '''--  sortQuery = " Dep_Name ASC"
        qrystring = Request.ServerVariables("QUERY_STRING")
        r = qrystring.IndexOf("sort")
        If r > 0 Then
            '  Response.Write("<BR>gg=" & Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)) & "<BR>")
            sortQuery1 = Replace(Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)), "&", ",")
            sortQuery1 = Replace(sortQuery1, "=", " ")
            sortQuery1 = Replace(sortQuery1, "sort_10", "supplier_Name")
            sortQuery1 = Replace(sortQuery1, "sort_1", "LASTNAME")
            sortQuery1 = Replace(sortQuery1, "sort_2", "Series_Name")
            sortQuery1 = Replace(sortQuery1, "sort_3", "Departure_Date")
            sortQuery1 = Replace(sortQuery1, "sort_4", "Departure_Date_End")
            sortQuery1 = Replace(sortQuery1, "sort_5", "Guide_LName")
            sortQuery1 = Replace(sortQuery1, "sort_6", "Departure_Itinerary")
            sortQuery1 = Replace(sortQuery1, "sort_7", "Departure_DateBrief")
            sortQuery1 = Replace(sortQuery1, "sort_8", "Departure_DateGroupMeeting")
            sortQuery1 = Replace(sortQuery1, "sort_9", "DateMeetingAfterTrip")


            sortQuery1 = sortQuery1.Substring(1)

            sortQuery = sortQuery1
            'Response.Write("<BR>" & sortQuery1 & "<BR>")
            ' Response.End() '   sortQuery = sortQuery1
        End If


        If Not Page.IsPostBack Then
            'Response.Write(Request.Cookies("bizpegasus")("UserId") & "<BR>")
            '  Response.Write(Request.Cookies("bizpegasus")("Chief") & "<BR>")
            If Request.Cookies("bizpegasus")("Chief") <> "1" Then
                SeriasId = func.GetSeriasId(Request.Cookies("bizpegasus")("UserId"))

              
            End If

            '  Response.Write(Request.Cookies("bizpegasus")("Chief"))
            GetSeries(SeriasId)
            GetUsers()
            GetGuides()
            GetSuppliers()

            If IsNumeric(Request.QueryString("tab")) Then
                tab = Request.QueryString("tab")
                ' Response.Write(tab)
                Select Case tab

                    Case "1"
                        query = "and  DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date) < 0 and DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date_End)>= 0 "
                        tabName = "כרגע בחו""ל"
                    Case "2"
                        query = "and DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_DateBrief) =0 "
                        tabName = "בריפים היום"
                    Case "3"
                        query = "and  DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date) < 0 and DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date_End)>= 0 "
                        tabName = "בחו""ל ולא התקשרו"
                        ' SELECT  dbo.Tours_Departures.Departure_Id FROM   dbo.Tours_Departures
                        '	INNER JOIN Tours ON Tours_Departures.Tour_Id = Tours.Tour_Id INNER JOIN Tours_Categories ON Tours.Category_Id = Tours_Categories.Category_Id 
                        'left JOIN Departments ON Tours.DeparureId = Departments.Dep_Id 
                        'left join bizpower_pegasus_test.dbo.Series BS on BS.Series_Id=Tours.SeriasId left JOIN bizpower_pegasus_test.dbo.Users BU on BS.User_Id=BU.User_Id
                        'left join bizpower_pegasus_test.dbo.GuideMessages GM on GM.Departure_Id=Tours_Departures.Departure_Id
                        ' WHERE 1=1 and  ( DATEDIFF(dd, GetDate(), Tours_Departures.Departure_Date)>=-30 or DATEDIFF(dd, GetDate(), Tours_Departures.Departure_Date) >=0)

                        'GROUP  BY dbo.Tours_Departures.Departure_Id
                        'HAVING COUNT(GM.Messages_Id) >0

                    Case "4"
                        query = " and len(ISNULL(cast(Departure_DateBrief as varchar),''))=0 "
                        tabName = "לא תואם בריף החזרה"
                    Case "5"
                        query = "  and ( DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date) < 0 and DATEDIFF(dd, GetDate(), Tours_Departures.Departure_Date)>=-30 ) and  	 len(ISNULL(Departure_Costing,'') )=0 "
                        tabName = "הטיולים שיצאו ואין תמחור"
                    Case "6"
                        query = " and ( DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_DateBrief) >=0 ) "
                        tabName = "לא תואם בריף החזרה"
                    Case "7"
                        'תיולי החודש הנוכחי. תאריך יציאה שלכם
                        query = "  and ( DATEDIFF(mm, GETDATE(), Tours_Departures.Departure_Date) = 0 ) "
                        tabName = "טיולי החודש"
                    Case "8"
                        query = " and ( DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date) < 0 and DATEDIFF(dd, GetDate(), Tours_Departures.Departure_Date)>=-30 ) and Vouchers_Provider<>'מותאם' "
                        tabName = "שובר ללא התאמה"
                    Case "9"
                        query = " and (DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date_End)<= 0 and len(ISNULL(cast(DateMeetingAfterTrip as varchar),''))=0 ) "
                        tabName = "לא תואם בריף"
                    Case "10"
                        tabName = "רשימת בריפים קדימה"
                        query = " and DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_DateBrief)>= 0 "
                    Case "11"
                        query = " and ( DATEDIFF(dd, GETDATE(), Tours_Departures.DateMeetingAfterTrip) = 0 ) "
                        tabName = "החזרת טיולים היום"
                End Select

            End If

            bindList()
        End If


    End Sub

    Public Sub bindList()
        ' Response.Write(Request.Form("sPayFromDate"))
        ' Response.End()
        If Request.Form("sGilboaHotel") <> "" Then
            pGilboaHotel = Request.Form("sGilboaHotel")
        Else
            pGilboaHotel = ""
        End If
        If Request.Form("sVoucher_Group") <> "" Then
            pVoucher_Group = Request.Form("sVoucher_Group")
        Else
            pVoucher_Group = ""
        End If
        If Request.Form("codeTiul") <> "" Then
            pcodeTiul = Request.Form("codeTiul")
        Else
            pcodeTiul = ""
        End If

        If Request.Form("sVoucher_Simultaneous") <> "" Then
            pVoucher_Simultaneous = Request.Form("sVoucher_Simultaneous")
        Else
            pVoucher_Simultaneous = ""
        End If
        If Request.Form("sDeparture_Costing") <> "" Then
            pDeparture_Costing = Request.Form("sDeparture_Costing")
        Else
            pDeparture_Costing = ""
        End If
        If Request.Form("sVouchers_Provider") <> "" Then
            pVouchers_Provider = Request.Form("sVouchers_Provider")
        Else
            pVouchers_Provider = ""
        End If
        If Request.Form("sStatus") <> "" Then
            pStatus = Request.Form("sStatus")
        Else
            pStatus = ""
        End If
        ' Response.Write("j=" & pStatus)
        ' Response.End()
        If IsDate(Request.Form("sFromDateEnd")) Then 'תאריך חזרה
            FromDateEnd = Request.Form("sFromDateEnd")
        Else
            FromDateEnd = ""
        End If
        If IsDate(Request.Form("sToDateEnd")) Then 'תאריך חזרה
            ToDateEnd = Request.Form("sToDateEnd")
        Else
            ToDateEnd = ""
        End If

        If IsDate(Request.Form("sFromDate")) Then 'תאריך יציאה
            FromDate = Request.Form("sFromDate")
        Else
            FromDate = ""
        End If
        If IsDate(Request.Form("sToDate")) Then 'תאריך חזרה
            ToDate = Request.Form("sToDate")
        Else
            ToDate = ""
        End If
        If IsDate(Request.Form("sFromItineraryDate")) Then 'תאריך  קבלת Itinerary
            pFromItineraryDate = Request.Form("sFromItineraryDate")
        Else
            pFromItineraryDate = ""
        End If
        If IsDate(Request.Form("sToItineraryDate")) Then 'תאריך  קבלת Itinerary
            pToItineraryDate = Request.Form("sToItineraryDate")
        Else
            pToItineraryDate = ""
        End If
        If IsDate(Request.Form("sFromBrifDate")) Then
            pFromBrifDate = Request.Form("sFromBrifDate")
        Else
            pFromBrifDate = ""
        End If
        If IsDate(Request.Form("sToBrifDate")) Then
            pToBrifDate = Request.Form("sToBrifDate")
        Else
            pToBrifDate = ""
        End If
        If IsDate(Request.Form("sFromGroupMeetingDate")) Then 'תאריך  מפגש קבוצה 
            pFromGroupMeetingDate = Request.Form("sFromGroupMeetingDate")
        Else
            pFromGroupMeetingDate = ""
        End If
        If IsDate(Request.Form("sToGroupMeetingDate")) Then 'תאריך  מפגש קבוצה 
            pToGroupMeetingDate = Request.Form("sToGroupMeetingDate")
        Else
            pToGroupMeetingDate = ""
        End If
        '''If IsDate(Request.Form("sMeetingAfterTripDate")) Then 'תאריך פגישה לאחר טיול
        '''    pMeetingAfterTripDate = Request.Form("sMeetingAfterTripDate")
        '''Else
        '''    pMeetingAfterTripDate = ""
        '''End If
        If IsDate(Request.Form("sFromMeetingAfterTripDate")) Then 'תאריך פגישה לאחר טיול
            pFromMeetingAfterTripDate = Request.Form("sFromMeetingAfterTripDate")
        Else
            pFromMeetingAfterTripDate = ""
        End If
        If IsDate(Request.Form("sToMeetingAfterTripDate")) Then 'תאריך פגישה לאחר טיול
            pToMeetingAfterTripDate = Request.Form("sToMeetingAfterTripDate")
        Else
            pToMeetingAfterTripDate = ""
        End If


        'If IsDate(Request.QueryString("sFromD")) Then
        '    FromDate = Request.QueryString("sFromD")
        'End If


        'If IsDate(Request.QueryString("sToD")) Then
        '    ToDate = Request.QueryString("sToD")
        'End If
        '      Response.Write(sCountries.Value)

        ''--------- cmdSelect = New SqlCommand("GetScreenAdminStatic", conPegasus)
        '   Response.Write("s=" & Request.Form("sSuppliers"))

        '  Response.End()
        ' pSuppliers = Request.Form("sSuppliers")
        '  Response.Write("pSuppliers=" & pSuppliers & ":")
        ' Response.End()
        '  If pSuppliers = "" Then pSuppliers = 0

        If query <> "" Then
            pnlPages.Visible = False
            cmdSelect = New SqlCommand("GetWorkScreen_tab", conPegasus)
            cmdSelect.CommandType = CommandType.StoredProcedure
            If SeriasId <> "" Then
                cmdSelect.Parameters.Add("@StringSeriasId", SqlDbType.VarChar, 300).Value = SeriasId
            Else
                cmdSelect.Parameters.Add("@StringSeriasId", SqlDbType.VarChar, 300).Value = "0"
            End If
            cmdSelect.Parameters.Add("@PageSize", DbType.Int32).Value = Trim(PageSize.SelectedValue)
            cmdSelect.Parameters.Add("@PageNumber", DbType.Int32).Value = CInt(CurrentIndex) + 1
            cmdSelect.Parameters.Add("@query", SqlDbType.VarChar, 500).Value = query
            cmdSelect.Parameters.Add("@tab", DbType.Int32).Value = CInt(tab)
            cmdSelect.Parameters.Add("@CountRecords", System.Data.SqlDbType.Int).Value = 0
            cmdSelect.Parameters("@CountRecords").Direction = ParameterDirection.Output


        Else

            cmdSelect = New SqlCommand("GetWorkScreen", conPegasus)
            cmdSelect.CommandType = CommandType.StoredProcedure
            If SeriasId <> "" Then
                cmdSelect.Parameters.Add("@StringSeriasId", SqlDbType.VarChar, 300).Value = SeriasId
            Else
                cmdSelect.Parameters.Add("@StringSeriasId", SqlDbType.VarChar, 300).Value = "0"
            End If


            cmdSelect.Parameters.Add("@userId", SqlDbType.Int).Value = sUsers.Value
            cmdSelect.Parameters.Add("@SeriesId", SqlDbType.Int).Value = sSeries.Value
            cmdSelect.Parameters.Add("@GuideId", SqlDbType.Int).Value = sGuides.Value
            'If Len(pSuppliers) > 0 Then
            '    cmdSelect.Parameters.Add("@SupplierId", SqlDbType.VarChar, 100).Value = pSuppliers ' pSuppliers 'sSuppliers.Value
            'Else
            '    cmdSelect.Parameters.Add("@SupplierId", SqlDbType.VarChar, 100).Value = "" ' pSuppliers 'sSuppliers.Value

            'End If
            cmdSelect.Parameters.Add("@SupplierId", SqlDbType.Int, 100).Value = sSuppliers.Value

            cmdSelect.Parameters.Add("@pGilboaHotel", SqlDbType.VarChar, 2).Value = pGilboaHotel
            cmdSelect.Parameters.Add("@pVoucher_Simultaneous", SqlDbType.VarChar, 2).Value = pVoucher_Simultaneous
            cmdSelect.Parameters.Add("@pVoucher_Group", SqlDbType.VarChar, 2).Value = pVoucher_Group
            cmdSelect.Parameters.Add("@pDeparture_Costing", SqlDbType.VarChar, 2).Value = pDeparture_Costing
            cmdSelect.Parameters.Add("@pVouchers_Provider", SqlDbType.VarChar, 5).Value = pVouchers_Provider
            cmdSelect.Parameters.Add("@pStatus", SqlDbType.VarChar, 15).Value = pStatus

            cmdSelect.Parameters.Add("@FromDateEnd", SqlDbType.VarChar, 30).Value = FromDateEnd
            cmdSelect.Parameters.Add("@ToDateEnd", SqlDbType.VarChar, 30).Value = ToDateEnd

            cmdSelect.Parameters.Add("@fromDate", SqlDbType.VarChar, 30).Value = FromDate
            cmdSelect.Parameters.Add("@toDate", SqlDbType.VarChar, 30).Value = ToDate

            cmdSelect.Parameters.Add("@pFromItineraryDate", SqlDbType.VarChar, 30).Value = pFromItineraryDate
            cmdSelect.Parameters.Add("@pToItineraryDate", SqlDbType.VarChar, 30).Value = pToItineraryDate

            cmdSelect.Parameters.Add("@pFromBrifDate", SqlDbType.VarChar, 30).Value = pFromBrifDate
            cmdSelect.Parameters.Add("@pToBrifDate", SqlDbType.VarChar, 30).Value = pToBrifDate

            cmdSelect.Parameters.Add("@pFromGroupMeetingDate", SqlDbType.VarChar, 30).Value = pFromGroupMeetingDate
            cmdSelect.Parameters.Add("@pToGroupMeetingDate", SqlDbType.VarChar, 30).Value = pToGroupMeetingDate

            cmdSelect.Parameters.Add("@pFromMeetingAfterTripDate", SqlDbType.VarChar, 30).Value = pFromMeetingAfterTripDate
            cmdSelect.Parameters.Add("@pToMeetingAfterTripDate", SqlDbType.VarChar, 30).Value = pToMeetingAfterTripDate
            cmdSelect.Parameters.Add("@pcodeTiul", SqlDbType.VarChar, 100).Value = pcodeTiul



            cmdSelect.Parameters.Add("@sortQuery", SqlDbType.VarChar, 500).Value = sortQuery

            cmdSelect.Parameters.Add("@PageSize", DbType.Int32).Value = Trim(PageSize.SelectedValue)
            cmdSelect.Parameters.Add("@PageNumber", DbType.Int32).Value = CInt(CurrentIndex) + 1
            cmdSelect.Parameters.Add("@CountRecords", System.Data.SqlDbType.Int).Value = 0
            cmdSelect.Parameters("@CountRecords").Direction = ParameterDirection.Output
        End If
        conPegasus.Open()
        Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
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
        ' End If


    End Sub
    Sub GetGuides()
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT Guide_ID, Guide_FName +' ' + Guide_LName as  GuideName from Guides  where [Guide_Vis]=1 order by Guide_FName asc ,Guide_LName asc", conPegasus)
        conPegasus.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtGuides)
        conPegasus.Close()
        primKeydtGuides(0) = dtGuides.Columns("Guide_ID")
        dtGuides.PrimaryKey = primKeydtGuides
        sGuides.Items.Clear()
        Dim list1 As New ListItem("הכל", "0")
        sGuides.Items.Add(list1)
        For i As Integer = 0 To dtGuides.Rows.Count - 1
            Dim list As New ListItem(dtGuides.Rows(i)("GuideName"), dtGuides.Rows(i)("Guide_ID"))
            If Request.QueryString("sGuides") > 0 And Request.QueryString("sGuides") = dtGuides.Rows(i)("Guide_ID") Then
                list.Selected = True

            End If
            sGuides.Items.Add(list)
        Next



    End Sub
    Sub GetSuppliers()
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT  supplier_Id,supplier_Name " & _
        "FROM Suppliers order by supplier_Name", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtSuppliers)
        con.Close()
        primKeydtSuppliers(0) = dtSuppliers.Columns("supplier_Id")
        dtSuppliers.PrimaryKey = primKeydtSuppliers
        sSuppliers.Items.Clear()
        Dim list1 As New ListItem("הכל", "0")
        'If Request.QueryString("sSuppliers") = "0" Or Request.QueryString("sSuppliers") = "" Then list1.Selected = True

        sSuppliers.Items.Add(list1)
        For i As Integer = 0 To dtSuppliers.Rows.Count - 1
            Dim list As New ListItem(dtSuppliers.Rows(i)("supplier_Name"), dtSuppliers.Rows(i)("supplier_Id"))
            'If Request.QueryString("sSuppliers") > 0 And Request.QueryString("sSuppliers") = dtSuppliers.Rows(i)("supplier_Id") Then
            '    list.Selected = True

            'End If
            sSuppliers.Items.Add(list)
        Next



    End Sub
    Sub GetUsers()
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT USER_ID, FIRSTNAME + ' ' + LASTNAME as Username from Users  where [ACTIVE]=1 and job_id=471 order by FIRSTNAME asc ,LASTNAME asc", con)
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
            'If Request.QueryString("sUser") > 0 And Request.QueryString("sUser") = dtUsers.Rows(i)("USER_ID") Then
            '    list.Selected = True

            'End If
            sUsers.Items.Add(list)
        Next
    End Sub
    Sub GetSeries(ByVal SeriasId)
        Dim cmdSelect As New SqlClient.SqlCommand
        If SeriasId <> "" Then
            cmdSelect = New SqlClient.SqlCommand("select  Series_Id,Series_Name from Series where Series_Id in (" & SeriasId & ") ORDER BY Series_Name", con)

        Else
            cmdSelect = New SqlClient.SqlCommand("select  Series_Id,Series_Name from Series ORDER BY Series_Name", con)
        End If
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtSeries)
        con.Close()
        primKeydtSeries(0) = dtSeries.Columns("Series_Id")
        dtSeries.PrimaryKey = primKeydtSeries
        sSeries.Items.Clear()
        Dim list1 As New ListItem("הכל", "0")
        sSeries.Items.Add(list1)
        For i As Integer = 0 To dtSeries.Rows.Count - 1
            Dim list As New ListItem(dtSeries.Rows(i)("Series_Name"), dtSeries.Rows(i)("Series_Id"))
            If Request.QueryString("sSer") > 0 And Request.QueryString("sSer") = dtSeries.Rows(i)("Series_Id") Then
                list.Selected = True
            End If
            sSeries.Items.Add(list)
        Next



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
        'sDepartments.Value = 0
        sUsers.Value = 0
        sSeries.Value = 0
        sGuides.Value = 0
        sSuppliers.Value = 0

        'sStatus.Value = 0
        'sCountries.Value = 0
        sFromDate.Value = ""
        sToDate.Value = ""
        FromDateEnd = ""
        ToDateEnd = ""

        FromDate = ""
        ToDate = ""
        sFromItineraryDate.Value = ""
        pFromItineraryDate = ""
        sToItineraryDate.Value = ""
        pToItineraryDate = ""
        sFromBrifDate.Value = ""
        pFromBrifDate = ""
        sToBrifDate.Value = ""
        pToBrifDate = ""
        sFromGroupMeetingDate.Value = ""
        pFromGroupMeetingDate = ""
        sToGroupMeetingDate.Value = ""
        pToGroupMeetingDate = ""
        sFromMeetingAfterTripDate.Value = ""
        sToMeetingAfterTripDate.Value = ""
        pFromMeetingAfterTripDate = ""
        pToMeetingAfterTripDate = ""
        pGilboaHotel = ""
        pVoucher_Simultaneous = ""
        pVoucher_Group = ""
        pDeparture_Costing = ""
        pVouchers_Provider = ""
        pStatus = ""
        sortQuery = ""

        'qrystring = ""
        'bindList()
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

    Private Sub Page_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.DataBinding

    End Sub
End Class
