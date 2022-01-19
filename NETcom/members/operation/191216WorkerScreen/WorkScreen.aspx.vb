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
    Protected sPayFromDate, sPayToDate, sItineraryDate, sBrifDate, sGroupMeetingDate, sMeetingAfterTripDate As System.Web.UI.HtmlControls.HtmlInputText
  
    Public dtSeries, dtUsers, dtGuides, dtSuppliers As New DataTable
    Dim primKeydtSeries(0), primKeydtUsers(0), primKeydtGuides(0), primKeydtSuppliers(0) As Data.DataColumn
    Protected WithEvents sSeries, sUsers, sGuides, sSuppliers As System.Web.UI.HtmlControls.HtmlSelect
    Protected FromDate, ToDate, pItineraryDate, pBrifDate, pGroupMeetingDate, pMeetingAfterTripDate As String
    Protected pGilboaHotel, pVoucher_Simultaneous, pVoucher_Group, pDeparture_Costing, pVouchers_Provider, pStatus As String




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

            If Request.Cookies("bizpegasus")("Chief") <> "1" Then
                SeriasId = func.GetSeriasId(Request.Cookies("bizpegasus")("UserId"))
                '   Response.Write(SeriasId)

            End If
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

                    Case "2"
                        query = "and DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_DateBrief) =0 "
                    Case "3"
                        query = "and  DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date) < 0 and DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date_End)>= 0 "

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


                    Case "5"
                        query = "  and ( DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date) < 0 and DATEDIFF(dd, GetDate(), Tours_Departures.Departure_Date)>=-30 ) and  	 len(ISNULL(Departure_Costing,'') )=0 "

                    Case "6"
                        query = " and ( DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_DateBrief) >=0 ) "

                    Case "7"

                    Case "8"
                        query = " and ( DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date) < 0 and DATEDIFF(dd, GetDate(), Tours_Departures.Departure_Date)>=-30 ) and Vouchers_Provider<>'מותאם' "
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



        If IsDate(Request.Form("sPayFromDate")) Then 'תאריך יציאה
            FromDate = Request.Form("sPayFromDate")
        Else
            FromDate = ""
        End If
        If IsDate(Request.Form("sPayToDate")) Then 'תאריך חזרה
            ToDate = Request.Form("sPayToDate")
        Else
            ToDate = ""
        End If
        If IsDate(Request.Form("sItineraryDate")) Then 'תאריך  קבלת Itinerary
            pItineraryDate = Request.Form("sItineraryDate")
        Else
            pItineraryDate = ""
        End If
        If IsDate(Request.Form("sBrifDate")) Then
            pBrifDate = Request.Form("sBrifDate")
        Else
            pBrifDate = ""
        End If
        If IsDate(Request.Form("sGroupMeetingDate")) Then 'תאריך  מפגש קבוצה 
            pGroupMeetingDate = Request.Form("sGroupMeetingDate")
        Else
            pGroupMeetingDate = ""
        End If
        If IsDate(Request.Form("sMeetingAfterTripDate")) Then 'תאריך פגישה לאחר טיול
            pMeetingAfterTripDate = Request.Form("sMeetingAfterTripDate")
        Else
            pMeetingAfterTripDate = ""
        End If

        'If IsDate(Request.QueryString("sFromD")) Then
        '    FromDate = Request.QueryString("sFromD")
        'End If


        'If IsDate(Request.QueryString("sToD")) Then
        '    ToDate = Request.QueryString("sToD")
        'End If
        '      Response.Write(sCountries.Value)

        ''--------- cmdSelect = New SqlCommand("GetScreenAdminStatic", conPegasus)
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
            cmdSelect.Parameters.Add("@SupplierId", SqlDbType.Int).Value = sSuppliers.Value

            cmdSelect.Parameters.Add("@pGilboaHotel", SqlDbType.VarChar, 2).Value = pGilboaHotel
            cmdSelect.Parameters.Add("@pVoucher_Simultaneous", SqlDbType.VarChar, 2).Value = pVoucher_Simultaneous
            cmdSelect.Parameters.Add("@pVoucher_Group", SqlDbType.VarChar, 2).Value = pVoucher_Group
            cmdSelect.Parameters.Add("@pDeparture_Costing", SqlDbType.VarChar, 2).Value = pDeparture_Costing
            cmdSelect.Parameters.Add("@pVouchers_Provider", SqlDbType.VarChar, 5).Value = pVouchers_Provider
            cmdSelect.Parameters.Add("@pStatus", SqlDbType.VarChar, 15).Value = pStatus

            cmdSelect.Parameters.Add("@fromDate", SqlDbType.VarChar, 30).Value = FromDate
            cmdSelect.Parameters.Add("@toDate", SqlDbType.VarChar, 30).Value = ToDate

            cmdSelect.Parameters.Add("@pItineraryDate", SqlDbType.VarChar, 30).Value = pItineraryDate
            cmdSelect.Parameters.Add("@pBrifDate", SqlDbType.VarChar, 30).Value = pBrifDate
            cmdSelect.Parameters.Add("@pGroupMeetingDate", SqlDbType.VarChar, 30).Value = pGroupMeetingDate
            cmdSelect.Parameters.Add("@pMeetingAfterTripDate", SqlDbType.VarChar, 30).Value = pMeetingAfterTripDate

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
        sSuppliers.Items.Add(list1)
        For i As Integer = 0 To dtSuppliers.Rows.Count - 1
            Dim list As New ListItem(dtSuppliers.Rows(i)("supplier_Name"), dtSuppliers.Rows(i)("supplier_Id"))
            If Request.QueryString("sSuppliers") > 0 And Request.QueryString("sSuppliers") = dtSuppliers.Rows(i)("supplier_Id") Then
                list.Selected = True

            End If
            sSuppliers.Items.Add(list)
        Next



    End Sub
    Sub GetUsers()
        Dim cmdSelect As New SqlClient.SqlCommand("SELECT USER_ID, FIRSTNAME + ' ' + LASTNAME as Username from Users  where [ACTIVE]=1 order by FIRSTNAME asc ,LASTNAME asc", con)
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
            If Request.QueryString("sUser") > 0 And Request.QueryString("sUser") = dtUsers.Rows(i)("USER_ID") Then
                list.Selected = True

            End If
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
        sPayFromDate.Value = ""
        sPayToDate.Value = ""
        FromDate = ""
        ToDate = ""
        sItineraryDate.Value = ""
        pItineraryDate = ""
        sBrifDate.value = ""
        pBrifDate = ""
        sGroupMeetingDate.Value = ""
        pGroupMeetingDate = ""
        sMeetingAfterTripDate.value = ""
        pMeetingAfterTripDate = ""
        pGilboaHotel = ""
        pVoucher_Simultaneous = ""
        pVoucher_Group = ""
        pDeparture_Costing = ""
        pVouchers_Provider = ""
        pStatus = ""
        'sortQuery = ""

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

End Class
