Imports System.Data.SqlClient
Public Class LogMembersTotalToursScreenByMembers
    Inherits System.Web.UI.Page
    Protected WithEvents rptTitle As Repeater
    Protected WithEvents rptLogs As Repeater
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected totalRows As Integer = 0
    Protected TotalPages As Integer = 0
    Protected CurrentIndex As Integer = 0
    Protected sContactName, sPhone, sEmail, sLastEnterFromDate, sLastEnterToDate, sCountEnters As String
    Public sSubCategoriesTours, sSeries, sCountries, sExistsAppleal16735, sExistsAppleal16504 As String
    Public dtSubCategoriesTours, dtSeries, dtCountries As New DataTable

    Dim cmdSelect As New SqlClient.SqlCommand
    Dim primKeydtSubCategoriesTours(0), primKeydtSeries(0), primKeydtCountries(0) As Data.DataColumn
    Protected WithEvents pnlPages, pnlSearchMess As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents PageSize, pageList As Web.UI.WebControls.DropDownList
    Protected lblCount, lblTotalPages As Web.UI.WebControls.Label
    Protected WithEvents cmdPrev, cmdNext, btnIsPaid, btnSearch As Web.UI.WebControls.LinkButton
    Protected WithEvents btnSaveChecked As Web.UI.WebControls.Button
    Protected FromDate, ToDate, sortQuery, sortQuery1 As String

    Public qrystring As String
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

        setTitles()
        If Not Page.IsPostBack Then
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

    Sub setSortFields()
        sortQuery = ""
        Dim r As Integer
        '''--  sortQuery = " Dep_Name ASC"
        qrystring = Request.ServerVariables("QUERY_STRING")
        r = qrystring.IndexOf("sort")
        If r > 0 Then
            '  Response.Write("<BR>gg=" & Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)) & "<BR>")
            sortQuery1 = Replace(Mid(qrystring, qrystring.IndexOf("sort"), Len(qrystring)), "&", ",")
            sortQuery1 = Replace(sortQuery1, "=", " ")
            ' Response.Write("<BR>" & dtScreenColumns.Rows.Count & "<BR>")
            For i As Integer = 0 To dtScreenColumns.Rows.Count - 1
                sortQuery1 = Replace(sortQuery1, "sort_" & dtScreenColumns.Rows(i)("Column_Id"), dtScreenColumns.Rows(i)("Column_DBName"))
            Next

            sortQuery1 = sortQuery1.Substring(1)

            sortQuery = sortQuery1
            '   sortQuery = sortQuery1
        End If
    End Sub

    Sub setSearchFields()

        If func.dbNullFix(Request.Form("sContactName")) <> "" Then
            sContactName = Request.Form("sContactName")
        ElseIf func.dbNullFix(Request.QueryString("sContactName")) <> "" Then
            sContactName = Request.QueryString("sContactName")
        Else
            sContactName = ""
        End If

        If func.dbNullFix(Request.Form("sPhone")) <> "" Then
            sPhone = Request.Form("sPhone")
        ElseIf func.dbNullFix(Request.QueryString("sPhone")) <> "" Then
            sPhone = Request.QueryString("sPhone")
        Else
            sPhone = ""
        End If

        If func.dbNullFix(Request.Form("sEmail")) <> "" Then
            sEmail = Request.Form("sPhone")
        ElseIf func.dbNullFix(Request.QueryString("sEmail")) <> "" Then
            sEmail = Request.QueryString("sEmail")
        Else
            sEmail = ""
        End If

        If func.dbNullFix(Request.Form("sSubCategoriesTours")) <> "" Then
            sSubCategoriesTours = Request.Form("sSubCategoriesTours")
        ElseIf func.dbNullFix(Request.QueryString("sSubCategoriesTours")) <> "" Then
            sSubCategoriesTours = Request.QueryString("sSubCategoriesTours")
        Else
            sSubCategoriesTours = ""
        End If

        If func.dbNullFix(Request.Form("sCountries")) <> "" Then
            sCountries = Request.Form("sCountries")
        ElseIf func.dbNullFix(Request.QueryString("sCountries")) <> "" Then
            sCountries = Request.QueryString("sCountries")
        Else
            sCountries = ""
        End If

        If func.dbNullFix(Request.Form("sSeries")) <> "" Then
            sSeries = Request.Form("sSeries")
        ElseIf func.dbNullFix(Request.QueryString("sSeries")) <> "" Then
            sSeries = Request.QueryString("sSeries")
        Else
            sSeries = ""
        End If

        If IsDate(Request.Form("sLastEnterFromDate")) Then
            FromDate = Request.Form("sLastEnterFromDate")
        ElseIf IsDate(Request.QueryString("sLastEnterFromDate")) Then
            FromDate = Request.QueryString("sLastEnterFromDate")
        Else
            FromDate = ""
        End If
        sLastEnterFromDate = FromDate

        If IsDate(Request.Form("sLastEnterToDate")) Then
            ToDate = Request.Form("sLastEnterToDate")
        ElseIf IsDate(Request.QueryString("sLastEnterToDate")) Then
            ToDate = Request.QueryString("sLastEnterToDate")
        Else
            ToDate = ""
        End If
        sLastEnterToDate = ToDate


        If func.dbNullFix(Request.Form("sSeries")) <> "" Then
            sSeries = Request.Form("sSeries")
        ElseIf func.dbNullFix(Request.QueryString("sSeries")) <> "" Then
            sSeries = Request.QueryString("sSeries")
        Else
            sSeries = ""
        End If


        If func.dbNullFix(Request.Form("sCountEnters")) <> "" Then
            sCountEnters = Request.Form("sCountEnters")
        ElseIf func.dbNullFix(Request.QueryString("sCountEnters")) <> "" Then
            sCountEnters = Request.QueryString("sCountEnters")
        Else
            sCountEnters = ""
        End If

        If func.dbNullFix(Request.Form("sExistsAppleal16735")) <> "" Then
            sExistsAppleal16735 = Request.Form("sExistsAppleal16735")
        ElseIf func.dbNullFix(Request.QueryString("sExistsAppleal16735")) <> "" Then
            sExistsAppleal16735 = Request.QueryString("sExistsAppleal16735")
        Else
            sExistsAppleal16735 = ""
        End If

        If func.dbNullFix(Request.Form("sExistsAppleal16504")) <> "" Then
            sExistsAppleal16504 = Request.Form("sExistsAppleal16504")
        ElseIf func.dbNullFix(Request.QueryString("sExistsAppleal16504")) <> "" Then
            sExistsAppleal16504 = Request.QueryString("sExistsAppleal16504")
        Else
            sExistsAppleal16504 = ""
        End If

    End Sub
    Public Sub bindList()
        ' Response.Write(Request.Form("sPayFromDate"))
        ' Response.End()


        setSearchFields()

        setSortFields()

        cmdSelect = New SqlCommand("get_LogContactsToursInterestingFilterBycontact", con)
        cmdSelect.CommandType = CommandType.StoredProcedure
        'cmdSelect.Parameters.Add("@TourId", SqlDbType.Int).Value = drpTours.SelectedValue

        cmdSelect.Parameters.Add("@siteDBName", SqlDbType.VarChar, 20).Value = ConfigurationSettings.AppSettings("PegasusSiteDBName")
        cmdSelect.Parameters.Add("@ExistsAppleal16504", SqlDbType.VarChar, 20).Value = sExistsAppleal16504
        cmdSelect.Parameters.Add("@ExistsAppleal16735", SqlDbType.VarChar, 20).Value = sExistsAppleal16735
        cmdSelect.Parameters.Add("@CountEnters", SqlDbType.VarChar, 20).Value = sCountEnters
        cmdSelect.Parameters.Add("@seriesId", SqlDbType.VarChar).Value = sSeries
        cmdSelect.Parameters.Add("@countryId", SqlDbType.VarChar).Value = sCountries
        cmdSelect.Parameters.Add("@subCatId", SqlDbType.VarChar).Value = sSubCategoriesTours
        cmdSelect.Parameters.Add("@fromDate", SqlDbType.VarChar, 30).Value = FromDate
        cmdSelect.Parameters.Add("@toDate", SqlDbType.VarChar, 30).Value = ToDate
        cmdSelect.Parameters.Add("@Email", SqlDbType.VarChar, 100).Value = sEmail
        cmdSelect.Parameters.Add("@Phone", SqlDbType.VarChar, 100).Value = sPhone
        cmdSelect.Parameters.Add("@ContactName", SqlDbType.VarChar, 100).Value = sContactName

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




End Class

