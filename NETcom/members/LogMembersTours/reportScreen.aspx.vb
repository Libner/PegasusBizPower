Public Class LogMembersToursScreenReport
    Inherits System.Web.UI.Page
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim func As New bizpower.cfunc
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected sortQuery, sortQuery1, FromDate, ToDate As String
    Public XLSfile As New System.Text.StringBuilder("")
    Public qrystring As String
    Public dtScreenColumns As New DataTable

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
        Response.Buffer = False
        Server.ScriptTimeout = 600
        'Response.Charset = "windows-1255"





        XLSfile.Append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x = 'urn:schemas-microsoft-com:office:excel'  xmlns = 'http://www.w3.org/TR/REC-html40' dir=RTL>")
        XLSfile.Append("<head>")
        XLSfile.Append("<meta name=ProgId content=Excel.Sheet>")
        XLSfile.Append("<meta name=Generator content=""Microsoft Excel 9"">")
        XLSfile.Append("<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; CHARSET=windows-1255"">")
        XLSfile.Append("<style>")
        XLSfile.Append("<!--table")
        XLSfile.Append("	{mso-displayed-decimal-separator:""\."";")
        XLSfile.Append("	mso-displayed-thousand-separator:""\,"";}")
        XLSfile.Append("@page")
        XLSfile.Append("{ margin:1.0in .75in 1.0in .75in;")
        XLSfile.Append("	mso-header-margin:.5in;")
        XLSfile.Append("	mso-footer-margin:.5in;")
        XLSfile.Append(" mso-page-orientation:landscape;}")
        XLSfile.Append("br")
        XLSfile.Append("	{mso-data-placement:same-cell;}")
        XLSfile.Append(".xl24")
        XLSfile.Append("	{mso-style-parent:style0;")
        XLSfile.Append("	font-family:Arial, sans-serif;")
        XLSfile.Append("	font-size:12.0pt;")
        XLSfile.Append("	font-weight:700;")
        XLSfile.Append("	text-align:center;")
        XLSfile.Append("	mso-height-source: userset;")
        XLSfile.Append("    height:20pt;")
        XLSfile.Append("	white-space:normal;}")
        XLSfile.Append(".xl25")
        XLSfile.Append("	{")
        XLSfile.Append("	color:#000000;")
        XLSfile.Append("	font-size:10.0pt;")
        XLSfile.Append("	font-weight:700;")
        XLSfile.Append("	font-family:Arial, sans-serif;")
        XLSfile.Append("	text-align:center;")
        XLSfile.Append("	background:#C0C0C0;")
        XLSfile.Append("	border: 0.5pt solid black;")
        XLSfile.Append("	mso-pattern:auto none;")
        XLSfile.Append("	mso-height-source: userset;")
        XLSfile.Append("	height:22pt;")
        XLSfile.Append("    white-space:normal;}")
        XLSfile.Append(".xl26")
        XLSfile.Append("	{")
        XLSfile.Append("	font-weight:400;")
        XLSfile.Append("	font-size:10.0pt;")
        XLSfile.Append("	font-family:Arial, sans-serif;")
        XLSfile.Append("	border: 0.5pt solid black;")
        XLSfile.Append("	mso-height-source: userset;")
        XLSfile.Append("    white-space:nowrap;")
        XLSfile.Append("    direction:rtl;")
        XLSfile.Append("	vertical-align:top;")
        XLSfile.Append("	text-align:right;}")
        XLSfile.Append("-->")
        XLSfile.Append("</style>")
        XLSfile.Append("<!--[if gte mso 9]><xml>")
        XLSfile.Append("<x:ExcelWorkbook>")
        XLSfile.Append("<x:ExcelWorksheets>")
        XLSfile.Append(" <x:ExcelWorksheet>")
        XLSfile.Append("    <x:Name>דוח לקוחות גולשים באתר</x:Name>")
        XLSfile.Append("    <x:WorksheetOptions>")
        XLSfile.Append("    <x:Print>")
        XLSfile.Append("    <x:ValidPrinterInfo/>")
        XLSfile.Append("    <x:PaperSizeIndex>9</x:PaperSizeIndex>")
        XLSfile.Append("    <x:HorizontalResolution>600</x:HorizontalResolution>")
        XLSfile.Append("    <x:VerticalResolution>600</x:VerticalResolution>")
        XLSfile.Append("    </x:Print>")
        XLSfile.Append("    <x:Selected/>")
        XLSfile.Append("    <x:DisplayRightToLeft/>")
        XLSfile.Append("    <x:ProtectContents>False</x:ProtectContents>")
        XLSfile.Append("    <x:ProtectObjects>False</x:ProtectObjects>")
        XLSfile.Append("    <x:ProtectScenarios>False</x:ProtectScenarios>")
        XLSfile.Append("    </x:WorksheetOptions>")
        XLSfile.Append("</x:ExcelWorksheet>")
        XLSfile.Append("</x:ExcelWorksheets>")
        XLSfile.Append("<x:WindowHeight>7605</x:WindowHeight>")
        XLSfile.Append("<x:WindowWidth>15360</x:WindowWidth>")
        XLSfile.Append("<x:WindowTopX>0</x:WindowTopX>")
        XLSfile.Append("<x:WindowTopY>645</x:WindowTopY>")
        XLSfile.Append("<x:ProtectStructure>False</x:ProtectStructure>")
        XLSfile.Append("<x:ProtectWindows>False</x:ProtectWindows>")
        XLSfile.Append("</x:ExcelWorkbook>")
        XLSfile.Append("</xml><![endif]-->")
        XLSfile.Append("</head>")
        XLSfile.Append("<body>")
        XLSfile.Append("<table style='border-collapse:collapse;table-layout:fixed' dir=ltr>")
        XLSfile.Append("<tr><td colspan=5 class=xl24>דוח לקוחות גולשים באתר </td></tr>")
        XLSfile.Append("<tr><td colspan=5 class=xl24> ") : XLSfile.Append(String.Format("{0:dd/MM/yy HH:mm}", Now())) : XLSfile.Append(" נכון לתאריך</td></tr>")

        XLSfile.Append("<tr><td colspan=5 class=xl24>&nbsp;</td></tr>")

        setTitles()

        XLSfile.Append("<tr>")
        
        For i As Integer = 0 To dtScreenColumns.Rows.Count - 1
            XLSfile.Append("<td class=xl25>") : XLSfile.Append(dtScreenColumns.Rows(i)("Column_Title")) : XLSfile.Append("</td>")
        Next
        XLSfile.Append("</tr>")

        GetdataList()

        XLSfile.Append("</TABLE></body></html>")

        Dim filestring As String = "report_logToursScreen_" & CStr(Now.Hour) & "_" & CStr(Now.Minute) & "_" & CStr(Now.Second) & ".xls"

        Response.Clear()
        Response.AddHeader("content-disposition", "inline; filename=" & filestring)
        Response.ContentType = "application/vnd.ms-excel"
        Response.Write(XLSfile)
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
            ' Response.Write("<BR>" & sortQuery1 & "<BR>")
            '         Response.End() '   sortQuery = sortQuery1
        End If
    End Sub
    Sub setSearchFields()
        If IsDate(Server.UrlDecode(Request.Form("sLastEnterFromDate"))) Then
            FromDate = Server.UrlDecode(Request.Form("sLastEnterFromDate"))
        ElseIf IsDate(Server.UrlDecode(Request.QueryString("sLastEnterFromDate"))) Then
            FromDate = Server.UrlDecode(Request.QueryString("sLastEnterFromDate"))
        Else
            FromDate = ""
        End If

        If IsDate(Server.UrlDecode(Request.Form("sLastEnterToDate"))) Then
            ToDate = Server.UrlDecode(Request.Form("sLastEnterToDate"))
        ElseIf IsDate(Server.UrlDecode(Request.QueryString("sLastEnterToDate"))) Then
            ToDate = Server.UrlDecode(Request.QueryString("sLastEnterToDate"))
        Else
            ToDate = ""
        End If


    End Sub
    Sub GetdataList()

        setSearchFields()

        setSortFields()


        cmdSelect = New SqlClient.SqlCommand("get_LogContactsToursInteresting_Report", con)
        cmdSelect.CommandType = CommandType.StoredProcedure
        'cmdSelect.Parameters.Add("@TourId", SqlDbType.Int).Value = drpTours.SelectedValue

        cmdSelect.Parameters.Add("@siteDBName", SqlDbType.VarChar, 20).Value = ConfigurationSettings.AppSettings("PegasusSiteDBName")
        cmdSelect.Parameters.Add("@ExistsAppleal16504", SqlDbType.VarChar, 20).Value = Request("sExistsAppleal16504")
        cmdSelect.Parameters.Add("@ExistsAppleal16735", SqlDbType.VarChar, 20).Value = Request("sExistsAppleal16735")
        cmdSelect.Parameters.Add("@CountEnters", SqlDbType.VarChar, 20).Value = Request.QueryString("sCountEnters")
        cmdSelect.Parameters.Add("@seriesId", SqlDbType.VarChar).Value = Request.QueryString("sSeries")
        cmdSelect.Parameters.Add("@countryId", SqlDbType.VarChar).Value = Request("sCountries")
        cmdSelect.Parameters.Add("@subCatId", SqlDbType.VarChar).Value = Request("sSubCategoriesTours")
        cmdSelect.Parameters.Add("@fromDate", SqlDbType.VarChar, 30).Value = FromDate
        cmdSelect.Parameters.Add("@toDate", SqlDbType.VarChar, 30).Value = ToDate
        cmdSelect.Parameters.Add("@Email", SqlDbType.VarChar, 100).Value = Server.UrlDecode(Request("sEmail"))
        cmdSelect.Parameters.Add("@Phone", SqlDbType.VarChar, 100).Value = Server.UrlDecode(Request("sPhone"))
        cmdSelect.Parameters.Add("@ContactName", SqlDbType.VarChar, 100).Value = Server.UrlDecode(Request("sContactName"))

        'return value - total records count
        cmdSelect.Parameters.Add("@sortQuery", SqlDbType.VarChar, 500).Value = sortQuery
        'Response.Write(sortQuery)
        'Response.End()
        'For Each p As SqlClient.SqlParameter In cmdSelect.Parameters
        '    Response.Write("<br>" & p.ParameterName & "=" & p.Value)
        'Next
        'Response.End()
        con.Open()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()

        While dr.Read()
            XLSfile.Append("<tr>")
            'If dr("Checked_ForBizpower") = "true" Then
            '    XLSfile.Append("<td class=xl26>כן</td>")
            'Else
            '    XLSfile.Append("<td class=xl26>לא</td>")
            'End If
            For i As Integer = 0 To dtScreenColumns.Rows.Count - 1
                If dtScreenColumns.Rows(i)("Column_Name") = "ExistsAppleal16504" Or dtScreenColumns.Rows(i)("Column_Name") = "ExistsAppleal16735" Then
                    If func.dbNullBool(dr(dtScreenColumns.Rows(i)("Column_Name"))) Then
                        XLSfile.Append("<td class=xl26>") : XLSfile.Append("כן") : XLSfile.Append("</td>")
                    Else
                        XLSfile.Append("<td class=xl26>") : XLSfile.Append("לא") : XLSfile.Append("</td>")
                    End If
                Else
                    XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.dbNullFix(dr(dtScreenColumns.Rows(i)("Column_DBName")))) : XLSfile.Append("</td>")
                End If
            Next

            XLSfile.Append("</tr>")
        End While
        dr.Close() : conPegasus.Close()


    End Sub
End Class
