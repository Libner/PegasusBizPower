Public Class reportScreen
    Inherits System.Web.UI.Page
    Protected sortQuery, sortQuery1, FromDate, ToDate As String
    Public XLSfile As New System.Text.StringBuilder("")
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
        Response.Buffer = False
        Server.ScriptTimeout = 600
        Response.Charset = "windows-1255"

        Dim func As New bizpower.cfunc

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


        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmdSelect As New SqlClient.SqlCommand




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
        XLSfile.Append("    <x:Name>דוח מיקודים</x:Name>")
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
        XLSfile.Append("<tr><td colspan=5 class=xl24>דוח מיקודים </td></tr>")
        XLSfile.Append("<tr><td colspan=5 class=xl24> ") : XLSfile.Append(String.Format("{0:dd/MM/yy HH:mm}", Now())) : XLSfile.Append(" נכון לתאריך</td></tr>")

        XLSfile.Append("<tr><td colspan=5 class=xl24>&nbsp;</td></tr>")
        Dim sqlstr As String = "SELECT Column_Id, Column_Name  FROM ScreenSetting order by Column_Order"
        ''Response.Write(sqlstr)
        ''Response.End()

        cmdSelect = New SqlClient.SqlCommand(sqlstr, con)
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()
        XLSfile.Append("<tr>")
        XLSfile.Append("<td class=xl25>") : XLSfile.Append(" מופיע במסך צפיה") : XLSfile.Append("</td>")

        While dr.Read()

            XLSfile.Append("<td class=xl25>") : XLSfile.Append(dr("Column_Name")) : XLSfile.Append("</td>")
        End While
        dr.Close() : con.Close()

        XLSfile.Append("</tr>")
        GetdataList()

        XLSfile.Append("</TABLE></body></html>")

        Dim filestring As String = "report_ScreenAdmin_" & CStr(Now.Hour) & "_" & CStr(Now.Minute) & "_" & CStr(Now.Second) & ".xls"

        Response.Clear()
        Response.AddHeader("content-disposition", "inline; filename=" & filestring)
        Response.ContentType = "application/vnd.ms-excel"
        Response.Write(XLSfile)
    End Sub
    Sub GetdataList()
        If IsDate(Request.QueryString("sFromD")) Then
            FromDate = Request.QueryString("sFromD")
        Else
            FromDate = ""
        End If

        If IsDate(Request.Form("sPayToDate")) Then
            ToDate = Request.Form("sPayToDate")
        Else
            ToDate = ""
        End If
        If IsDate(Request.QueryString("sToD")) Then
            ToDate = Request.QueryString("sToD")
        End If
        Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))

        Dim cmdSelect As New SqlClient.SqlCommand("GetScreenAdminStatic_Report", conPegasus)

        cmdSelect.CommandType = CommandType.StoredProcedure
        cmdSelect.Parameters.Add("@statusId", SqlDbType.Int).Value = Request.QueryString("sStatus")
        cmdSelect.Parameters.Add("@countryId", SqlDbType.Int).Value = Request.QueryString("sCountry")
        cmdSelect.Parameters.Add("@seriesId", SqlDbType.Int).Value = Request.QueryString("sSer")
        cmdSelect.Parameters.Add("@depId", SqlDbType.Int).Value = Request.QueryString("sdep")
        cmdSelect.Parameters.Add("@userId", SqlDbType.Int).Value = Request.QueryString("sUser")
        cmdSelect.Parameters.Add("@fromDate", SqlDbType.VarChar, 30).Value = FromDate
        cmdSelect.Parameters.Add("@toDate", SqlDbType.VarChar, 30).Value = ToDate
        cmdSelect.Parameters.Add("@sortQuery", SqlDbType.VarChar, 500).Value = sortQuery
        ' Response.Write(sortQuery)
        ' Response.End()
        conPegasus.Open()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()

        While dr.Read()
            XLSfile.Append("<tr>")
            If dr("Checked_ForBizpower") = "true" Then
                XLSfile.Append("<td class=xl26>כן</td>")
            Else
                XLSfile.Append("<td class=xl26>לא</td>")
            End If

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Status_Name")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Series_Name")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Date_Begin")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("CountryName")) : XLSfile.Append("</td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("CountMembers")) : XLSfile.Append("</td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Dep_Name")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("CountMembersBitulim")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("LastDateSale")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("CountlastWeek")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Countlast2Week")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("CountlastMonth")) : XLSfile.Append("</td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("User_Name")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Flying_Company")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Price")) : XLSfile.Append("</td>")

            XLSfile.Append("</tr>")
        End While
        dr.Close() : conPegasus.Close()


    End Sub
End Class
