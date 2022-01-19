Public Class changesSales
    Inherits System.Web.UI.Page

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
        Session.CodePage = 1255
        Dim date_start, date_end, ChangeType, ChangeTable As String
        date_start = Trim(Request.QueryString("date_startSales"))
        date_end = Trim(Request.QueryString("date_endSales"))
        ChangeTable = Trim(Request.QueryString("cmbChangeTableSales"))
        If ChangeTable = "" Then
            ChangeTable = 0
        End If
        Dim XLSfile As New System.Text.StringBuilder("")
        XLSfile.Append("<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<head>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1255"">") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<meta name=ProgId content=Excel.Sheet>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<meta name=Generator content=""Microsoft Excel 9"">") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<style>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<!--table") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	{mso-displayed-decimal-separator:""\."";") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-displayed-thousand-separator:""\,"";}") : XLSfile.Append(vbCrLf)
        XLSfile.Append("@page") : XLSfile.Append(vbCrLf)
        XLSfile.Append("{ margin:1.0in .75in 1.0in .75in;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-header-margin:.5in;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-footer-margin:.5in;") : XLSfile.Append(vbCrLf)
        XLSfile.Append(" mso-page-orientation:landscape;}") : XLSfile.Append(vbCrLf)
        XLSfile.Append("tr") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	{mso-height-source:auto;}") : XLSfile.Append(vbCrLf)
        XLSfile.Append("br") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	{mso-data-placement:same-cell;}") : XLSfile.Append(vbCrLf)
        XLSfile.Append(".style0") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	{mso-number-format:General;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	vertical-align:bottom;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	white-space:nowrap;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-rotate:0;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-background-source:auto;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-pattern:auto;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	color:windowtext;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-size:10.0pt;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-weight:400;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-style:normal;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	text-decoration:none;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-family:Arial;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-generic-font-family:auto;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-font-charset:0;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	border:none;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-protection:locked visible;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	text-align:center;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-style-name:Normal;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-height-source:auto userset;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	height:20pt;}") : XLSfile.Append(vbCrLf)
        XLSfile.Append(".xl24") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	{mso-style-parent:style0;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-family:Arial, sans-serif;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-size:12.0pt;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-weight:700;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	text-align:center;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-height-source: userset;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	height:18pt;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	white-space:normal;}") : XLSfile.Append(vbCrLf)
        XLSfile.Append(".xl25") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	{mso-style-parent:style0;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	color:#000000;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-size:10.0pt;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	margin-right:5pt;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	margin-left:5pt;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-weight:700;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-family:Arial, sans-serif;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	text-align:center;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	background:#FFFF99;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	border: 0.5pt solid black;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-pattern:auto none;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-height-source: userset;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	height:16pt;") : XLSfile.Append(vbCrLf)
        XLSfile.Append(" white-space:normal;}") : XLSfile.Append(vbCrLf)
        XLSfile.Append(".xl26") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	{mso-style-parent:style0;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-weight:400;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-size:10.0pt;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	margin-right:5pt;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	margin-left:5pt;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	font-family:Arial, sans-serif;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	border: 0.5pt solid black;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	mso-height-source: userset;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	vertical-align:top;") : XLSfile.Append(vbCrLf)
        XLSfile.Append("	text-align:right;}") : XLSfile.Append(vbCrLf)
        XLSfile.Append("-->") : XLSfile.Append(vbCrLf)
        XLSfile.Append("</style>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<!--[if gte mso 9]><xml>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:ExcelWorkbook>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:ExcelWorksheets>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:ExcelWorksheet>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:Name>דוח לוגים</x:Name>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:WorksheetOptions>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:Print>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("      <x:ValidPrinterInfo/>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("      <x:PaperSizeIndex>9</x:PaperSizeIndex>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("      <x:HorizontalResolution>600</x:HorizontalResolution>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("      <x:VerticalResolution>600</x:VerticalResolution>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("</x:Print>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:Selected/>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:ProtectContents>False</x:ProtectContents>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:ProtectObjects>False</x:ProtectObjects>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:ProtectScenarios>False</x:ProtectScenarios>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:WorksheetOptions>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:Selected/>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:DisplayRightToLeft/>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:Panes>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:Pane>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:Number>1</x:Number>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:ActiveRow>1</x:ActiveRow>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("</x:Pane>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("</x:Panes>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:ProtectContents>False</x:ProtectContents>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:ProtectObjects>False</x:ProtectObjects>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:ProtectScenarios>False</x:ProtectScenarios>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("</x:WorksheetOptions>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("</x:ExcelWorksheet>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("</x:ExcelWorksheets>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:WindowTopX>360</x:WindowTopX>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<x:WindowTopY>60</x:WindowTopY>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("</x:ExcelWorkbook>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("</xml><![endif]-->") : XLSfile.Append(vbCrLf)
        XLSfile.Append("</head>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<body>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<table style='border-collapse:collapse;table-layout:fixed' dir=ltr>") : XLSfile.Append(vbCrLf)
        XLSfile.Append("<tr><td colspan='11' class='xl24'>דוח לוגים על ביצוע פעולות בדש בורד מכירות</td></tr>")
        XLSfile.Append("<tr><td colspan='11' class='xl24'>") : XLSfile.Append(date_start) : XLSfile.Append(" - ") : XLSfile.Append(date_end) : XLSfile.Append("</td></tr>")

        If False Then ' ChangeTable > 0 Then
            XLSfile.Append("<tr><td colspan='11' class='xl24'>סוג אובייקט: ") : XLSfile.Append(ChangeTable) : XLSfile.Append("</td></tr>")
        End If
        XLSfile.Append("<tr><td colspan='11' class='xl24'> ")
        XLSfile.Append(String.Format("{0:dd/MM/yy HH:mm}", Now())) : XLSfile.Append(" נכון לתאריך</td></tr>") : XLSfile.Append(vbCrLf)


        XLSfile.Append("<tr><td>&nbsp;</td></tr>")
        XLSfile.Append("<tr style='mso-height-source:userset;height:28.5pt'><td class=xl25>אובייקט</td><td class=xl25>שם טבלה</td><td class=xl25>שם שדה</td><td class=xl25>יום</td><td class=xl25>חודש</td><td class=xl25>שנה</td>")
        XLSfile.Append("<td class=xl25>מחלקה</td><td class=xl25>תאריך שינוי</td><td class=xl25>שם משתמש</td><td class=xl25>כתובת IP</td></tr>") : XLSfile.Append(vbCrLf)

        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmdSelect As New SqlClient.SqlCommand("get_report_changesSales", con)
        cmdSelect.CommandType = CommandType.StoredProcedure
        If Len(date_start) > 0 Then
            cmdSelect.Parameters.Add("@date_start", Trim(date_start))
        Else
            cmdSelect.Parameters.Add("@date_start", DBNull.Value)
        End If
        If Len(date_end) > 0 Then
            cmdSelect.Parameters.Add("@date_end", Trim(date_end))
        Else
            cmdSelect.Parameters.Add("@date_end", DBNull.Value)
        End If
        cmdSelect.Parameters.Add("@ChangeTable", Trim(ChangeTable))
        con.Open()
        Dim dtTmp As New DataTable
        Dim da As New SqlClient.SqlDataAdapter(cmdSelect)
        dtTmp.BeginLoadData()
        da.Fill(dtTmp)
        dtTmp.EndLoadData()
        cmdSelect.Dispose()
        con.Close()

        If dtTmp.Rows.Count > 0 Then
            For pp As Integer = 0 To dtTmp.Rows.Count - 1
                XLSfile.Append("<tr><td class=xl26 x:str>") : XLSfile.Append(Trim(dtTmp.Rows(pp)("Change_Table"))) : XLSfile.Append("</td>")
                XLSfile.Append("<td class=xl26 x:str>") : XLSfile.Append(Trim(dtTmp.Rows(pp)("ChangeSubTable"))) : XLSfile.Append("</td>")
                XLSfile.Append("<td class=xl26 x:str>") : XLSfile.Append(Trim(dtTmp.Rows(pp)("Object_Title"))) : XLSfile.Append("</td>")
                XLSfile.Append("<td class=xl26 x:str>") : XLSfile.Append(Trim(dtTmp.Rows(pp)("Day"))) : XLSfile.Append("</td>")
                XLSfile.Append("<td class=xl26 x:str>") : XLSfile.Append(Trim(dtTmp.Rows(pp)("Month"))) : XLSfile.Append("</td>")
                XLSfile.Append("<td class=xl26 x:str>") : XLSfile.Append(Trim(dtTmp.Rows(pp)("Year"))) : XLSfile.Append("</td>")
                XLSfile.Append("<td class=xl26 x:str>") : XLSfile.Append(Trim(dtTmp.Rows(pp)("DepartmentName"))) : XLSfile.Append("</td>")
                XLSfile.Append("<td class=xl26 x:str>") : XLSfile.Append(Trim(dtTmp.Rows(pp)("Change_Date"))) : XLSfile.Append("</td>")
                XLSfile.Append("<td class=xl26 x:str>") : XLSfile.Append(Trim(dtTmp.Rows(pp)("UserName"))) : XLSfile.Append("</td>")
                If Not IsDBNull(dtTmp.Rows(pp)("IP")) Then
                    XLSfile.Append("<td class=xl26 x:str>") : XLSfile.Append(Trim(dtTmp.Rows(pp)("IP"))) : XLSfile.Append("</td></tr>") : XLSfile.Append(vbCrLf)
                Else
                    XLSfile.Append("<td class=xl26 x:str></td></tr>") : XLSfile.Append(vbCrLf)
                End If
            Next
        End If
        XLSfile.Append("</TABLE></body></html>")

        Dim filestring As String = "changesSales_" & CStr(Now.Hour) & "_" & CStr(Now.Minute) & "_" & CStr(Now.Second) & ".xls"

        dtTmp.Dispose()

        Response.Clear()
        Response.AddHeader("content-disposition", "inline; filename=" & filestring)
        Response.ContentType = "application/vnd.ms-excel"
        Response.Write(XLSfile)


    End Sub

End Class
