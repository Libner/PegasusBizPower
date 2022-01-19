Public Class report_orders_all
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
        Response.Buffer = False
        Server.ScriptTimeout = 600
        Response.Charset = "windows-1255"

        Dim func As New bizpower.cfunc
        Dim UserId As Integer = func.fixNumeric(Request.Cookies("bizpegasus")("UserId"))
        If (UserId = 0) Then
            Response.End()
        End If

        Dim dateStart As String = Trim(Request.QueryString("dateStart"))
        Dim dateEnd As String = Trim(Request.QueryString("dateEnd"))
        Dim reportUserId As Integer = func.fixNumeric(Request.QueryString("UserId"))
        Dim reportUserName As String = ""

        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmdSelect As New SqlClient.SqlCommand

        If reportUserId > 0 Then
            cmdSelect = New SqlClient.SqlCommand("SELECT (FIRSTNAME + ' ' + LASTNAME) UserName FROM Users WHERE (User_ID = " & reportUserId & ") ", con)
            cmdSelect.CommandType = CommandType.Text
            con.Open()
            reportUserName = func.dbNullFix(cmdSelect.ExecuteScalar())
            cmdSelect.Dispose() : con.Close()
        End If

        Dim XLSfile As New System.Text.StringBuilder("")
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
        XLSfile.Append("    <x:Name>דוח הזמנות כל הנציגים</x:Name>")
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
        XLSfile.Append("<tr><td colspan=5 class=xl24>דוח הזמנות כל הנציגים</td></tr>")
        XLSfile.Append("<tr><td colspan=5 class=xl24> ") : XLSfile.Append(String.Format("{0:dd/MM/yy HH:mm}", Now())) : XLSfile.Append(" נכון לתאריך</td></tr>")
        If Len(dateStart) > 0 Then
            XLSfile.Append("<tr><td colspan=5 class=xl24> ") : XLSfile.Append(dateStart) : XLSfile.Append(" מתאריך</td></tr>")
        End If
        If Len(dateEnd) > 0 Then
            XLSfile.Append("<tr><td colspan=5 class=xl24> ") : XLSfile.Append(dateEnd) : XLSfile.Append(" עד תאריך</td></tr>")
        End If
        If Len(reportUserName) > 0 Then
            XLSfile.Append("<tr><td colspan=5 class=xl24> ") : XLSfile.Append(reportUserName) : XLSfile.Append(" עובד</td></tr>")
        End If

        XLSfile.Append("<tr><td colspan=5 class=xl24>&nbsp;</td></tr>")

        'Dim sqlstr As String = "SET DATEFORMAT DMY; SELECT Appeal_Id as [מספר טופס], Appeal_Date as [תאריך], " & _
        '" FIRSTNAME as [נציג], TourNum as [קוד הטיול], Docket as [מספר דוקט], CountPeople as [כמות האנשים], " & _
        '" CASE isNULL(Appeal_Id_G, 0) WHEN 0 THEN 'לא' ELSE 'כן' END [הוזן תקין] , " & _
        '" CASE Docket_Status WHEN 1 THEN 'לא קיים' WHEN 2 THEN 'דרוס תקין' WHEN 3 THEN 'תקין' WHEN 4 THEN 'בוטל' " & _
        '" ELSE 'לא נבדק' END as [סטטוס אימות] " & _
        '" FROM dbo.tbl_all_orders() O LEFT OUTER JOIN dbo.Users U ON O.User_Id = U.User_Id WHERE (1=1) "
        Dim sqlstr As String = "SET DATEFORMAT DMY; SELECT Appeal_Id as [מספר טופס], Appeal_Date as [תאריך], " & _
   " (FIRSTNAME + ' ' + LASTNAME) as [נציג],departmentName as [מחלקה], TourNum as [קוד הטיול], Docket as [מספר דוקט], CountPeople as [כמות האנשים], " & _
   " CASE isNULL(Appeal_Id_G, 0) WHEN 0 THEN 'לא' ELSE 'כן' END [הוזן תקין] , " & _
   " CASE Docket_Status WHEN 1 THEN 'לא קיים' WHEN 2 THEN 'דרוס תקין' WHEN 3 THEN 'תקין' WHEN 4 THEN 'בוטל' " & _
   " ELSE 'לא נבדק' END as [סטטוס אימות] " & _
   " FROM dbo.tbl_all_orders() O LEFT OUTER JOIN dbo.Users U ON O.User_Id = U.User_Id LEFT OUTER JOIN dbo.Departments D on D.departmentId=U.Department_Id  WHERE (1=1) "

        If Len(dateStart) > 0 Then
            sqlstr = sqlstr & " AND (datediff(dd, O.Appeal_Date, '" & dateStart & "') <=0 )"
        End If
        If Len(dateEnd) > 0 Then
            sqlstr = sqlstr & " AND (datediff(dd, O.Appeal_Date, '" & dateEnd & "') > 0 )"
        End If
        If reportUserId > 0 Then
            sqlstr = sqlstr & " AND (O.User_Id = " & reportUserId & ")"
        End If
        sqlstr = sqlstr & " ORDER BY 1"
        Response.Write(sqlstr)
        Response.End()

        cmdSelect = New SqlClient.SqlCommand(sqlstr, con)
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()
        XLSfile.Append("<tr>")
        For iCol As Integer = 0 To dr.FieldCount - 1
            XLSfile.Append("<td class=xl25>") : XLSfile.Append(dr.GetName(iCol).ToString()) : XLSfile.Append("</td>")
        Next
        XLSfile.Append("</tr>")

        Dim countGood As Integer = 0 : Dim countBad As Integer = 0
        Dim isGood As Boolean = True

        While dr.Read()
            XLSfile.Append("<tr>")
            For i As Integer = 0 To dr.FieldCount - 1
                Dim Val = func.dbNullFix(dr(i))
                XLSfile.Append("<td ")
                If (Val = "לא" And i = dr.FieldCount - 1) Then
                    XLSfile.Append("style='background:#FF99CC;'")
                    isGood = False
                Else
                    isGood = True
                End If
                XLSfile.Append("class=xl26 x:str>") : XLSfile.Append(Val) : XLSfile.Append("</td>")
            Next
            If isGood Then
                countGood = countGood + 1
            Else
                countBad = countBad + 1
            End If
            XLSfile.Append("</tr>")
        End While
        dr.Close() : con.Close()

        XLSfile.Append("<tr>")
        XLSfile.Append("<td colspan='2' class=xl26 x:str>&nbsp;סיכום של כל הרישומים התקינים</td>")
        XLSfile.Append("<td colspan='4' class=xl26 x:str>") : XLSfile.Append(countGood) : XLSfile.Append("</td>")
        XLSfile.Append("</tr>")

        XLSfile.Append("<tr>")
        XLSfile.Append("<td colspan='2' class=xl26 x:str>&nbsp;סיכום של כל הרישומים שאינם תקינים</td>")
        XLSfile.Append("<td colspan='4' class=xl26 x:str>") : XLSfile.Append(countBad) : XLSfile.Append("</td>")
        XLSfile.Append("</tr>")

        XLSfile.Append("<tr><td colspan='6' x:str></td></tr>")
        XLSfile.Append("<tr><td colspan='6' x:str>טופס חייב לעמוד בתנאים כדלקמן: </td></tr>")
        XLSfile.Append("<tr><td colspan='6' x:str>לטופס מוזן מספר דוקט תקין (בעל 6 ספרות)</td></tr>")
        XLSfile.Append("<tr><td colspan='6' x:str>מספר הדוקט לא מופיע בטופס רישום אחר</td></tr>")
        XLSfile.Append("<tr><td colspan='6' x:str>הטופס אינו במצב בוטל</td></tr>")
        XLSfile.Append("<tr><td colspan='6' x:str>לטופס מצורף קובץ PDF – זהו טופס הרישום החתום.</td></tr>")

        XLSfile.Append("</TABLE></body></html>")

        Dim filestring As String = "report_ord_" & UserId & "_" & CStr(Now.Hour) & "_" & CStr(Now.Minute) & "_" & CStr(Now.Second) & ".xls"

        Response.Clear()
        Response.AddHeader("content-disposition", "inline; filename=" & filestring)
        Response.ContentType = "application/vnd.ms-excel"
        Response.Write(XLSfile)
    End Sub

End Class
