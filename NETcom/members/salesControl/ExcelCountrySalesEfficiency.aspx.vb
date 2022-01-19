Imports System.Data.SqlClient
Public Class ExcelCountrySalesEfficiency
    Inherits System.Web.UI.Page
    Protected sortQuery, sortQuery1, FromDate, ToDate As String
    Public XLSfile As New System.Text.StringBuilder("")
    Public qrystring As String
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim con1 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmdSelect As New SqlClient.SqlCommand

    Protected dateStart, dateEnd As String
    Public dep, usr, depName, RadioType, country As String
    Protected pdateStart, pdateEnd As String
    Protected WorkDays As Integer
    Protected sumNumberDaysWork, SumVar1, SumVar16504, SumVar16470_40811_out, SumVar16470_40811_in, SumVar16735_40660, SumVar16735_40660Bitul, sumColumn5, sum16735To16504 As Integer
    Protected SumVarnumberOf16735_16470totalperUser As Integer
    Dim departmentName As String
    Dim var1 As String
    Dim var16504 As Integer
    Dim var16470_40811_out As Integer
    Dim var16470_40811_in As Integer
    Dim var16735_40660 As Integer
    Dim var16735_40660Bitul As Integer
    Dim var16470_40105 As Integer   '-  כמות טפסי "תיעודי שיחה" יוצאת כמות שמכילה טקסטים של "אין מענה" 
    Dim var16735To16504 As Integer ''--numberOf16735To16504 	 מתוכם בכמה בוצעה מכירה column 8 
    Dim varnumberOf16735_16470totalperUser As Integer  'column 12 כמות מכירות ב"תהליך מלא

    Dim cls As String

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
        cls = "cl1"
        Response.Buffer = False
        '  Server.ScriptTimeout = 600
        Response.Charset = "windows-1255"
        Response.Clear()
        Response.ClearContent()
        Response.ClearHeaders()
        dateStart = Request("dateStartEx")
        If IsDate(Request("dateEndEx")) Then 'תאריך חזרה
            dateEnd = Request("dateEndEx")
        End If
        'Response.Write("selCountryEx=" & Request("selCountryEx") & ":" & Left(Request("selCountryEx"), Request("selCountryEx").Length - 1) & "<BR>")
        'Response.Write("dateStartEx=" & Request("dateStartEx") & "<BR>")
        'Response.Write("dateEndEx=" & Request("dateEndEx") & "<BR>")
        'Response.Write("RadioTypeEx=" & Request("RadioTypeEx") & "<BR>")
        'Response.Write("seldepEx=" & Request("seldepEx") & "<BR>")
        'Response.Write("selUserEx=" & Request("selUserEx") & "<BR>")
        'Response.End()

        If Request("seldepEx") = "" And Request("selUserEx") = "" Then
            dep = "2"
        Else
            dep = Request("seldepEx")
            dep = Left(dep, dep.Length - 1)

        End If
        If Request("selUserEx") <> "" Then
            usr = Request("selUserEx")
            usr = Left(usr, usr.Length - 1)
        End If
        If Request("selCountryEx") <> "" Then
            country = Request("selCountryEx")
            country = Left(country, country.Length - 1)
        End If
        '  Response.Write("country=" & country)
        '  Response.End()
        If Request("RadioTypeEx") <> "" Then
            RadioType = Request("RadioTypeEx")
        Else
            RadioType = 1
        End If
        depName = func.GetSelectDepName(dep)


        sortQuery = ""

        pdateStart = Year(dateStart) & "/" & Month(dateStart) & "/" & Day(dateStart)
        pdateEnd = Year(dateEnd) & "/" & Month(dateEnd) & "/" & Day(dateEnd)




        'Dim cmdSelect = New SqlCommand("SET DATEFORMAT YMD;SELECT count(id) as WorkDays " & _
        '              "  FROM DimDate  WHERE DateDiff(d,DateKey, convert(smalldatetime,'" & pdateStart & "',101)) <= 0  " & _
        '              "  AND DateDiff(d,DateKey, convert(smalldatetime,'" & pdateEnd & "',101)) >= 0  and DayTypeId<>'0'", con)
        '''and  IsHoliday=0 
        ''cmdSelect.Parameters.Add("@start_date", SqlDbType.VarChar, 10).Value = Year(dateStart) & "/" & Month(dateStart) & "/" & Day(dateStart)
        ''cmdSelect.Parameters.Add("@end_date", SqlDbType.VarChar, 10).Value = Year(dateEnd) & "/" & Month(dateEnd) & "/" & Day(dateEnd)

        ''''   Response.Write(cmdSelect.CommandText)
        ''Response.End()
        'con.Open()
        'WorkDays = cmdSelect.ExecuteScalar
        ''Response.Write("WorkDays=" & WorkDays)
        'con.Close()







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
        XLSfile.Append("	font-size:10.0pt;")
        XLSfile.Append("	font-weight:700;")
        XLSfile.Append("	text-align:center;")
        XLSfile.Append("	background:#C0C0C0;")
        XLSfile.Append("	mso-height-source: userset;")
        XLSfile.Append("	border: 0.5pt solid black;")
        XLSfile.Append("    height:20pt;")
        XLSfile.Append("	white-space:normal;}")
        XLSfile.Append(".cl1")
        XLSfile.Append("	{")
        XLSfile.Append("	color:#000000;")
        XLSfile.Append("	background:#e1e1e1;)}")

        XLSfile.Append(".cl2")
        XLSfile.Append("	{")
        XLSfile.Append("	color:#000000;")
        XLSfile.Append("	background:#a0a0a0;)}")

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

        XLSfile.Append(".xl26") ' title cell
        XLSfile.Append("	{mso-style-parent:style0;")
        XLSfile.Append("	color:#000000;")
        XLSfile.Append("	font-size:10.0pt;")
        XLSfile.Append("	margin-right:5pt;")
        XLSfile.Append("	margin-left:5pt;")
        XLSfile.Append("	font-weight:400;")
        XLSfile.Append("	font-family:Arial, sans-serif;")
        XLSfile.Append("	text-align:center;")
        XLSfile.Append("	background:#e1e1e1;")
        XLSfile.Append("   width:60px;")
        XLSfile.Append("	border: 0.5pt solid black;")
        XLSfile.Append("	mso-pattern:auto none;")
        XLSfile.Append("	mso-height-source: auto;")
        XLSfile.Append(" white-space:normal;}")
        XLSfile.Append(".xl26S") ' title cell
        XLSfile.Append("	{mso-style-parent:style0;")
        XLSfile.Append("	color:#000000;")
        XLSfile.Append("	font-size:8.0pt;")
        XLSfile.Append("	margin-right:1pt;")
        XLSfile.Append("	margin-left:1pt;")
        XLSfile.Append("	font-weight:400;")
        XLSfile.Append("	font-family:Arial, sans-serif;")
        XLSfile.Append("	text-align:center;")
        XLSfile.Append("	background:#e1e1e1;")
        XLSfile.Append("    width:40px;")
        XLSfile.Append("	border: 0.5pt solid black;")
        XLSfile.Append("	mso-pattern:auto none;")
        XLSfile.Append("	mso-height-source: auto;")
        XLSfile.Append(" white-space:normal;}")
        XLSfile.Append(".xl26Country") ' title cell
        XLSfile.Append("	{mso-style-parent:style0;")
        XLSfile.Append("	color:#660000;")
        XLSfile.Append("	font-size:10.0pt;")
        XLSfile.Append("	margin-right:1pt;")
        XLSfile.Append("	margin-left:1pt;")
        XLSfile.Append("	font-weight:400;")
        XLSfile.Append("	font-family:Arial, sans-serif;")
        XLSfile.Append("	text-align:center;")
        XLSfile.Append("	background:#e0e0eb;")
        XLSfile.Append("    width:40px;")
        XLSfile.Append("	border: 0.5pt solid black;")
        XLSfile.Append("	mso-pattern:auto none;")
        XLSfile.Append("	mso-height-source: auto;")
        XLSfile.Append("	border: 1pt solid #217346;")
        XLSfile.Append(" white-space:normal;}")
        XLSfile.Append(".xl23") ' title cell
        XLSfile.Append("	{mso-style-parent:style0;")
        XLSfile.Append("	color:#000000;")
        XLSfile.Append("	background:#ffd011;")
        XLSfile.Append("	font-size:10.0pt;")
        XLSfile.Append("	margin-right:5pt;")
        XLSfile.Append("	margin-left:5pt;")
        XLSfile.Append("	font-weight:400;")
        XLSfile.Append("	font-family:Arial, sans-serif;")
        XLSfile.Append("	text-align:center;")
        XLSfile.Append("   width:60px;")
        XLSfile.Append("	border: 0.5pt solid black;")
        XLSfile.Append("	mso-pattern:auto none;")
        XLSfile.Append("	mso-height-source: auto;")
        XLSfile.Append(" white-space:normal;}")


        XLSfile.Append("-->")
        XLSfile.Append("</style>")
        XLSfile.Append("<!--[if gte mso 9]><xml>")
        XLSfile.Append("<x:ExcelWorkbook>")
        XLSfile.Append("<x:ExcelWorksheets>")
        XLSfile.Append("<x:ExcelWorksheet>")
        XLSfile.Append("<x:Name>נצילות נציג מכירות</x:Name>")
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
        XLSfile.Append("<table dir=ltr width=100%>")
        ''XLSfile.Append("<tr><td colspan=5 class=xl24>אופרציה  - מסך עבודה</td></tr>")
        XLSfile.Append("<tr><td  class=xl24 colspan=15>  ") : XLSfile.Append(String.Format("{0:dd/MM/yy HH:mm}", Now())) : XLSfile.Append(" נכון לתאריך</td></tr>")

        XLSfile.Append("<tr><td  class=xl24  colspan=15><span style=COLOR: #6F6DA6;font-size:14pt>")
        If RadioType = 1 Then
            XLSfile.Append("מחלקת " & depName)
        Else
            XLSfile.Append("נציגים")
        End If
        XLSfile.Append("</tr></table><BR clear=all>")
        XLSfile.Append("<table dir=ltr style='boder:1px solid #ff0000'>")

        XLSfile.Append("<tr>")
  
        XLSfile.Append("<TD class=xl26Country>&nbsp;</td><td class=xl26Country>&nbsp;</td><TD class=xl26Country>&nbsp;</td>")
        Dim cmdSelect = New SqlCommand("select Country_Name,Country_Id From dbo.Countries  where Country_Id in (" & country & ") order by Country_Name", con)
        con.Open()
        Dim i = 0
        Dim j As Integer
        Dim pRow = 0

        Dim rs_r = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        While rs_r.Read()
            i = i + 1
            XLSfile.Append("<td colspan=7 class=xl26Country  align=center>") : XLSfile.Append(rs_r("Country_Name")) : XLSfile.Append("</tD>")
        End While
        con.Close()
        XLSfile.Append("</tr>")
        XLSfile.Append("<tr><td class=xl26S>&nbsp;</td>")
        XLSfile.Append("<td class=xl26S>נציג</td>")
        XLSfile.Append("<td class=xl26S>מחלקה</td>")
        '  Response.Write("i=" & i)
            For j = 1 To i
            XLSfile.Append("<td class=xl26S>כמות <BR>מכירות כללית<BR> של הנציג</td>")
            XLSfile.Append("<td class=xl26S>כמות <BR>לקוחות שביטלו<BR> הרשמה</td>")
            XLSfile.Append("<td class=xl26S>כמות<BR> מכירות בתהליך<BR> מלא</td>")
            XLSfile.Append("<td class=xl26S>% סגירה<BR> של מכירות<BR> בתהליך מלא </td>")
            XLSfile.Append("<td class=xl26S>כמות <BR>מכירות בתהליך <BR>לא מלא</td>")
            XLSfile.Append("<td class=xl26S>סגירה<BR> של מכירות <BR>בתהליך לא מלא</td>")
            XLSfile.Append("<td class=xl26S>כמות <BR>טפסי <BR>המתעניין</td>")
            '
            'Response.Write("j=" & j & "<BR>")
        Next
        '   Response.End()

        XLSfile.Append("</tr>")
        XLSfile.Append("<tr><TD colspan=3><table border=1>")
        Dim sqlUser As String
        If RadioType = 2 Then
            '   Response.Write("Select USER_ID, FIRSTNAME + char(32) + LASTNAME as UserName,departmentName  from Users left  join Departments on Departments.DepartmentId=Users.Department_Id where  ACTIVE = 1 and  USER_ID in (" & usr & ") &  order by FIRSTNAME,LASTNAME")
            sqlUser = "Select USER_ID, FIRSTNAME + char(32) + LASTNAME as UserName,departmentName  from Users left  join Departments on Departments.DepartmentId=Users.Department_Id where  ACTIVE = 1 and  USER_ID in (" & usr & ")   order by DepartmentName,FIRSTNAME + Char(32) + LASTNAME"
        Else
            sqlUser = "Select USER_ID, FIRSTNAME + char(32) + LASTNAME as UserName,departmentName  from Users left  join Departments on Departments.DepartmentId=Users.Department_Id where  ACTIVE = 1 and  Department_Id in (" & dep & ") order by DepartmentName,FIRSTNAME + Char(32) + LASTNAME"
            '   Response.Write("Select USER_ID, FIRSTNAME + char(32) + LASTNAME as UserName,departmentName  from Users left  join Departments on Departments.DepartmentId=Users.Department_Id where  ACTIVE = 1 and  Department_Id in (" & dep & ")   order by FIRSTNAME,LASTNAME")
        End If
        Dim cmdSelectUser = New SqlCommand(sqlUser, con)
        con.Open()
        '   Response.End()
        Dim rs_Users = cmdSelectUser.ExecuteReader(CommandBehavior.CloseConnection)
        While rs_Users.Read()
            pRow = pRow + 1
            If IsDBNull(rs_Users("departmentName")) Then
                departmentName = ""
            Else
                departmentName = rs_Users("departmentName")

            End If

            XLSfile.Append("<tr><TD>") : XLSfile.Append(pRow) : XLSfile.Append("</td><td>") : XLSfile.Append(rs_Users("UserName")) : XLSfile.Append("</td><td>") : XLSfile.Append(departmentName) : XLSfile.Append("</td>")

        End While
        con.Close()
        XLSfile.Append("</tr></table></td>")
        '''columns by country
        Dim counryId As Integer
        cmdSelect = New SqlCommand("select Country_Name,Country_Id From dbo.Countries  where Country_Id in (" & country & ") order by Country_Name", con)
        con.Open()
        rs_r = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        While rs_r.Read()
            counryId = rs_r("Country_Id")
            getDataByCountry(counryId)
        End While
        con.Close()
        '''end column by country


        con.Close()
        XLSfile.Append("</table></td><TD colspan=10></tD></tr>")
        '   Response.Write(XLSfile)
        '   Response.End()

        '   GetdataList()
        XLSfile.Append("</table></td></tr>")
        XLSfile.Append("</TABLE></body></html>")
        '   Response.Write(XLSfile)
        '  Response.End()
        Dim filestring As String = "Excel_SalesEfficiency_" & CStr(Now.Hour) & "_" & CStr(Now.Minute) & "_" & CStr(Now.Second) & ".xls"

        Response.Clear()
        Response.AddHeader("content-disposition", "inline; filename=" & filestring)
        Response.ContentType = "application/vnd.ms-excel"
        Response.Write(XLSfile)
    End Sub
    Sub getDataByCountry(ByVal CountryId)
        If cls = "cl1" Then
            cls = "cl2"
        Else
            cls = "cl1"
        End If
        If RadioType = 2 Then
            cmdSelect = New SqlCommand("get_SalesEfficiencyByUserUserCountry", con1)
            cmdSelect.CommandType = CommandType.StoredProcedure
            cmdSelect.Parameters.Add("@CountryId", SqlDbType.VarChar, 20).Value = CountryId
            cmdSelect.Parameters.Add("@user", SqlDbType.VarChar, 200).Value = usr
            cmdSelect.Parameters.Add("@dateStart", SqlDbType.VarChar, 10).Value = pdateStart
            cmdSelect.Parameters.Add("@dateEnd", SqlDbType.VarChar, 10).Value = pdateEnd
            con1.Open()
            Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            Dim i = 0
            Dim field5 = 0
            Dim field6 = 0 '%בתהליך  לא מלא		(numberOf16735_40660-numberOf16735_16470totalperUser)/numberOf16735_40660)*100
            Dim field4 = 0 '	%בתהליך מלא		  (numberOf16735_16470totalperUser/numberOf16735_40660)*100




            XLSfile.Append("<td colspan=7 width=100%><table border=1 width=100%>")

            While dr.Read()
                field5 = CInt(dr("numberOf16735_40660")) - CInt(dr("numberOf16735_16470totalperUser"))

                If field5 > 0 And dr("numberOf16735_40660") <> 0 Then
                    field6 = Math.Round(CInt(field5) / CInt(dr("numberOf16735_40660")), 2) * 100 & "%"
                Else
                    field6 = ""
                End If
                If dr("numberOf16735_40660") > 0 Then
                    field4 = Math.Round(CInt(dr("numberOf16735_16470totalperUser")) / CInt(dr("numberOf16735_40660")), 2) * 100 & "%"
                End If






                '  XLSfile.Append("<TR><TD>") : XLSfile.Append(dr("numberOf16735_40660")) : XLSfile.Append("</td><TD>") : XLSfile.Append(dr("numberOf16735_40660Bitul")) : XLSfile.Append("</td><TD>") : XLSfile.Append(dr("numberOf16735_16470totalperUser")) : XLSfile.Append("</td><TD>") : XLSfile.Append(field4) : XLSfile.Append("</td><TD>") : XLSfile.Append(field5) : XLSfile.Append("</td><TD>") : XLSfile.Append(field6) : XLSfile.Append("</td><TD>") : XLSfile.Append(dr("numberOf16504")) : XLSfile.Append("</td></tr>")
                XLSfile.Append("<TR><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(dr("numberOf16735_40660")) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(dr("numberOf16735_40660Bitul")) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(dr("numberOf16735_16470totalperUser")) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(field4) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(field5) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(field6) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(dr("numberOf16504")) : XLSfile.Append("</td></tr>")

            End While

            con1.Close()
            XLSfile.Append("</table></td>")


        Else
            cmdSelect = New SqlCommand("get_SalesEfficiencyByUserCountry", con1)
            cmdSelect.CommandType = CommandType.StoredProcedure
            cmdSelect.Parameters.Add("@CountryId", SqlDbType.VarChar, 20).Value = CountryId
            cmdSelect.Parameters.Add("@depId", SqlDbType.VarChar, 20).Value = dep
            cmdSelect.Parameters.Add("@dateStart", SqlDbType.VarChar, 10).Value = pdateStart
            cmdSelect.Parameters.Add("@dateEnd", SqlDbType.VarChar, 10).Value = pdateEnd
            con1.Open()
            Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            Dim i = 0
            Dim field5 = 0
            Dim field6 = 0 '%בתהליך  לא מלא		(numberOf16735_40660-numberOf16735_16470totalperUser)/numberOf16735_40660)*100
            Dim field4 = 0 '	%בתהליך מלא		  (numberOf16735_16470totalperUser/numberOf16735_40660)*100




            XLSfile.Append("<td colspan=7 width=100%><table border=1 width=100%>")

            While dr.Read()
                field5 = CInt(dr("numberOf16735_40660")) - CInt(dr("numberOf16735_16470totalperUser"))

                If field5 > 0 And dr("numberOf16735_40660") <> 0 Then
                    field6 = Math.Round(CInt(field5) / CInt(dr("numberOf16735_40660")), 2) * 100 & "%"
                Else
                    field6 = ""
                End If
                If dr("numberOf16735_40660") > 0 Then
                    field4 = Math.Round(CInt(dr("numberOf16735_16470totalperUser")) / CInt(dr("numberOf16735_40660")), 2) * 100 & "%"
                Else
                    field4 = ""
                End If





                XLSfile.Append("<TR><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(dr("numberOf16735_40660")) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(dr("numberOf16735_40660Bitul")) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(dr("numberOf16735_16470totalperUser")) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(field4) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(field5) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(field6) : XLSfile.Append("</td><TD class=") : XLSfile.Append(cls) : XLSfile.Append(">") : XLSfile.Append(dr("numberOf16504")) : XLSfile.Append("</td></tr>")

            End While

            con1.Close()
            XLSfile.Append("</table></td>")
        End If

    End Sub

    Sub GetdataList()

        Dim s As String = ""

        If RadioType = 2 Then
            Dim cmdSelectProc As New SqlClient.SqlCommand("get_SalesEfficiencyByUser_User", con)
            cmdSelectProc.CommandType = CommandType.StoredProcedure
            cmdSelectProc.Parameters.Add("@user", SqlDbType.VarChar, 200).Value = usr
            cmdSelectProc.Parameters.Add("@dateStart", SqlDbType.VarChar, 10).Value = pdateStart
            cmdSelectProc.Parameters.Add("@dateEnd", SqlDbType.VarChar, 10).Value = pdateEnd
            con.Open()
        Else

            cmdSelect = New SqlCommand("get_SalesEfficiencyByUser", con)
            cmdSelect.CommandType = CommandType.StoredProcedure
            cmdSelect.Parameters.Add("@depId", SqlDbType.VarChar, 20).Value = dep
            cmdSelect.Parameters.Add("@dateStart", SqlDbType.VarChar, 10).Value = pdateStart
            cmdSelect.Parameters.Add("@dateEnd", SqlDbType.VarChar, 10).Value = pdateEnd
            con.Open()

            'For Each item As SqlClient.SqlParameter In cmdSelect.Parameters
            '    Response.Write(item.ParameterName & ":" & item.Value() & "<BR>")
            'Next
            'Response.Write("UserId=" & UserId & ":")
            ' Response.Write("SeriasId=" & SeriasId & ":")

            'Response.End()
        End If


        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        Dim i = 0

        While dr.Read()
            i = i + 1
            sumNumberDaysWork = sumNumberDaysWork + CInt(dr("numberOfDays")) '' sum day work column1 

            var1 = CInt(WorkDays) - CInt(dr("numberOfDays"))
            SumVar1 = SumVar1 + var1
            var16504 = dr("numberOf16504")
            SumVar16504 = SumVar16504 + var16504

            var16470_40811_in = dr("numberOf16470_40811_in")
            SumVar16470_40811_in = SumVar16470_40811_in + var16470_40811_in
            var16470_40811_out = dr("numberOf16470_40811_out")
            SumVar16470_40811_out = SumVar16470_40811_out + var16470_40811_out
            If IsNumeric(dr("numberOf16470_40105")) Then
                var16470_40105 = dr("numberOf16470_40105")
            Else
                var16470_40105 = 0
            End If

            If IsNumeric(dr("numberOf16735_40660")) Then
                var16735_40660 = dr("numberOf16735_40660")
                SumVar16735_40660 = SumVar16735_40660 + var16735_40660
            End If

            If IsNumeric(dr("numberOf16735_40660Bitul")) Then
                var16735_40660Bitul = dr("numberOf16735_40660Bitul")
                SumVar16735_40660Bitul = SumVar16735_40660Bitul + var16735_40660Bitul
            End If




            sumColumn5 = sumColumn5 + var16504 + var16470_40811_in + var16470_40105
            If IsNumeric(dr("numberOf16735To16504")) Then
                var16735To16504 = dr("numberOf16735To16504")
                sum16735To16504 = sum16735To16504 + var16735To16504
            End If
            If IsNumeric(dr("numberOf16735_16470totalperUser")) Then
                varnumberOf16735_16470totalperUser = dr("numberOf16735_16470totalperUser")
                SumVarnumberOf16735_16470totalperUser = SumVarnumberOf16735_16470totalperUser + varnumberOf16735_16470totalperUser
            End If


            XLSfile.Append("<tr>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(i) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("User_Name")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("departmentName")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(CInt(WorkDays) - dr("numberOfDays")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("numberOf16470_40811_in")) : XLSfile.Append("</td>")
            If dr("numberOf16470_40105") = 0 Then
                XLSfile.Append("<td class=xl26>&nbsp;</td>")
            Else
                XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("numberOf16470_40105")) : XLSfile.Append("&nbsp;/&nbsp;") : XLSfile.Append(CInt(dr("numberOf16470_40811_out")) - CInt(dr("numberOf16470_40105"))) : XLSfile.Append("</td>")
            End If


            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("numberOf16504")) : XLSfile.Append("</td>")
            ' XLSfile.Append("<td class=xl26>&nbsp;8888</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(CInt(dr("numberOf16504")) + CInt(dr("numberOf16470_40811_in")) + CInt(dr("numberOf16470_40105"))) : XLSfile.Append("</td>")
            If dr("numberOfDays") > 0 Then
                XLSfile.Append("<td class=xl26>") : XLSfile.Append(Math.Round((CInt(dr("numberOf16504")) + CInt(dr("numberOf16470_40811_in")) + CInt(dr("numberOf16470_40105"))) / dr("numberOfDays"), 2)) : XLSfile.Append("</td>")
            Else
                XLSfile.Append("<td class=xl26>&nbsp;</td>")
            End If
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("numberOf16504")) : XLSfile.Append("</td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("numberOf16735To16504")) : XLSfile.Append("</td>")

            If dr("numberOf16504") > 0 Then
                XLSfile.Append("<td class=xl26>") : XLSfile.Append(Math.Round(CInt(dr("numberOf16735To16504")) / CInt(dr("numberOf16504")), 2) * 100) : XLSfile.Append("%</td>")
            Else
                XLSfile.Append("<td class=xl26>&nbsp;</td>")
            End If
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("numberOf16735_40660")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("numberOf16735_40660Bitul")) : XLSfile.Append("</td>")


            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("numberOf16735_16470totalperUser")) : XLSfile.Append("</td>")
            If dr("numberOf16735_40660") > 0 Then
                XLSfile.Append("<td class=xl26>") : XLSfile.Append(Math.Round(CInt(dr("numberOf16735_16470totalperUser")) / CInt(dr("numberOf16735_40660")), 2) * 100) : XLSfile.Append("%</td>")
            Else
                XLSfile.Append("<td class=xl26>&nbsp;</td>")
            End If

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(CInt(dr("numberOf16735_40660")) - CInt(dr("numberOf16735_16470totalperUser"))) : XLSfile.Append("</td>")
            If CInt(dr("numberOf16735_40660")) - CInt(dr("numberOf16735_16470totalperUser")) > 0 Then
                XLSfile.Append("<td class=xl26>") : XLSfile.Append(Math.Round((CInt(dr("numberOf16735_40660")) - CInt(dr("numberOf16735_16470totalperUser"))) / CInt(dr("numberOf16735_40660")), 2) * 100) : XLSfile.Append("%</td>")
            Else
                XLSfile.Append("<td class=xl26>&nbsp;</td>")
            End If

            XLSfile.Append("</tr>")


        End While
        'End If
        dr.Close()
        con.Close()
        XLSfile.Append("<tr style=height:30px>")

        XLSfile.Append("<td class=xl23>&nbsp;</td>")
        XLSfile.Append("<td class=xl23>&nbsp;</td>")
        XLSfile.Append("<td class=xl23>&nbsp;</td>")
        XLSfile.Append("<td class=xl23>") : XLSfile.Append(SumVar1) : XLSfile.Append("</td>")
        XLSfile.Append("<td class=xl23>") : XLSfile.Append(SumVar16470_40811_in) : XLSfile.Append("</td>")
        XLSfile.Append("<td class=xl23>") : XLSfile.Append(SumVar16470_40811_out) : XLSfile.Append("</td>")
        XLSfile.Append("<td class=xl23>") : XLSfile.Append(SumVar16504) : XLSfile.Append("</td>")
        XLSfile.Append("<td class=xl23>") : XLSfile.Append(sumColumn5) : XLSfile.Append("</td>")
        XLSfile.Append("<td class=xl23>") : XLSfile.Append(Math.Round(sumColumn5 / sumNumberDaysWork, 2)) : XLSfile.Append("</td>")
        XLSfile.Append("<td class=xl23>") : XLSfile.Append(SumVar16504) : XLSfile.Append("</td>")
        XLSfile.Append("<td class=xl23>") : XLSfile.Append(sum16735To16504) : XLSfile.Append("</td>")
        If SumVar16504 > 0 Then
            XLSfile.Append("<td class=xl23>") : XLSfile.Append(Math.Round(sum16735To16504 / SumVar16504, 2) * 100) : XLSfile.Append("%</td>")
        Else
            XLSfile.Append("<td class=xl23>&nbsp;</td>")
        End If

        XLSfile.Append("<td class=xl23>") : XLSfile.Append(SumVar16735_40660) : XLSfile.Append("</td>")
        XLSfile.Append("<td class=xl23>") : XLSfile.Append(SumVar16735_40660Bitul) : XLSfile.Append("</td>")
        XLSfile.Append("<td class=xl23>") : XLSfile.Append(SumVarnumberOf16735_16470totalperUser) : XLSfile.Append("</td>")
        If SumVar16735_40660 > 0 Then
            XLSfile.Append("<td class=xl23>") : XLSfile.Append(Math.Round(SumVarnumberOf16735_16470totalperUser / SumVar16735_40660, 2) * 100) : XLSfile.Append("%</td>")
        Else
            XLSfile.Append("<td class=xl23>&nbsp;</td>")
        End If

        XLSfile.Append("<td class=xl23>") : XLSfile.Append(SumVar16735_40660 - SumVarnumberOf16735_16470totalperUser) : XLSfile.Append("</td>")
        If SumVar16735_40660 > 0 Then
            XLSfile.Append("<td class=xl23>") : XLSfile.Append(Math.Round((SumVar16735_40660 - SumVarnumberOf16735_16470totalperUser) / SumVar16735_40660, 2) * 100) : XLSfile.Append("%</td>")
        Else
            XLSfile.Append("<td class=xl23>&nbsp;</td>")
        End If



        XLSfile.Append("</tr>")

    End Sub
End Class
