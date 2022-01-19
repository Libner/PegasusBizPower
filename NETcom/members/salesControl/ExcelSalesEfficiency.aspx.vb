Imports System.Data.SqlClient

Public Class ExcelSalesEfficiency
    Inherits System.Web.UI.Page
    Protected sortQuery, sortQuery1, FromDate, ToDate As String
    Public XLSfile As New System.Text.StringBuilder("")
    Public qrystring As String
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmdSelect As New SqlClient.SqlCommand

    Protected dateStart, dateEnd As String
    Public dep, usr, depName, RadioType As String
    Protected pdateStart, pdateEnd As String
    '  Protected WorkDays As Integer
    Protected WorkDays, sumNumberDaysWork, SumVar1 As Decimal

    Protected SumVar16504, SumVar16470_40811_out, SumVar16470_40811_in, SumVar16735_40660, SumVar16735_40660Bitul, sumColumn5, sum16735To16504 As Integer
    Protected SumVarnumberOf16735_16470totalperUser As Integer

    Dim var1 As String
    Dim var16504 As Integer
    Dim var16470_40811_out As Integer
    Dim var16470_40811_in As Integer
    Dim var16735_40660 As Integer
    Dim var16735_40660Bitul As Integer
    Dim var16470_40105 As Integer   '-  כמות טפסי "תיעודי שיחה" יוצאת כמות שמכילה טקסטים של "אין מענה" 
    Dim var16735To16504 As Integer ''--numberOf16735To16504 	 מתוכם בכמה בוצעה מכירה column 8 
    Dim varnumberOf16735_16470totalperUser As Integer  'column 12 כמות מכירות ב"תהליך מלא


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
        '  Server.ScriptTimeout = 600
        Response.Charset = "windows-1255"
        Response.Clear()
        Response.ClearContent()
        Response.ClearHeaders()
        dateStart = Request("dateStartEx")
        If IsDate(Request("dateEndEx")) Then 'תאריך חזרה
            dateEnd = Request("dateEndEx")
        End If

        If Request("seldepEx") = "" And Request("selUserEx") = "" Then
            dep = "2"
        Else
            dep = Request("seldepEx")

        End If
        If Request("selUserEx") <> "" Then
            usr = Request("selUserEx")

        End If
        If Request("RadioTypeEx") <> "" Then
            RadioType = Request("RadioTypeEx")
        Else
            RadioType = 1
        End If
        depName = func.GetSelectDepName(dep)


        sortQuery = ""

        pdateStart = Year(dateStart) & "/" & Month(dateStart) & "/" & Day(dateStart)
        pdateEnd = Year(dateEnd) & "/" & Month(dateEnd) & "/" & Day(dateEnd)

        Dim cmdSelect = New SqlCommand("SET DATEFORMAT YMD;SELECT sum(cast(DayTypeId as  DECIMAL(9,1))) as WorkDays " & _
                      "  FROM DimDate  WHERE DateDiff(d,DateKey, convert(smalldatetime,'" & pdateStart & "',101)) <= 0  " & _
                      "  AND DateDiff(d,DateKey, convert(smalldatetime,'" & pdateEnd & "',101)) >= 0  and DayTypeId<>'0'", con)
        ''and  IsHoliday=0 
        'cmdSelect.Parameters.Add("@start_date", SqlDbType.VarChar, 10).Value = Year(dateStart) & "/" & Month(dateStart) & "/" & Day(dateStart)
        'cmdSelect.Parameters.Add("@end_date", SqlDbType.VarChar, 10).Value = Year(dateEnd) & "/" & Month(dateEnd) & "/" & Day(dateEnd)

        '''   Response.Write(cmdSelect.CommandText)
        'Response.End()
        con.Open()
        WorkDays = cmdSelect.ExecuteScalar
        'Response.Write("WorkDays=" & WorkDays)
        con.Close()







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
        XLSfile.Append("   width:150px;")
        XLSfile.Append("	border: 0.5pt solid black;")
        XLSfile.Append("	mso-pattern:auto none;")
        XLSfile.Append("	mso-height-source: auto;")
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
        XLSfile.Append("   width:150px;")
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
        XLSfile.Append("<table style='border-collapse:collapse;table-layout:fixed' dir=ltr>")
        ''XLSfile.Append("<tr><td colspan=5 class=xl24>אופרציה  - מסך עבודה</td></tr>")
        XLSfile.Append("<tr><td colspan=18 class=xl24> ") : XLSfile.Append(String.Format("{0:dd/MM/yy HH:mm}", Now())) : XLSfile.Append(" נכון לתאריך</td></tr>")

        XLSfile.Append("<tr><td colspan=18 class=xl24><span style=COLOR: #6F6DA6;font-size:14pt>")
        If RadioType = 1 Then
            XLSfile.Append("מחלקת " & depName)
        Else
            XLSfile.Append("נציגים")
        End If
        XLSfile.Append(" בין התאריכים " & dateStart & " -  " & dateEnd & "</span></td></tr>")
        XLSfile.Append("<tr><TD colspan=3 class=xl24></td>")

        XLSfile.Append("<TD class=xl24>נוכחות</td>")

        XLSfile.Append("<TD colspan=5 class=xl24>כמויות שיחות</td>")

        XLSfile.Append("<TD colspan=3 class=xl24>כמויות מתעניינים</td>")

        XLSfile.Append("<TD colspan=6 class=xl24>כמויות המכירות</td></tr>")


        XLSfile.Append("<tr>")
        
  

        XLSfile.Append("<td class=xl26>&nbsp;</td>")
        XLSfile.Append("<td class=xl26>נציג</td>")
        XLSfile.Append("<td class=xl26>מחלקה</td>")
        XLSfile.Append("<td class=xl26 width=50>כמות ימים שלא נכח בעבודה</td>")
        XLSfile.Append("<td class=xl26>כמות תיעודי שיחה נכנסת</td>")
        XLSfile.Append("<td class=xl26>כמות תיעודי שיחה יוצאת</td>")
        XLSfile.Append("<td class=xl26>כמות טפסי המתעניין</td>")
        XLSfile.Append("<td class=xl26>כמות כללית של סך השיחות</td>")
        XLSfile.Append("<td class=xl26>ממוצע שיחות יומי</td>")
        XLSfile.Append("<td class=xl26>כמות טפסי המתעניין</td>")
        XLSfile.Append("<td class=xl26>מתוכם בכמה בוצעה מכירה</td>")
        XLSfile.Append("<td class=xl26>% סגירה של הנציג</td>")
        XLSfile.Append("<td class=xl26>כמות מכירות כללית של הנציג</td>")
        XLSfile.Append("<td class=xl26>כמות לקוחות שביטלו הרשמה</td>")
        XLSfile.Append("<td class=xl26>כמות מכירות בתהליך מלא</td>")
        XLSfile.Append("<td class=xl26>% סגירה של מכירות בתהליך מלא </td>")
        XLSfile.Append("<td class=xl26>כמות מכירות בתהליך לא מלא</td>")
        XLSfile.Append("<td class=xl26>סגירה של מכירות בתהליך לא מלא</td>")
        XLSfile.Append("</tr>")
        GetdataList()

        XLSfile.Append("</TABLE></body></html>")

        Dim filestring As String = "Excel_SalesEfficiency_" & CStr(Now.Hour) & "_" & CStr(Now.Minute) & "_" & CStr(Now.Second) & ".xls"

        Response.Clear()
        Response.AddHeader("content-disposition", "inline; filename=" & filestring)
        Response.ContentType = "application/vnd.ms-excel"
        Response.Write(XLSfile)
    End Sub
    Sub GetdataList()
        ''SELECT APPEAL_ID, FIELD_ID, FIELD_VALUE FROM FORM_VALUE WHERE (FIELD_ID = 40660) 
        'AND (isnumeric(FIELD_VALUE) = 0)
        Dim s As String = ""
        Dim dr As SqlClient.SqlDataReader
        If RadioType = 2 Then
            Dim cmdSelect As New SqlClient.SqlCommand("get_SalesEfficiencyByUser_User", con)
            cmdSelect.CommandType = CommandType.StoredProcedure
            cmdSelect.Parameters.Add("@user", SqlDbType.VarChar, 200).Value = usr
            cmdSelect.Parameters.Add("@dateStart", SqlDbType.VarChar, 10).Value = pdateStart
            cmdSelect.Parameters.Add("@dateEnd", SqlDbType.VarChar, 10).Value = pdateEnd
            con.Open()
            dr = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

        Else

            Dim cmdSelect As New SqlCommand("get_SalesEfficiencyByUser", con)
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
            dr = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

        End If


           Dim i = 0

        While dr.Read()
            i = i + 1
            sumNumberDaysWork = sumNumberDaysWork + CDbl(dr("numberOfDays")) '' sum day work column1 

            var1 = CDbl(WorkDays) - CDbl(dr("numberOfDays"))
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
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(CDbl(WorkDays) - func.fixNumeric(dr("numberOfDays"))) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("numberOf16470_40811_in")) : XLSfile.Append("</td>")
            If dr("numberOf16470_40105") = 0 Then
                XLSfile.Append("<td class=xl26>&nbsp;</td>")
            Else
                '  XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("numberOf16470_40105")) : XLSfile.Append("&nbsp;/&nbsp;") : XLSfile.Append(CInt(dr("numberOf16470_40811_out")) - CInt(dr("numberOf16470_40105"))) : XLSfile.Append("</td>")
                XLSfile.Append("<td class=xl26>") : XLSfile.Append(CInt(dr("numberOf16470_40811_out")) - CInt(dr("numberOf16470_40105"))) : XLSfile.Append("&nbsp;/&nbsp;") : XLSfile.Append(dr("numberOf16470_40105")) : XLSfile.Append("</td>")

            End If


            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("numberOf16504")) : XLSfile.Append("</td>")
            ' XLSfile.Append("<td class=xl26>&nbsp;8888</td>")
            ''070419           XLSfile.Append("<td class=xl26>") : XLSfile.Append(CInt(dr("numberOf16504")) + CInt(dr("numberOf16470_40811_in")) + CInt(dr("numberOf16470_40105"))) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(CInt(dr("numberOf16504")) + CInt(dr("numberOf16470_40811_in")) + (CInt(dr("numberOf16470_40811_out")) - CInt(dr("numberOf16470_40105")))) : XLSfile.Append("</td>")



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
        If sumNumberDaysWork > 0 Then
            XLSfile.Append("<td class=xl23>") : XLSfile.Append(Math.Round(sumColumn5 / sumNumberDaysWork, 2)) : XLSfile.Append("</td>")
        Else
            XLSfile.Append("<td class=xl23></td>")
        End If
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
