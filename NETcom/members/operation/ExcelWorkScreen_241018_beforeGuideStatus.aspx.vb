Imports System.Data.SqlClient
Public Class ExcelWorkScreen
    Inherits System.Web.UI.Page
    Protected sortQuery, sortQuery1, FromDate, ToDate As String
    Public XLSfile As New System.Text.StringBuilder("")
    Public qrystring, GuidesId As String
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected WithEvents sSeries, sUsers, sGuides, sSuppliers As System.Web.UI.HtmlControls.HtmlSelect
    Protected FromDateEnd, ToDateEnd, pFromItineraryDate, pToItineraryDate, pFromBrifDate, pToBrifDate, pFromGroupMeetingDate, pToGroupMeetingDate, pFromMeetingAfterTripDate, pToMeetingAfterTripDate As String
    Protected pcodeTiul, pdateTiul, pGilboaHotel, pisSendLetterGroupMeeting, pisSendToGuide, pisSendTicket, pVoucher_Simultaneous, pVoucher_Group, pDeparture_Costing, pVouchers_Provider, pStatus, pSuppliers As String
    Protected UsersSelect, SeriesSelect, SuppliersSelect, StatusSelect, GuidesSelect, Vouchers_ProviderSelect, ShortTab As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected SeriasId, query, tabOrder As String
    Protected tab As Integer
    Protected UserId As String = "0"
    Protected tabName As String
    Protected DateMeetingAfterTrip As String



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
        FromDate = Request("sFromDateH")
        If IsDate(Request("sToDateH")) Then 'תאריך חזרה
            ToDate = Request("sToDateH")
        Else
            If pcodeTiul <> "" Then
                ToDate = DateAdd(DateInterval.Year, 3, Now())

            Else
                ToDate = ""
            End If
        End If
        If Request("codeTiulH") <> "" Then
            pcodeTiul = Request("codeTiulH")
        Else
            pcodeTiul = ""
        End If

        If Request("dateTiulH") <> "" Then
            pdateTiul = Request("dateTiulH")
        Else
            pdateTiul = ""
        End If
        '   Response.Write("date=" & Request("sFromDateH") & "<BR>")

        ' Response.Write("dss=" & Request("sisSendLetterGroupMeetingH"))
        ' Response.End()
        If Request("sisSendLetterGroupMeetingH") <> "" Then
            pisSendLetterGroupMeeting = Request("sisSendLetterGroupMeetingH")
            If pisSendLetterGroupMeeting = "כן" Then
                pisSendLetterGroupMeeting = "1"
            ElseIf pisSendLetterGroupMeeting = "לא" Then
                pisSendLetterGroupMeeting = "0"
            End If

        Else
            pisSendLetterGroupMeeting = ""
        End If
        If Request("sisSendToGuideH") <> "" Then
            pisSendToGuide = Request("sisSendToGuideH")
            If pisSendToGuide = "כן" Then
                pisSendToGuide = "1"
            ElseIf pisSendToGuide = "לא" Then
                pisSendToGuide = "0"
            End If
        Else
            pisSendToGuide = ""
        End If
        If Request("SuppliersSelectH") <> "" Then
            pSuppliers = Request("SuppliersSelectH")
        Else
            pSuppliers = ""
        End If
        If IsDate(Request("sFromMeetingAfterTripDateH")) Then 'תאריך פגישה לאחר טיול
            pFromMeetingAfterTripDate = Request("sFromMeetingAfterTripDateH")
        Else
            pFromMeetingAfterTripDate = ""
        End If
        If Request("sDeparture_CostingH") <> "" Then
            pDeparture_Costing = Request("sDeparture_CostingH")
        Else
            pDeparture_Costing = ""
        End If
        If Request("sVoucher_GroupH") <> "" Then
            pVoucher_Group = Request("sVoucher_GroupH")
        Else
            pVoucher_Group = ""
        End If
        If Request("sVoucher_SimultaneousH") <> "" Then
            pVoucher_Simultaneous = Request("sVoucher_SimultaneousH")
        Else
            pVoucher_Simultaneous = ""
        End If
        If Request("sGilboaHotelH") <> "" Then
            pGilboaHotel = Request("sGilboaHotelH")
        Else
            pGilboaHotel = ""
        End If
        If IsDate(Request("sFromGroupMeetingDateH")) Then 'תאריך  מפגש קבוצה 
            pFromGroupMeetingDate = Request("sFromGroupMeetingDateH")
        Else
            pFromGroupMeetingDate = ""
        End If
        If IsDate(Request("sToGroupMeetingDateH")) Then 'תאריך  מפגש קבוצה 
            pToGroupMeetingDate = Request("sToGroupMeetingDateH")
        Else
            pToGroupMeetingDate = ""
        End If
        If IsDate(Request("sFromBrifDateH")) Then
            pFromBrifDate = Request("sFromBrifDateH")
        Else
            pFromBrifDate = ""
        End If


        If IsDate(Request("sToBrifDateH")) Then
            pToBrifDate = Request("sToBrifDateH")
        Else
            pToBrifDate = ""
        End If
        If IsDate(Request("sFromItineraryDateH")) Then 'תאריך  קבלת Itinerary
            pFromItineraryDate = Request("sFromItineraryDateH")
        Else
            pFromItineraryDate = ""
        End If

        If IsDate(Request("sToItineraryDateH")) Then 'תאריך  קבלת Itinerary
            pToItineraryDate = Request("sToItineraryDateH")
        Else
            pToItineraryDate = ""
        End If

        If IsDate(Request("sFromDateEndH")) Then 'תאריך חזרה
            FromDateEnd = Request("sFromDateEndH")
        Else

            FromDateEnd = ""
        End If
        If IsDate(Request("sToDateEndH")) Then 'תאריך חזרה
            ToDateEnd = Request("sToDateEndH")
        Else


            ToDateEnd = ""
        End If


        If Request("sisSendTicketH") <> "" Then
            pisSendTicket = Request("sisSendTicketH")
            If pisSendTicket = "כן" Then
                pisSendTicket = "1"
            ElseIf pisSendTicket = "לא" Then
                pisSendTicket = "0"
            End If
        Else
            pisSendTicket = ""
        End If
        If Request("UsersSelectH") <> "" Then
            UserId = Request("UsersSelectH")
        End If
        If Request("SeriesSelectH") <> "" Then
            SeriasId = Request("SeriesSelectH")
        End If
        If Request("StatusSelectH") <> "" Then '?כרגע בחול
            pStatus = Request("StatusSelectH")
        Else
            pStatus = ""
        End If
        If Request("GuidesSelectH") <> "" Then
            GuidesId = Request("GuidesSelectH")
        Else
            GuidesId = ""
        End If


        'Response.Write("pisSendToGuide=" & pisSendToGuide)
        'Response.End()

        Dim func As New bizpower.cfunc
        If IsNumeric(Request("sTabH")) Then
            tab = Request("sTabH")
            '  Response.Write("tab=" & tab)
            ' Response.End()
            Select Case tab

                Case "1"
                    query = "and  DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date) < 0 and DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date_End)>= 0 "
                    tabName = "כרגע בחו""ל"
                    tabOrder = " order by Departure_Date"
                Case "2"
                    query = "and DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_DateBrief) =0 "
                    tabName = "בריפים היום"
                    tabOrder = " order by  convert(datetime,Departure_TimeBrief,8)"
                Case "3"
                    query = "and  DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date) < 0 and DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date_End)>= 0 "
                    tabName = "בחו""ל ולא התקשרו"
                    tabOrder = " order by  dbo.Tours_Departures.Departure_Date"
                    ' SELECT  dbo.Tours_Departures.Departure_Id FROM   dbo.Tours_Departures
                    '	INNER JOIN Tours ON Tours_Departures.Tour_Id = Tours.Tour_Id INNER JOIN Tours_Categories ON Tours.Category_Id = Tours_Categories.Category_Id 
                    'left JOIN Departments ON Tours.DeparureId = Departments.Dep_Id 
                    'left join bizpower_pegasus.dbo.Series BS on BS.Series_Id=Tours.SeriasId left JOIN bizpower_pegasus.dbo.Users BU on BS.User_Id=BU.User_Id
                    'left join bizpower_pegasus.dbo.GuideMessages GM on GM.Departure_Id=Tours_Departures.Departure_Id
                    ' WHERE 1=1 and  ( DATEDIFF(dd, GetDate(), Tours_Departures.Departure_Date)>=-30 or DATEDIFF(dd, GetDate(), Tours_Departures.Departure_Date) >=0)

                    'GROUP  BY dbo.Tours_Departures.Departure_Id
                    'HAVING COUNT(GM.Messages_Id) >0

                Case "4"
                    query = " and len(ISNULL(cast(Departure_DateBrief as varchar),''))=0 "
                    tabName = "לא תואם בריף החזרה"
                    tabOrder = " order by Departure_Date_End"

                Case "5"
                    query = "  and ( DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date) < 0 and DATEDIFF(dd, GetDate(), Tours_Departures.Departure_Date)>=-30 ) and  	 len(ISNULL(Departure_Costing,'') )=0 "
                    tabName = "הטיולים שיצאו ואין תמחור"
                    tabOrder = " order by Departure_Date"

                    ' Case "6"
                    '    query = " and ( DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_DateBrief) >=0 ) "
                    '   tabName = "לא תואם בריף החזרה"
                    '  tabOrder = " order by Departure_Date_End"
                Case "7"
                    'תיולי החודש הנוכחי. תאריך יציאה שלכם
                    query = "  and ( DATEDIFF(mm, GETDATE(), Tours_Departures.Departure_Date) = 0 ) "
                    tabName = "טיולי החודש"
                    tabOrder = " order by Departure_Date"
                Case "8"
                    query = " and ( DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date) < 0 and DATEDIFF(dd, GetDate(), Tours_Departures.Departure_Date)>=-30 ) and Vouchers_Provider<>'מותאם' "
                    tabName = "שובר ללא התאמה"
                    tabOrder = " order by Departure_Date"
                Case "9"
                    ' query = " and (DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_Date_End)<= 0 and len(ISNULL(cast(DateMeetingAfterTrip as varchar),''))=0 ) "
                    query = " and (len(ISNULL(cast(Departure_DateBrief as varchar),''))=0 ) "
                    tabOrder = " order by Departure_Date"

                    tabName = "לא תואם בריף"
                Case "10"
                    tabName = "רשימת בריפים קדימה"
                    query = " and DATEDIFF(dd, GETDATE(), Tours_Departures.Departure_DateBrief)>= 0 "
                    tabOrder = " order by Departure_DateBrief"

                Case "11"
                    query = " and ( DATEDIFF(dd, GETDATE(), Tours_Departures.DateMeetingAfterTrip) = 0 ) "
                    tabName = "החזרת טיולים היום"
                    tabOrder = " order by  TimeMeetingAfterTrip"
                Case Else

                    tabName = ""
            End Select

        End If

        sortQuery = ""

        sortQuery = ""



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
        XLSfile.Append("    mso-number-format:\@;")
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
        XLSfile.Append("    <x:Name>אופרציה - מסך עבודה</x:Name>")
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
        XLSfile.Append("<tr><td colspan=5 class=xl24>אופרציה  - מסך עבודה</td></tr>")
        XLSfile.Append("<tr><td colspan=5 class=xl24> ") : XLSfile.Append(String.Format("{0:dd/MM/yy HH:mm}", Now())) : XLSfile.Append(" נכון לתאריך</td></tr>")
        If tabName <> "" Then
            XLSfile.Append("<tr><td colspan=5 class=xl24>") : XLSfile.Append(tabName) : XLSfile.Append("</td></tr>")

        End If
        XLSfile.Append("<tr><td colspan=5 class=xl24>&nbsp;</td></tr>")

        XLSfile.Append("<tr>")
        XLSfile.Append("<td class=xl25>אופריטור</td>")

        XLSfile.Append("<td class=xl25>טיול</td>")
        XLSfile.Append("<td class=xl25>תאריך</td>")
        XLSfile.Append("<td class=xl25>קוד טיול</td>")
        XLSfile.Append("<td class=xl25>תאריך יציאה</td>")
        XLSfile.Append("<td class=xl25>תאריך חזרה</td>")
        XLSfile.Append("<td class=xl25>כרגע בחול?</td>")
        XLSfile.Append("<td class=xl25>שם המדריך</td>")
        XLSfile.Append("<td class=xl25>טלפון של מדריך</td>")

        XLSfile.Append("<td class=xl25>ספק</td>")
        XLSfile.Append("<td class=xl25>Itinerary תאריך קבלת</td>")
        XLSfile.Append("<td class=xl25>תאריך בריף</td>")
        XLSfile.Append("<td class=xl25>שעת בריף</td>")
        XLSfile.Append("<td class=xl25>תאריך מפגש קבוצה</td>")
        XLSfile.Append("<td class=xl25>שעת מפגש קבוצה</td>")
        XLSfile.Append("<td class=xl25>הוקלדו מלונות בגלבוע?</td>")
        XLSfile.Append("<td class=xl25>שובר לסימולטני?</td>")
        XLSfile.Append("<td class=xl25>ומדריך  קבוצה שוברי הוצאות?</td>")
        XLSfile.Append("<td class=xl25>תמחיר</td>")
        XLSfile.Append("<td class=xl25>שוברי ספק קרקע?</td>")
        XLSfile.Append("<td class=xl25>-$ כספית הערכות</td>")
        XLSfile.Append("<td class=xl25>-€ כספית הערכות</td>")
        XLSfile.Append("<td class=xl25>תאריך פגישה לאחר טיול</td>")
        XLSfile.Append("<td class=xl25>שעת פגישה לאחר טיול</td>")
        XLSfile.Append("<td class=xl25>נשלח מכתב למפגש קבוצה</td>")
        XLSfile.Append("<td class=xl25>הועבר איטינררי למדריך</td>")
        XLSfile.Append("<td class=xl25>הועברו כרטיסי טיסה</td>")
        XLSfile.Append("<td class=xl25>טיפול הסתיים?</td>")
        XLSfile.Append("<td class=xl25>כמות שיחות</td>")

        XLSfile.Append("</tr>")
        GetdataList()

        XLSfile.Append("</TABLE></body></html>")

        Dim filestring As String = "Excel_WorkScreen_" & CStr(Now.Hour) & "_" & CStr(Now.Minute) & "_" & CStr(Now.Second) & ".xls"

        Response.Clear()
        Response.AddHeader("content-disposition", "inline; filename=" & filestring)
        Response.ContentType = "application/vnd.ms-excel"
        Response.Write(XLSfile)
    End Sub
    Sub GetdataList()
      
        cmdSelect = New SqlCommand("GetWorkScreenExcel_Pages", conPegasus)
        cmdSelect.CommandType = CommandType.StoredProcedure
        cmdSelect.Parameters.Add("@SeriesId", SqlDbType.VarChar, 300).Value = SeriasId
        cmdSelect.Parameters.Add("@userId", SqlDbType.VarChar, 1000).Value = UserId
        cmdSelect.Parameters.Add("@tab", SqlDbType.Int).Value = tab
        cmdSelect.Parameters.Add("@GuideId", SqlDbType.VarChar, 300).Value = GuidesId 'sGuides ' GuidesSelect.Value
     
        cmdSelect.Parameters.Add("@SupplierId", SqlDbType.VarChar, 100).Value = pSuppliers 'sSuppliers
        cmdSelect.Parameters.Add("@pGilboaHotel", SqlDbType.VarChar, 2).Value = pGilboaHotel
        cmdSelect.Parameters.Add("@pisSendLetterGroupMeeting", SqlDbType.Char, 1).Value = pisSendLetterGroupMeeting
        cmdSelect.Parameters.Add("@pisSendToGuide", SqlDbType.Char, 1).Value = pisSendToGuide
        cmdSelect.Parameters.Add("@pisSendTicket", SqlDbType.Char, 1).Value = pisSendTicket

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
        cmdSelect.Parameters.Add("@pdateTiul", SqlDbType.VarChar, 4).Value = pdateTiul


        cmdSelect.Parameters.Add("@sortQuery", SqlDbType.VarChar, 500).Value = sortQuery


            conPegasus.Open()

        'For Each item As SqlClient.SqlParameter In cmdSelect.Parameters
        '    Response.Write(item.ParameterName & ":" & item.Value() & "<BR>")
        'Next
        'Response.Write("UserId=" & UserId & ":")
        ' Response.Write("SeriasId=" & SeriasId & ":")

        'Response.End()



        'Response.Write("UserId=" & cmdSelect.Parameters.Item("@userId").Value() & "<BR>")
        'Response.Write("@SeriesId=" & cmdSelect.Parameters.Item("@SeriesId").Value() & "<BR>")
        '  Response.Write("@StringSeriasId=" & cmdSelect.Parameters.Item("@StringSeriasId").Value() & "<BR>")
        'Response.End()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        ' If dr.Read() Then
        While dr.Read()
            XLSfile.Append("<tr>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("User_Name")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Series_Name")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.dbNullDate(dr("Departure_Date"), "MMdd")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Departure_Code")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.dbNullDate(dr("Departure_Date"), "dd/MM/yy")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.dbNullDate(dr("Departure_Date_End"), "dd/MM/yy")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.TourStatusForOperation(dr("Departure_Id"), func.dbNullFix(dr("Departure_Date")), func.dbNullFix(dr("Departure_Date_End")))) : XLSfile.Append("</td>")
            'XLSfile.Append("<td class=xl26></td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("GuideName")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Departure_GuideTelphone")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.GetSelectSupplierName(func.sFix(dr("supplier_Id")))) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26></td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.dbNullDate(dr("Departure_Itinerary"), "dd/MM/yy")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.dbNullDate(dr("Departure_DateBrief"), "dd/MM/yy")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Departure_TimeBrief")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.dbNullDate(dr("Departure_DateGroupMeeting"), "dd/MM/yy")) : XLSfile.Append("</td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Departure_TimeGroupMeeting")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("GilboaHotel")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Voucher_Simultaneous")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Voucher_Group")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Departure_Costing")) : XLSfile.Append("</td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.GetVouchers_ProviderStatus(dr("Departure_Id"))) : XLSfile.Append("</td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Financial_Dolar")) : XLSfile.Append("</td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Financial_Euro")) : XLSfile.Append("</td>")

            ' "DateMeetingAfterTrip", "{0:dd/MM/yy}")


            XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.dbNullDate(dr("DateMeetingAfterTrip"), "dd/MM/yy")) : XLSfile.Append("</td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("TimeMeetingAfterTrip")) : XLSfile.Append("</td>")



            If dr("isSendLetterGroupMeeting") = "1" Then
                XLSfile.Append("<td class=xl26>כן</td>")
            Else
                XLSfile.Append("<td class=xl26>&nbsp;</td>")
            End If

            If dr("isSendToGuide") = "1" Then
                XLSfile.Append("<td class=xl26>כן</td>")
            Else
                XLSfile.Append("<td class=xl26>&nbsp;</td>")
            End If
            If dr("isSendTicket") = "1" Then
                XLSfile.Append("<td class=xl26>כן</td>")
            Else
                XLSfile.Append("<td class=xl26>&nbsp;</td>")
            End If

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.StatusOperation(dr("Status"))) : XLSfile.Append("</td>")
            'XLSfile.Append("<td class=xl26></td>")

            'XLSfile.Append("<td class=xl26>") : XLSfile.Append(func.GetCountGuideMessages(dr("Departure_Id"))) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("CountGuideMessages")) : XLSfile.Append("</td>")

            XLSfile.Append("</tr>")
            XLSfile.Append("</tr>")

        End While
        'End If
        dr.Close()
        conPegasus.Close()

    End Sub
End Class
