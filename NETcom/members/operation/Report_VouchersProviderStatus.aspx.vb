Public Class Report_VouchersProviderStatus
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
        XLSfile.Append("    <x:Name>רשימת חשבוניות שטרם בוצע עבורם התאמה</x:Name>")
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
        XLSfile.Append("<tr><td colspan=5 class=xl24>רשימת חשבוניות שטרם בוצע עבורם התאמה </td></tr>")
        XLSfile.Append("<tr><td colspan=5 class=xl24> ") : XLSfile.Append(String.Format("{0:dd/MM/yy HH:mm}", Now())) : XLSfile.Append(" נכון לתאריך</td></tr>")

        XLSfile.Append("<tr><td colspan=5 class=xl24>&nbsp;</td></tr>")
   
        XLSfile.Append("<tr>")
        XLSfile.Append("<td class=xl25>אופריטור</td>")

        XLSfile.Append("<td class=xl25>טיול</td>")
        XLSfile.Append("<td class=xl25>תאריך טיול</td>")
        XLSfile.Append("<td class=xl25>קוד טיול</td>")
        XLSfile.Append("<td class=xl25>ספק</td>")


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
        '''If IsDate(Request.QueryString("sFromD")) Then
        '''    FromDate = Request.QueryString("sFromD")
        '''Else
        '''    FromDate = ""
        '''End If

        '''If IsDate(Request.Form("sPayToDate")) Then
        '''    ToDate = Request.Form("sPayToDate")
        '''Else
        '''    ToDate = ""
        '''End If
        '''If IsDate(Request.QueryString("sToD")) Then
        '''    ToDate = Request.QueryString("sToD")
        '''End If
        Dim conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        Dim cmdSelect As New SqlClient.SqlCommand("Report_VouchersProviderStatus", conB)

        cmdSelect.CommandType = CommandType.StoredProcedure
  
        conB.Open()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()

        While dr.Read()
            XLSfile.Append("<tr>")
          

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("UName")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Series_Name")) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(Right(dr("Departure_Code"), 4)) : XLSfile.Append("</td>")
            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("Departure_Code")) : XLSfile.Append("</td>")

            XLSfile.Append("<td class=xl26>") : XLSfile.Append(dr("supplier_Name")) : XLSfile.Append("</td>")

              XLSfile.Append("</tr>")
        End While
        dr.Close() : conB.Close()


    End Sub
End Class
