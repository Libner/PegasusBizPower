Imports System.Drawing
Imports WebChart
Public Class ifrmInsureChartWithReasons
    Inherits System.Web.UI.Page
    Dim func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected sdate, enddate As String


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
        If IsDate(Request("dateStart")) Then
            sdate = Request("dateStart")
        Else
            sdate = Now.AddDays(-7).ToShortDateString()
        End If
        If IsDate(Request.QueryString("dateEnd")) Then
            enddate = Request.QueryString("dateEnd")
        Else
            enddate = Now().ToShortDateString()
        End If

        Dim wcEng As New ChartEngine
        wcEng.Size = New Size(1800, 550)
        Dim wcCharts As New ChartCollection(wcEng)
        wcEng.Charts = wcCharts
        cmdSelect = New SqlClient.SqlCommand("getInsureReport", con)
        'SELECT  top 40 FIRSTNAME + CHAR(32) + LASTNAME as LASTNAME,  InsuranceSend_Counter as YY,dbo.GetInsuranseCountNotSend(USER_ID) as NN FROM  USERS
        ' WHERE     (Add_Insurance = 1)
        cmdSelect.CommandType = CommandType.StoredProcedure
        cmdSelect.Parameters.Add("@dateStart", SqlDbType.Char, 10).Value = sdate
        cmdSelect.Parameters.Add("@dateEnd", SqlDbType.Char, 10).Value = enddate

        con.Open()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()
        Dim chart1 As New ColumnChart
        chart1.Fill.Color = Color.FromArgb(80, 163, 195, 18) : chart1.Line.Color = Color.FromArgb(80, 163, 195, 18)
        chart1.DataLabels.Font = New System.Drawing.Font("Arial", 10, FontStyle.Regular)
        chart1.Legend = "נשלח" : chart1.ShowLegend = True
        chart1.DataLabels.Visible = True : chart1.Shadow.Visible = False


        Dim chart2 As New ColumnChart
        chart2.Line.Color = Color.FromArgb(186, 186, 0, 0)
        chart2.Legend = "לא נשלח" : chart2.ShowLegend = True
        chart2.MaxColumnWidth = "10"
        chart2.DataLabels.Visible = True : chart2.Shadow.Visible = False
        chart2.Line.Width = "10"
        While dr.Read()


            '    AddSeries()
            chart1.Data.Add(New ChartPoint(dr("LASTNAME").ToString, CType(dr("YY"), Integer)))
            chart2.Data.Add(New ChartPoint(dr("LASTNAME").ToString, CType(dr("NN"), Integer)))


            'Dim point As New ChartPoint
            'point.XValue = dr("LASTNAME")
            'point.YValue = dr("NN")
            'chart1.Data.Add(point)



            '    ' wcEng.XTitle.Text = dr("LASTNAME").ToString
            '    'chart1.Data.Add(New ChartPoint(dr("LASTNAME").ToString, CType(dr("תקין"), Integer)))
            '    'chart2.Data.Add(New ChartPoint(dr("FIRSTNAME").ToString, CType(dr("לא תקין"), Integer)))
            '    'chart3.Data.Add(New ChartPoint(dr("FIRSTNAME").ToString, CType(dr("בוטל"), Integer)))
        End While
        ''Response.Write(dr.HasRows)
        dr.Close() : con.Close()
        wcCharts.Add(chart1)
        wcCharts.Add(chart2)



        SetMoreProperties(wcEng)
        Dim objImageCodecInfo As System.Drawing.Imaging.ImageCodecInfo
        Dim objEncoder0 As System.Drawing.Imaging.Encoder
        Dim objEncoderParameter0 As System.Drawing.Imaging.EncoderParameter
        Dim objEncoderParameters As System.Drawing.Imaging.EncoderParameters

        'Get an ImageCodecInfo object that represents the PNG codec. 
        objImageCodecInfo = GetEncoderInfo("image/png")

        objEncoderParameters = New System.Drawing.Imaging.EncoderParameters(1)

        'Create an Encoder object based on the GUID 
        'for the Quality parameter category. 
        objEncoder0 = System.Drawing.Imaging.Encoder.Quality

        'Save the bitmap as a JPEG file with quality level 92. 
        objEncoderParameter0 = New System.Drawing.Imaging.EncoderParameter(objEncoder0, 92L)
        objEncoderParameters.Param(0) = objEncoderParameter0

        Dim bmp As Bitmap
        Dim memStream As New System.IO.MemoryStream
        bmp = wcEng.GetBitmap()
        bmp.SetResolution(72, 72)
        bmp.Save(memStream, objImageCodecInfo, objEncoderParameters)
        memStream.WriteTo(Response.OutputStream)
        Response.End()

    End Sub
    Private Sub SetMoreProperties(ByVal engine As ChartEngine)
        engine.HasChartLegend = True : engine.Legend.Position = LegendPosition.Bottom : engine.Legend.Width = 30

        ' Set-up the XTitle
        Dim myFont As New System.Drawing.Font("Arial", 10, FontStyle.Regular)
        engine.XTitle = New ChartText
        Dim horizontalFormat As New StringFormat
        horizontalFormat.Trimming = StringTrimming.Character
        horizontalFormat.FormatFlags = StringFormatFlags.DirectionVertical
        horizontalFormat.LineAlignment = StringAlignment.Center
        horizontalFormat.Alignment = StringAlignment.Center
        engine.XAxisFont.StringFormat = horizontalFormat

        'horizontalFormat.FormatFlags = StringFormatFlags.NoWrap
        'engine.XTitle.StringFormat = horizontalFormat
        '  engine.XTitle.Text = "רשימת נציגים"
        engine.XTitle.Font = myFont

        ' Set-up the YTitle
        engine.YTitle = New ChartText
        engine.YTitle.Font = myFont
        Dim verticalFormat As New StringFormat
        verticalFormat.Alignment = StringAlignment.Center
        'verticalFormat.LineAlignment = StringAlignment.Near
        verticalFormat.Trimming = StringTrimming.Character
        verticalFormat.FormatFlags = StringFormatFlags.DirectionVertical
        verticalFormat.Alignment = StringAlignment.Near
        engine.YTitle.StringFormat = verticalFormat
        engine.YTitle.Text = "(כמות השליחות (לידים"

        ' Set-up the Title
        engine.Title = New ChartText : engine.Title.Text = " דוח ביטוחים  " & sdate & "-" & enddate : engine.Title.Font = New Font("Arial", 12, FontStyle.Bold)

        ' Specify show XValues
        engine.ShowXValues = True : engine.ShowYValues = True

        ' Some padding
        engine.Padding = 2 : engine.ChartPadding = 30 : engine.BottomChartPadding = 30 : engine.TopChartPadding = 10

        ' some color
        engine.Background.Color = Color.FromArgb(30, Color.White)

        ' some color
        engine.PlotBackground.Color = Color.White

        engine.GridLines = GridLines.Horizontal
        engine.Border.Color = Drawing.Color.FromArgb(150, 150, 150)
        engine.Border.LineJoin = Drawing2D.LineJoin.MiterClipped
        engine.Border.StartCap = Drawing2D.LineCap.Square

        engine.Legend.Border.Color = Drawing.Color.FromArgb(150, 150, 150)
        engine.Legend.Border.LineJoin = Drawing2D.LineJoin.MiterClipped
        engine.Legend.Border.StartCap = Drawing2D.LineCap.Square
    End Sub
    Function GetEncoderInfo(ByRef mimeType As String) As System.Drawing.Imaging.ImageCodecInfo

        For Each ice As System.Drawing.Imaging.ImageCodecInfo In System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
            If ice.MimeType.Equals(mimeType) Then
                Return ice
            End If
        Next
    End Function


End Class
