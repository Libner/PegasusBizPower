Imports System.Drawing
Imports WebChart

Public Class img_graph
    Inherits System.Web.UI.Page
    Dim reportUserName As String = ""
    Dim func As New bizpower.cfunc
    'Protected ChartControl1 As New ChartControl

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
        Dim dateStart As String = Trim(Request.QueryString("dateStart"))
        Dim dateEnd As String = Trim(Request.QueryString("dateEnd"))
        Dim reportUserId As Integer = func.fixNumeric(Request.QueryString("UserId"))

        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmdSelect As New SqlClient.SqlCommand

        If reportUserId > 0 Then
            cmdSelect = New SqlClient.SqlCommand("SELECT (FIRSTNAME + ' ' + LASTNAME) UserName FROM Users WHERE (User_ID = " & reportUserId & ") ", con)
            cmdSelect.CommandType = CommandType.Text
            con.Open()
            reportUserName = func.dbNullFix(cmdSelect.ExecuteScalar())
            cmdSelect.Dispose() : con.Close()
        End If

        'ChartControl1.BackColor = System.Drawing.Color.White
        'ChartControl1.ChartFormat = ChartImageFormat.Png

        Dim wcEng As New ChartEngine
        wcEng.Size = New Size(954, 550)

        Dim wcCharts As New ChartCollection(wcEng)
        wcEng.Charts = wcCharts

        Dim chart1 As New ColumnChart
        chart1.MaxColumnWidth = 0
        chart1.Fill.Color = Color.FromArgb(80, 163, 195, 18) : chart1.Line.Color = Color.FromArgb(80, 163, 195, 18)
        chart1.Legend = "תקין" : chart1.ShowLegend = True
        chart1.DataLabels.Visible = True : chart1.Shadow.Visible = False

        Dim chart2 As New ColumnChart
        chart2.MaxColumnWidth = 0
        chart2.Fill.Color = Color.FromArgb(80, 234, 86, 60) : chart2.Line.Color = Color.FromArgb(80, 234, 86, 60)
        chart2.Shadow.Visible = False
        chart2.Legend = "לא תקין" : chart2.ShowLegend = True
        chart2.DataLabels.Visible = True

        Dim chart3 As New ColumnChart
        chart3.MaxColumnWidth = 0
        chart3.Fill.Color = Color.FromArgb(80, 128, 128, 128) : chart3.Line.Color = Color.FromArgb(80, 128, 128, 128)
        chart3.Shadow.Visible = False
        chart3.Legend = "בוטל" : chart3.ShowLegend = True
        chart3.DataLabels.Visible = True

        cmdSelect = New SqlClient.SqlCommand("orders_graph_report", con)
        cmdSelect.CommandType = CommandType.StoredProcedure
        cmdSelect.Parameters.Add("@dateStart", SqlDbType.Char, 10).Value = dateStart
        cmdSelect.Parameters.Add("@dateEnd", SqlDbType.Char, 10).Value = dateEnd
        cmdSelect.Parameters.Add("@reportUserId", SqlDbType.Int).Value = reportUserId
        con.Open()
        Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()
        While dr.Read()
            chart1.Data.Add(New ChartPoint(dr("FIRSTNAME").ToString, CType(dr("תקין"), Integer)))
            chart2.Data.Add(New ChartPoint(dr("FIRSTNAME").ToString, CType(dr("לא תקין"), Integer)))
            chart3.Data.Add(New ChartPoint(dr("FIRSTNAME").ToString, CType(dr("בוטל"), Integer)))
        End While
        'Response.Write(dr.HasRows)
        dr.Close() : con.Close()

        wcCharts.Add(chart1)
        wcCharts.Add(chart2)
        wcCharts.Add(chart3)
        'ChartControl1.RedrawChart()

        SetMoreProperties(wcEng)
        'Set the image quality 
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

    ' This are all Optional properties, You can call this method to change the look of your chart
    Private Sub SetMoreProperties(ByVal engine As ChartEngine)
        engine.HasChartLegend = True : engine.Legend.Position = LegendPosition.Bottom : engine.Legend.Width = 30

        ' Set-up the XTitle
        Dim myFont As New System.Drawing.Font("Arial", 10)
        engine.XTitle = New ChartText
        Dim horizontalFormat As New StringFormat
        horizontalFormat.Trimming = StringTrimming.Character
        horizontalFormat.FormatFlags = StringFormatFlags.DirectionVertical
        horizontalFormat.LineAlignment = StringAlignment.Center
        horizontalFormat.Alignment = StringAlignment.Center
        engine.XAxisFont.StringFormat = horizontalFormat

        'horizontalFormat.FormatFlags = StringFormatFlags.NoWrap
        'engine.XTitle.StringFormat = horizontalFormat
        engine.XTitle.Text = ""
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
        engine.YTitle.Text = "מספר הזמנות"

        ' Set-up the Title
        engine.Title = New ChartText : engine.Title.Text = "דוח הזמנות פר נציג" : engine.Title.Font = New Font("Arial", 12, FontStyle.Bold)

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
