Imports System.Drawing
Imports WebChart
Imports System.Web.UI
Public Class reportGraph_Guide
    Inherits System.Web.UI.Page
    Protected ChartControl1 As WebChart.AreaChart
    Dim func As New bizpower.cfunc
    Protected guideId, currentYear, FromDate, ToDate, Guide_Name, Guide_Phone, Guide_Email As String

    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim rs_Guide As System.Data.SqlClient.SqlDataReader

    Dim sqlstrGuide As New SqlClient.SqlCommand


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
        guideId = Request("guide_id")
        currentYear = Request("currentYear")

        FromDate = "1/01/" & currentYear
        ToDate = "31/12/" & currentYear
        '	 response.Write FromDate &":"& ToDate
        '	 response.end
        conPegasus.Open()
        Dim sqlstrGuide As New SqlClient.SqlCommand("SELECT Guide_Id, (Guide_FName + ' - ' + Guide_LName) as Guide_Name,Guide_Phone,Guide_Email  FROM Guides  where Guide_Id=" & guideId, conPegasus)
        rs_Guide = sqlstrGuide.ExecuteReader(CommandBehavior.CloseConnection)
        While rs_Guide.Read()

            If Not IsDBNull(rs_Guide("Guide_Name")) Then
                Guide_Name = rs_Guide("Guide_Name")
            End If
            If Not IsDBNull(rs_Guide("Guide_Phone")) Then
                Guide_Phone = rs_Guide("Guide_Phone")
            End If
            If Not IsDBNull(rs_Guide("Guide_Email")) Then
                Guide_Email = rs_Guide("Guide_Email")
            End If
        End While
        conPegasus.Close()

        Dim wcEng As New ChartEngine
        '  Dim test As New DataLabelPosition

        wcEng.Size = New Size(1800, 600)

        Dim wcCharts As New ChartCollection(wcEng)
        wcEng.Charts = wcCharts

        Dim chart1 As New ColumnChart

        chart1.MaxColumnWidth = 0
        ' chart1.Fill.WrapMode.TileFlipX = Drawing2D.WrapMode.TileFlipX




        Dim myFontChart As New System.Drawing.Font("Arial", 12, FontStyle.Bold)
        chart1.DataLabels.Font = myFontChart 'result font

        'chart1.DataXValueField = System.Drawing.Drawing2D.WrapMode.TileFlipX
        'chart1.DataXValueField.StartsWith(22)


        ' chart1.DataXValueField = System.Drawing.Drawing2D.WrapMode.Clamp



        chart1.Fill.Color = Color.FromArgb(80, 163, 195, 18) 'chart1.Line.Color = Color.FromArgb(80, 163, 195, 18)
        ' chart1.Legend = "תקין" : chart1.ShowLegend = True
        chart1.DataLabels.Visible = True : chart1.Shadow.Visible = False
        chart1.DataLabels.Separator.PadRight(5)

        chart1.DataLabels.ShowXTitle = False
        chart1.ShowLegend = False  'לא להציג שורה עם הסברים

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
        '  Response.Write("SetMoreProperties=" & wcEng.Title.Text)
        ' Response.End()

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
        engine.LeftChartPadding = 10 'padding- left 
        engine.RightChartPadding = 20
        engine.BottomChartPadding = 120





        ' Set-up the XTitle
        Dim myFont As New System.Drawing.Font("Arial", 15, FontStyle.Bold)
        'Dim myFontX As New System.Drawing.Font("Arial", 14, FontStyle.Bold)
        Dim myFontX As New System.Drawing.Font("Arial", 12, FontStyle.Bold)

        engine.XAxisFont.Font = myFontX   'text data coocrdinate x
        'engine.XAxisFont.ForeColor = Drawing.Color.FromArgb(199, 61, 89)
        engine.XAxisFont.ForeColor = Drawing.Color.FromArgb(0, 0, 0)


        engine.XTitle = New ChartText
        Dim horizontalFormat As New StringFormat
        horizontalFormat.Trimming = StringTrimming.Character
        ' horizontalFormat.FormatFlags = StringFormatFlags.DirectionRightToLeft 'x data horizontal or vertical

        'horizontalFormat.FormatFlags = StringFormatFlags.DirectionVertical
        'horizontalFormat.FormatFlags = System.Drawing.RotateFlipType.Rotate270FlipX


        'horizontalFormat.Clone=System.Drawing
        ' horizontalFormat.FormatFlags = StringFormatFlags.NoClip

        '  horizontalFormat.FormatFlags = StringFormatFlags.NoFontFallback

        ' horizontalFormat.FormatFlags = StringFormatFlags.MeasureTrailingSpaces
        ' horizontalFormat.FormatFlags = StringFormatFlags.DisplayFormatControl


        horizontalFormat.LineAlignment = StringAlignment.Near

        horizontalFormat.Alignment = StringAlignment.Near
        engine.XAxisFont.StringFormat = horizontalFormat


        ' horizontalFormat.FormatFlags = StringFormatFlags.NoWrap
        'engine.XTitle.StringFormat = horizontalFormat

        '  engine.XTitle.Text = "יעדים הנבחרים"
        'engine.XTitle.Font = myFontX   'title coordinate X

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
        engine.YTitle.Text = "אחוז סגירה"




        ' Set-up the Title
        engine.Title = New ChartText
        engine.Title.Text = "  משובים דוח שנתי  למדריך   " & FromDate & " - " & ToDate
        Dim t = "1st line\n2ndline"
        engine.Title.Text = t
        engine.Title.Font = New Font("Arial", 12, FontStyle.Bold)
        '  engine.Title.Text = engine.Title.Text & "/bbbbbbb"


        ' Specify show XValues
        engine.ShowXValues = True : engine.ShowYValues = True

        ' Some padding
        engine.Padding = 2 : engine.ChartPadding = 30 : engine.BottomChartPadding = 30 : engine.TopChartPadding = 10
        engine.YValuesInterval = 10 'interval in Y coordinate;
        engine.YCustomEnd = 100 'max coordinate  Y



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
