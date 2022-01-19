Imports System.Drawing
Imports WebChart
Imports System.Web.UI

'Imports System.Web.UI.WebControls
'Imports System.Web.UI.HtmlControls





Public Class reportRegistrationGraphWithReasons

    Inherits System.Web.UI.Page
    Protected ChartControl1 As WebChart.AreaChart
    Dim func As New bizpower.cfunc
    Protected dateStart, dateEnd, countryid, sqlQuery As String
    Protected quest_id, s16504, s16735 As Integer

    Dim reasons As String

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
        dateStart = Trim(Request("dateStart"))
        dateEnd = Trim(Request("dateEnd"))
        countryid = Trim(Request("country_id"))

        reasons = ""
        'only for tofes mitan'en

        'הגורם ליצירת הטופס
        '---------------P2932 - type of form- REASONS ---------------
        'select creation reason-------------------------------
        'added by Mila 22/10/2019-----------------------------	
        reasons = Request.Form("reasons")
        reasons = Replace(reasons, " ", "")

        '-----------------------------------------------------------
        quest_id = 16735
        sqlQuery = " and appeal_CountryId   in (" & countryid & ")"
        ' Response.Write(sqlQuery)
        ' Response.End()
        ' The data for the bar chart
        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim con1 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        Dim cmdSelect As New SqlClient.SqlCommand

        ' dim  sql = "Exec dbo.get_reportRegistration  @start_date='" & dateStart & "', @end_date='" & dateEnd & "'," & _
        '" @sqlQuery='" & sqlQuery & "', @productID=" & quest_id)
        '   Response.Write(sql)
        '   Response.End()
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
        If countryid <> "0" Then
            cmdSelect = New System.Data.SqlClient.SqlCommand("Select Country_Id, Country_Name FROM Countries  where Country_Id in (" & countryid & ") Order BY Country_Name ", con)
            con.Open()
            Dim drNew = cmdSelect.ExecuteReader()
            Do While drNew.Read()
                Dim sss As String
                'sss = "SELECT count(appeal_id)  as Count16504 FROM dbo.APPEALS APP  where  (DateDiff(d,APP.appeal_date, convert(smalldatetime, '" & dateStart & "',103)) <= 0 OR (Len('" & dateStart & "') = 0 ))	" & _
                '                       " AND (DateDiff(d,APP.appeal_date, convert(smalldatetime, '" & dateEnd & "' ,103)) >= 0 OR (Len('" & dateEnd & "') = 0)) " & _
                '                       " AND (APP.questions_id = 16504) and appeal_CountryId=" & drNew("country_id")
                'Response.Write(sss)

                Dim sqlStr As String
                If reasons = "" Then
                    sqlStr = "SELECT count(appeal_id)  as Count16504 FROM dbo.APPEALS APP  where  (DateDiff(d,APP.appeal_date, convert(smalldatetime, '" & dateStart & "',103)) <= 0 OR (Len('" & dateStart & "') = 0 ))	" & _
                       " AND (DateDiff(d,APP.appeal_date, convert(smalldatetime, '" & dateEnd & "' ,103)) >= 0 OR (Len('" & dateEnd & "') = 0)) " & _
                       " AND (APP.questions_id = 16504) and APP.appeal_CountryId=" & drNew("country_id")
                Else
                    sqlStr = "SELECT count(appeal_id)  as Count16504 FROM dbo.APPEALS APP  where  (DateDiff(d,APP.appeal_date, convert(smalldatetime, '" & dateStart & "',103)) <= 0 OR (Len('" & dateStart & "') = 0 ))	" & _
                       " AND (DateDiff(d,APP.appeal_date, convert(smalldatetime, '" & dateEnd & "' ,103)) >= 0 OR (Len('" & dateEnd & "') = 0)) " & _
                       " AND (APP.questions_id = 16504) and APP.appeal_CountryId=" & drNew("country_id") & " and APP.Reason_Id in (" & reasons & ") "
                End If

                Dim cmd16504 = New System.Data.SqlClient.SqlCommand(sqlStr, con1)
                con1.Open()
                If Not IsDBNull(cmd16504.ExecuteScalar()) Then
                    s16504 = cmd16504.ExecuteScalar()
                Else
                    s16504 = 0
                End If
                con1.Close()
                If reasons = "" Then
                    sqlStr = "SELECT count(appeal_id)  as Count16735 FROM dbo.APPEALS APP  where  (DateDiff(d,APP.appeal_date, convert(smalldatetime, '" & dateStart & "',103)) <= 0 OR (Len('" & dateStart & "') = 0 ))	" & _
                       " AND (DateDiff(d,APP.appeal_date, convert(smalldatetime, '" & dateEnd & "' ,103)) >= 0 OR (Len('" & dateEnd & "') = 0)) " & _
                       " AND (APP.questions_id = 16735) and appeal_CountryId=" & drNew("country_id")
                Else
                    sqlStr = "SELECT COUNT(APP.APPEAL_ID) as Count16735 FROM APPEALS APP inner join APPEALS app16504 on APP.RelationId=app16504.APPEAL_ID " & _
                    " where  (DateDiff(d,APP.appeal_date, convert(smalldatetime, '" & dateStart & "',103)) <= 0 OR (Len('" & dateStart & "') = 0 ))	" & _
                                          " AND (DateDiff(d,APP.appeal_date, convert(smalldatetime, '" & dateEnd & "' ,103)) >= 0 OR (Len('" & dateEnd & "') = 0)) " & _
                                          " AND (APP.questions_id = 16735) and APP.appeal_CountryId=" & drNew("country_id") & _
 " AND app16504.Reason_Id in (" & reasons & ") "

                End If
                Dim cmd16735 = New System.Data.SqlClient.SqlCommand(sqlStr, con1)
                con1.Open()
                If Not IsDBNull(cmd16735.ExecuteScalar()) Then
                    s16735 = cmd16735.ExecuteScalar()
                Else
                    s16735 = 0
                End If
                con1.Close()
                Dim res As Decimal
                If s16504 > 0 Then
                    res = Math.Round(s16735 / s16504 * 100)
                Else
                    res = 0
                End If



                '   Response.Write("s16504=" & s16504)
                '   Response.End()


                chart1.Data.Add(New ChartPoint(drNew("Country_Name").ToString & s16735 & "/" & s16504, res))
                wcCharts.Add(chart1)

            Loop

            con.Close()
            drNew.Close()




        End If


        '''''''cmdSelect = New SqlClient.SqlCommand("get_reportRegistration", con)
        '''''''cmdSelect.CommandType = CommandType.StoredProcedure
        '''''''cmdSelect.Parameters.Add("@start_date", SqlDbType.Char, 10).Value = dateStart
        '''''''cmdSelect.Parameters.Add("@end_date", SqlDbType.Char, 10).Value = dateEnd
        '''''''cmdSelect.Parameters.Add("@productID", SqlDbType.Int).Value = quest_id
        '''''''cmdSelect.Parameters.Add("@sqlQuery", SqlDbType.VarChar, 500).Value = sqlQuery
        '''''''con.Open()
        '''''''Dim dr As SqlClient.SqlDataReader = cmdSelect.ExecuteReader()
        '''''''While dr.Read()
        '''''''    'chart1.Data.Add(New ChartPoint(dr("Country_CRMName").ToString, CType(dr("Count22_30"), Integer)))
        '''''''    ' chart1.DataXValueField = dr("Country_CRMName").ToString
        '''''''    'chart1.DataYValueField = "25"

        '''''''    chart1.Data.Add(New ChartPoint(dr("Country_CRMName").ToString, "25"))
        '''''''    wcCharts.Add(chart1)

        '''''''    '  chart1.Fill.Color = Color.FromArgb(10, 113, 115, 18)

        '''''''    '''chart1.Data.Add(New ChartPoint("berlin", "30"))
        '''''''    ''''   chart1.Fill.Color = Color.FromArgb(0, 58, 58, 22)
        '''''''    '''wcCharts.Add(chart1)
        '''''''    '''chart1.Data.Add(New ChartPoint("russia", "60"))
        '''''''    ''''  chart1.Fill.Color = Color.FromArgb(0, 0, 0, 0)

        '''''''    '''wcCharts.Add(chart1)
        '''''''End While
        ''''''''Response.Write(dr.HasRows)
        '''''''dr.Close() : con.Close()

        '  wcCharts.Add(chart1)
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

        horizontalFormat.FormatFlags = StringFormatFlags.DirectionVertical
        horizontalFormat.FormatFlags = System.Drawing.RotateFlipType.Rotate270FlipX


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
        Dim titleStr As String
        titleStr = "  אחוז הסגירה ביעדים הנבחרים בין התאריכים   " & dateStart & " - " & dateEnd

        '---------------P2932 - type of form- REASONS ---------------
        If reasons <> "" Then
            Dim i As Integer = 0
            titleStr = titleStr & ",  הגורם ליצירת הטופס: "
            Dim sqlstr As String = "SELECT  Reason_Title FROM  Appeals_CreationReasons where QUESTIONS_ID=16504 and Reason_Id in (" & reasons & ") order by Reason_Order"

            Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            Dim cmdtitlte As New System.Data.SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            Dim rs_Reason As SqlClient.SqlDataReader
            rs_Reason = cmdtitlte.ExecuteReader
            Do While rs_Reason.Read
                If i > 0 Then
                    titleStr = titleStr & ", "
                End If
                titleStr = titleStr & rs_Reason("Reason_Title")
                i = i + 1
            Loop
            rs_Reason.Close()
            con.Close()
        End If

        engine.Title = New ChartText : engine.Title.Text = titleStr : engine.Title.Font = New Font("Arial", 12, FontStyle.Bold)



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
