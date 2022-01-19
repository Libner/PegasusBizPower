<%@ Page Language="vb" AutoEventWireup="false" Codebehind="testGraph.aspx.vb" Inherits="bizpower_pegasus.testGraph"%>
<script runat=server>
    Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
 
        Dim ds As DataSet = GetDataSet()
        Dim view As DataView = ds.Tables(0).DefaultView
            
        Dim chart As New SmoothLineChart()
        chart.Line.Color = Color.SteelBlue
        chart.Legend = "Value 1"
        chart.DataSource = view
        chart.DataXValueField = "Description"
        chart.DataYValueField = "Value1"
        chart.DataBind()
        ChartControl1.Charts.Add(chart)
 
        Dim chart1 As New SmoothLineChart()
        chart1.Line.Color = Color.Red
        chart1.Legend = "Value 2"
        chart1.DataSource = view
        chart1.DataXValueField = "Description"
        chart1.DataYValueField = "Value2"
        chart1.DataBind()
        ChartControl1.Charts.Add(chart1)
        
        ConfigureColors()
        
        ChartControl1.RedrawChart()
    End Sub
     ' Just create a simple dataset
    Private Function GetDataSet() As DataSet
        Dim ds As New DataSet()
        Dim table As DataTable = ds.Tables.Add("Data")
        table.Columns.Add("Description")
        table.Columns.Add("Value1", GetType(Integer))
        table.Columns.Add("Value2", GetType(Integer))
        Dim rnd As New Random()
        Dim i As Integer
        For i = 0 To 20
            Dim row As DataRow = table.NewRow()
            row("Description") = i.ToString()
            row("Value1") = rnd.Next(100)
            row("Value2") = rnd.Next(100)
            table.Rows.Add(row)
        Next
        Return ds
    End Function
    
    ' Configure some colors for the Chart, this could be done declaratively also
    Sub ConfigureColors()
        ChartControl1.Background.Color = Color.DarkGray
        ChartControl1.Background.Type = InteriorType.Hatch
        ChartControl1.Background.HatchStyle = System.Drawing.Drawing2D.HatchStyle.DiagonalBrick
        ChartControl1.Background.ForeColor = Color.Yellow
        ChartControl1.Background.EndPoint = New Point(500, 350)
        ChartControl1.Legend.Position = LegendPosition.Left
        ChartControl1.Legend.Width = 80
 
        ChartControl1.YAxisFont.ForeColor = Color.Red
        ChartControl1.XAxisFont.ForeColor = Color.Green
        
        ChartControl1.ChartTitle.Text = "WebChart Control Sample"
        ChartControl1.ChartTitle.ForeColor = Color.DarkBlue
    
        ChartControl1.Border.Color = Color.SteelBlue
        ChartControl1.BorderStyle = BorderStyle.Ridge
    End Sub
</script>
<html>
<head><title>WebChart Sample</title></head>
<body>
      <Web:ChartControl Width="500" Height="350" id="ChartControl1" runat="Server" />
</body>
</html>
 