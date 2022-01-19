<%@Language=VBScript%>
<%''http://www.dundas.com/products/onlinedocs/Chart2D/DSChart2D.htm
Response.Buffer = true  'enable buffering so that ALL browsers will save image as a JPEG when
						'  a user right-clicks over it and saves it to disk
%>
<!-- #INCLUDE FILE ="ChartConst.inc" -->
<%
strDataSeries = Request.QueryString("strDataSeries")
ArrDataSeries = split(strDataSeries,";;")
Title = Request.QueryString("Title")
Title = Replace(Title,";and;","&")
strDataLabels = Request.QueryString("strDataLabels")
ArrDataLabels = split(strDataLabels,";;")

if ( UBound(ArrDataSeries) <> UBound(ArrDataLabels) or UBound(ArrDataSeries) = -1 or UBound(ArrDataLabels) = -1 ) then
	Response.End
end if

set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
objChart.CodePage = 1255

for ctr = 0 to ubound(ArrDataLabels)-1
	if ArrDataSeries(ctr) <> "0" then
		objChart.AddData ArrDataSeries(ctr), 0, ArrDataSeries(ctr)& "% " & ArrDataLabels(ctr)
	end if
next

objChart.ChartArea(0).AddChart PIE_CHART, 0, 0

objChart.ShowPieLabels 
objChart.ShowPieLabels "Arial", 11, 0, 000, 40
'objChart.ChartArea(0).SetShadow rgb(128,128,128)
objChart.ChartArea(0).Transparent = true
objChart.SetRelativeCoordinates 600,250

objChart.AddStaticText "BIZPOWER", 0, 1085, 000, "Arial", 16, 1, 0,0
objChart.AntiAlias ()
objChart.SetExploded 0,0

'increase size of chart area to decrease whitespace around pie chart
objChart.ChartArea(0).SetPosition 250, 300, 950, 1100
objChart.BackgroundColor = rgb(255,255,255)
if trim(Title) <> "" then
	objChart.SetTitle Title,"Arial","13",FONT_BOLD
end if	
'Step 5: Send a 400 x 400 pixels JPEG
objChart.SendJpeg 640,350,0
set objChart = nothing%>