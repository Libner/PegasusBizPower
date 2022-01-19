<%@Language=VBScript%>
<%Option Explicit

Response.Buffer = true  'enable buffering so that ALL browsers will save image as a JPEG when
						'  a user right-clicks over it and saves it to disk
%>
<!-- #INCLUDE FILE ="ChartConst.inc" -->
<%session("uName") = "1" 
if session("uName") = "" then
	%>
		<br><br><br>
		<p align="center"><font size="5">Not authorized access</font></p>
	<%	
	else
dim objChart		'Dundas Chart 2D object
dim param,param2,param3
'param = Request.QueryString(1) 'percent to be displayed
param=30
param2=50
param3=20
'Step 1: Create a Dundas Chart 2D object
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

'Step 2: Add ALL data to be used by the chart. 
objChart.AddData param, 0, param & "%"	'Adds value 12 to Data Series 0 
objChart.AddData param2, 0, param2 & "%"	'Adds value 12 to Data Series 0 
objChart.AddData param3, 0, param3 & "%"	'Adds value 12 to Data Series 0 

									'and assigns a data label "Data 1" to this data (for pie charts used for legend items).
							 'BBGGRR 				
objChart.AddData 0, 0, "", &Hcccccc


'Step 3: Select data from Data Series 0 to plot a pie chart then add this chart to 
'		 ChartArea 0. The constant PIE_CHART is defined in the ChartConst.inc file.
objChart.ChartArea(0).AddChart PIE_CHART, 0, 0

objChart.ShowPieLabels

'Step 4: Apply Antialiasing effect
objChart.AntiAlias 

'use relative coordinates, which results in the top-left corner of JPEG at 0,0
'  and the bottom-right corner represented by 1000,1000
objChart.SetRelativeCoordinates 400,400

'increase size of chart area to decrease whitespace around pie chart
objChart.ChartArea(0).SetPosition 200, 200, 900, 900

'Step 5: Send a 400 x 400 pixels JPEG
objChart.SendJpeg 400,400

set objChart = nothing
end if
%>