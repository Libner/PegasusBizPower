<!-- #INCLUDE FILE ="../../chart/ChartConst.inc" -->
<!--#include file="../../connect.asp"-->
<%	pageID = trim(Request.QueryString("pageID"))
	categID = trim(Request.QueryString("categID"))
	
	sql = "Exec dbo.getCategsPages '" & pageID & "','" & categID & "'"
	set rs_tmp=con.getRecordSet(sql)
	if not rs_tmp.EOF then
		arr_count = rs_tmp.GetRows()
	end if
	Set rs_tmp = Nothing
	
	set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
	objChart.CodePage = 1255
	objChart.Rectangle3DEffect()
	objChart.SetRelativeCoordinates 780, 250	
	
	arr_colors = Array(Rgb(255,115,0),Rgb(0,132,206),Rgb(222,40,0),Rgb(165,99,205),Rgb(123,156,189),Rgb(123,182,42),_
	Rgb(1,197,255),Rgb(254,206,0),Rgb(0,99,190),Rgb(173,100,156),Rgb(90,73,115),Rgb(214,40,89),Rgb(231,82,0),Rgb(172,197,132))
	'Array(Rgb(255,115,0),"0084CE","DE2800","A563CD","7B9CBD","7BB62A","01C5FF","FECE00","0063BE","AD649C","5A4973","D62859","E75200","ACC584")
	
	If isArray(arr_count) Then	
		max_count = 0
		For i=0 To Ubound(arr_count,2)			
			st_count = cInt(trim(arr_count(0,i)))
			If st_count > max_count Then
				max_count = st_count
			End If
			st_percent = trim(arr_count(1,i))
			st_name = trim(arr_count(2,i))
			If Len(st_name) = 0 Or isNULL(st_name) Then
				st_name = "ללא מקור"
			End If			
			objChart.AddData st_count, 1, "(" & st_name & " (" & cStr(st_count) & "", arr_colors(i Mod 15)
			objChart.AddData st_count, 0, st_name & " " & st_percent & "%", arr_colors(i Mod 15)
		Next
	
		'one data series per bar
		objChart.SetTitle catName, "Arial", "15", 1 , Rgb(115,107,166) , 500, 5
		
		objChart.ChartArea(0).AddChart BAR_CHART, 1, 1
		objChart.ChartArea(2).AddChart PIE_CHART, 0, 0
		
		'Step 6: Set chart area sizes and positions
		objChart.ChartArea(0).SetPosition 160, 100, 460, 800
		objChart.SetColorFromPoint(1) 
		objChart.ChartArea(2).SetPosition 550, 100, 900, 800
		objChart.ChartArea(0).SetGradient rgb(209,209,209), rgb(255,255,255), 1
		
		objChart.ShowPieLabels "Arial", "10", 1, Rgb(0,0,0), 0
		objChart.AntiAlias ()
		objChart.SetExploded 0,0
		objChart.ChartArea(2).SetShadow(Rgb(200,200,200))
					
		'objChart.ChartArea(0).Axis(1).Interval  = 1
		'objChart.ChartArea(0).Axis(1).Maximum = max_count + 1
		objChart.ChartArea(0).Axis(1).ShowEndLabels = True

		objChart.ChartArea(0).GridHEnabled = False

		objChart.ChartArea(0).Axis(0).FontSize = 10
		objChart.ChartArea(0).Axis(0).FontStyle = 1
		objChart.ChartArea(0).Axis(1).FontSize = 10
		objChart.ChartArea(0).Axis(1).FontStyle = 1
		objChart.ChartArea(0).Axis(1).Angle = 0
		objChart.ChartArea(0).Axis(1).FontName = "Arial"
		objChart.ChartArea(0).SetGradient rgb(175,175,175), rgb(255,255,255), 1
		'objChart.ShowDataLabels 0,-1,"Arial",10,1,,0,0
		objChart.ShowDataLabels 1,-1,"Arial",10,1, Rgb(0,0,0),0,90
		
		objChart.SendJPEG 780, 250
	End If

	set objChart = nothing
	Set con = Nothing%>