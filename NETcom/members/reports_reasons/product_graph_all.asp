<!-- #INCLUDE FILE ="../../chart/ChartConst.inc" -->
<!--#include file="../../connect.asp"--><%	
	prodID = trim(Request.QueryString("prodID"))
	start_date = trim(Request.QueryString("start_date"))
	end_date = trim(Request.QueryString("end_date"))	
	
	sql = "Exec dbo.getStatProd '" & prodID & "','" & start_date & "','" & end_date & "'"
	set rs_tmp=con.getRecordSet(sql)
	if not rs_tmp.EOF then
		arr_st = rs_tmp.GetRows()
	end if
	Set rs_tmp = Nothing	
	
	set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
	objChart.CodePage = 1255
	objChart.Rectangle3DEffect()
	objChart.SetRelativeCoordinates 780, 250	
	
	If isArray(arr_st) Then	
		max_count = 0
		For i=0 To Ubound(arr_st,2)
			st_count = cInt(trim(arr_st(0,i)))
			If st_count > max_count Then
				max_count = st_count
			End If
			st_percent = trim(arr_st(1,i))
			st_name = trim(arr_st(2,i))	
			If trim(st_name) = "" Or isNULL(st_name) Then
				st_name = "ללא מקור פרסום"
			End If
			
			'objChart.AddData st_count, 1, "(" & st_name & " (" & st_count & ""
			objChart.AddData st_count, 0, st_name & " " & st_percent & "%"
		Next
	
		'one data series per bar
		'objChart.SetTitle catName, "Arial", "15", 1 , Rgb(115,107,166) , 500, 5
		
		'objChart.ChartArea(0).AddChart BAR_CHART, 1, 1
		objChart.ChartArea(2).AddChart PIE_CHART, 0, 0
		
		'Step 6: Set chart area sizes and positions
		'objChart.ChartArea(0).SetPosition 100, 100, 400, 800
		objChart.SetColorFromPoint(1) 
		objChart.ChartArea(2).SetPosition 100, 100, 870, 880
		'objChart.ChartArea(0).SetGradient rgb(209,209,209), rgb(255,255,255), 1
		
		objChart.ShowPieLabels "Arial", "10", 1, Rgb(0,0,0), 0
		objChart.AntiAlias ()
		'objChart.SetExploded 0,0
		objChart.ChartArea(2).SetShadow(Rgb(200,200,200))
					
		'objChart.ChartArea(0).Axis(1).Interval  = 1
		'objChart.ChartArea(0).Axis(1).Maximum = max_count + 1
		'objChart.ChartArea(0).Axis(1).ShowEndLabels = True

		'objChart.ChartArea(0).GridHEnabled = False

		'objChart.ChartArea(0).Axis(0).FontSize = 10
		'objChart.ChartArea(0).Axis(0).FontStyle = 1
		'objChart.ChartArea(0).Axis(1).FontSize = 10
		'objChart.ChartArea(0).Axis(1).FontStyle = 1
		'objChart.ChartArea(0).Axis(1).Angle = 0
		'objChart.ChartArea(0).Axis(1).FontName = "Arial"
		'objChart.ChartArea(0).SetGradient rgb(175,175,175), rgb(255,255,255), 1
		'objChart.ShowDataLabels 0,-1,"Arial",10,1,,0,0
		objChart.ShowDataLabels 1,-1,"Arial",10,1, Rgb(0,0,0),0,90
		
		objChart.SendJPEG 780, 250
	End If

	set objChart = nothing
	Set con = Nothing%>