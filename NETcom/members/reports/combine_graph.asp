<!-- #INCLUDE FILE ="../../chart/ChartConst.inc" -->
<!--#include file="../../connect.asp"-->
<%	catID = trim(Request.QueryString("catID"))
	prodId = trim(Request.QueryString("prodId"))
	start_date = trim(Request.QueryString("start_date"))
	end_date = trim(Request.QueryString("end_date"))
	
	If Len(catID) = 0 Then
		catID = "-1" 'כולם
	End If
	
	sql = "Exec dbo.getCategsProd '" & prodId & "','" & catID & "','" & start_date & "','" & end_date & "'"
	set rs_tmp=con.getRecordSet(sql)
	if not rs_tmp.EOF then
		arr_st = rs_tmp.GetRows()
	end if
	Set rs_tmp = Nothing	
	
	set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
	objChart.CodePage = 1255
	objChart.Rectangle3DEffect()
	objChart.SetRelativeCoordinates 780, 250
	
	arr_Status = Array("","חדש","בטיפול","סגור")
	arr_color = Array("",Rgb(255,208,17),Rgb(255,128,64),Rgb(123,182,42)) 
	
	If isArray(arr_st) Then	
		max_count = 0
		For i=0 To Ubound(arr_st,2)
			st_count = cInt(trim(arr_st(0,i)))
			If st_count > max_count Then
				max_count = st_count
			End If
			st_percent = trim(arr_st(1,i))
			app_status = trim(arr_st(2,i))
			Closing_ID = trim(arr_st(3,i)) 
						
			If app_status <> "3" Then
				st_name = arr_Status(app_status)
				st_color = arr_color(app_status)				
			Else
				sqlStr = "Select Closing_Name, RIGHT(Closing_Color, LEN(Closing_Color)+2) From Product_Closings "&_
				" WHERE Product_ID = " & prodID & " And Closing_ID = " & Closing_ID
				set rs_closings = con.GetRecordSet(sqlStr)
				If not rs_closings.eof Then
					st_name = trim(rs_closings(0))
					tmp = trim(rs_closings(1))
					color1=Mid(tmp,1,2) : color2=Mid(tmp, 3, 2) : color3=Mid(tmp, 5, 2)					
					color1=CLng("&H" & color1) : color2=CLng("&H" & color2) : color3=CLng("&H" & color3)
					st_color = Rgb(color1, color2, color3)
				End If	
				Set rs_closings = Nothing				    
			End If				
			objChart.AddData st_count, 1, "(" & st_name & " (" & st_count & "", st_color
			objChart.AddData st_count, 0, st_name & " " & st_percent & "%", st_color
		Next
	
		'one data series per bar
		'objChart.SetTitle catName, "Arial", "15", 1 , Rgb(115,107,166) , 500, 5
		
		objChart.ChartArea(0).AddChart BAR_CHART, 1, 1
		objChart.ChartArea(2).AddChart PIE_CHART, 0, 0
		
		'Step 6: Set chart area sizes and positions
		objChart.ChartArea(0).SetPosition 100, 100, 400, 800
		objChart.SetColorFromPoint(1) 
		objChart.ChartArea(2).SetPosition 500, 20, 870, 880
		objChart.ChartArea(0).SetGradient rgb(209,209,209), rgb(255,255,255), 1
		
		objChart.ShowPieLabels "Arial", "10", 1, Rgb(0,0,0), 0
		objChart.AntiAlias ()
		'objChart.SetExploded 0,0
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