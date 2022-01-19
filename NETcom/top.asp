

<div class="biz_nav_section">
    <div class="biz_buttons">
    <%If Len(numOftab) = 0 Then
		numOftab = -1
	 End If %>
	 <a href="<%=strLocal%>/netCom/" class="bizlink_level_1 <%If trim(numOfTab) = "-1" Then%>current <%End if%>w-inline-block">
        <div>ראשי</div>
      </a>
<%arr_bars_top = session("arr_bars_top")
	  links_array = session("links_array")
	  If IsArray(arr_bars_top) Then    
      For i=0 To Ubound(arr_bars_top)-1
		arr_bar = Split(arr_bars_top(i),",")
		barID = trim(arr_bar(0))
		barTitle = trim(arr_bar(1))
		barVisible = arr_bar(2)	
		If trim(barVisible) = "1" Then
			'href = "href='" & Application("VirDir") & "/" & trim(links_array(i+1)) &"' "
			href = "href='" & trim(links_array(i+1)) &"' "
		
			dis = "" %>  
			<a <%=href%> <%=dis%> target=_self class="bizlink_level_1 <%If trim(barID) = trim(numOfTab+1) Then%>current <%End if%>w-inline-block">
        <div><%=barTitle%></div>
      </a>  	
	<%End If%>
	<%Next%>
	<%End If%>
    </div>

    <div class="biz_links">
    <%
		        count = 0
		        arr_bars = session("arr_bars")
		        If IsArray(arr_bars) Then
		        For j=0 To Ubound(arr_bars)-1
				arr_bar = Split(arr_bars(j),",")
				barID = trim(arr_bar(0))
				barTitle = trim(arr_bar(1))
				barVisible = trim(arr_bar(2))
				barParent = trim(arr_bar(3))
				barURL = trim(arr_bar(4))
				barParentOrder = trim(arr_bar(5))
				
				If cInt(barParent) = cint(numOfTab)+1 Then
				If trim(barVisible) = "1" Then
					'href = "href='" & Application("VirDir") & "/" & trim(barURL) &"' "
					href = "href='" &  trim(barURL) &"' "
				
					dis = ""	%>   
    <a <%=href%> <%=dis%> target="_self" class="bizlink_level_2<%If trim(count) = trim(numOfLink) Then%> current<%End If%>"><%=barTitle%></a>
    <%			count = count + 1	
			End If
			End If%>
			<%Next%>
			
			<%End If%>
    </div>



</div>