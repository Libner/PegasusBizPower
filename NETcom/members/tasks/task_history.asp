<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 29 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language=javascript>
<!--
	function closeTask(contactID,companyID,taskID)
	{
		h = parseInt(470);
		w = parseInt(420);
		window.open("closetask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskId=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=50,left=100,width="+w+",height="+h+",align=center,resizable=0");
	}
	
	function addtask(contactID,companyID,taskID)
	{
		h = parseInt(520);
		w = parseInt(420);
		window.open("addtask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskID=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}
//-->
</script>

</head>
<% taskID = trim(Request("taskID"))
   UserID = trim(Request.Cookies("bizpegasus")("UserID"))
%>
<body style="margin:0px">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td class="page_title"><span id=word22 name=word22><!--היסטורית--><%=arrTitles(22)%></span> <%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></td></tr>		   
<tr><td width=100%>
<%   
   If trim(taskID) <> "" Then

		function getTree(taskID, byREF strTmp)
			sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_id & "','" & taskId & "'"
			set rs_task = con.getRecordSet(sqlstr)
			If not rs_task.eof Then
		   		nextTaskID = trim(rs_task("task_ID"))			 
				taskParentID = trim(rs_task("parent_id"))								
			End If
			set rs_task = nothing	
			
		   
			If trim(taskParentID) <> "" And IsNull(taskParentID) = false Then	     	              
			        strTmp = strTmp & taskParentID & ","
					getTree taskParentID,strTmp
			Else	       	        
					Exit function
			End If 	
			   
			getTree = strTmp
		end function	
	
	parents_list = getTree(taskID,"")
	'Response.Write "parents_list=" & getTree(taskID,"")
        
    urlSort = "task_history.asp?taskID=" & taskID
	if lang_id = "1" then
		arr_Status = Array("","חדש","בטיפול","סגור")	
	else
		arr_Status = Array("","new","active","close")	
	end if

    function insertTaskID(chield_id,pos,lvl,stack)    
		stack(pos+1,0) = chield_id
		stack(pos+1,1) = lvl + 1	              
		insertTaskID = pos + 1
    end function
    
    function deleteTaskID(current_id,lvl,stack,pos)   
		for i=0 to pos
			if stack(i,1) = lvl and stack(i,0) = current_id then
				stack(i,1) = ""
				stack(i,0) = ""
			end if	
		next
		deleteTaskID = pos - 1
    end function
    
    function getTaskID(lvl,stack,pos)   
		for i=0 to pos
			if stack(i,1) = lvl then
				getTaskID = stack(i,0)
				exit function
			end if	
		next
		getTaskID = 0
    end function
    
    function getLevel(lvl,stack,pos)
		for i=0 to pos
			if stack(i,1) = lvl then
				getLevel = true
				exit function
			end if	
		next
		getLevel = false
    end function
    
    function getChildren (current_id)
    DIM lvl,task_ids
    task_ids = ""
    dim stack(2,1000)
    
    ' Insert current node to the stack.
    pos = 0
    stack(pos,0) = current_id
    stack(pos,1) = 1    
   
    lvl = 1
 				
	WHILE lvl > 0					'From the top level going down.	    
   
	    IF getLevel(lvl,stack,pos) Then
	            
	            current_id = getTaskID(lvl,stack,pos)	           
	            task_ids = task_ids & current_id & ","
	            
	            'Response.Write task_ids
	            'Response.End
	            pos = deleteTaskID(current_id,lvl,stack,pos)            
	            
	            
	            set rs_parent = con.getRecordSet("Select task_id FROM tasks WHERE parent_id = " & current_id)
	            if not rs_parent.eof then
					 chield_id = rs_parent(0)					
					 pos = insertTaskID(chield_id,pos,lvl,stack)
                     lvl = lvl + 1	'If no nodes are added, check its brother nodes.
                end if
                set rs_parent = Nothing     
		
    	ELSE
	      	lvl = lvl - 1		'Back to the level immediately above.
	    End IF
	      	
	 WEND
     getChildren = task_ids
 End FUNCTION
    
    childrens_list = getChildren(taskID)
    
    If Len(childrens_list) > 2 Then        
  
 %>      
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100% dir="<%=dir_var%>">
  <tr>
<td width="100%" valign=top>
   <table width=100% dir="<%=dir_var%>" border=0 bordercolor=red cellpadding=0 cellspacing=1 bgcolor="#FFFFFF">   	    
    <tr> 
      <td align="<%=align_var%>" class="title_sort" width=100% nowrap><span id="word4" name=word4><!--תוכן--><%=arrTitles(4)%></span>&nbsp;</td>                  
      <!--td width=110 nowrap class="title_sort" align="<%=align_var%>">&nbsp;סוגי משימה&nbsp;</td-->   
      <!--td align="<%=align_var%>" width=110 nowrap valign=top class="title_sort">&nbsp;פרויקט&nbsp;</td-->
      <!--td align="<%=align_var%>" width=90 nowrap class="title_sort">&nbsp;איש קשר&nbsp;</td-->
      <td width="150" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;</td>
      <td width="90" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort">&nbsp;<span id="word6" name=word6><!--אל--><%=arrTitles(6)%></span>&nbsp;</td>
      <td width="90" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort">&nbsp;<span id="word7" name=word7><!--מאת--><%=arrTitles(7)%></span>&nbsp;</td>      
      <td align=center width=65 nowrap class="title_sort"><span id="word8" name=word8><!--תאריך יעד--><%=arrTitles(8)%></span></td>          
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=40 nowrap class="title_sort">&nbsp;<span id="word9" name=word9><!--'סט--><%=arrTitles(9)%></span>&nbsp;</td>
      <td width=20 nowrap class="title_sort">&nbsp;</td>
    </tr>   
 <%     
		current_date = Day(date()) & "/" & Month(date()) & "/" & Year(date())
		dim	 IS_DESTINATION
		If Len(parents_list) > 0 Then
			list = parents_list & childrens_list
		Else
			list = childrens_list
		End If		    
		list = Left(list, Len(list) - 1)   		
		
		sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & list & "'"
		'Response.Write sqlstr
		'Response.End  
		set rs_task = con.getRecordSet(sqlstr)
		while not rs_task.eof
		   	currTaskID = trim(rs_task("task_ID"))
			CONTACT_NAME = trim(rs_task("CONTACT_NAME"))
			companyName = trim(rs_task("company_Name"))
			project_name = trim(rs_task("project_name"))				
			task_types = trim(rs_task("task_types"))
			task_status = trim(rs_task("task_status"))
			sender_name = trim(rs_task("sender_name"))
			reciver_name = trim(rs_task("reciver_name"))
			ReciverID = trim(rs_task("reciver_id")) 
			SenderID = trim(rs_task("User_ID")) 			
			attachment = trim(rs_task("attachment"))
			
			If cInt(strScreenWidth) > 800 Then
				numOfLetters = 150
			Else
				numOfLetters = 55
			End If
			tel_text = trim(rs_task("task_content"))
			If Len(tel_text) > numOfLetters Then
				tel_text_short = Left(tel_text , numOfLetters-2) & ".."
			Else tel_text_short = tel_text	
			End If
			task_date = trim(rs_task("task_date"))
			If isDate(task_date) Then
				d_s = Day(task_date) & "/" & Month(task_date) & "/" & Right(Year(task_date),2)
				if DateDiff("d",d_s,current_date) >= 0 then
					IS_DESTINATION = true
				else
					IS_DESTINATION = false
				end if
			Else
			   d_s=""
			   IS_DESTINATION = false
			End If
			
			class_ = ""
			
			If trim(Request("taskID")) = trim(currTaskID) Then
				class_ = "8"
			ElseIf trim(UserID) = trim(SenderID)  Then
				class_ = "4"
			ElseIf  trim(UserID) = trim(ReciverID) Then
				class_ = "7"
		    End if
		    
			If trim(UserID) = trim(SenderID) AND trim(task_status) = "1" Then
				href = "href=""javascript:addtask('" & contactID & "','" & companyID & "','" & currTaskID & "')"""   
			Else      
				href = "href=""javascript:closeTask('" & contactID & "','" & companyID & "','" & currTaskID & "')"""     
			End If				
	   
      %>      
      <tr>  
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%> title="<%=vFix(tel_text)%>"><%=tel_text_short%></a></td>
      <!--td align="<%=align_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%>><%=task_types_name%></a></td-->             
      <!--td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%>><%=project_name%></a></td-->
      <!--td align="<%=align_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%>><%=CONTACT_NAME%></a></td-->
      <td class="card<%=class_%>" valign=top align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>><%=companyName%></a></td>
      <td align="<%=align_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>"><%=reciver_name%></a></td>
      <td align="<%=align_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>"><%=sender_name%></a></td>      
      <td align=center class="card<%=class_%>" valign=top nowrap dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%> <%if IS_DESTINATION and task_status <> 3 then%> title="תאריך יעד עבר"><span style="width:9px;COLOR: #FFFFFF;BACKGROUND-COLOR: #FF0000;text-align:center"><b>!</b></span><%else%>><%end if%>&nbsp;<%=d_s%>&nbsp;</a></td>            
      <td align=center class="card<%=class_%>" valign=top><a class="task_status_num<%=task_status%>"><%=arr_Status(task_status)%></A></td>
      <td align=center class="card<%=class_%>"><img src="../../images/hets3.gif" vspace=0 hspace=0></td>
      </tr>
      <%	
		rs_task.moveNext
		Wend
		set rs_task = nothing	  		
    End If   
 End If   
 %>	
</body>
<%set con=Nothing%>
</html>
