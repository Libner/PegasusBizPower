<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%pUserId= trim(Request.QueryString("pUserId"))

	sqlstr = "Select  FIRSTNAME + '  ' + LASTNAME as UserName from Users  Where User_id = " & pUserId 				
	set rsname=con.getRecordSet(sqlstr)	
	if not rsname.eof then
	UserName=rsname(0)
	end if
	set rsname=Nothing
status = trim(Request.QueryString("status"))
task_status= trim(Request.QueryString("status"))
  if lang_id = "1" then
    arr_Status = Array("","חדש","בטיפול","סגור")	
    self_name = UserName
 else
    arr_Status = Array("","new","active","close")	
    self_name = UserName
 end if
Ttype= trim(Request.QueryString("Ttype"))
taskTypeID=trim(Request.QueryString("Ttype"))
if TType="2026" then
	TTypeName="משימות קישורית"
	elseif TType="2005" then
	TTypeName="משימות מעקב"
	elseif  TType="2027" then
	TTypeName="	רישום ופתיחת טופס או עדכון כרטיס"
end if
reciver_id = pUserId		
where_reciver = " AND reciver_id = " & reciver_id  
where_status = " And (task_status in (" & sFix(task_status) & "))"
  If Request("start_date") <> nil Then	
	start_date = Request("start_date")
	start_date_ = Month(start_date) & "/" & Day(start_date) & "/" & Year(start_date)   
 Else
	start_date_ = ""  
 End If

 If Request("end_date") <> nil Then
	end_date = Request("end_date")
    end_date_ = Month(end_date) & "/" & Day(end_date) & "/" & Year(end_date)
 Else
    end_date_ = ""
 End If
 
 if trim(Request.QueryString("page"))<>"" then
    page=Request.QueryString("page")
 else
    Page=1
 end if  
 
 if trim(Request.QueryString("numOfRow"))<>"" then
    numOfRow=Request.QueryString("numOfRow")
 else    
    numOfRow = 1
 end if  
  sort = Request.QueryString("sort")	
 if trim(sort)="" then  sort=0 end if  

dim sortby(12)	

sortby(0) = "task_status, task_date, task_id DESC"

sortby(1) = "company_name, task_date DESC, task_id DESC"
sortby(2) = "company_name DESC, task_date DESC, task_id DESC"
sortby(3) = "task_date, task_id DESC"
sortby(4) = "task_date DESC, task_id DESC"
sortby(5) = "contact_name, task_date DESC, task_id DESC"
sortby(6) = "contact_name DESC, task_date DESC, task_id DESC"
sortby(7) = "U.FIRSTNAME, U.LASTNAME, task_date DESC,task_id DESC"
sortby(8) = "U.FIRSTNAME DESC, U.LASTNAME DESC,task_date DESC,task_id DESC"
sortby(9) = "U1.FIRSTNAME,  U1.LASTNAME,task_date DESC,task_id DESC"
sortby(10) = "U1.FIRSTNAME DESC,  U1.LASTNAME DESC,task_date DESC,task_id DESC"
sortby(11) = "project_name,task_date DESC,task_id DESC"
sortby(12) = "project_name DESC,task_date DESC,task_id DESC"
urlSort="taskDetails.asp?status="&status &"&Ttype="&Ttype &"&pUserId=" & pUserId &"&search_company="& Server.URLEncode(search_company) &"&amp;search_contact="& Server.URLEncode(search_contact)&"&amp;search_project="&Server.URLEncode(search_project)&"&amp;start_date="&start_date&"&amp;end_date="&end_date & "&amp;task_status="&task_status & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id & "&amp;tasktypeID=" & tasktypeID & "&amp;T=" & T

%>
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name=vs_defaultClientScript content="JavaScript">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name=ProgId content=VisualStudio.HTML>
<meta name=Originator content="Microsoft Visual Studio .NET 7.1">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript" src="../../tooltip.js"></script>
<script>
function cball_onclick() {
	var strid = new String(document.getElementById("ids").value);
	var arrid = strid.split(',');
	for (i=0;i<arrid.length;i++)
		document.getElementById('cb'+ arrid[i]).checked = document.getElementById("cb_all").checked;	
}
	function DoCal(elTarget)
	{
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
    }
	function checkMoveTasks()
	{
		var fl = 0;
		document.getElementById("trapp").value = '';
		if(document.getElementById("ids"))
		{
		var strid = new String(document.getElementById("ids").value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.getElementById('cb'+ arrid[i]).checked)
				{	document.getElementById("trapp").value = document.getElementById("trapp").value + arrid[i] + ',';
					fl = 1;
				}	
			}		  
			if (fl && confirm("? האם ברצונך להעביר את המשימות הנבחרות בצורה גורפת")){
				var txtnew = new String(document.getElementById("trapp").value);
				document.getElementById("trapp").value = txtnew.substr(0,txtnew.length - 1);
				
				h = parseInt(520);
				w = parseInt(520);
				window.open("../tasks/movetasks.asp?tasksId=" + document.getElementById("trapp").value, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
				return true;
			}
			else if (fl) return false;
		}
		window.alert("! נא לסמן משימות על מנת לבצע העברה גורפת");
		return false;
	}	
	}
</script>
</head>
<body style="margin: 0px;" onload="self.focus();"  >
<div id="ToolTip"></div>

   <%
    'Response.Write sqlStr
	'set app=con.GetRecordSet(sqlStr)
	'app_count = app.RecordCount	
	'if Request("page_app")<>"" then
	'	page_app=request("page_app")
	'else
	'	page_app=1
	'end if
	'if not app.eof then
	'	app.PageSize = 15
	'	app.AbsolutePage=page_app
	'	recCount=app.RecordCount 		
'		NumberOfPagesApp = app.PageCount		
'		i=1
'		j=0
'		ids = "" 'list of appeal_id
'	end if%>
	<%'if not app.eof Or search = true then %>
  <table cellpadding="0" cellspacing="0" dir="<%=dir_var%>" width="100%" border="0" ID="Table1">  
  <tr><td height=20></td></tr>
    <tr>
    <td valign="top" class="title_form" width="100%" align="<%=align_var%>" dir=<%=dir_obj_var%>>&nbsp;פרטי &nbsp;<b><%=TTypeName%>  - אל: &nbsp;<%=UserName%></b>&nbsp;</td>
  </tr> 
<tr><td align=center bgcolor=#FFD011>
<FORM action="taskDetails.asp?status=<%=status%>&Ttype=<%=Ttype%>&pUserId=<%=pUserId%>&sort=<%=sort%>&amp;T=<%=T%>" method=POST id=form_search name=form_search target="_self">   

<table border=0 cellpadding=5 cellspacing=5>
<tr>
<td nowrap align="center"><a class="button_edit_1" style="width:80;" href="#" onclick="document.form_search.submit();">&nbsp;<!--עדכן תאריך--><%=arrTitles(14)%>&nbsp;</a></td>
<td align="center" nowrap><input dir="ltr" class="texts" type="text" id="end_date" name="end_date" value="<%=end_date%>" style="width:70" onclick="return DoCal(this);" readonly></td>		
<td  align="<%=align_var%>" >&nbsp;<b><!--עד תאריך--><%=arrTitles(13)%></b>&nbsp;</td>
<td align="center" nowrap><input dir="ltr" class="texts" type="text" id="start_date" name="start_date" value="<%=start_date%>" style="width:70" onclick="return DoCal(this);" readonly></td>					
<td align="<%=align_var%>" ><b><!--מתאריך--><%=arrTitles(12)%></b>&nbsp;</td>
</tr>


</table>
</form>
</td></tr>

  <form id="form1" name="form1" action="taskDetails.asp?status=<%=status%>&Ttype=<%=Ttype%>&pUserId=<%=pUserId%>" method="post">
    <TR><td align=right bgcolor=#FFD011><a class="button_edit_1" style="width:78px;" href='javascript:void(0)' onclick="return checkMoveTasks();">העבר משימות</a></td></tr>

<tr><td width=100%>
<table border=0 cellpadding=1 cellspacing=1 width=100%>

   
    <tr style="line-height:18px"> 	  
       
          <td align="center" class="title_sort" width=29 nowrap>&nbsp;</td>   
  	      <td width="145" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;תוכן סגירה&nbsp;</td>	                         
  	      <td width="145" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;תוכן משימה&nbsp;</td>	          
	      <td id="td_type" width="150" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort"><!--סוג--><%=arrTitles(5)%>&nbsp;</td>
	      <!--td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="11" OR trim(sort)="12" then%>_act<%end if%>"><%if trim(sort)="11" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="12" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=11"  title="<%=arrTitles(26)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="11" then%>bot<%elseif trim(sort)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td-->
	      <td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="5" OR trim(sort)="6" then%>_act<%end if%>"><%if trim(sort)="5" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="6" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=5"  title="<%=arrTitles(26)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>bot<%elseif trim(sort)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>	      
	         <td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort">&nbsp;<!--אל--><%=arrTitles(6)%>&nbsp;</td>
          <td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="7" OR trim(sort)="8" then%>_act<%end if%>"><%if trim(sort)="7" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="8" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=7"  title="<%=arrTitles(26)%>"><%end if%>&nbsp;<!--מאת--><%=arrTitles(7)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="7" then%>bot<%elseif trim(sort)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	      <td width="100" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=3"  title="<%=arrTitles(26)%>"><%end if%><!--תאריך יעד--><%=arrTitles(8)%><img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	      <td width="40" nowrap id="td_status" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort">&nbsp;סטטוס&nbsp;</td>
	    <td class="card<%=class_%>" align="center" valign="top"><%If trim(contactID) <> "" And trim(contactID) <> "0" Then%><img src="../../images/forms_icon.gif" border="0" 
			style="cursor: pointer;"	alt="היסטוריית טפסים של הלקוח" 
			onclick="window.open('../appeals/contact_appeals.asp?contactID=<%=contactID%>','winCA','top=20, left=10, width=950, height=500, scrollbars=1');"  ><%End If%></td>
	    
	      <td class="card<%=class_%>" align="center" valign="top"><INPUT type="checkbox" LANGUAGE="javascript" onclick="return cball_onclick()" title="<%=arrTitles(20)%>" id="cb_all" name="cb_all"></td>
	    
    </tr>  
    	<%
    	
	  PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
   If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
     	PageSize = 10
   End If	
  
  Set tasksList = Server.CreateObject("ADODB.RECORDSET")
   sqlstr = "exec dbo.get_tasks_paging " & Page & "," & PageSize & ",'" & sFix(search_company) & "','" & sFix(search_contact) & "','" & sFix(search_project) & "','" & status & "','" & UserID & "','" & OrgID & "','" & lang_id & "','" & taskTypeID & "','" & reciver_id & "','" & sender_id & "','" & sortby(sort) & "','" & start_date_ & "','" & end_date_ & "'"
   'Response.Write sqlstr
   'Response.End
   set  tasksList = con.getRecordSet(sqlstr)
   current_date = Day(date()) & "/" & Month(date()) & "/" & Year(date())
   dim	IS_DESTINATION 
   ids = ""
   If not tasksList.EOF then		
        recCount = tasksList("CountRecords")		
		       
   while not tasksList.EOF       
       taskId = trim(tasksList(1))   :  ids = ids & taskId 		
       companyID =  trim(tasksList(2))
       contactID = trim(tasksList(3))
       companyName = trim(tasksList(4))      
       contactName = trim(tasksList(5))
       task_date = trim(tasksList(6))
       projectName = trim(tasksList(7))  
       status = trim(tasksList(8))       
	   sender_name = trim(tasksList(9))
	   reciver_name = trim(tasksList(10))      
       task_content = trim(tasksList(11))          
       parentID = trim(tasksList(12))  
       childID = trim(tasksList(15))  
       task_replay = trim(tasksList(16))  
       
       If Len(task_content) > 25 Then
			task_content_short = Left(task_content,22) & "..."
	   Else
		   task_content_short = task_content 		
       End If
       
       If Len(task_replay) > 25 Then
			task_replay_short = Left(task_replay,22) & "..."
	   Else
		   task_replay_short = task_replay 		
       End If       
     
       If IsDate(trim(task_date)) Then
		  task_date = Day(trim(task_date)) & "/" & Month(trim(task_date)) & "/" & Right(Year(trim(task_date)),2)
		  if DateDiff("d",task_date,current_date) >= 0 then
			IS_DESTINATION = true
		  else
			IS_DESTINATION = false
		  end if
	   Else
		   task_date=""
		   IS_DESTINATION = false
	   End If     				
	   task_types_n = ""
	   sqlstr = "exec dbo.get_task_types '"&taskID&"','"&OrgID&"'"
	   set rs_task_types = con.getRecordSet(sqlstr)
	   If not rs_task_types.eof Then
		    task_types_n = rs_task_types.getString(,,"<br>","<br>")
	   Else
			task_types_n = ""
	   End If
%>
    <tr>
       <td align="center" class="card<%=class_%>" valign="top">
           <%If trim(taskID) <> "" And trim(childID) <> "" Then%>
           <input type=image src="../../images/hets4.gif" onclick='window.open("task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' align="absmiddle" ID="Image1" NAME="Image1">
           <%End If%>
           <%If trim(parentID) <> "" And trim(taskID) <> "" Then%>
           <input type=image src="../../images/hets4a.gif" onclick='window.open("task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' align="absmiddle" ID="Image2" NAME="Image2">
           <%End If%>
           </td>  
	      <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>" onMouseover="EnterContent('ToolTip','<%=arrTitles(4)%>','<%=Escape(breaks(task_replay))%>','<%=dir_obj_var%>'); Activate();" onMouseout="deActivate()">&nbsp;<%=breaks(trim(task_replay_short))%>&nbsp;</td>	      	      	      
	      <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>" onMouseover="EnterContent('ToolTip','<%=arrTitles(4)%>','<%=Escape(breaks(task_content))%>','<%=dir_obj_var%>'); Activate();" onMouseout="deActivate()">&nbsp;<%=breaks(trim(task_content_short))%>&nbsp;</td>	      	      	      
          <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>><%=task_types_n%></a></td>	
		  <!--td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=projectName%>&nbsp;</a></td-->
          <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=contactName%>&nbsp;</a></td>
          <%If trim(is_companies) = "1" Then%>
          <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=companyName%>&nbsp;</a></td>
          <%End If%>
          <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=reciver_name%>&nbsp;</a></td>
          <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=sender_name%>&nbsp;</a></td>
          <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%> style="font-size: 8pt;" <%if IS_DESTINATION and status <> 3 then%> name=word27 title="<%=arrTitles(27)%>"><span style="width:9px;COLOR: #FFFFFF;BACKGROUND-COLOR: #FF0000;text-align:center"><b>!</b></span><%else%>><%end if%><%If isDate(task_date) Then%>&nbsp;<%=task_date%>&nbsp;<%End If%></a></td>         
	      <td class="card<%=class_%>" valign="top" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num<%=status%>" <%=href%>><%=arr_Status(status)%></a></td>	  
  <td class="card<%=class_%>" align="center" valign="top"><%If trim(contactID) <> "" And trim(contactID) <> "0" Then%><img src="../../images/forms_icon.gif" border="0" 
			style="cursor: pointer;"	alt="היסטוריית טפסים של הלקוח" 
			onclick="window.open('../appeals/contact_appeals.asp?contactID=<%=contactID%>','winCA','top=20, left=10, width=950, height=500, scrollbars=1');"  ><%End If%></td>
	       <td class="card<%=class_%>" align="center" valign="top"><INPUT type="checkbox" id="cb<%=taskId%>" name="cb<%=taskId%>"></td>
	    </tr>
      <%	tasksList.MoveNext
		if not tasksList.eof then
		ids = ids & ","
		end if	  
	  Wend
	  
	  NumberOfPages = Fix((recCount / PageSize)+0.9)
	    if NumberOfPages > 1 then
	  urlSort = urlSort & "&amp;sort=" & sort 	  %>
	  <tr>
		<td width="100%" align="center" colspan="12" nowrap class="card<%=class_%>">
			<table border="0" cellspacing="0" cellpadding="2" dir="ltr" ID="Table2">
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" align="<%=align_var%>"><A class="pageCounter" name=word28 title="<%=arrTitles(28)%>" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  align="<%=align_var%>"><A class="pageCounter" href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(29)%>" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>
	<!--tr><td  class="title_form" align="center" colspan=9><%=recCount%> :ואצמנ תומושר כ"הס</td></tr-->										 
	<%End If%> 
	<tr>
	   <td colspan="12" height=18 class="card<%=class_%>" align="center" dir="ltr" style="color:#6E6DA6;font-weight:600"><!--נמצאו--><%=arrTitles(10)%>&nbsp;<%=recCount%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%> &nbsp;</td>
	</tr>
	<%Else%>
	<tr><td colspan="12" class="card<%=class_%>" align="center">&nbsp;&nbsp;</td></tr>
<% End If
   set tasksList = Nothing%>


     </table></td></tR>
</table><input type="hidden" id="ids" value="<%=ids%>" NAME="ids">
<input type="hidden" name="trapp" value="" ID="trapp">
</form>
<%'end if%>
<script language="javascript" type="text/javascript">
<!--
document.body.onmousemove = overhere;
//-->
</script>
</body>
</html>