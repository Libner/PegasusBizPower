<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%  
  projectID = trim(Request("projectID"))  
  companyID = trim(Request("companyID")) 
  
  If trim(lang_id) = "1" Then
	arr_status_pr = Array("","עתידי","בביצוע","סגור")
  Else
	arr_status_pr = Array("","new","active","close")	
  End If	
  
  found_project = false 
  sqlStr = "SELECT COMPANY_ID, PROJECT_CODE, PROJECT_NAME, PROJECT_DESCRIPTION, START_DATE, END_DATE, "&_
  " STATUS FROM PROJECTS WHERE project_ID=" & projectID & " AND ORGANIZATION_ID = " & OrgID  
  ''Response.Write sqlStr
  set rs_projects = con.GetRecordSet(sqlStr)
  if not rs_projects.eof then
       If trim(rs_projects("company_id")) <> "" And trim(rs_projects("company_id")) <> "0" Then
			companyID = rs_projects("company_id")
			project_title = trim(Request.Cookies("bizpegasus")("Projectone"))
			page_title = trim(Request.Cookies("bizpegasus")("ProjectMulti"))
			numOfLink = 2
			topLevel2 = 11 'current bar ID in top submenu - added 03/10/2019
	   Else
			project_title = trim(Request.Cookies("bizpegasus")("ActivitiesOne"))
			page_title = trim(Request.Cookies("bizpegasus")("ActivitiesMulti"))
			numOfLink = 3
			topLevel2 = 12 'current bar ID in top submenu - added 03/10/2019
	   End If
	   project_code = rs_projects("project_code")
	   project_name = rs_projects("project_name")
	   project_description = rs_projects("project_description")
	   dateStart = rs_projects("start_date")
	   dateEnd = rs_projects("end_date")
	   project_status = rs_projects("status")	
	   found_project = true
  else
	   found_project = false	       
  end if
  set rs_projects = nothing 
 
  If trim(companyID) <> "" Then
  sqlStr = "SELECT company_name, address, zip_code, city_Name, phone1, phone2, fax1, fax2, url, email"&_
  ", date_update, status, company_desc, private_flag FROM companies WHERE company_id="& companyID 
  'Response.Write "test"
  'Response.End
  set pr=con.GetRecordSet(sqlStr)
  if not pr.EOF then	
	companyName   =pr("company_name")
	address	      =pr("address")
	cityName	  =pr("city_Name")	
	zip_code	  =pr("zip_code")	
	phone1	      =pr("phone1")
	phone2	      =pr("phone2")
	fax           =pr("fax1")
	fax2	      =pr("fax2")
	company_email =pr("email")
	url	          =pr("url")	
	date_update	  =pr("date_update")			   
    status_company=pr("status")
	company_desc  =pr("company_desc")
	private_flag  = pr("private_flag")
  	If Len(company_desc) > 43 Then
  		company_desc_short = Left(company_desc,40) & ".."
  	Else
  		company_desc_short = company_desc
  	End If    
  end if
  set pr = nothing  
  End If 
 
  if trim(lang_id) = "1" then
	   arr_status_comp = Array("","עתידי","פעיל","סגור","פונה")
  else
	   arr_status_comp = Array("","new","active","close","appeal")
  end if  
  
  	'עדכן סטטוס פגישות שעברו
	sqlstr = " execute UpdateMeetingStatus '" & OrgID & "'"
    'Response.Write sqlstr
    con.ExecuteQuery(sqlstr)   

	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 10 Order By word_id"				
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
	
    delMechID=trim(Request("delMechID"))  
    If delMechID<>nil And delMechID<>"" Then 	
		con.ExecuteQuery "Delete From Hours WHERE Mechanism_Id = " & delMechID	
		con.executeQuery "Delete From Mechanism WHERE Mechanism_Id = " &  delMechID		
        Response.Redirect "project.asp?companyId=" & companyId & "&projectID=" & projectID
    End If    	
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
var oPopup = window.createPopup();
function appealDropDown(obj)
{
    oPopup.document.body.innerHTML = appeal_Popup.innerHTML; 
    oPopup.document.charset="windows-1255";    
    oPopup.show(0, 17, obj.offsetWidth, 68, obj);    
    
}

function taskDropDown(obj)
{
    oPopup.document.body.innerHTML = task_Popup.innerHTML; 
    oPopup.document.charset="windows-1255";
    oPopup.show(0, 17, obj.offsetWidth, 68, obj);    
}

function DoCal(elTarget){
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

var comp_list=null;
function openWindow()
{
	comp_list = window.open("companies_list.asp","List","width=500,height=380,top=50,left=50");		
}

function addtask(projectID,companyID,taskID)
{
	h = parseInt(530);
	w = parseInt(470);
	window.open("../tasks/addtask.asp?project_id=" + projectID + "&companyId=" + companyID + "&taskID=" + taskID, "AddTask" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
}

function closeTask(projectID,companyID,taskID)
{
	h = parseInt(480);
	w = parseInt(470);
	window.open("../tasks/closetask.asp?project_id=" + projectID + "&companyId=" + companyID + "&taskId=" + taskID, "CloseTask" ,"scrollbars=1,toolbar=0,top=20,left=100,width="+w+",height="+h+",align=center,resizable=0");
}

function addMechanism(projectID,companyID,mechId)
{
	h = parseInt(250);
	w = parseInt(500);
	window.open("addmechanism.asp?projectID=" + projectID + "&companyID=" + companyID + "&mechId=" + mechId, "AddMech" ,"scrollbars=1,toolbar=0,top=100,left=150,width="+w+",height="+h+",align=center,resizable=0");
}

function CheckDelMech(projectID,companyID,mechId)
{
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את המנגנון"
     Else
		str_confirm = "Are you sure want to delete the mechanism?"
     End If   
  %>
  if (window.confirm("<%=str_confirm%>"))
  {         
     document.location.href = "project.asp?projectID="+projectID+"&companyId="+companyID+"&delMechID="+mechId;
    
  }
  return false;  	
}

function descDropDown(obj)
{
  document.all["div_comp_desc"].style.zIndex=11; 
  document.all["div_comp_desc"].style.display='inline';
  document.all["div_comp_desc"].style.visibility="visible"; 
  return false; 
}

function closeDesc()
{
   document.all["div_comp_desc"].style.display='none';
   document.all["div_comp_desc"].style.visibility="hidden"; 
   return false;
}
	function addmeeting(meetingID)
	{
		h = parseInt(500);
		w = parseInt(465);
		window.open("../meetings/addmeeting.asp?meetingID=" + meetingID + "&meeting_date=<%=Date()%>&participant_id=<%=UserID%>&companyId=<%=companyID%>&project_id=<%=projectID%>", "AddMeeting" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}
	function closemeeting(meetingID)
	{
		h = parseInt(490);
		w = parseInt(465);
		window.open("../meetings/addmeeting.asp?meetingID=" + meetingID + "", "CloseMeeting" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}			
//-->
</script>  
</head>
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF>
<tr><td width="100%">
<!-- #include file="../../logo_top.asp" -->
</td></tr>
<%numOftab = 0%>
   <%topLevel2 = 11 'current bar ID in top submenu - added 03/10/2019%>
<tr><td width="100%">
<!--#include file="../../top_in.asp"-->
</td></tr>
<%If found_project Then%>
<tr><td width="100%" class="page_title" dir="<%=dir_obj_var%>"><font color="#6F6DA6"><%=project_title%></font>&nbsp;<%=project_name%>&nbsp; >> &nbsp;<%If trim(companyID)<>"" then%><a class=normalB <%If private_flag = "0" Then%>href="../companies/company.asp?companyID=<%=companyID%>"<%Else%>nohref<%End If%> target=_self><font color="#6F6DA6"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></font>&nbsp;<%=companyName%></a><%End If%></td></tr>
<tr><td width="100%" valign=top>
<table cellpadding=0 cellspacing=1 width=100% bgcolor=white dir="<%=dir_var%>">
<tr><td width="100%" valign=top>
<table cellpadding=0 bgcolor="#808080"  dir="ltr" cellspacing=1 width=100% border=0 style="border-collpase:collapse;" ID="Table2">
<tr>          
  <td align="<%=align_var%>" width="50%" bgcolor=#E6E6E6 nowrap valign="top">  
  <input type="hidden" name="companyID" id="companyID" value="<%=companyID%>">  
  <input type="hidden" name="projectID" id="projectID" value="<%=projectID%>">  
   <table border="0" cellpadding="0" cellspacing="1" width="100%" dir="<%=dir_var%>">
   <tr><td colspan=2 class="title_form" align="<%=align_var%>"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;</td></tr>
   <tbody style="padding-right:15px">  
      <tr><td colspan=2 height=5 nowrap></td></tr> 
      <tr>
         <td width="100%" align="<%=align_var%>"><span dir="<%=dir_obj_var%>" style="width:250" class="Form_R"><%=trim(companyName)%></span></td>
         <td width="100" nowrap align="<%=align_var%>">&nbsp;<!--שם--><%=arrTitles(4)%>&nbsp;</td> 
      </tr>           
      <tr>
        <td align="<%=align_var%>"><span dir="<%=dir_obj_var%>" style="width:250" class="Form_R"><%=trim(address)%></span></td>
        <td align="<%=align_var%>">&nbsp;<!--כתובת--><%=arrTitles(5)%>&nbsp;</td>
     </tr>             
     <tr>
         <td align="<%=align_var%>" nowrap>                    
           <span dir="<%=dir_obj_var%>" class="Form_R" style="width:130"><%=cityName%></span>
         </td>
         <td align="<%=align_var%>">&nbsp;<!--עיר--><%=arrTitles(6)%>&nbsp;</td>
      </tr> 
      <tr>
         <td align="<%=align_var%>" nowrap><span dir="ltr" class="Form_R" style="width:130"><%=zip_code%></span></td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--מיקוד--><%=arrTitles(58)%>&nbsp;</td>
      </tr>                              
      <tr>
          <td align="<%=align_var%>" dir=ltr>             
			<span dir="ltr" class="Form_R" style="width:130"><%=trim(phone1)%></span>
           </td>
          <td align="<%=align_var%>">&nbsp;<!--1 טלפון--><%=arrTitles(7)%>&nbsp;</td>
      </tr>                                
      <tr>  
          <td align="<%=align_var%>" dir=ltr>            
			<span dir="ltr" class="Form_R" style="width:130"><%=trim(phone2)%></span>
          </td>
          <td align="<%=align_var%>">&nbsp;<!--2 טלפון--><%=arrTitles(8)%>&nbsp;</td>
      </tr>     
      <tr>  
          <td align="<%=align_var%>" dir=ltr>                       
			<span dir="ltr" class="Form_R" style="width:130"><%=trim(fax)%></span>
           </td>
          <td align="<%=align_var%>">&nbsp;<!--פקס--><%=arrTitles(9)%>&nbsp;</td>
      </tr>                          
      <tr>
         <td align="<%=align_var%>">
         <%If Len(company_email) > 0 Then%>
         <a href="mailto:<%=company_email%>" dir="<%=dir_obj_var%>" style="width:250;direction:ltr;text-align:left" class="Form_R hand">
         <%=trim(company_email)%></a>
		 <%Else%>
         <span style="width:250" class="Form_R"></span>         
         <%End If%>
         </td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;Email&nbsp;</td>
      </tr>       
      <tr>
         <td align="<%=align_var%>">
         <%If Len(url) > 0 Then%>
         <a href="<%=url%>" target=_blank dir="<%=dir_obj_var%>" style="width:250;direction:ltr;text-align:left"  class="Form_R hand">         
         <%=trim(url)%></a>
         <%Else%>
         <span style="width:250" class="Form_R"></span>
         <%End If%>
         </td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<span id="word10" name=word10><!--אתר--><%=arrTitles(10)%></span>&nbsp;</td>
      </tr> 
      <tr>
           <td align="<%=align_var%>">
           <span class="Form_R" dir="<%=dir_obj_var%>" style="width:250;line-height:16px"><%=types%></span>
           </td>
           <td align="<%=align_var%>" style="padding-right:15px" valign=top>&nbsp;<span id="word11" name=word11><!--קבוצה--><%=arrTitles(11)%></span>&nbsp;</td>
      </tr> 
      <tr>
      <td align="<%=align_var%>" valign=bottom>
      <%If Len(trim(company_desc)) > 0 Then%><input type=image src="../../images/popup.gif" onclick="return descDropDown(this)" name=word51 title="<%=arrTitles(51)%>" border=0 hspace=0 vspace=0 ID="Image2" NAME="Image2"><%End if%>
      <span class="Form_R" dir="<%=dir_obj_var%>" style="width:250;line-height:16px"><%=trim(company_desc_short)%></span>      
      </td>
      <td align="<%=align_var%>" style="padding-right:15px" valign=top>&nbsp;<span id="word12" name=word12><!--פרטים נוספים--><%=arrTitles(12)%></span>&nbsp;</td>
      </tr>            
      <tr>
           <td align="<%=align_var%>"><span style="width:60;text-align:center" class="status_num<%=status_company%>">
           <%=arr_status_comp(status_company)%></span>
           </td>
           <td align="<%=align_var%>" valign=top>&nbsp;<span id="word13" name=word13><!--סטטוס--><%=arrTitles(13)%></span>&nbsp;</td>
      </tr> 
      <tr><td colspan="2" height="5" nowrap></td></tr> 
      </tbody>                        
</table>
</td>
<td width=50% nowrap valign="top"  bgcolor=E6E6E6>
   <table border="0" cellpadding="0" cellspacing="1" width="100%" align=center dir="<%=dir_var%>">   
   <tr><td colspan=7 class="title_form" align="<%=align_var%>"><%=project_title%>&nbsp;</td></tr> 
   <tbody style="padding-right:15px" dir="<%=dir_var%>">
   <tr><td colspan=2 height=5 nowrap></td></tr>
      <tr>
         <td width="100%" align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">           
         <span style="width:250" class="Form_R" dir="<%=dir_obj_var%>"><%=project_name%></span>      
         </td>
         <td width="100" nowrap align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir="<%=dir_obj_var%>"><!--שם--><%=arrTitles(25)%></td>
      </tr>      
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">                             
	       <span class="Form_R" dir="<%=dir_obj_var%>" style="width:250;line-height:16px"><%=breaks(trim(project_description))%></span>
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top"><!--תיאור--><%=arrTitles(17)%></td>
      </tr>
      <!-- start fields dynamics -->		
	  <!--#INCLUDE FILE="project_fields.asp"-->	
	  <!-- end fields dynamics -->	
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">                        
         <span dir="<%=dir_obj_var%>" style="width:250" class="Form_R"><%=trim(project_code)%></span>
         </td>                    
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top"><!--קוד--><%=arrTitles(18)%></td>
      </tr> 
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">                        
          <span dir="ltr" style="width:80" class="Form_R"><%=trim(dateStart)%></span>
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top"><!--תאריך פתיחה--><%=arrTitles(26)%></td>
      </tr> 
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">                        
          <span dir="ltr" style="width:80" class="Form_R"><%=trim(dateEnd)%></span>
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top"><!--תאריך סגירה--><%=arrTitles(27)%></td>
      </tr> 
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top"><span style="width:80;text-align:center" class="status_num<%=project_status%>"><%=arr_status_pr(project_status)%></span></td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top"><!--סטטוס--><%=arrTitles(41)%></td>
      </tr> 
      <tr><td colspan="2" height="5" nowrap></td></tr>    
      </tbody>     
      <tr><td colspan=2 class="title_form" align="<%=align_var%>">&nbsp;<!--מנגנונים--><%=arrTitles(66)%>&nbsp;</td></tr>
      <tr><td colspan=2>
      <table cellpadding=2 cellspacing=1 width=100% align=center bgcolor=white dir="<%=dir_var%>">      
      <%urlSort="project.asp?companyID="& companyID & "&projectID="&projectID
		PageSize = 5
		if trim(Request.QueryString("pageMech"))<>"" then
			pageMech=Request.QueryString("pageMech")
		else
			pageMech=1
		end if  
		 
		if trim(Request.QueryString("numOfRowMech"))<>"" then
			numOfRowMech=Request.QueryString("numOfRowMech")
		else
			numOfRowMech = 1
		end if  

		sql="SELECT mechanism_id,mechanism_name,(Select Count(mechanism_id) FROM hours WHERE mechanism_id=mechanism.mechanism_id And organization_id = " & trim(OrgID) & ") FROM mechanism "&_
		" WHERE project_id = " & projectID & " AND ORGANIZATION_ID = " & OrgID
		If trim(companyID) <> "" Then
		   sql=sql&" And company_id = " & companyID
		End If
		sql=sql & " ORDER BY mechanism_name"		  
		set listMech=con.GetRecordSet(sql)
		if not listMech.EOF then
			listMech.PageSize=PageSize
			listMech.AbsolutePage=pageMech
			recCount=listMech.RecordCount 
			NumberOfPages=listMech.PageCount
			i=1	
		%>
		<tr>     
		    <td nowrap class="title_sort" align="<%=align_var%>"><%=arrTitles(40)%></td>     
			<td nowrap class="title_sort" align="<%=align_var%>"><%=arrTitles(39)%></td>     
			<td width="100%" align="<%=align_var%>" class="title_sort"><!--מנגנון--><%=arrTitles(65)%></td>
		</tr> 
		<% do while (not listMech.EOF and i<=listMech.PageSize)
			mechID=trim(listMech(0))			
			mechName=trim(listMech(1))
			countHours=trim(listMech(2))
	           
		%>      
			<tr>            
			    <td class="card" align="center"><%If cInt(countHours) = 0 Then%><input type=image src="../../images/delete_icon.gif" style="border: none" hspace=0 vspace=0 onclick="return CheckDelMech('<%=projectID%>','<%=companyID%>','<%=mechID%>')"><%Else%><input type=image SRC="../../images/delete_icon.gif" BORDER=0 Onclick="window.alert('<%=Space(24)%>שים לב, לא ניתן למחוק את המנגנון\n\nמפני שקיימות שעות עבודה על מנגנון זה במערכת\n\n<%=Space(38)%>על מנת למחוק את המנגנון\n\n<%=Space(11)%>אליך למחוק תחילה את שעות העבודה הנל');return false;"><%End If%></td>			    
				<td class="card" align="center"><input type=image src="../../images/edit_icon.gif" style="border: none" hspace=0 vspace=0 onclick="addMechanism('<%=projectID%>','<%=companyID%>','<%=mechID%>')"></td>         
				<td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=mechName%>&nbsp;</td>
			</tr>
		<%  listMech.MoveNext
			i=i+1
			loop
			if NumberOfPages > 1 then
		%>
	  <tr>
		<td width="100%" align=middle colspan="8" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr>               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>	         
	         <tr>
	         <%if numOfRowMech <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word53 title="<%=arrTitles(53)%>" href="<%=urlSort%>&pageMech=<%=10*(numOfRowMech-1)-9%>&amp;numOfRowMech=<%=numOfRowMech-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowMech-1)) <= NumberOfPages Then
	                  if CInt(pageMech)=CInt(i+10*(numOfRowMech-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRowMech-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&pageMech=<%=i+10*(numOfRowMech-1)%>&amp;numOfRowMech=<%=numOfRowMech%>" ><%=i+10*(numOfRowMech-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRowMech) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word54 title="<%=arrTitles(54)%>" href="<%=urlSort%>&pageMech=<%=10*(numOfRowMech) + 1%>&amp;numOfRowMech=<%=numOfRowMech+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	<%listMech.close 
	set listMech=Nothing
	End If
	End If
	%>
      </table></td></tr>
</table>
</td>
</tr>
<%If trim(projectID)<>"" Then   
   
		if lang_id = "1" then
			arr_Status = Array("","חדש","בטיפול","סגור")	
		else
			arr_Status = Array("","new","active","close")	
		end if
	        
		if lang_id = "1" then
			arr_StatusT = Array("","חדש","בטיפול","סגור")	
        else
			arr_StatusT = Array("","new","active","close")	
        end if
        	        
		dim sortby_task(13)			
		sortby_task(1) = "CONTACT_NAME"
		sortby_task(2) = "CONTACT_NAME DESC "
		sortby_task(3) = "task_date"
		sortby_task(4) = "task_date DESC"
		sortby_task(5) = "task_status,task_date DESC"
		sortby_task(6) = "task_status DESC,task_date DESC"
		sortby_task(7) = "U.FIRSTNAME, U.LASTNAME"
		sortby_task(8) = "U.FIRSTNAME DESC, U.LASTNAME DESC"
		sortby_task(9) = "U1.FIRSTNAME,  U1.LASTNAME"
		sortby_task(10) = "U1.FIRSTNAME DESC,  U1.LASTNAME DESC"
		sortby_task(11) = "company_Name"
		sortby_task(12) = "company_Name DESC "	
      
    sort_task = trim(Request.QueryString("sort_task"))
    If Len(sort_task) = 0 Then
		sort_task = 5
	Else sort_task = trim(Request.QueryString("sort_task"))	
    End If    
  	
	If trim(Request("taskTypeID")) <> nil Then
		  taskTypeID = trim(Request("taskTypeID"))
		  search = true 
	Else  taskTypeID = ""
		  search = false	
	End If	          

	if trim(Request.QueryString("page_task"))<>"" then
		page_task=Request.QueryString("page_task")
	else
		page_task=1
	end if 
	
	task_status = trim(Request("task_status")) ' הצגת כל סטטוסי המשימות   
	if trim(Request.QueryString("row_task"))<>"" then
		row_task=Request.QueryString("row_task")
	else
		row_task = 1
	end if  
	PageSize = 5
	task_status = trim(Request("task_status")) ' הצגת כל סטטוסי המשימות 		
	title_tasks = trim(Request.Cookies("bizpegasus")("TasksMulti"))		
  
	sqlstr = "EXECUTE get_tasks_paging " & page_task & "," & PageSize & ",'','','','" & task_status & "','" & UserID & "','" & OrgID & "','" & lang_id & "','" & taskTypeID & "','','','" & sortby_task(sort_task) & "','','','" & companyID & "','','" & projectID & "'"
	'Response.Write sqlstr
	'Response.End   
	set tasksList = con.getRecordSet(sqlstr)
	urlSort = urlSort & "&allTasks=" & allTasks
    If not tasksList.eof Then
		recCount = tasksList("CountRecords")
	End If		
    If not tasksList.eof Or search = true Then           
%>
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">  
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100%>
  <tr><td width=100%><A name="table_tasks"></A>  
  <table cellpadding=0 cellspacing=0 width=100% border=0>
  <tr>  
  <td class="title_form" width=100%>&nbsp;<%=title_tasks%>&nbsp;<font color="#E6E6E6">(<%=companyName & " - " & project_name%>)</font>&nbsp;</td>
  </tr>
  </table></td></tr>  	  
  <tr>
  <td width="100%" valign=top>
   <table width=100% dir="<%=dir_var%>" border=0 bordercolor=red cellpadding=0 cellspacing=1 bgcolor="#FFFFFF">   	    
     <tr> 
      <td align=center class="title_sort" width=26 nowrap>&nbsp;</td>     
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort" width=215 nowrap><span id=word19 name=word19><!--תוכן--><%=arrTitles(19)%></span>&nbsp;</td>                  
      <td width=120 nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>" id=td_types name=td_types>&nbsp;<span id="word20" name=word20><!--סוגים--><%=arrTitles(20)%></span>&nbsp;<IMG name=word52 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(52)%>" align=absmiddle onmousedown="taskDropDown(td_types)"></td>
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=90 nowrap class="title_sort<%if trim(sort_task)="1" OR trim(sort_task)="2" then%>_act<%end if%>"><%if trim(sort_task)="1" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word53 title="<%=arrTitles(53)%>"><%elseif trim(sort_task)="2" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word54 title="<%=arrTitles(54)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=1#table_tasks" name=word54 title="<%=arrTitles(54)%>"><%end if%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="1" then%>bot<%elseif trim(sort_task)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=100 nowrap class="title_sort<%if trim(sort_task)="11" OR trim(sort_task)="12" then%>_act<%end if%>"><%if trim(sort_task)="11" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" title="<%=arrTitles(53)%>"><%elseif trim(sort_task)="12" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word54 title="<%=arrTitles(54)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=11#table_tasks" name=word54 title="<%=arrTitles(54)%>"><%end if%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="11" then%>bot<%elseif trim(sort_task)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>
      <td width="90" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_task)="9" OR trim(sort_task)="10" then%>_act<%end if%>"><%if trim(sort_task)="9" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>" name=word53 title="<%=arrTitles(53)%>"><%elseif trim(sort_task)="10" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>" name=word54 title="<%=arrTitles(54)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=9" name=word54 title="<%=arrTitles(54)%>"><%end if%><span id="word21" name=word21><!--אל--><%=arrTitles(21)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="9" then%>bot<%elseif trim(sort_task)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
      <td width="90" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_task)="7" OR trim(sort_task)="8" then%>_act<%end if%>"><%if trim(sort_task)="7" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>" name=word53 title="<%=arrTitles(53)%>"><%elseif trim(sort_task)="8" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>" name=word54 title="<%=arrTitles(54)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=7" name=word54 title="<%=arrTitles(54)%>"><%end if%><span id="word22" name=word22><!--מאת--><%=arrTitles(22)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="7" then%>bot<%elseif trim(sort_task)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>      
      <td align="<%=align_var%>" dir="ltr" width=80 nowrap class="title_sort<%if trim(sort_task)="3" OR trim(sort_task)="4" then%>_act<%end if%>"><%if trim(sort_task)="3" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word53 title="<%=arrTitles(53)%>"><%elseif trim(sort_task)="4" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word54 title="<%=arrTitles(54)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=3#table_tasks" name=word54 title="<%=arrTitles(54)%>"><%end if%><span id="word23" name=word23><!--תאריך יעד--><%=arrTitles(23)%></span><img src="../../images/arrow_<%if trim(sort_task)="3" then%>bot<%elseif trim(sort_task)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>          
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=45 nowrap class="title_sort<%if trim(sort_task)="5" OR trim(sort_task)="6" then%>_act<%end if%>"><%if trim(sort_task)="5" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word53 title="<%=arrTitles(53)%>"><%elseif trim(sort_task)="6" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word54 title="<%=arrTitles(54)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=5#table_tasks" name=word54 title="<%=arrTitles(54)%>"><%end if%>&nbsp;<span id="word24" name=word24><!--'סט--><%=arrTitles(24)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="5" then%>bot<%elseif trim(sort_task)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>
    </tr>    
 <%	  current_date = Day(date()) & "/" & Month(date()) & "/" & Year(date())
	  dim	 IS_DESTINATION
      task_types_name = ""
      while not tasksList.EOF       
		taskId = trim(tasksList(1))		
		company_Name = trim(tasksList(4))      
		contact_Name = trim(tasksList(5))
		task_date = trim(tasksList(6))
		projectName = trim(tasksList(7))  
		task_status = trim(tasksList(8))	 
		sender_name = trim(tasksList(9))
		reciver_name = trim(tasksList(10))      
		task_content = trim(tasksList(11))          
		parentID = trim(tasksList(12))  
		ReciverID = trim(tasksList(13))
	    SenderID = trim(tasksList(14))
	    childID = trim(tasksList(15)) 	
		
		task_types_names=""	    
		sqlstr = "Exec dbo.get_task_types '"&taskID&"','"&OrgID&"'"
		set rs_task_types = con.getRecordSet(sqlstr)
		If not rs_task_types.eof Then
			task_types_names = rs_task_types.getString(,,",",",")
		Else
			task_types_names = ""
		End If		
		
		If Len(task_types_names) > 0 Then
			task_types_names = Left(task_types_names,(Len(task_types_names)-1))
		End If
		
		If cInt(strScreenWidth) > 800 Then
			numOfLetters = 150
		Else
			numOfLetters = 55
		End If
		tel_text = trim(tasksList("task_content"))
		If Len(tel_text) > numOfLetters Then
			tel_text_short = Left(tel_text , numOfLetters-2) & ".."
		Else tel_text_short = tel_text	
		End If
		task_date = trim(tasksList("task_date"))
		If isDate(task_date) Then
			d_s = Day(task_date) & "/" & Month(task_date) & "/" & Right(Year(task_date),2)
			if DateDiff("d",d_s,current_date) >= 0 then
				IS_DESTINATION = true
			else
				IS_DESTINATION = false
			end if
		else
			d_s = ""
			IS_DESTINATION = false  
		End If
		If trim(UserID) = trim(SenderID) Then
			class_ = "4"
		ElseIf trim(UserID) = trim(ReciverID) Then
			class_ = "7"
		Else
			class_ = ""	
	    End if	
	    If trim(UserID) = trim(SenderID) AND trim(task_status) = "1" Then
			href = "href=""javascript:addtask('" & projectID & "','" & companyID & "','" & taskID & "')"""   
        Else      
			href = "href=""javascript:closeTask('" & projectID & "','" & companyID & "','" & taskID & "')"""     
        End If        
      %>      
      <tr>  
		<td align=center class="card<%=class_%>" valign=middle>
		<%If trim(taskID) <> "" And trim(childID) <> "" Then%>
		<input type=image src="../../images/hets4.gif" border=0 hspace=0 vspace=0 onclick='window.open("../tasks/task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' ID="Image1" NAME="Image1">
		<%End If%>
		<%If trim(parentID) <> "" And trim(taskID) <> "" Then%>
		<input type=image src="../../images/hets4a.gif" border=0 hspace=0 vspace=0 onclick='window.open("../tasks/task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' ID="Image4" NAME="Image4">
		<%End If%>
		</td>              
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>" title="<%=vFix(tel_text)%>">&nbsp;<%=breaks(trim(tel_text_short))%>&nbsp;</a></td>
		<td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=task_types_names%>&nbsp;</a></td>
		<td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=CONTACT_NAME%>&nbsp;</a></td>
		<td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=Company_NAME%>&nbsp;</a></td>
		<td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=reciver_name%>&nbsp;</a></td>                 
		<td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=sender_name%>&nbsp;</a></td>                         
		<td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> <%if IS_DESTINATION and task_status <> 3 then%> name=word55 title="<%=arrTitles(55)%>"><span style="width:9px;COLOR: #FFFFFF;BACKGROUND-COLOR: #FF0000;text-align:center"><b>!</b></span><%else%>><%end if%>&nbsp;<%=d_s%>&nbsp;</a></td>            
		<td align=center class="card<%=class_%>" valign=top><a class="task_status_num<%=task_status%>" style="line-height:120%;"><%=arr_Status(task_status)%></A></td>
      </tr>
<%
    tasksList.MoveNext
	Wend
	  
	NumOfPagesTasks = Fix((recCount / PageSize)+0.9)
    urlSort = urlSort & "&sort_task=" & sort_task
	if NumOfPagesTasks > 1 then	
%>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr>               
	        <% If NumOfPagesTasks > 10 Then 
	              num = 10 : NumOfRowstask = cInt(NumOfPagesTasks / 10)
	           else num = NumOfPagesTasks : NumOfRowstask = 1    	                      
	           End If	         
            %>
	         
	         <tr>
	         <%if row_task <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word56 title="<%=arrTitles(56)%>" href="<%=urlSort%>&page_task=<%=10*(row_task-1)-9%>&amp;row_task=<%=row_task-1%>#table_tasks" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(NumOfRowstask-1)) <= NumOfPagesTasks Then
	                  if CInt(page_task)=CInt(i+10*(row_task-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(row_task-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&page_task=<%=i+10*(row_task-1)%>&amp;row_task=<%=row_task%>#table_tasks" ><%=i+10*(row_task-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumOfPagesTasks > cint(num * row_task) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word57 title="<%=arrTitles(57)%>" href="<%=urlSort%>&page=<%=10*(row_task) + 1%>&amp;row_task=<%=row_task+1%>#table_tasks" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	<%	End If %>
</table></td></tr>
<%If tasksList.recordCount = 0 Then%>
<tr><td align=center class=card1>&nbsp;</td></tr>									
<%End If%>	 	
</table></td></tr> 
	<%End If%>   
 <%set tasksList = Nothing%>
<%End If%>       
<%
	dim sortby_app(16)	
	sortby_app(1) = "appeal_date"
	sortby_app(2) = "appeal_date DESC"
	sortby_app(3) = "appeal_id"
	sortby_app(4) = "appeal_id DESC"
	sortby_app(5) = "User_Name"
	sortby_app(6) = "User_Name DESC"
	sortby_app(7) = "product_name,appeal_date"
	sortby_app(8) = "product_name DESC,appeal_date"
	sortby_app(9) = "CONTACT_NAME"
	sortby_app(10) = "CONTACT_NAME DESC"
	sortby_app(11) = "status_order, appeal_date"
	sortby_app(12) = "status_order DESC,appeal_date"
	sortby_app(13) = "Company_NAME,appeal_date"
	sortby_app(14) = "Company_NAME DESC,appeal_date"

	sort_app = Request("sort_app")	
	if sort_app = nil then
		sort_app = 2
	end if
	
	search = false
	If trim(Request("productID")) <> nil Then
		productID = trim(Request("productID"))
		where_product = " AND QUESTIONS_ID = " & productID
		search = true
	Else 
	    productID = ""	
	    where_product = ""
	    search = false
	End If			

	sqlstr = "Exec dbo.get_appeals '','','','','" & OrgID & "','" & sortby_app(sort_app) & "','','','" & companyID & "','" & contactID & "','" & projectID & "','','" & productID & "','" & UserID & "','','','" & is_groups & "'"
    'Response.Write sqlStr
	set app=con.GetRecordSet(sqlStr)
	app_count = app.RecordCount
	if Request("page_app")<>"" then
		page_app=request("page_app")
	else
		page_app=1
	end if
	if not app.eof then
		app.PageSize = 5
		app.AbsolutePage=page_app
		recCount=app.RecordCount 		
		NumberOfPagesApp = app.PageCount
		i=1
		j=0
		ids = "" 'list of appeal_id
	end if
	if not app.eof Or search = true then %>
<input type="hidden" name="trapp" value="" ID="trapp">		
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100% border=0>  
  <tr><td width=100%><A name="table_appeals"></A>  
  <table cellpadding=0 cellspacing=0 width=100% border=0 dir="<%=dir_var%>">
  <tr>  
  <td class="title_form" width=100% align="<%=align_var%>" dir=<%=dir_obj_var%>>&nbsp;<span id=word28 name=word28><!--טפסים מצורפים--><%=arrTitles(28)%></span>&nbsp;<font color="#E6E6E6">(<%=companyName & " - " & project_name%>)</font>&nbsp;</td>
  </tr>
  </table></td></tr>  	  	
  <tr>
	<td width="100%" align="center" valign=top dir="<%=dir_var%>">
	<table BORDER="0" CELLSPACING="1" CELLPADDING="1" bgcolor="#FFFFFF">	
    <tr>
	    <td nowrap class="title_sort" align=center><!--הדפס--><%=arrTitles(29)%></td>
		<td nowrap class="title_sort" dir="<%=dir_obj_var%>" align=center><%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%></td>
		<!--td nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=5 or sort_app=6 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=6 then%>5<%elseif sort_app=5 then%>6<%else%>6<%end if%>#table_appeals" target="_self">&nbsp;עובד&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="5" then%>bot<%elseif trim(sort_app)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td-->	
		<td nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=9 or sort_app=10 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=10 then%>9<%elseif sort_app=9 then%>10<%else%>10<%end if%>#table_appeals" target="_self">&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="9" then%>bot<%elseif trim(sort_app)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>		
		<td nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort" nowrap>&nbsp;<%=arrTitles(65)%>&nbsp;</td>			
		<td width="100%" id=td_app_type name=td_app_type align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort" nowrap>&nbsp;<!--סוג טופס--><%=arrTitles(31)%>&nbsp;<IMG name=word52 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(52)%>" align=absmiddle onmousedown="appealDropDown(td_app_type)"></td>
		<td nowrap align=center dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=1 or sort_app=2 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=2 then%>1<%elseif sort_app=1 then%>2<%else%>2<%end if%>#table_appeals" target="_self">&nbsp;<!--תאריך--><%=arrTitles(32)%>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="1" then%>bot<%elseif trim(sort_app)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
		<%key_table_width=250%>
		<td class="title_sort">&nbsp;</td>		
		<td nowrap align=center dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=3 or sort_app=4 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=4 then%>3<%elseif sort_app=3 then%>4<%else%>4<%end if%>#table_appeals" target="_self">&nbsp;ID&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="3" then%>bot<%elseif trim(sort_app)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td nowrap align=center dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=11 or sort_app=12 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=12 then%>11<%elseif sort_app=11 then%>12<%else%>12<%end if%>#table_appeals" target="_self">&nbsp;<!--'סט--><%=arrTitles(33)%>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="11" then%>bot<%elseif trim(sort_app)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	</tr>
<%do while (not app.eof and j<app.PageSize)
		appid = trim(app("appeal_id"))
		If j Mod 2 = 0 Then
			tr_bgcolor = "#E6E6E6"
		Else
			tr_bgcolor = "#C9C9C9"
		End If				
		COMPANY_NAME = app("COMPANY_NAME")
		If trim(COMPANY_NAME) = "" Or IsNull(COMPANY_NAME) Then
			COMPANY_NAME = ""
		End If	
		CONTACT_NAME = app("CONTACT_NAME")
		If trim(CONTACT_NAME) = "" Or IsNull(CONTACT_NAME) Then
			CONTACT_NAME = ""
		End If	
		PROJECT_NAME = app("PROJECT_NAME")
		If trim(PROJECT_NAME) = "" Or IsNull(PROJECT_NAME) Then
			PROJECT_NAME = ""
		End If
		product_name = app("product_name")
		If trim(product_name) = "" Or IsNull(product_name) Then
			product_name = ""
		End If	
		if len(product_name) > 30 then
			product_name = left(product_name,27) & "..."		
		end if			
		User_Name = app("User_Name")		
		prod_id = app("product_id")
		quest_id = trim(app("questions_id"))
		mechanismId = app("mechanism_id")		
		If trim(mechanismId) <> "" Then
			sqlstr = "Select mechanism_Name from mechanism WHERE mechanism_Id = " & mechanismId
			set rs_name = con.getRecordSet(sqlstr)
			If not rs_name.eof Then
				mechanismName = trim(rs_name.Fields(0))
			End If
			set rs_name = Nothing
		Else
			mechanismName = ""	
		End If			
		sqlstr = "EXECUTE get_appeal_tasks '" & OrgID & "','" & appid & "'"	
	    set rsmess = con.getRecordSet(sqlstr)
		if not rsmess.eof then
			mes_new = rsmess("mes_new")
			mes_work = rsmess("mes_work")
			mes_close = rsmess("mes_close") 
		else
			mes_new = 0  :  mes_work = 0  :  mes_close = 0
		end if
		If trim(SURVEYS)  = "1" Then
		    href_ = " HREF=""../appeals/appeal_card.asp?quest_id=" & quest_id & "&appid=" & appid & """"
		Else
			href_ = " nohref"
		End If 
		appeal_status = trim(app("appeal_status"))	 
		appeal_status_name = trim(app("appeal_status_name"))	
		appeal_status_color = trim(app("appeal_status_color"))	%>
		<tr bgcolor="<%=tr_bgcolor%>" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='<%=tr_bgcolor%>';">
		    <td align="center" nowrap>&nbsp;<a  href="#" onclick="javascript:window.open('../appeals/view_appeal.asp?quest_id=<%=quest_id%>&appid=<%=appid%>','','top=100,left=100,width=500,height=500,scrollbars=1,resizable=1,menubar=1')"><IMG SRC="../../images/print_icon.gif" BORDER=0 hspace=0 vspace=0></a>&nbsp;</td>
			<td align="center" nowrap>						
			&nbsp;
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=3" target="_self" style="WIDTH:10pt;" class="task_status_num3" title="<%=arr_StatusT(3)%>"><%=mes_close%></a>
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=2" target="_self" style="WIDTH:10pt;" class="task_status_num2" title="<%=arr_StatusT(2)%>"><%=mes_work%></a>
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=1" target="_self" style="WIDTH:10pt;" class="task_status_num1" title="<%=arr_StatusT(1)%>"><%=mes_new%></a>
			&nbsp;		
			</td>		
			<!--td nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=User_NAME%>&nbsp;</a></td-->
			<td nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=CONTACT_NAME%>&nbsp;</a></td>
			<td nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=mechanismName%>&nbsp;</a></td>
			<td nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=product_name%>&nbsp;</a></td>
			<td align=center><a class="link_categ" <%=href_%> target="_self">&nbsp;<%=day(app("appeal_date"))%>/<%=month(app("appeal_date"))%>/<%=mid(year(app("appeal_date")),3,2)%>&nbsp;</a></td>
			<%key_table_width=250%>
			<td align="<%=align_var%>">			
		    <!--#include file="../appeals/key_fields_t.asp"-->
		    </td>
			<td nowrap align=center><a class="link_categ" <%=href_%> target="_self"><%=appid%></a></td>
			<td nowrap align=center><a class="status_num" style="background-color:<%=trim(appeal_status_color)%>" <%=href_%> target="_self"><%=appeal_status_name%></a></td>
		</tr>
<%		app.movenext
		j=j+1
		if not app.eof and j <> app.PageSize then
		ids = ids & ","
		end if
		loop 
		%>
		</table>		
		<input type="hidden" name="ids" value="<%=ids%>" ID="ids">
		</td></tr>
		</form>				
		<% if NumberOfPagesApp > 1 then%>
		<tr class="card">
		<td width="100%" align=center nowrap class="card" dir="<%=dir_var%>">
			<table border="0" cellspacing="0" cellpadding="2" ID="Table7">               
	        <% If NumberOfPagesApp > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPagesApp / 10)
	           else num = NumberOfPagesApp : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRowApp") <> nil Then
	               numOfRowApp = Request.QueryString("numOfRowApp")
	           Else numOfRowApp = 1
	           End If
	         %>
	         
	         <tr>
	         <%if numOfRowApp <> 1 then%> 
			 <td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=10*(numOfRowApp-1)-9%>&numOfRowApp=<%=numOfRowApp-1%>#table_appeals" name=word56 title="<%=arrTitles(56)%>"><b><<</b></a></td>			                
			  <%end if%>
	             <td><font size="2" color="#060165">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowApp-1)) <= NumberOfPagesApp Then
	                  if CInt(page_app)=CInt(i+10*(numOfRowApp-1)) then %>
		                 <td align="center"><span class="pageCounter"><%=i+10*(numOfRowApp-1)%></span></td>
	                  <%else%>
	                     <td align="center"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=i+10*(numOfRowApp-1)%>&numOfRowApp=<%=numOfRowApp%>#table_appeals"><%=i+10*(numOfRowApp-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td><font size="2" color="#060165">]</font></td>
				<%if NumberOfPagesApp > cint(num * numOfRowApp) then%>  
					<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=10*(numOfRowApp) + 1%>&numOfRowApp=<%=numOfRowApp+1%>#table_appeals" name=word57 title="<%=arrTitles(57)%>"><b>>></b></a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>
	<%End If%>										
	<%If app.recordCount = 0 Then%>
	<tr><td align=center class=card1>&nbsp;</td></tr>									
	<%End If%>			 			 
	</table></td></tr>
	<%End If%>	
<%set app = nothing	%>
<%If trim(is_meetings) = "1" Then%>
<!-------------------------------------------------מפגשים-------------------------------------------------->
<%
	if lang_id = "1" then
		arr_Status_M = Array("","עתידית","הסתיימה","הוכנס סיכום","נדחתה")	
	else
		arr_Status_M = Array("","Future","Done","Summary added","Postponed")	
	end if			
   
	if Request("PageM")<>"" then
		PageM=request("PageM")
	else
		PageM=1
	end if
	PageSizeM = 5
	sqlstr = "EXECUTE get_meetings_paging " & PageM & "," & PageSizeM & ",'','','','','','" & OrgID & "',' meeting_date, start_time, end_time','','','" & CompanyID & "','" & ContactID & "','" & projectID & "'"
	'Response.Write sqlStr
	'Response.End
	set rs_meetings = con.GetRecordSet(sqlStr)
	if not rs_meetings.EOF then
	    recCountM = rs_meetings("CountRecords")
%>
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6"><A name="table_meetings"></A>  
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100%>  
  <tr>  
  <td class="title_form" width=100%>&nbsp;<!--פגישות--><%=arrTitles(59)%>&nbsp;<font color="#E6E6E6">(<%=companyName & " - " & project_name%>)</font>&nbsp;</td>
  </tr>
<tr>    
    <td width="100%" valign="top" align="center" colspan=3>    
	<table BORDER="0" CELLSPACING="1" CELLPADDING="1" bgcolor="#FFFFFF" dir="<%=dir_var%>">	
	<tr>		
	<td width="100%" nowrap align="center" class="title_sort"><!--משתתפים--><%=arrTitles(63)%></td>
	<td width="100" nowrap align="center" valign="top" class="title_sort"><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td>
	<td width="150" nowrap align="center" valign="top" class="title_sort"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
	<td width="70" nowrap align="center" class="title_sort"><!--שעת סיום--><%=arrTitles(60)%></td>
	<td width="70" nowrap align="center" class="title_sort"><!--שעת התחלה--><%=arrTitles(61)%></td>
	<td width="60" nowrap align="center" valign="top" class="title_sort"><!--תאריך--><%=arrTitles(32)%></td>
	<td width="70" align="center" dir="<%=dir_obj_var%>" nowrap class="title_sort"><!--'סט--><%=arrTitles(13)%>&nbsp;</td>
	</tr>
<% while not rs_meetings.EOF
	meetingID = rs_meetings(1)
	company_name = rs_meetings(4)
	contact_name = rs_meetings(5)
	meetingDate = rs_meetings(6) 
	status = rs_meetings(7)
	startTime = rs_meetings(8)
	endTime = rs_meetings(9)
	
	users_names= ""
	sqlstr = "Select FIRSTNAME + ' ' + LASTNAME From meeting_to_users Inner Join Users On Users.User_ID = meeting_to_users.User_ID " &_
	" Where meeting_ID = " & meetingID & " ORDER BY FIRSTNAME + ' ' + LASTNAME"
	set rs_participants = con.getRecordSet(sqlstr)
	if not rs_participants.eof then
		users_names= rs_participants.getString(,,",",",")
	end if
	set rs_participants = nothing
	If Len(users_names) > 1 Then
		users_names= Left(users_names,Len(users_names)-1)
	End If	
    If trim(status) = "1" Then
       href = "href=""javascript:addmeeting('" & meetingID & "')"""   
    Else
       href = "href=""javascript:closemeeting('" & meetingID & "')"""     
    End If	
%>
	<tr class="card">
		<td align="<%=align_var%>" valign="top" ><a class="link_categ" <%=href%>><%=users_names%></a></td>	    
		<td align="center" valign="top"><a class="link_categ" <%=href%>><%=contact_name%></a></td>
		<td align="center" valign="top"><a class="link_categ" <%=href%>><%=company_name%></a></td>
		<td align="center" valign="top"><a class="link_categ" <%=href%>><%=endTime%></a></td>
		<td align="center" valign="top"><a class="link_categ" <%=href%>><%=startTime%></a></td>
		<td align="<%=align_var%>" valign="top"><a class="link_categ" <%=href%>><%=FormatMediumDateShort(meetingDate)%></a></td>
		<td align="center" valign="top" dir="<%=dir_obj_var%>" nowrap><a class="task_status_num<%=status%>" <%=href%> style="width:64"><%=arr_Status_M(status)%></a></td>	  
	</tr>	
<% 
   rs_meetings.moveNext
   Wend  
   NumberOfPages = Fix((recCountM / PageSizeM)+0.9)
   if NumberOfPages > 1 then
	  %>
	  <tr>
		<td width="100%" align=middle colspan=11 nowrap class="card">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr ID="Table1">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRowM") <> nil Then
	               numOfRowM = Request.QueryString("numOfRowM")
	           Else numOfRowM = 1
	           End If	           
            %>	         
	         <tr>
	         <%if numOfRowM <> 1 then%> 
			 <td valign="center" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(20)%>" href="<%=urlSort%>&PageM=<%=10*(numOfRowM-1)-9%>&numOfRowM=<%=numOfRowM-1%>#table_meetings" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowM-1)) <= NumberOfPages Then
	                  if CInt(PageM)=CInt(i+10*(numOfRowM-1)) then %>
		                 <td align="middle" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRowM-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&PageM=<%=i+10*(numOfRowM-1)%>&amp;numOfRowM=<%=numOfRowM%>#table_meetings" ><%=i+10*(numOfRowM-1)%></a></td>
	                  <%end if
	                  end if
	               next%>	            
					<td align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRowM) then%>  
					<td valign="center" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(19)%>" href="<%=urlSort%>&PageM=<%=10*(numOfRowM) + 1%>&numOfRowM=<%=numOfRowM+1%>#table_meetings" >&gt;&gt;</a></td>
				<%end if%>
				</tr> 				  	             
			</table>
		</td>
	</tr>
	<!--tr><td  class="title_form" align=center colspan=9><%=recCount%> :ואצמנ תומושר כ"הס</td></tr-->										 
	<%End If%> 
 </table></td></tr>
 </table></td></tr> 	
 <%  
 End If
 set rs_meetings = Nothing 
 End If
%>	
</table></td>
<td width=110 nowrap align="<%=align_var%>" valign=top valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100% border=0>
<tr><td align="<%=align_var%>" colspan=2 height="18" nowrap></td></tr>
<tr><td align="center"><a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick="addtask('<%=projectID%>','<%=companyID%>','')"><!--הוסף--><%=arrTitles(38)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a></td></tr>
<%If trim(is_meetings) = "1" Then%>
<tr><td align="center"><a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick="addmeeting('')">&nbsp;<!--הוסף פגישה --><%=arrTitles(62)%>&nbsp;</a></td></tr>  
<%End If%>
<tr><td align="center"><a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick="addMechanism('<%=projectID%>','<%=companyID%>','')">&nbsp;<!--הוסף מנגנון --><%=arrTitles(64)%>&nbsp;</a></td></tr>  
<tr><td nowrap align="center">
<%if numOfLink = 1 then%>
<a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick="javascript:window.open('editProject.asp?companyID=<%=companyID%>&project_ID=<%=projectID%>','editProject','top=50,left=120,resizable=1,scrollbars=1,width=480,height=500');">
<%else%>
<a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick="javascript:window.open('editProject_action.asp?companyID=<%=companyID%>&project_ID=<%=projectID%>','editProject_action','top=120,left=120,resizable=1,scrollbars=1,width=480,height=450');">
<%end if%>
&nbsp;<span id="word43" name=word43><!--עדכן--><%=arrTitles(43)%></span>&nbsp;<%=project_title%>&nbsp;</a>
</td></tr>
<tr><td nowrap align="center">
<%
     If trim(lang_id) = "1" Then
        str_delete = "? האם ברצונך למחוק את ה" & trim(project_title)
     Else
		str_delete = "Are you sure want to delete the " & trim(project_title) & " ?"
     End If   
  %>
<%if numOfLink = 2 then%>
<a class="button_edit_1" style="width:106;" onclick="return window.confirm('<%=str_delete%>')" href="default.asp?delPROJECT_ID=<%=projectID%>" target=_self>
<%else%>
<a class="button_edit_1" style="width:106;" onclick="return window.confirm('<%=str_delete%>')" href="default_action.asp?delPROJECT_ID=<%=projectID%>" target=_self>
<%end if%>
&nbsp;<span id="word45" name=word45><!--מחק--><%=arrTitles(45)%></span> <%=project_title%>&nbsp;</a>
</td></tr>
<%If trim(SURVEYS)  = "1" Then%>
<tr><td height=5 nowrap></td></tr>
<tr><td align="center" colspan=2><a class="button_edit_2" href="javascript:void(0)" onclick="tfasimDropDown(this)"><img hspace=0 vspace=0 border=0 src="../../images/back_arrow.gif" <%If trim(lang_id) = "2" Then%>style="Filter: FlipH"<%End If%>>&nbsp;<!--צרף טופס--><%=arrTitles(42)%></a></td></tr>
<%End If%>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table></td></tr>
</table>
<DIV ID="task_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:100%; height:68; overflow:scroll; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #E6E6E6; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%  sqlstr = "Select activity_type_id, activity_type_name from activity_types Where ORGANIZATION_ID = "&OrgID&" Order By activity_type_id"
	set rsactivity = con.getRecordSet(sqlstr)
	If not rsactivity.eof Then
		TaskTypes = rsactivity.getRows()
	End If
	set rsactivity = Nothing	
	If IsArray(TaskTypes) Then
	For i=0 To Ubound(TaskTypes,2)%> 
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; text-align:right; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>&taskTypeID=<%=TaskTypes(0,i)%>#table_tasks'">
    <%=TaskTypes(1,i)%></DIV>
	<% Next
	   End If %>
	<DIV dir="<%=dir_obj_var%>"  onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>#table_tasks'"><!--כל הרשימה--><%=arrTitles(47)%></DIV>
</div>
</DIV>
<DIV ID="appeal_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:100%; height:68; overflow:scroll; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #E6E6E6; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%
	If is_groups = 0 Then
	sqlstr = "Select product_id, product_name from Products Where "&_
	" product_number = '0' and ORGANIZATION_ID=" & OrgID & " order by product_name"
	' משתמש אשר שייך לקבוצה אבל אינו אחראי באף קבוצה
	Else
	sqlstr = "Execute get_products_list '" & OrgID & "','" & UserID & "'"
	End If
	'Response.Write sqlstr
	'Response.End
	set rs_products = con.GetRecordSet(sqlstr)
	if not rs_products.eof then 
		ResProductsList = rs_products.getRows()		
	end if
	set rs_products=nothing				
	If IsArray(ResProductsList) Then
	For i=0 To Ubound(ResProductsList,2)%>
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; text-align:right; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>&productID=<%=ResProductsList(0,i)%>#table_appeals'">
    <%=ResProductsList(1,i)%>
    </DIV>
<%	Next	
	End If	
%>
	<DIV dir="<%=dir_obj_var%>"  onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>#table_appeals'"><!--כל הרשימה--><%=arrTitles(48)%></DIV>
</div>
</DIV>
<div id=div_comp_desc name=div_comp_desc style="display:none;visibility:hidden;position:absolute;left:30;top:435;width:350;height:100;z-index:11;">
<table cellpadding=0 cellspacing=1 border=1 bgcolor="#ffffff" width=100%>
 <tr>
        <td width="100%" bgcolor="#0F2771" align="<%=align_var%>">
            <table border="0" width="100%" cellspacing="0" cellpadding="0">    
            <tr>
				<td colspan=2 width="100%" bgcolor="#616161" height="1"></td>
			</tr>	
            <tr>
                <td width=20 nowrap align=center class=title_form><INPUT type=image src="../../images/close_icon.gif" border="0" onClick="return closeDesc();" vspace=0 hspace=0>              
                </td>            
                <td width=330 nowrap class=title_form align="<%=align_var%>">&nbsp;<!--פרטים נוספים--><%=arrTitles(49)%>&nbsp;</td>
             </tr>
             <tr>
				<td colspan=2 width="100%" bgcolor="#616161" height="1"></td>
			</tr>
			<tr>
            <td width="100%" colspan=2 align="<%=align_var%>" class=card style="padding:5px" dir="<%=dir_obj_var%>"><%=trim(company_desc)%></td>
            </tr>	
            </table>
        </td>  
  </tr>                     
</table>
</div>
<%If trim(SURVEYS)  = "1" Then%>
<!--#include file="../companies/tfasim_inc.asp"-->
<%End If%>	
<%End If%>
</body>
<%set con=Nothing%>
</html>

