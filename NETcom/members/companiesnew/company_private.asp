<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%   
  
  contactID = trim(Request("contactID"))
  companyID = trim(Request("companyID"))  
  
  If trim(companyID)<>"" then 
   if trim(lang_id) = "1" then
	   arr_status_comp = Array("","עתידי","פעיל","סגור","פונה")
   else
	   arr_status_comp = Array("","new","active","close","appeal")
   end if	   
   found_company = false
  urlSort="company_private.asp?companyID="& companyID & "&contactID=" & contactID
  sqlStr = "SELECT COMPANY_NAME,ADDRESS,CITY_NAME,PREFIX_PHONE,PHONE,PREFIX_FAX,FAX"&_
  ", DATE_UPDATE, STATUS,CONTACT_ID FROM COMPANIES_PRIVATE_VIEW WHERE COMPANY_ID="& companyID &_
  " And ORGANIZATION_ID = " & orgID
  'Response.Write "test"
  'Response.End
  set pr=con.GetRecordSet(sqlStr)
  if not pr.EOF then	
	company_name  =pr("company_name")  
	address	      =pr("address")		
	cityName	  =pr("city_Name")
	prefix_phone  =pr("prefix_phone")	
	prefix_fax	  =pr("prefix_fax")	
	phone	      =pr("phone")	
	fax           =pr("fax")			
	date_update	  =pr("date_update")
	status_company =pr("status")
	contactID	   =pr("contact_id")	
	found_company = true		    
    If isNumeric(trim(companyID)) Then
		types=""		
		sqlstr="Select type_Name From company_to_types_view Where company_id = " & companyID & " Order By type_id"
		set rssub = con.getRecordSet(sqlstr)			
		If not rssub.eof Then
			types = rssub.getString(,,",",",") 				
		End If
		set rssub=Nothing
		If Len(types) > 0 Then
			types = Left(types,(Len(types)-1))
		End If					
      	
      	sqlStr = "SELECT company_desc FROM companies WHERE company_id="& companyID 
		'Response.Write "test"
		'Response.End
		set rs_d=con.GetRecordSet(sqlStr)
		If not rs_d.EOF then				
  			company_desc = trim(rs_d(0))
  			If Len(company_desc) > 43 Then
  				company_desc_short = Left(company_desc,40) & ".."
  			Else
  			    company_desc_short = company_desc
  			End If
		End if 
		set rs_d = Nothing      
      End If 
  Else
	found_company = false             
  End If   
 End if 
 
 If trim(contactID)<>"" then 
   sql="SELECT company_ID,contact_ID,email,prefix_phone,phone,prefix_cellular,prefix_fax,fax,"&_
   "cellular,CONTACT_NAME,messanger_name,date_update FROM contacts"&_
   " WHERE contact_ID=" & contactID
   set listContact=con.GetRecordSet(sql)
   if not listContact.EOF then 
      contactID=listContact("contact_ID") 
      companyID=listContact("company_ID") 
      CONTACT_NAME=listContact("CONTACT_NAME")  
      contacter = trim(CONTACT_NAME)
      email=listContact("email")
      prefix_phone=listContact("prefix_phone")
      phone=listContact("phone")
      prefix_cellular=listContact("prefix_cellular")
      cellular=listContact("cellular")
      prefix_fax=listContact("prefix_fax")
      fax=listContact("fax")       
      messangerName=listContact("messanger_name")     
    End if   
    set  listContact = Nothing
  End If  
    
 delContactID=trim(Request("delContactID"))  
 If delContactID<>nil And delContactID<>"" Then 	
	con.ExecuteQuery "delete from Contacts WHERE Contact_Id=" & delContactID	
	con.executeQuery "delete From tasks WHERE task_id IN (Select task_id FROM activities Where contact_id=" &  delContactID & ")"' delete contact also from activities	
	con.executeQuery "delete From activities Where contact_id=" &  delContactID ' delete contact also from activities	
    Response.Redirect "contacts.asp"
 End If   
  
 delActivityID = trim(Request("delActivityID"))
 If delActivityID<>nil And delActivityID<>"" Then 			
	'con.executeQuery "delete From tasks WHERE task_id IN (Select task_id FROM activities Where activity_id=" &  delActivityID & ")"' delete contact also from activities	
	con.executeQuery "delete From activities Where activity_id=" &  delActivityID ' delete contact also from activities	
    Response.Redirect "company_private.asp?companyID=" & companyId & "&contactID=" & contactID              
 End If
  
 delTaskID = trim(Request("delTaskID"))
 If delTaskID<>nil And delTaskID<>"" Then 			
	con.executeQuery "DELETE From tasks WHERE task_id = " & delTaskID
	'con.executeQuery "delete From activities Where activity_id=" &  delActivityID ' delete contact also from activities	
    Response.Redirect "company_private.asp?companyID=" & companyId & "&contactID=" & contactID              
 End If 
 
 delprojectId=trim(Request("delprojectId"))  
  If delprojectId<>nil And delprojectId<>"" Then 	
    con.ExecuteQuery "delete from HOURS where PROJECT_ID=" & delprojectId
	con.ExecuteQuery "delete from PROJECTS where PROJECT_ID=" & delprojectId
	con.ExecuteQuery "Delete From pricing_to_jobs where pricing_id=" & delprojectId	
	con.ExecuteQuery "Delete From pricing_to_projects where project_id=" & delprojectId
    Response.Redirect "company_private.asp?companyID=" & companyId & "&contactID=" & contactID              
  End If  
  
  delDocumentID=trim(Request("delDocumentID"))
  if delDocumentID <> nil then   
		set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
'-----deleting files----
		sqlstr = "select * from DOCUMENTS where document_id = "& delDocumentID
		'Response.Write sqlstr
		'Response.End
		set files = con.getRecordSet(sqlstr)
		do while not files.eof
			file_path="../../../download/documents/" & trim(files("document_file"))
			if fs.FileExists(server.mappath(file_path)) then
				set a = fs.GetFile(server.mappath(file_path))
				a.delete '------------------------------
			else
				Response.Write(file_path)	
			end if	
		files.movenext
		loop
		set files =nothing
		set fs = nothing
		
		con.executeQuery("delete from DOCUMENTS where document_id = " & delDocumentID)
		con.executeQuery("delete from company_documents where document_id = " & delDocumentID)
		Response.Redirect "company_private.asp?companyID=" & companyId & "&contactID=" & contactID
  End if     
  
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 8 Order By word_id"				
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
<!--#include file="../../../include/title_meta_inc.asp"-->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">

var oPopup = window.createPopup();
function appealDropDown(obj)
{
    oPopup.document.body.innerHTML = appeal_Popup.innerHTML; 
    oPopup.document.charset="windows-1255";
    oPopup.show(0, 20, obj.offsetWidth, 68, obj);    
}

function taskDropDown(obj)
{
    oPopup.document.body.innerHTML = task_Popup.innerHTML; 
    oPopup.document.charset="windows-1255";
    oPopup.show(0, 17, obj.offsetWidth, 68, obj);    
}

function CheckDel() {
   <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את ה" & trim(Request.Cookies("bizpegasus")("CompaniesOne"))
     Else
		str_confirm = "Are you sure want to delete the " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & " ?"
     End If   
  %>
  if (window.confirm("<%=str_confirm%>"))  
  {
	 document.location.href = "companies_private.asp?delid=<%=companyID%>";
	 return true;	
  }
  return false;
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

function CheckDelProject(projectId,companyId) {
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את ה" & trim(Request.Cookies("bizpegasus")("Projectone"))
     Else
		str_confirm = "Are you sure want to delete the " & trim(Request.Cookies("bizpegasus")("Projectone")) & " ?"
     End If   
  %>
  if (confirm("<%=str_confirm%>"))
  {         
     document.location.href = "company_private.asp?delprojectId="+projectId+"&companyId="+companyId;
    
  }
  return false;  
}

function CheckDelDocument(companyId,documentId) {
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את המסמך"
     Else
		str_confirm = "Are you sure want to delete the document?"
     End If   
  %>
  if (confirm("<%=str_confirm%>"))
  {         
     document.location.href = "company_private.asp?companyId="+companyId+"&delDocumentID="+documentId;
    
  }
  return false;  
}

function addtask(contactID,companyID,taskID)
{
	h = parseInt(520);
	w = parseInt(420);
	window.open("../tasks/addtask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskID=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
}

function closeTask(contactID,companyID,taskID)
{
		h = parseInt(470);
		w = parseInt(420);
		window.open("../tasks/closetask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskId=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=50,left=100,width="+w+",height="+h+",align=center,resizable=0");
}

function openTask(contactID,companyID,taskID)
{
	h = parseInt(400);
	w = parseInt(740);
	window.open("../tasks/edittask.asp?contactID=" + contactID + "&companyId=" + companyID + "&oldtaskId=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=70,left=20,width="+w+",height="+h+",align=center,resizable=0");
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

//-->
</script>  
</head>
<body>
<!-- #include file="../../logo_top.asp" -->
<%numOftab = 0%>
<%numOfLink = 2%>
<!--#include file="../../top_in.asp"-->
<%If found_company Then%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF  dir="<%=dir_var%>">
<tr><td width="100%" class="page_title" dir="<%=dir_obj_var%>"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<font color="#6F6DA6"><%=company_name%></font>&nbsp;</td></tr>         
<tr><td width="100%">
<table cellpadding=0 cellspacing=1 width=100% bgcolor=white><tr><td width="100%" valign=top>
<table cellpadding=0 bgcolor="#808080" dir="ltr" cellspacing=1 width=100% border=0 style="border-collpase:collapse;" ID="Table13">
<tr>          
  <td align="<%=align_var%>" width="50%" nowrap valign="top" bgcolor=#E6E6E6>  
  <input type="hidden" name="companyID" id="companyID" value="<%=companyID%>">  
  <input type="hidden" name="contactID" id="contactID" value="<%=contactID%>">  
   <table border="0" cellpadding="0" cellspacing="1" align=center dir="<%=dir_var%>">   
   <tr><td colspan=2 class="title_form" align="<%=align_var%>"><span id=word3 name=word3><!--פרטים--><%=arrTitles(3)%></span>&nbsp;</td></tr>
      <tr>
         <td width="100%" align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">           
         <input type=text dir="<%=dir_obj_var%>" name="CONTACT_NAME" value="<%=vFix(CONTACT_NAME)%>" maxlength=20 style="width:250;" class="Form_R" ReadOnly>           
         </td>
         <td width="120" nowrap align="<%=align_var%>" style="padding-right:15px" bgcolor="#e6e6e6" valign="top" dir="<%=dir_obj_var%>">&nbsp;<span id=word4 name=word4><!--שם--><%=arrTitles(4)%></span>&nbsp;</td>
      </tr>     
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">                        
         <INPUT name="messanger" dir="<%=dir_obj_var%>" style="FONT-FAMILY: Arial;" style="width:250;" class="Form_R" ReadOnly value="<%=vFix(messangerName)%>">
         </td>                    
         <td align="<%=align_var%>" style="padding-right:15px" bgcolor="#e6e6e6" valign="top">&nbsp;<span id="word10" name=word10><!--תפקיד--><%=arrTitles(10)%></span>&nbsp;</td>
      </tr> 
      <tr>
        <td align="<%=align_var%>"><input type=text dir="<%=dir_obj_var%>" name="address" value="<%=vFix(address)%>"  style="width:250" class="Form_R" ReadOnly ID="Text2"></td>
        <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<span id="word5" name=word5><!--כתובת--><%=arrTitles(5)%></span>&nbsp;</td>
     </tr>             
     <tr>
         <td align="<%=align_var%>" nowrap>                    
           <input type=text dir="<%=dir_obj_var%>" class="Form_R" name="cityName" id="cityName" value="<%=cityName%>"  ReadOnly>
         </td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<span id="word6" name=word6><!--עיר--><%=arrTitles(6)%></span>&nbsp;</td>
      </tr>     
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir=ltr>                            
                <input type=text dir="ltr" name="phone" value="<%=vFix(phone)%>" maxlength="20" size="20" class="Form_R" ReadOnly>                
         </td>
         <td align="<%=align_var%>" style="padding-right:15px" bgcolor="#e6e6e6" valign="top"><span id="word7" name=word7><!--טלפון--><%=arrTitles(7)%></span>&nbsp;</td>
      </tr>
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir=ltr>                             
            <input type=text dir="ltr" name="cellular" value="<%=vFix(cellular)%>" size="20" class="Form_R" ReadOnly>              
         </td>
         <td align="<%=align_var%>" style="padding-right:15px" bgcolor="#e6e6e6" valign="top"><span id="word8" name=word8><!--טלפון נייד--><%=arrTitles(8)%></span>&nbsp;</td>
      </tr>
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir=ltr>                                 
              <input type=text dir="ltr" name="fax" value="<%=vFix(fax)%>" size="20" class="Form_R" ReadOnly></td>           
         </td>
         <td align="<%=align_var%>" style="padding-right:15px" bgcolor="#e6e6e6" valign="top"><span id="word9" name=word9><!--פקס--><%=arrTitles(9)%></span>&nbsp;</td>
      </tr>
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top"> 
         <%If Len(email) > 0 Then%>
         <a href="mailto:<%=email%>" dir=ltr style="width:250" class="Form_R hand">                       
         <%=trim(email)%>           
         </a>
         <%Else%>
         <span style="width:250" class="Form_R"></span>
         <%End If%>
         </td>
         <td align="<%=align_var%>" style="padding-right:15px" bgcolor="#e6e6e6" valign="top">Email&nbsp;</td>
      </tr> 
      <tr>
           <td align="<%=align_var%>">
           <span class="Form_R" dir="<%=dir_obj_var%>" style="width:250;line-height:16px"><%=types%></span>
           </td>
           <td align="<%=align_var%>" style="padding-right:15px" valign=top>&nbsp;<span id="word11" name=word11><!--קבוצה--><%=arrTitles(11)%></span>&nbsp;</td>
      </tr> 
      <tr>
      <td align="<%=align_var%>" valign=bottom>
      <%If Len(trim(company_desc)) > 0 Then%><input type=image src="../../images/popup.gif" onclick="return descDropDown(this)" name=word50 title="<%=arrTitles(50)%>" border=0 hspace=0 vspace=0 ID="Image6" NAME="Image6"><%End if%>
      <span class="Form_R" dir="<%=dir_obj_var%>" style="width:250;line-height:16px"><%=trim(company_desc_short)%></span>      
      </td>
      <td align="<%=align_var%>" style="padding-right:15px" valign=top>&nbsp;<span id="word12" name=word12><!--פרטים נוספים--><%=arrTitles(12)%></span>&nbsp;</td>
      </tr>      
      <tr>
           <td align="<%=align_var%>"><span style="width:60;text-align:center" class="status_num<%=status_company%>">
           <%=arr_status_comp(status_company)%></span>
           </td>
           <td align="<%=align_var%>" style="padding-right:15px" valign=top>&nbsp;<span id="word13" name=word13><!--סטטוס--><%=arrTitles(13)%></span>&nbsp;</td>
      </tr>
      <tr><td colspan="2" height="5" nowrap></td></tr>    
</table>
</td>
<td width=50% align=center nowrap valign="top" bgcolor=#E6E6E6>
<table cellpadding=0 cellspacing=1 width=100% align=center bgcolor=white dir="<%=dir_var%>">
<%if trim(companyID)<>"" then
 urlSort="company_private.asp?companyID="& companyID 
 dim sortby_pr(7)	
 sortby_pr(1) = "project_name"
 sortby_pr(2) = "project_name DESC"
 sortby_pr(3) = "project_code"
 sortby_pr(4) = "project_code DESC"

 sort = trim(Request.QueryString("sort"))
 If trim(sort) = "" Then
	sort = 1
 End If

'PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
 If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
	PageSize = 5
 End If	

 if trim(Request.QueryString("pageProj"))<>"" then
    pageProj=Request.QueryString("pageProj")
 else
    pageProj=1
 end if  
 
 if trim(Request.QueryString("numOfRowProj"))<>"" then
    numOfRowProj=Request.QueryString("numOfRowProj")
 else
    numOfRowProj = 1
 end if  

  sql="SELECT projects.project_id, projects.project_code, projects.project_name FROM projects "&_
  " WHERE company_ID = " & companyID & " AND projects.ORGANIZATION_ID = " & OrgID & " ORDER BY "& sortby_pr(sort)
  
 set list_pr=con.GetRecordSet(sql)
 if not list_pr.EOF then
	list_pr.PageSize=PageSize
	list_pr.AbsolutePage=pageProj
	recCount=list_pr.RecordCount 
	NumberOfPages=list_pr.PageCount
	i=1	
%>
<tr><td colspan=4 class="title_form" align="<%=align_var%>"><%=trim(Request.Cookies("bizpegasus")("ProjectMulti"))%>&nbsp;</td></tr>
<tr>     
     <td width="35"   align="center" nowrap class="title_sort"><span id="word16" name=word16><!--מחק--><%=arrTitles(16)%></span></td>        
     <td width="125"  align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort"><span id="word17" name=word17><!--קוד--><%=arrTitles(17)%></span>&nbsp;</td>     
     <td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" name=word52 title="<%=arrTitles(52)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
</tr> 
<% do while (not list_pr.EOF and i<=list_pr.PageSize)
      PROJECT_ID=trim(list_pr("PROJECT_ID"))
      project_code=trim(list_pr("project_code"))      
      project_name=trim(list_pr("project_name"))                  
%>      
       <tr>   
         <td class="card" align="center"><input type=image src="../../images/delete_icon.gif" name=word53 title="<%=arrTitles(53)%>" style="border: none" hspace=0 vspace=0 onclick="return CheckDelProject('<%=PROJECT_ID%>','<%=companyId%>')" ID="Image2" NAME="Image2"></td>         
         <td class="card" align="<%=align_var%>"><a href="../projects/project.asp?projectID=<%=PROJECT_ID%>&companyID=<%=companyID%>" class="link_categ" target=_self>&nbsp;<%=project_code%>&nbsp;</a></td>         
         <td class="card" align="<%=align_var%>"><a href="../projects/project.asp?projectID=<%=PROJECT_ID%>&companyID=<%=companyID%>" class="link_categ" target=_self>&nbsp;<%=project_name%>&nbsp;</a></td>
      </tr>
<%  list_pr.MoveNext
	i=i+1
	loop
	if NumberOfPages > 1 then
	urlSort = urlSort & "&sort=" & sort
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
	         <%if numOfRowProj <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word54 title="<%=arrTitles(54)%>" href="<%=urlSort%>&pageProj=<%=10*(numOfRowProj-1)-9%>&amp;numOfRowProj=<%=numOfRowProj-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowProj-1)) <= NumberOfPages Then
	                  if CInt(pageProj)=CInt(i+10*(numOfRowProj-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRowProj-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&pageProj=<%=i+10*(numOfRowProj-1)%>&amp;numOfRowProj=<%=numOfRowProj%>" ><%=i+10*(numOfRowProj-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRowProj) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word55 title="<%=arrTitles(55)%>" href="<%=urlSort%>&pageProj=<%=10*(numOfRowProj) + 1%>&amp;numOfRowProj=<%=numOfRowProj+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	<%list_pr.close 
	set list_pr=Nothing
  End if  
 End If%>
<%
dim sortby_pr_cl(7)	
 sortby_pr_cl(1) = "project_name"
 sortby_pr_cl(2) = "project_name DESC"
 sortby_pr_cl(3) = "project_code"
 sortby_pr_cl(4) = "project_code DESC"

 sort_pr_cl = trim(Request.QueryString("sort_pr_cl"))
 If trim(sort_pr_cl) = "" Then
	sort_pr_cl = 1
 End If

'PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
 If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
	PageSize = 5
 End If	

 if trim(Request.QueryString("pageProjCl"))<>"" then
    pageProjCl=Request.QueryString("pageProjCl")
 else
    pageProjCl=1
 end if  
 
 if trim(Request.QueryString("numOfRowProjCl"))<>"" then
    numOfRowProjCl=Request.QueryString("numOfRowProjCl")
 else
    numOfRowProjCl = 1
 end if  

  sql="SELECT projects.project_id, projects.project_code, projects.project_name FROM projects "&_
  " WHERE company_ID = 0 AND projects.ORGANIZATION_ID = " & OrgID & " ORDER BY "& sortby_pr_cl(sort_pr_cl)
  
 set list_pr=con.GetRecordSet(sql)
 if not list_pr.EOF then
	list_pr.PageSize=PageSize
	list_pr.AbsolutePage=pageProjCl
	recCount=list_pr.RecordCount 
	NumberOfPages=list_pr.PageCount
	i=1	
%>
<tr><td colspan=4 class="title_form" align="<%=align_var%>"><%=trim(Request.Cookies("bizpegasus")("ActivitiesMulti"))%>&nbsp;</td></tr>
<tr>     
     <td colspan=2 width=160 nowrap class="title_sort" align="<%=align_var%>"><span id=word18 name=word18><!--קוד--><%=arrTitles(18)%></span>&nbsp;</td>     
     <td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort_pr_cl)="1" OR trim(sort_pr_cl)="2" then%>_act<%end if%>"><%if trim(sort_pr_cl)="1" then%><a class="title_sort" href="<%=urlSort%>&sort_pr_cl=<%=sort_pr_cl+1%>" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_pr_cl)="2" then%><a class="title_sort" href="<%=urlSort%>&sort_pr_cl=<%=sort_pr_cl-1%>" name=word52 name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_pr_cl=1" name=word52 title="<%=arrTitles(52)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("ActivitiesOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_pr_cl)="1" then%>bot<%elseif trim(sort_pr_cl)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
</tr> 
<% do while (not list_pr.EOF and i<=list_pr.PageSize)
      PROJECT_ID=trim(list_pr("PROJECT_ID"))
      project_code=trim(list_pr("project_code"))      
      project_name=trim(list_pr("project_name"))                  
%>      
       <tr>            
         <td colspan=2 class="card" align="<%=align_var%>"><a href="../projects/project.asp?projectID=<%=PROJECT_ID%>&companyID=<%=companyID%>" class="link_categ" target=_self>&nbsp;<%=project_code%>&nbsp;</a></td>         
         <td class="card" align="<%=align_var%>"><a href="../projects/project.asp?projectID=<%=PROJECT_ID%>&companyID=<%=companyID%>" class="link_categ" target=_self>&nbsp;<%=project_name%>&nbsp;</a></td>
      </tr>
<%  list_pr.MoveNext
	i=i+1
	loop
	if NumberOfPages > 1 then
	urlSort = urlSort & "&sort=" & sort
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
	         <%if numOfRowProjCl <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word54 title="<%=arrTitles(54)%>" href="<%=urlSort%>&pageProjCl=<%=10*(numOfRowProjCl-1)-9%>&amp;numOfRowProjCl=<%=numOfRowProjCl-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowProjCl-1)) <= NumberOfPages Then
	                  if CInt(pageProjCl)=CInt(i+10*(numOfRowProjCl-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRowProjCl-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&pageProjCl=<%=i+10*(numOfRowProjCl-1)%>&amp;numOfRowProjCl=<%=numOfRowProjCl%>" ><%=i+10*(numOfRowProjCl-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRowProjCl) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word55 title="<%=arrTitles(55)%>" href="<%=urlSort%>&pageProjCl=<%=10*(numOfRowProjCl) + 1%>&amp;numOfRowProjCl=<%=numOfRowProjCl+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	<%list_pr.close 
	set list_pr=Nothing
  End if 
%> 
<% End If%>
<% End If%>
</table></td></tr>
<%
   If trim(companyID) <> "" Then
		if lang_id = "1" then
			arr_Status = Array("","חדש","בטיפול","סגור")	
        else
			arr_Status = Array("","new","active","close")	
        end if
        
       dim sortby_task(12)			
	   sortby_task(1) = "CONTACT_NAME"
	   sortby_task(2) = "CONTACT_NAME DESC "
	   sortby_task(3) = "task_date"
	   sortby_task(4) = "task_date DESC"
	   sortby_task(5) = "task_status,task_date DESC"
	   sortby_task(6) = "task_status DESC,task_date DESC"
	   sortby_task(7) = "sender_name"
	   sortby_task(8) = "sender_name DESC"
	   sortby_task(9) = "reciver_name"
	   sortby_task(10) = "reciver_name DESC"
	   sortby_task(11) = "project_name"
	   sortby_task(12) = "project_name DESC "
		
      
       sort_task = trim(Request.QueryString("sort_task"))
       If Len(sort_task) = 0 Then
		sort_task = 5
	   Else sort_task = trim(Request.QueryString("sort_task"))	
       End If    
       
		
	   If trim(Request("taskTypeID")) <> nil Then
			  taskTypeID = trim(Request("taskTypeID"))
			  search = true 
		Else taskTypeID = ""
		     search = false	
		End If	           		
		
		task_status = trim(Request("task_status")) ' הצגת כל סטטוסי המשימות 		
		title_tasks = trim(Request.Cookies("bizpegasus")("TasksMulti"))   
    
		if trim(Request.QueryString("page_task"))<>"" then
			page_task=Request.QueryString("page_task")
		else
			page_task=1
		end if  
	 
		if trim(Request.QueryString("row_task"))<>"" then
			row_task=Request.QueryString("row_task")
		else
			row_task = 1
		end if  
		PageSize = 5
		'sqlstr = "Select company_Name,CONTACT_NAME,task_Id,task_date,project_name, " &_
		'" task_status,task_types,sender_name,reciver_name,task_content,User_ID,reciver_id " &_
		'" from tasks_view Where company_ID = " & companyID
		'If taskTypeId <> "" Then
		'	sqlstr=sqlstr& " And CHARINDEX('"& taskTypeId &"',task_types) > 0"
		'End if
		'If trim(task_status) <> "" Then		
		'sqlstr=sqlstr& " AND task_status = " & task_status
		'End If
		'sqlstr=sqlstr& " Order By " & sortby_task(sort_task)
		'Response.Write sqlstr
		'Response.End      
		sqlstr = "EXECUTE get_tasks_paging " & page_task & "," & PageSize & ",'','','','" & task_status & "','" & OrgID & "','" & taskTypeID & "','','','" & sortby_task(sort_task) & "','','','" & companyID & "'"
		'Response.Write sqlstr
		'Response.End   
		set tasksList = con.getRecordSet(sqlstr)
		
		If not tasksList.eof Then
			recCount = tasksList("CountRecords")
		End If		
  	    If not tasksList.eof Or search = true Then    
%>
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100%>
  <tr><td width=100%><A name="table_tasks"></A>  
  <table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_obj_var%>">
  <tr>  
  <td class="title_form" width=100% align="<%=align_var%>">&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%>&nbsp;<font color="#E6E6E6">(<%=company_name%>)</font>&nbsp;</td>
  </tr>
  </table></td></tr>  	  
  <tr>
  <td width="100%" valign=top dir="<%=dir_var%>">
   <table width=100% border=0 cellpadding=0 cellspacing=1 bgcolor="#FFFFFF" dir="<%=dir_var%>">   	    
    <tr> 
      <td align=center class="title_sort" width=29 nowrap>&nbsp;</td>     
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort" width=265 nowrap><span id=word19 name=word19><!--תוכן--><%=arrTitles(19)%></span>&nbsp;</td>                  
      <td width=120 nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>" id=td_types name=td_types>&nbsp;<span id="word20" name=word20><!--סוגים--><%=arrTitles(20)%></span>&nbsp;<IMG name=word56 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(56)%>" align=absmiddle onmousedown="taskDropDown(td_types)"></td>
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=120 nowrap valign=top class="title_sort<%if trim(sort_task)="11" OR trim(sort_task)="12" then%>_act<%end if%>"><%if trim(sort_task)="11" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_task)="12" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=11#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%end if%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("Projectone"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="11" then%>bot<%elseif trim(sort_task)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>      
      <td width="90" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_task)="9" OR trim(sort_task)="10" then%>_act<%end if%>"><%if trim(sort_task)="9" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_task)="10" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=9" name=word52 title="<%=arrTitles(52)%>"><%end if%><span id="word21" name=word21><!--אל--><%=arrTitles(21)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="9" then%>bot<%elseif trim(sort_task)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
      <td width="90" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_task)="7" OR trim(sort_task)="8" then%>_act<%end if%>"><%if trim(sort_task)="7" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_task)="8" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=7" name=word52 title="<%=arrTitles(52)%>"><%end if%><span id="word22" name=word22><!--מאת--><%=arrTitles(22)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="7" then%>bot<%elseif trim(sort_task)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>      
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=75 nowrap class="title_sort<%if trim(sort_task)="3" OR trim(sort_task)="4" then%>_act<%end if%>"><%if trim(sort_task)="3" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_task)="4" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=3#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%end if%><span id="word23" name=word23><!--תאריך יעד--><%=arrTitles(23)%></span><img src="../../images/arrow_<%if trim(sort_task)="3" then%>bot<%elseif trim(sort_task)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>          
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=45 nowrap class="title_sort<%if trim(sort_task)="5" OR trim(sort_task)="6" then%>_act<%end if%>"><%if trim(sort_task)="5" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_task)="6" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=5#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%end if%>&nbsp;<span id="word24" name=word24><!--'סט--><%=arrTitles(24)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="5" then%>bot<%elseif trim(sort_task)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>
    </tr>   
 <%		current_date = Day(date()) & "/" & Month(date()) & "/" & Year(date())
		dim	 IS_DESTINATION
		task_types_name = ""
		while not tasksList.EOF       
			taskId = trim(tasksList(1))			
			contact_Name = trim(tasksList(5))
			task_date = trim(tasksList(6))
			project_Name = trim(tasksList(7))  
			task_status = trim(tasksList(8))
			task_types = trim(tasksList(9))      
			sender_name = trim(tasksList(10))
			reciver_name = trim(tasksList(11))      
			task_content = trim(tasksList(12))          
			parentID = trim(tasksList(13))  
			ReciverID = trim(tasksList(14))
			SenderID = trim(tasksList(15))
			task_types_name=""	
			short_types_name=""
				
			If Len(task_types) > 0 Then					
				sqlstr="Select activity_type_name From activity_types Where activity_types.activity_type_id IN (" & task_types & ") Order By activity_types.activity_type_id"
				'Response.Write sqlstr
				'Response.End
				set rssub = con.getRecordSet(sqlstr)		   
					
				While not rssub.eof
					task_types_name = task_types_name & rssub(0) & "<br>"
					rssub.moveNext
				Wend		    			
				set rssub=Nothing								

				If Len(task_types_name) > 0 Then
					task_types_name = Left(task_types_name,(Len(task_types_name)-4))
				End If			
			  
			Else
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
				href = "href=""javascript:addtask('" & contactID & "','" & companyID & "','" & taskID & "')"""   
			Else      
				href = "href=""javascript:closeTask('" & contactID & "','" & companyID & "','" & taskID & "')"""     
			End If
			If trim(taskID) <> "" Then
				sqlstr = "Select TOP 1 task_id from tasks WHERE parent_id = " & taskID
				set rs_par = con.getRecordSet(sqlstr)
				if not rs_par.eof then
					is_parent = true
				else
					is_parent = false	
				end if
				set rs_par = nothing
			End If 		
      %>      
      <tr>  
		<td align=center class="card<%=class_%>" valign=middle>
		<%If trim(taskID) <> "" And is_parent Then%>
		<input type=image src="../../images/hets4.gif" border=0 hspace=0 vspace=0 onclick='window.open("../tasks/task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' ID="Image1" NAME="Image1">
		<%End If%>
		<%If trim(parentID) <> "" And trim(taskID) <> "" Then%>
		<input type=image src="../../images/hets4a.gif" border=0 hspace=0 vspace=0 onclick='window.open("../tasks/task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' ID="Image4" NAME="Image4">
		<%End If%>
		</td>              
       <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card<%=class_%>" valign=top><a class="link_categ" title="<%=vFix(tel_text)%>" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=tel_text_short%>&nbsp;</a></td>
      <td align="<%=align_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=task_types_name%>&nbsp;</a></td>                   
      <td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=project_name%>&nbsp;</a></td>     
      <td align="<%=align_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=reciver_name%>&nbsp;</a></td>      
      <td align="<%=align_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=sender_name%>&nbsp;</a></td>      
      <td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> <%if IS_DESTINATION and task_status <> 3 then%> name=word57 title="<%=arrTitles(57)%>"><span style="width:9px;COLOR: #FFFFFF;BACKGROUND-COLOR: #FF0000;text-align:center"><b>!</b></span><%else%>><%end if%>&nbsp;<%=d_s%>&nbsp;</a></td>            
      <td align=center class="card<%=class_%>" valign=top><a class="task_status_num<%=task_status%>" <%=href%>><%=arr_Status(task_status)%></A></td>
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
			<table border="0" cellspacing="0" cellpadding="2"  dir=ltr>               
	        <% If NumOfPagesTasks > 10 Then 
	              num = 10 : row_task = cInt(NumOfPagesTasks / 10)
	           else num = NumOfPagesTasks : row_task = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if row_task <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word54 title="<%=arrTitles(54)%>" href="<%=urlSort%>&page_task=<%=10*(row_task-1)-9%>&amp;row_task=<%=row_task-1%>#table_tasks" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(row_task-1)) <= NumOfPagesTasks Then
	                  if CInt(page_task)=CInt(i+10*(row_task-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(row_task-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&page_task=<%=i+10*(row_task-1)%>&amp;row_task=<%=row_task%>#table_tasks" ><%=i+10*(row_task-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumOfPagesTasks > cint(num * row_task) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word55 title="<%=arrTitles(55)%>" href="<%=urlSort%>&page_task=<%=10*(row_task) + 1%>&amp;row_task=<%=row_task+1%>#table_tasks" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	
	<%
	End if 
%>
 </table></td></tr>
<%If tasksList.recordCount = 0 Then%>
<tr><td align=center class=card1>&nbsp;</td></tr>									
<%End If%>	 
 </table></td></tr>  
<%	
	End if 
    set tasksList = Nothing
    End If
    
   If trim(companyID) <> "" Then       
        
        dim sortby_doc(2)			
        sortby_doc(0) = "document_id desc"
		sortby_doc(1) = "document_name"
		sortby_doc(2) = "document_name DESC "		
      
        sort_doc = trim(Request.QueryString("sort_doc"))
        If Len(sort_doc) = 0 Then
			sort_doc = 0
		Else sort_doc = trim(Request.QueryString("sort_doc"))	
        End If    
    
		if trim(Request.QueryString("page_doc"))<>"" then
			page_doc=Request.QueryString("page_doc")
		else
			page_doc=1
		end if  
	 
		if trim(Request.QueryString("row_doc"))<>"" then
			row_doc=Request.QueryString("row_doc")
		else
			row_doc = 1
		end if  
		PageSize = 5
		sqlstr = "Select document_id,document_name,document_desc,document_file " &_
		" from company_documents_view Where company_ID = " & companyID		
		sqlstr=sqlstr& " Order By " & sortby_doc(sort_doc)
		'Response.Write sqlstr
		'Response.End      
		set rsdoc = con.getRecordSet(sqlstr)
	
		If not rsdoc.eof  Then
			rsdoc.PageSize=PageSize
			rsdoc.AbsolutePage=page_doc
			records_doc=rsdoc.RecordCount 
			pages_doc=rsdoc.PageCount
			i=1		
	   End If		
  	   If not rsdoc.eof Then    
%>
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100%>
  <tr><td width=100%><A name="table_doc"></A>  
  <table cellpadding=0 cellspacing=0 width=100%>
  <tr>  
  <td class="title_form" width=100%>&nbsp;<span id=word25 name=word25><!--מסמכים מצורפים--><%=arrTitles(25)%></span>&nbsp;<font color="#E6E6E6">(<%=company_name%>)</font>&nbsp;</td>
  </tr>
  </table></td></tr>  	  
  <tr>
  <td width="100%" valign=top dir="<%=dir_var%>">
   <table width=100% border=0 cellpadding=0 cellspacing=1 bgcolor="#FFFFFF" dir="<%=dir_var%>">   	    
    <tr> 
      <td align=center class="title_sort" width=40 nowrap><span id="word26" name=word26><!--מחיקה--><%=arrTitles(26)%></span></td>     
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort" width=100%>&nbsp;<span id="word27" name=word27><!--תיאור--><%=arrTitles(27)%></span>&nbsp;</td>                  
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=193 nowrap class="title_sort<%if trim(sort_doc)="1" OR trim(sort_doc)="2" then%>_act<%end if%>"><%if trim(sort_doc)="1" then%><a class="title_sort" href="<%=urlSort%>&sort_doc=<%=sort_doc+1%>#table_doc" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_doc)="2" then%><a class="title_sort" href="<%=urlSort%>&sort_doc=<%=sort_doc-1%>#table_doc" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_doc=1#table_doc" name=word52 title="<%=arrTitles(52)%>"><%end if%>&nbsp;<span id=word14 name=word14><!--כותרת--><%=arrTitles(14)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_doc)="1" then%>bot<%elseif trim(sort_doc)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>
    </tr>   
 <%
      task_types_name = ""
      do while (not rsdoc.EOF and i<=rsdoc.PageSize)       
		documentID = trim(rsdoc("document_ID"))			
		document_name = trim(rsdoc("document_name"))
	    document_desc = trim(rsdoc("document_desc"))
	    document_file = trim(rsdoc("document_file"))
	    href = " href = ""../../../download/documents/" & document_file & """"
		If cInt(strScreenWidth) > 800 Then
			numOfLetters = 150
		Else
			numOfLetters = 55
		End If
		
		If Len(document_desc) > numOfLetters Then
			document_desc_short = Left(document_desc , numOfLetters-2) & ".."
		Else document_desc_short = document_desc	
		End If
      %>      
      <tr>  
		<td align=center class="card" valign=middle><input type=image name=word53 src="../../images/delete_icon.gif" border=0 vspace=0 hspace=0 title="<%=arrTitles(53)%>" onclick="return CheckDelDocument('<%=companyId%>','<%=documentID%>')" ID="Image5" NAME="Image5"></td>              
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card" valign=top><a class="link_categ" <%=href%> target=_blank dir="<%=dir_obj_var%>" title="<%=vFix(document_desc)%>"><%=document_desc_short%></a></td>   
        <td align="<%=align_var%>" class="card" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> target=_blank dir="<%=dir_obj_var%>"><%=document_name%></a></td>                         
     </tr>
      <%
      rsdoc.moveNext
       i=i+1
	loop
    urlSort = urlSort & "&sort_doc=" & sort_doc
	if pages_doc > 1 then	
%>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr>               
	        <% If pages_doc > 10 Then 
	              num = 10 : row_doc = cInt(pages_doc / 10)
	           else num = pages_doc : row_doc = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if row_doc <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word54 title="<%=arrTitles(54)%>" href="<%=urlSort%>&page_doc=<%=10*(row_doc-1)-9%>&amp;row_doc=<%=row_doc-1%>#table_doc" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(row_doc-1)) <= pages_doc Then
	                  if CInt(page_doc)=CInt(i+10*(row_doc-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(row_doc-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&page_doc=<%=i+10*(row_doc-1)%>&amp;row_doc=<%=row_doc%>#table_doc" ><%=i+10*(row_doc-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if pages_doc > cint(num * row_doc) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word55 title="<%=arrTitles(55)%>" href="<%=urlSort%>&page_doc=<%=10*(row_doc) + 1%>&amp;row_doc=<%=row_doc+1%>#table_doc" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	
	<%
	End if 
%>
 </table></td></tr>	
 </table></td></tr>  
<% 
	End if 
    set rsdoc = Nothing
    End If

	dim sortby_app(16)	
	sortby_app(1) = "appeal_date"
	sortby_app(2) = "appeal_date DESC"
	sortby_app(3) = "appeal_id"
	sortby_app(4) = "appeal_id DESC"
	sortby_app(5) = "User_Name, appeal_id DESC"
	sortby_app(6) = "User_Name DESC, appeal_id DESC"
	sortby_app(7) = "product_name, appeal_id DESC"
	sortby_app(8) = "product_name DESC, appeal_id DESC"
	sortby_app(9) = "CONTACT_NAME, appeal_id DESC"
	sortby_app(10) = "CONTACT_NAME DESC, appeal_id DESC"
	sortby_app(11) = "appeal_status, appeal_id DESC"
	sortby_app(12) = "appeal_status DESC, appeal_id DESC"
	sortby_app(13) = "Company_NAME, appeal_id DESC"
	sortby_app(14) = "Company_NAME DESC, appeal_id DESC"

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
	
    if lang_id = 1 then
		arr_status_app = Array("","חדש","בטיפול","סגור")
	else
		arr_status_app = Array("","new","active","close")
	end if		
	
	
	'sqlStr = "SELECT appeal_id, CONTACT_ID, COMPANY_NAME, PROJECT_NAME, product_id, appeal_date, CONTACT_NAME, "&_
	'" questions_id, User_Name, appeal_status FROM appeals_view WHERE ORGANIZATION_ID=" & trim(OrgID) & where_product &_
	'" AND company_id = " & companyID & " ORDER BY " & sortby_app(sort_app)
	sqlstr = "EXECUTE get_appeals '','','','','" & OrgID & "','" & sortby_app(sort_app) & "','','','" & companyID & "','','','','" & productID & "','" & UserID & "','','','" & Groups & "','" & ResInGroups & "','" & is_groups & "'"
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
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100%>  
  <tr><td width=100%><A name="table_appeals"></A>  
  <table cellpadding=0 cellspacing=0 width=100% ID="Table1">
  <tr>  
  <td class="title_form" width=100%>&nbsp;<span id="word28" name=word28><!--טפסים מצורפים--><%=arrTitles(28)%></span>&nbsp;<font color="#E6E6E6">(<%=company_name%>)</font>&nbsp;</td>
  </tr>
  </table></td></tr>  	  	
  <tr>
	<td width="100%" align="center" valign=top dir="<%=dir_var%>">
	<table BORDER="0" CELLSPACING="1" CELLPADDING="1" bgcolor="#FFFFFF" dir="<%=dir_var%>">	
	<tr>
	    <td width="40" nowrap  class="title_sort" align=center><span id="word29" name=word29><!--הדפס--><%=arrTitles(29)%></span></td>
		<td width="60" nowrap class="title_sort" dir="<%=dir_obj_var%>" align=center><%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%></td>
		<!--td width="100" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=5 or sort_app=6 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=6 then%>5<%elseif sort_app=5 then%>6<%else%>6<%end if%>#table_appeals" target="_self">&nbsp;<span id="word30" name=word30>עובד</span>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="5" then%>bot<%elseif trim(sort_app)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td-->	
		<!--td width="150" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=13 or sort_app=14 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=14 then%>13<%elseif sort_app=13 then%>14<%else%>14<%end if%>#table_appeals" target="_self">&nbsp;<%=trim(Request.Cookies("bizpegasus")("Projectone"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="13" then%>bot<%elseif trim(sort_app)="14" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td-->
		<td width="150" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=13 or sort_app=14 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=14 then%>13<%elseif sort_app=13 then%>14<%else%>14<%end if%>#table_appeals" target="_self">&nbsp;<%=trim(Request.Cookies("bizpegasus")("Companiesone"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="13" then%>bot<%elseif trim(sort_app)="14" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>		
		<td width="100%" id=td_app_prod name=td_app_prod align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort" nowrap>&nbsp;<span id="word31" name=word31><!--סוג טופס--><%=arrTitles(31)%></span>&nbsp;<IMG name=word56 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(56)%>" align=absmiddle onmousedown="appealDropDown(td_app_prod)"></td>
		<td width="70" nowrap align=center dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=1 or sort_app=2 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=2 then%>1<%elseif sort_app=1 then%>2<%else%>2<%end if%>#table_appeals" target="_self">&nbsp;<span id="word32" name=word32><!--תאריך--><%=arrTitles(32)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="1" then%>bot<%elseif trim(sort_app)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
		<%key_table_width=250%>	
		<td class="title_sort" >&nbsp;</td>
		<td width="50" nowrap align=center dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=3 or sort_app=4 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=4 then%>3<%elseif sort_app=3 then%>4<%else%>4<%end if%>#table_appeals" target="_self">&nbsp;ID&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="3" then%>bot<%elseif trim(sort_app)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="50"  nowrap align=center dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=11 or sort_app=12 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=12 then%>11<%elseif sort_app=11 then%>12<%else%>12<%end if%>#table_appeals" target="_self">&nbsp;<span id="word33" name=word33><!--'סט--><%=arrTitles(33)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="11" then%>bot<%elseif trim(sort_app)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	</tr>
	<%	do while (not app.eof and j<app.PageSize)
		appid=app("appeal_id")			
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
		User_Name = app("User_Name")		
		prod_id = app("product_id")
		quest_id = trim(app("questions_id"))
		If trim(quest_id) <> "" Then
			sqlstr = "Select PRODUCT_NAME FROM PRODUCTS WHERE PRODUCT_ID = " & quest_id
			set rs_name = con.getRecordSet(sqlstr)
			If not rs_name.eof Then
				if len(rs_name("product_name")) > 30 then
					product_name = left(rs_name("product_name"),27) & "..."
				else
					product_name = rs_name("product_name")
				end if	
			End if
			set rs_name = Nothing
		End If	
		sqlstr = "EXECUTE get_appeal_tasks '" & OrgID & "','" & appid & "'"	
	    set rsmess = con.getRecordSet(sqlstr)
		'set rsmess = con.getRecordSet("Select mes_new,mes_work,mes_close From all_tasks_view Where appeal_id=" & appid)
		if not rsmess.eof then
			mes_new = rsmess("mes_new")
			mes_work = rsmess("mes_work")
			mes_close = rsmess("mes_close") 
		else
			mes_new = 0
			mes_work = 0
			mes_close = 0
		end if
		If trim(SURVEYS)  = "1" Then
		    href_ = " HREF=""../appeals/appeal_card.asp?quest_id=" & quest_id & "&appid=" & appid & """"
		Else
			href_ = " nohref"
		End If 
		appeal_status = trim(app("appeal_status"))			
		%>
		<tr>
		    <td align="center" class="card" nowrap>&nbsp;<a  href="#" onclick="javascript:window.open('../appeals/view_appeal.asp?quest_id=<%=quest_id%>&appid=<%=appid%>','','top=100,left=100,width=500,height=500,scrollbars=1,resizable=1,menubar=1')"><IMG SRC="../../images/print_icon.gif" BORDER=0 hspace=0 vspace=0></a>&nbsp;</td>
			<td align="center" class="card" nowrap>						
			&nbsp;
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=3" target="_self" style="WIDTH:10pt;" class="task_status_num3" title="<%=arr_status_app(3)%>"><%=mes_close%></a>
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=2" target="_self" style="WIDTH:10pt;" class="task_status_num2" title="<%=arr_status_app(2)%>"><%=mes_work%></a>
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=1" target="_self" style="WIDTH:10pt;" class="task_status_num1" title="<%=arr_status_app(1)%>"><%=mes_new%></a>
			&nbsp;		
			</td>
			<!--td class="card" nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> dir="<%=dir_obj_var%>">&nbsp;<%=User_NAME%>&nbsp;</a></td-->
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> dir="<%=dir_obj_var%>">&nbsp;<%=COMPANY_NAME%>&nbsp;</a></td>
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> dir="<%=dir_obj_var%>">&nbsp;<%=product_name%>&nbsp;</a></td>
			<td class="card" align=center><a class="link_categ" <%=href_%> >&nbsp;<%=day(app("appeal_date"))%>/<%=month(app("appeal_date"))%>/<%=mid(year(app("appeal_date")),3,2)%>&nbsp;</a></td>			
			<td class="card" align="<%=align_var%>">
		    <!--#include file="../appeals/key_fields.asp"-->
			</td>		
			<td class="card" nowrap align=center><a class="link_categ" <%=href_%>  dir="<%=dir_obj_var%>"><%=appid%></a></td>
			<td class="card" nowrap align=center><a class=status_num<%=appeal_status%> <%=href_%> target="_self"><%=arr_status_app(appeal_status)%></a></td>
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
		<% if NumberOfPagesApp > 1 then		   
		%>
		<tr class="card">
		<td width="100%" align=center nowrap class="card" dir=ltr>
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
			 <td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=10*(numOfRowApp-1)-9%>&numOfRowApp=<%=numOfRowApp-1%>#table_appeals" name=word54 title="<%=arrTitles(54)%>"><b><<</b></a></td>			                
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
					<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=10*(numOfRowApp) + 1%>&numOfRowApp=<%=numOfRowApp+1%>#table_appeals" name=word55 title="<%=arrTitles(55)%>"><b>>></b></a></td>
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
<%
	dim sortby_f(16)	
	sortby_f(1) = "appeal_date"
	sortby_f(2) = "appeal_date DESC"
	sortby_f(3) = "appeal_id"
	sortby_f(4) = "appeal_id DESC"
	sortby_f(5) = "PEOPLE_EMAIL"
	sortby_f(6) = "PEOPLE_EMAIL DESC"
	sortby_f(7) = "product_name,appeal_date"
	sortby_f(8) = "product_name DESC,appeal_date"
	sortby_f(9) = "PEOPLE_NAME"
	sortby_f(10) = "PEOPLE_NAME DESC"
	sortby_f(11) = "PEOPLE_COMPANY,appeal_date"
	sortby_f(12) = "PEOPLE_COMPANY DESC,appeal_date"
	sortby_f(13) = "appeal_status,appeal_date"
	sortby_f(14) = "appeal_status DESC,appeal_date"

	sort_f = Request("sort_f")	
	if sort_f = nil then
		sort_f = 2
	end if
	
	if lang_id = 1 then
		arr_status_f = Array("","חדש","בטיפול","סגור")
	else
		arr_status_f = Array("","new","active","close")
	end if

	'sqlStr = "SELECT appeal_id, PEOPLE_ID, PEOPLE_COMPANY, product_id, appeal_date, PEOPLE_NAME, "&_
	'" PEOPLE_EMAIL,questions_id,groupID,appeal_status FROM feedbacks_view WHERE ORGANIZATION_ID=" & trim(OrgID) &_
	'" AND COMPANY_ID = " & companyID & " ORDER BY " & sortby_f(sort_f)
	sqlstr = "EXECUTE get_feedbacks '','','','','" & OrgID & "','" & sortby_f(sort_f) & "','','','" & companyID & "','','','',''"
    'Response.Write sqlStr
	set app=con.GetRecordSet(sqlStr)
	app_count = app.RecordCount
	if Request("page_f")<>"" then
		page_f=request("page_f")
	else
		page_f=1
	end if
	%>		
	<%	if not app.eof then %>
  <tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100%>  
  <tr><td width=100%><A name="table_feedbacks"></A>  
  <table cellpadding=0 cellspacing=0 width=100%>
  <tr>  
  <td class="title_form" width=100%>&nbsp;<span id=word34 name=word34><!--משובים מדיוור--><%=arrTitles(34)%></span>&nbsp;<font color="#E6E6E6">(<%=company_name%>)</font>&nbsp;</td>
  </tr>
  </table></td></tr>  	  	
  <tr>  
	<td width="100%" align="center" valign=top dir=ltr>
	<table BORDER="0" CELLSPACING="1" CELLPADDING="1" bgcolor="#FFFFFF">	
	<tr>	    
		<td width="180" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_f=11 or sort_f=12 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=12 then%>11<%elseif sort_f=11 then%>12<%else%>12<%end if%>#table_appeals" target="_self">&nbsp;Email&nbsp;<img src="../../images/arrow_<%if trim(sort_f)="11" then%>bot<%elseif trim(sort_f)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
		<td width="150" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_f=5 or sort_f=6 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=6 then%>5<%elseif sort_f=5 then%>6<%else%>6<%end if%>#table_appeals" target="_self">&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_f)="5" then%>bot<%elseif trim(sort_f)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>			
		<td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_f=7 or sort_f=8 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=7 then%>8<%elseif sort_f=8 then%>7<%else%>7<%end if%>#table_appeals" target="_self">&nbsp;<span id="word35" name=word35><!--סוג טופס--><%=arrTitles(35)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_f)="7" then%>bot<%elseif trim(sort_f)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="60" nowrap dir="<%=dir_obj_var%>" align=center class="title_sort<%if sort_f=1 or sort_f=2 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=2 then%>1<%elseif sort_f=1 then%>2<%else%>2<%end if%>#table_appeals" target="_self">&nbsp;<span id="word36" name=word36><!--תאריך--><%=arrTitles(36)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_f)="1" then%>bot<%elseif trim(sort_f)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
		<td width="40" nowrap dir="<%=dir_obj_var%>" align=center class="title_sort<%if sort_f=3 or sort_f=4 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=4 then%>3<%elseif sort_f=3 then%>4<%else%>4<%end if%>#table_feedbacks" target="_self">&nbsp;ID&nbsp;<img src="../../images/arrow_<%if trim(sort_f)="3" then%>bot<%elseif trim(sort_f)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="45" nowrap dir="<%=dir_obj_var%>" align=center class="title_sort<%if sort_f=13 or sort_f=14 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=14 then%>13<%elseif sort_f=13 then%>14<%else%>14<%end if%>#table_feedbacks" target="_self">&nbsp;<span id="word37" name=word37><!--'סט--><%=arrTitles(37)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_f)="13" then%>bot<%elseif trim(sort_f)="14" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	</tr>
	<%	
		if isNumeric(trim(session("quantRows"))) = false then
			app.PageSize = 10
		else
			app.PageSize = session("quantRows")
		end if	
		app.AbsolutePage=page_f
		recCount=app.RecordCount 		
		NumberOfPagesF = app.PageCount
		i=1
		j=0
		ids = "" 'list of appeal_id
		do while (not app.eof and j<app.PageSize)
		appid=app("appeal_id")	
		PEOPLE_ID = app("PEOPLE_ID")		
		PEOPLE_EMAIL = app("PEOPLE_EMAIL")
		PEOPLE_COMPANY = app("PEOPLE_COMPANY")		
		If trim(PEOPLE_COMPANY) = "" Or IsNull(PEOPLE_COMPANY) Then
			PEOPLE_COMPANY = ""
		End If	
		PEOPLE_NAME = app("PEOPLE_NAME")
		If trim(PEOPLE_NAME) = "" Or IsNull(PEOPLE_NAME) Then
			PEOPLE_NAME = ""
		End If			
	
		groupID = app("groupID")		
		prod_id = app("product_id")
		quest_id = trim(app("questions_id"))
		If trim(quest_id) <> "" Then
			sqlstr = "Select PRODUCT_NAME FROM PRODUCTS WHERE PRODUCT_ID = " & quest_id
			set rs_name = con.getRecordSet(sqlstr)
			If not rs_name.eof Then
				if len(rs_name("product_name")) > 30 then
					product_name = left(rs_name("product_name"),27) & "..."
				else
					product_name = rs_name("product_name")
				end if	
			End if
			set rs_name = Nothing
		End If	
		If trim(SURVEYS)  = "1" Then
		    href_ = " HREF='../appeals/feedback.asp?quest_id=" & quest_id & "&appid=" & appid & "'"
		Else
			href_ = " nohref"
		End If
		feedback_status = trim(app("appeal_status"))	 
		%>
		<tr>			
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=PEOPLE_EMAIL%>&nbsp;</a></td>
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=PEOPLE_NAME%>&nbsp;</a></td>			
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=product_name%>&nbsp;</a></td>
			<td class="card" align=center><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=day(app("appeal_date"))%>/<%=month(app("appeal_date"))%>/<%=mid(year(app("appeal_date")),3,2)%>&nbsp;</a></td>
			<td class="card" nowrap align=center><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>"><%=appid%></a></td>
			<td class="card" nowrap align=center><a class=status_num<%=feedback_status%> <%=href_%> target="_self"><%=arr_status_f(feedback_status)%></a></td>
		</tr>
<%		app.movenext
		j=j+1
		if not app.eof and j <> app.PageSize then
		ids = ids & ","
		end if
		loop 
		%>
		<% if NumberOfPagesF > 1 then%>
		<tr class="card">
		<td width="100%" align=center nowrap class="card">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr>               
	        <% If NumberOfPagesF > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPagesF / 10)
	           else num = NumberOfPagesF : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRowF") <> nil Then
	               numOfRowF = Request.QueryString("numOfRowF")
	           Else numOfRowF = 1
	           End If
	         %>
	         
	         <tr>
	         <%if numOfRowF <> 1 then%> 
			 <td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sort_f=<%=sort_f%>&page_f=<%=10*(numOfRowF-1)-9%>&numOfRowF=<%=numOfRowF-1%>" name=word54 title="<%=arrTitles(54)%>"><b><<</b></a></td>			                
			  <%end if%>
	             <td><font size="2" color="#060165">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowF-1)) <= NumberOfPagesF Then
	                  if CInt(page_f)=CInt(i+10*(numOfRowF-1)) then %>
		                 <td align="center"><span class="pageCounter"><%=i+10*(numOfRowF-1)%></span></td>
	                  <%else%>
	                     <td align="center"><a class="pageCounter" href="<%=urlSort%>&sort_f=<%=sort_f%>&page_f=<%=i+10*(numOfRowF-1)%>&numOfRowF=<%=numOfRowF%>"><%=i+10*(numOfRowF-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td><font size="2" color="#060165">]</font></td>
				<%if NumberOfPagesF > cint(num * numOfRowF) then%>  
					<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sort_f=<%=sort_f%>&page_f=<%=10*(numOfRowF) + 1%>&numOfRowF=<%=numOfRowF+1%>" name=word55 title="<%=arrTitles(55)%>"><b>>></b></a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>				
	<%End If%>		
</table></td></tr>
<%	set app = nothing	%>
</table></td></tr>
<%    end if ' if not app.eof%>
</table></td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:106;" href="#" onclick="javascript:window.open('../projects/editProject.asp?companyID=<%=companyID%>&contactID=<%=contactID%>','addProject','top=120,left=120,resizable=1,scrollbars=1,width=480,height=450');"><span id=word38 name=word38><!--הוסף--><%=arrTitles(38)%></span>&nbsp;<%=trim(Request.Cookies("bizpegasus")("Projectone"))%></a></td></tr>
<tr><td align="center" nowrap colspan=2><a class="button_edit_1" style="width:106;" href="#" onclick='window.open("../tasks/addtask.asp?companyID=<%=companyID%>&contactID=<%=contactID%>", "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width=420,height=525,align=center,resizable=0");'><span id="word39" name=word39><!--הוסף--><%=arrTitles(39)%></span>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a></td></tr>
<tr><td align="center" nowrap colspan=2><a class="button_edit_1" style="width:106;" href="#" onclick='window.open("adddoc.asp?companyID=<%=companyID%>", "Documents" ,"scrollbars=1,toolbar=0,top=150,left=120,width=450,height=200,align=center,resizable=0");'><span id="word41" name=word41><!--הוסף מסמך--><%=arrTitles(41)%></span></a></td></tr>
<%If trim(SURVEYS)  = "1" Then%>
<%
	If is_groups = 0 Then
	sqlstr = "Select product_id, product_name from Products Where isNULL(FORM_CLIENT,0) = 1 And "&_
	" product_number = '0' and ORGANIZATION_ID=" & OrgID & " order by product_name"
	Else
	sqlstr = "Select DISTINCT Products.product_id, Products.product_name from Products Inner Join Users_To_Products "&_
	" ON Products.Product_ID = Users_To_Products.Product_ID WHERE Users_To_Products.User_ID = " & UserID &_
	" And isNULL(FORM_CLIENT,0)=1 And Products.product_number = '0' and Products.ORGANIZATION_ID=" & OrgID & " order by Products.product_name"
	End If
	'Response.Write sqlstr
	'Response.End
	set rs_products = con.GetRecordSet(sqlstr)
	if not rs_products.eof then 
		ClientProductsList = rs_products.getString(,,",",";")		   
		arr_products = Split(ClientProductsList,";")
	end if
	set rs_products=nothing	
	If IsArray(arr_products) Then
%>
<tr><td height=1 nowrap colspan=2 bgcolor="#808080"></td></tr>
<tr><td align="center" colspan=2 class="title_form"><span id="word42" name=word42><!--צרף טופס--><%=arrTitles(42)%></span></td></tr>
<%    
	For i=0 To Ubound(arr_products)-1	
	arr_prod = Split(arr_products(i),",")
	If IsArray(arr_prod) Then
    quest_id = arr_prod(0)   	
    product_name = arr_prod(1)
%>
<tr><td align="<%=align_var%>" colspan=2><a class="link1" href="#" style="width:106;line-height:110%;padding:4px;direction:rtl;font-weight:500" onclick="javascript:window.open('../appeals/edit_appeal.asp?companyID=<%=companyID%>&contactID=<%=contactID%>&quest_id=<%=quest_id%>','','top=100,left=100,width=500,height=500,scrollbars=1,resizable=1')" ><%=product_name%></a></td></tr>
<% End If
	Next	
	End If	
End If%>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:106;" href="#" onclick="javascript:window.open('editcompany_private.asp?companyID=<%=companyID%>&contactID=<%=contactID%>','','top=50,left=100,resizable=0,width=460,height=500');"><span id="word43" name=word43><!--עדכן פרטי--><%=arrTitles(43)%></span> <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></a></td></tr>	
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:106;" href="#" ONCLICK="return CheckDel()"><span id="word45" name=word45><!--מחק--><%=arrTitles(45)%></span> <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></a></td></tr>	
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:106;" href="#" onclick="javascript:window.open('printcompany_private.asp?companyID=<%=companyID%>&contactID=<%=contactID%>','','top=100,left=20,resizable=0,width=660,height=500,scrollbars=1,menubar=1');"><span id="word44" name=word44><!--הדפס--><%=arrTitles(44)%></span> <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></a></td></tr>
<tr><td colspan=2 height=10 nowrap></td></tr>
</table>
</td>
</tr>
</table>
<%End If%>
<DIV ID="task_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:100%; height:68; overflow:scroll; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #E6E6E6; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%  sqlstr = "Select activity_type_id, activity_type_name from activity_types Where ORGANIZATION_ID = "&OrgID&" Order By activity_type_id"
	set rsactivity = con.getRecordSet(sqlstr)
	while not rsactivity.eof %>
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; text-align:right; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>&taskTypeID=<%=rsactivity(0)%>#table_tasks'">
    <%=rsactivity(1)%>
    </DIV>
	<%
		rsactivity.moveNext
		Wend
		set rsactivity=Nothing
	%>
	<DIV dir="<%=dir_obj_var%>"  onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>#table_tasks'">
    <span id="word47" id=word47><!--כל הרשימה--><%=arrTitles(47)%></span>
    </DIV>
</div>
</DIV>
<DIV ID="appeal_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:100%; height:68; overflow:scroll; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #E6E6E6; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%
	If is_groups = 0 Then
	sqlstr = "Select product_id, product_name from Products Where "&_
	" product_number = '0' and ORGANIZATION_ID=" & OrgID & " order by product_name"
	' משתמש אשר שייך לקבוצה אבל אינו אחראי באף קבוצה
	ElseIf Len(Groups) > 0 And (Len(ResInGroups) = 0 Or trim(ResInGroups) = "-1") Then
		sqlstr = "Select DISTINCT Products.product_id, Products.product_name from Products Inner Join Appeals "&_
		" ON Products.Product_ID = Appeals.Questions_ID WHERE (Appeals.User_ID = " & UserID &_
		" Or (Appeals.questions_id IN (Select product_id From Users_To_Products WHERE Group_ID IN (" & Groups & "))" &_
 		" AND products.OPEN_FORM = 1 And appeals.User_Id IS NULL))" &_
 		" Or Products.Responsible = " & UserID &_
		" And Products.PRODUCT_NUMBER = '0' And Products.ORGANIZATION_ID=" & OrgID &_
		" Order By Products.product_name"
	Else
		sqlstr = "Select DISTINCT Products.product_id, Products.product_name from Products Inner Join Users_To_Products "&_
		" ON Products.Product_ID = Users_To_Products.Product_ID WHERE Users_To_Products.Group_ID IN (" &_	
		" Select DISTINCT Group_ID FROM Responsibles_To_Groups WHERE Responsible_ID = " & UserID & ")" &_
		" Or Products.Product_ID IN (Select DISTINCT Questions_ID From APPEALS WHERE User_ID = " & UserID & ")" &_ 
		" Or Products.Responsible = " & UserID &_
		" And Products.PRODUCT_NUMBER = '0' and Products.ORGANIZATION_ID=" & OrgID & " order by Products.product_name"
	End If
	'Response.Write sqlstr
	'Response.End
	set rs_products = con.GetRecordSet(sqlstr)
	if not rs_products.eof then 
		ResProductsList = rs_products.getString(,,",",";")		   
		arr_products = Split(ResProductsList,";")									
	end if
	set rs_products=nothing				
	If IsArray(arr_products) Then
	For i=0 To Ubound(arr_products)-1	
	arr_prod = Split(arr_products(i),",")
	If IsArray(arr_prod) Then
%>
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; text-align:right; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>&productID=<%=arr_prod(0)%>#table_appeals'">
    <%=arr_prod(1)%>
    </DIV>
<%
	End If
	Next	
	End If	
%>
	<DIV dir="<%=dir_obj_var%>"  onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>#table_appeals'">
    <span id="word48" id=word48><!--כל הרשימה--><%=arrTitles(48)%></span>
    </DIV>
</div>
</DIV>
<div id=div_comp_desc name=div_comp_desc style="display:none;visibility:hidden;position:absolute;left:30;top:430;width:350;height:100;z-index:11;">
<table cellpadding=0 cellspacing=1 border=1 bgcolor="#ffffff" width=100%>
 <tr>
        <td width="100%" bgcolor="#0F2771" align="<%=align_var%>">
            <table border="0" width="100%" cellspacing="0" cellpadding="0">    
            <tr>
				<td colspan=2 width="100%" bgcolor="#616161" height="1"></td>
			</tr>	
            <tr>
                <td width=20 nowrap align=center class=title_form><INPUT type=image src="../../images/close_icon.gif" border="0" onClick="return closeDesc();" vspace=0 hspace=0 ID="Image3" NAME="Image3">              
                </td>            
                <td width=330 nowrap class=title_form align="<%=align_var%>">&nbsp;<span id=word49 name=word49><!--פרטים נוספים--><%=arrTitles(49)%></span>&nbsp;</td>
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
</body>
<%set con=Nothing%>
</html>

