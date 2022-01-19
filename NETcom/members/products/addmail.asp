<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%	prodId=trim(Request("prodId"))		
	groups_id=trim(session("wgroupID"))
	If trim(groups_id) <> "" Then
		sqlstr="Select GROUPNAME FROM GROUPS WHERE GROUPID IN (" & groups_id & ")"
		set rsn = con.getRecordSet(sqlstr)
		If not rsn.eof Then
			groupname = trim(rsn.getString(,," ; "," ; "))
		End if
		set rsn = Nothing	
	
		If Len(groupname) > 1 Then
			groups_list = Left(groupname,Len(groupname)-1) 
		End If
		sqlstr="Select COUNT(PEOPLE_ID) FROM PEOPLES WHERE GROUPID IN (" & groups_id & ")"
		set rs_count = con.getRecordSet(sqlstr)
		If not rs_count.eof Then
			count_all = trim(rs_count(0))
		End If		
		set rs_count = Nothing
		count_all = "(" & count_all & " נמענים)"
	End If%>
<%	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 77 Order By word_id"				
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
	  
	  sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	  set rsbuttons = con.getRecordSet(sqlstr)
	  If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	  End If
	  set rsbuttons=nothing%>	
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--
function CheckField() {
 
  if(window.document.dataForm.groups_id.value == '')
  {
	 window.alert("! חובה לבחור נמענים להפצה");	
	 return false;
  }
  if(window.document.dataForm.EMAIL_SUBJECT.value == '')
  {
	 window.alert("! נא להכניס נושא של הפצה");
	 window.document.dataForm.EMAIL_SUBJECT.focus();
	 window.document.dataForm.EMAIL_SUBJECT.scrollIntoView();
	 return false;
  } 
  if (window.document.dataForm.attachment.value !='')
  {
		var fname=new String();
		var fext=new String();
		var extfiles=new String();
		fname=window.document.dataForm.attachment.value;
		indexOfDot = fname.lastIndexOf('.')
		fext=fname.slice(indexOfDot+1,-1)		
		fext=fext.toUpperCase();
		extfiles='HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT';		
		if ((extfiles.indexOf(fext)>-1) == false)
		{	
			window.alert(':סיומת הקובץ - אחת מרשימה' + '\n' + 'HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT');
		    return false;
		}    
  }    
  if (window.confirm("?האם ברצונך לתעד שליחת דואר לנמענים שנבחרו"))
  {
	  save_win = window.open("","save_win","scrollbars=1,toolbar=0,top=100,left=100,width=450,height=250,align=center,resizable=1")
	  window.document.dataForm.target = "save_win";
	  window.document.dataForm.submit();
  }
  return false;
}

function popupcal(elTarget){
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

function GetEnglish()
{
	var ch=event.keyCode;
	event.returnValue = ch < 126;
}
 
function openGroups()
{
	window.open("groups.asp?grp_type=2&groups_id="+window.document.all("groups_id").value,"","left=250,top=100,height=350,width=390,status=no,toolbar=no,scroll=no,menubar=no,location=no,scrollbars=yes")
}

function openPage(selObj)
{
	if(selObj.value != "")
	{	
		window.document.all('page_content').src='../pages/result.asp?pageId='+selObj.value;		
		window.document.all('page_content').style.display = "inline";
		//window.document.all('page_content').contentWindow.document.body.style.scrollbarBaseColor="#00BFFF";
	}
	else
		window.document.all('page_content').style.display = "none";
	return false;	
}
//-->
</script>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 2%>
<%numOfLink = 2%>
<!--#include file="../../top_in.asp"-->
<div align="<%=align_var%>"><center>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr><td class="page_title"><!--תיעוד שליחת דואר--><%=arrTitles(9)%>&nbsp;</td></tr>
<tr><td width="100%">
<table cellpadding=0 cellspacing=0 width=100% bgcolor="#E6E6E6">
<tr><td height=10 nowrap></td></tr>
<tr><td>
<FORM name="dataForm" ACTION="save_mail.asp" METHOD="post" ID="dataForm" enctype="multipart/form-data">
<table align=center border="0" cellpadding="1" cellspacing="1" width="650" ID="Table10">
<tr>
	<td align="<%=align_var%>" height=15 nowrap><INPUT type=hidden name="langu" id="langu" value="heb"></td>	
</tr>
<tr>
<td align="<%=align_var%>">
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
	<td width=100% align="<%=align_var%>" id="groups_name" name="groups_name" dir="<%=dir_obj_var%>">&nbsp;<span dir="<%=dir_obj_var%>"><%=groups_list%></span>&nbsp;<font dir="<%=dir_obj_var%>" color=red><%=count_all%></font></td>
	<input dir="ltr" type="hidden" id="groups_id" name="groups_id" value="<%=vFix(groups_id)%>">
	<td align="<%=align_var%>" width=100 dir="<%=dir_obj_var%>" nowrap>	
	<input type=button class="but_menu" value="<%=arrTitles(10)%>" onclick="return openGroups()" style="width:120" ID="Button1" NAME="Button1">
	</td>
</tr>
</table>
</td></tr>
<tr>
<td align="<%=align_var%>">
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
	<td align="<%=align_var%>" width=100%><input dir="<%=dir_obj_var%>" type="text" class="texts" style="width:350" id="EMAIL_SUBJECT" name="EMAIL_SUBJECT" value="<%=vFix(subjectEmail)%>" MaxLength=100></td>
	<td align="<%=align_var%>" width=100 nowrap dir="<%=dir_obj_var%>"><b>&nbsp;<!--נושא--><%=arrTitles(5)%>&nbsp;</b></td>
</tr>
</table>
</td></tr>
<tr>
<td align="<%=align_var%>">
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
    <td align="<%=align_var%>" width=100%><input type="file" name="attachment" ID="attachment" size=41 value=""></td>
	<td align="<%=align_var%>" width=100 nowrap dir="<%=dir_obj_var%>"><b><!--קובץ מצורף--><%=arrTitles(6)%></b></td></tr>
</table>
</td></tr>
<tr>
<td align="<%=align_var%>">
<table cellpadding=0 cellspacing=0 width=100% border=0>
<tr>
	<td align="<%=align_var%>" width=100%>
	  <select name="page_id" id="page_id" class="norm" style="width:350" dir="<%=dir_obj_var%>" onchange="return openPage(this)">
		<option value="">&nbsp;<!--בחר דף הפצה--><%=arrTitles(8)%>&nbsp;</OPTION>
<%  set prod = con.GetRecordSet("Select page_id, page_title from pages where ORGANIZATION_ID=" & trim(Request.Cookies("bizpegasus")("OrgID")) & " order by page_id desc")
    if not prod.eof then 
	  do while not prod.eof
    	page_Id = trim(prod(0))
    	pageTitle = trim(prod(1))
%>		  
		<option value="<%=page_Id%>" <%if trim(wpageID)=trim(page_Id) then%> selected<%end if%>>&nbsp;<%=pageTitle%>&nbsp;</OPTION>
<%		prod.MoveNext
	  loop
	end if
	set prod=nothing%>  
		</select>
	</td>
	<td align="<%=align_var%>" width=100 nowrap><b>&nbsp;<!--תוכן--><%=arrTitles(7)%>&nbsp;</b></td>
</tr>
</table>
</td></tr>
<tr>
<td align="<%=align_var%>">
<table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
<tr>
	<td align="<%=align_var%>">
	<iframe frameborder=0 style="border:1px solid" name="page_content" id="page_content" <% If trim(wpageID) <> "" Then%> src="../pages/result.asp?pageID=<%=wpageID%>" <%Else%> src="" style="display:none;" <%End If%> width=640 height=500 MARGINWIDTH=0 MARGINHEIGHT=0 hspace=0 vspace=0></iframe>
	</td>
</tr>
</table>
</td></tr>	
<tr><td height="10">
<input type="hidden" name="prodId" value="<%=prodId%>">
<input type="hidden" name="editFlag" value="yes">
</td></tr>
<tr><td colspan="2">
<table width=100% border=0 cellspacing=0 dir="<%=dir_var%>">
<tr><td width=48% align="<%=align_var%>">
<INPUT class="but_menu" style="width:90" type="button" value="<%=arrButtons(2)%>" id=button2 name=button2 onclick="document.location.href='products.asp'"></td>
<td width=4% nowrap></td>
<td width=48%  dir="<%=dir_var%>">
<input class="but_menu" style="width:90" type="button" value="<%=arrButtons(1)%>" onClick="return CheckField();"></td>
</tr>
</table>
</td></tr>
</table></form>
</td></tr></table>
</td></tr>
<tr><td height=10 nowrap></td></tr>
</table></center></div>
</body>
</html>
<%set con=nothing%>