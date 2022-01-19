<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
prod_id = Request.QueryString("prodId")
date_year = Year(date())
date_month = Month(date())
date_day = Day(date())
PathCalImage = "../../"
dateStart = Day(date()) & "/" & Month(date()) & "/" & Year(date())
project_title = trim(Request.Cookies("bizpegasus")("Projectone"))
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 15 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(trim(arr_title(0))) = arr_title(1)
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
	  set rsbuttons=nothing	  
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="javascript">
<!--
function checkfields(){
	return confirm('                                              ! שים לב \n\n.לא ניתן לשנות נרשמים לאחר השליחה\n\n       ? האם אתה בטוח שברצונך לשלוח')
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


//-->
</script>
</head>

<body style="margin:0px;background-color:#E5E5E5">
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr>
   <td align="left" valign="middle" nowrap>
	 <table width="100%" border="0" cellpadding="0" cellspacing="0">
	  <tr><td class="page_title"><%if project_ID<>nil then%><span id=word1 name=word1><!--עדכן--><%=arrTitles(1)%></span><%else%><span id=word2 name=word2><!--הוסף--><%=arrTitles(2)%></span><%end if%>&nbsp;<%=project_title%></td></tr>		   	  		       	
   </table>
</td></tr>  
<tr><td width=100% bgcolor="#E6E6E6" dir="<%=dir_var%>"> 
<TABLE WIDTH=99% BORDER=0 CELLSPACING=1 CELLPADDING=3 bgcolor=white align=center dir="<%=dir_var%>">
<FORM name="frmMain" ACTION="editProject.asp?project_ID=<%=project_ID%>" METHOD="post" onSubmit="return CheckFields()" ID="Form1">
<tr><td align="<%=align_var%>" colspan=2 style="padding-right:15px;padding-left:15px;" class=form bgcolor=#DADADA nowrap><b><span id=word4 name=word4><!--שם--><%=arrTitles(4)%></span></b></td></tr>
<tr><td align="<%=align_var%>" colspan=2 style="padding-right:15px;padding-left:15px;" width=100% bgcolor=#f0f0f0><input type=text dir="<%=dir_obj_var%>" name="project_name" value="<%=vFix(project_name)%>" style="width:300px;font-family:arial" ID="Text1"></td></tr>
<tr>
	<td align="<%=align_var%>" colspan=2 style="padding-right:15px;padding-left:15px;" class=form bgcolor=#DADADA nowrap valign=top><b><span id="word5" name=word5><!--תיאור--><%=arrTitles(5)%></span></b></td>
</tr>
<tr>
	<td align="<%=align_var%>" colspan=2 style="padding-right:15px;padding-left:15px;" bgcolor=#f0f0f0><textarea dir="<%=dir_obj_var%>" name="project_description" style="width:300px;font-family:arial" ID="Textarea1"><%=vFix(project_description)%></textarea></td>
</tr>
<tr>	
	<td align="<%=align_var%>" colspan=2 style="padding-right:15px;padding-left:15px;" class=form bgcolor=#DADADA nowrap><b><span id="word6" name=word6><!--קוד--><%=arrTitles(6)%></span></b></td>
</tr>
<tr>
	<td align="<%=align_var%>" colspan=2 style="padding-right:15px;padding-left:15px;" bgcolor=#f0f0f0><input type=text dir="<%=dir_obj_var%>" name="project_code" value="<%=vfix(project_code)%>" style="width:100px;font-family:arial" ID="project_code" maxlength=50></td>
</tr>

<% if project_ID=nil or project_ID="" then %>
<tr>
	<td align="<%=align_var%>" width=100% bgcolor=#f0f0f0 dir="<%=dir_obj_var%>"><b><span id=word7 name=word7><!--עתידי--><%=arrTitles(7)%></span></b><font color="#736BA6"></font></td>
	<td align=left width=100 class=form bgcolor=#DADADA nowrap><input type=radio dir="<%=dir_obj_var%>" name="project_status" <%If trim(project_ID) = "" Then%> checked <%End if%> onclick="date_body.style.display='none';dateStart.disabled=true;" value=0 ID="Radio1"></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=100% bgcolor=#f0f0f0 dir="<%=dir_obj_var%>"><b><span id="word8" name=word8><!--בביצוע--><%=arrTitles(8)%></span></b><font color="#736BA6"></font></td>
	<td align=left width=100 class=form bgcolor=#DADADA nowrap><input type=radio dir="<%=dir_obj_var%>" name="project_status" onclick="date_body.style.display='inline';dateStart.disabled=false;" value=1 ID="Radio2"></td>	
</tr>
<%End If%>
<tbody name="date_body" id="date_body" <%If trim(dateStart) = ""  Or cInt(status) = 1 Or trim(project_ID) = "" Then%> style="display:none" <%end if%>>
<tr>
	<td align="<%=align_var%>" width=100% bgcolor=#f0f0f0>
	<input type=image src='../../images/calend.gif' border=0  onclick='return popupcal(this.form.elements("dateStart"));' ID="Image1" NAME="Image1">&nbsp;<input dir="ltr" type="text" class="passw" style="width:85" id="dateStart" name="dateStart" value="<%=dateStart%>" <%If trim(project_ID) = "" Then%> disabled <%End If%>  maxlength=10 onclick="return popupcal(this);" readonly>
	</td>
	<td align="<%=align_var%>" width=100 class=form bgcolor=#DADADA nowrap><b><span id="word9" name=word9><!--תאריך פתיחה--><%=arrTitles(9)%></span></b></td>
</tr>
</tbody>
<%If trim(status) = "3" Then%>
<tr>
	<td align="<%=align_var%>" width=100% bgcolor=#f0f0f0>	
	<input type=image src='../../images/calend.gif' border=0  onclick='return popupcal(this.form.elements("dateEnd"));' ID="Image2" NAME="Image2">&nbsp;<input dir="ltr" type="text" class="passw" style="width:85" id="dateEnd" name="dateEnd" value="<%=dateEnd%>" maxlength=10 onclick="return popupcal(this);" readonly>
	</td>
	<td align="<%=align_var%>" width=100 nowrap class=form bgcolor=#DADADA><b><span id="word10" name=word10><!--תאריך סגירה--><%=arrTitles(10)%></span></b></td>
</tr>
<%End If%> 
<tr>	
	<td align="<%=align_var%>" colspan=2 style="padding-right:15px;padding-left:15px;" class=form bgcolor=#DADADA nowrap><b><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></b></td>
</tr>
<tr>
	<td align="<%=align_var%>" colspan=2 style="padding-right:15px;padding-left:15px;" bgcolor=#f0f0f0>
	<select dir="<%=dir_obj_var%>" name="company_id" style="width:300px;font-family:arial" <%If is_worked Then%> disabled <%End if%> ID="Select1">	
	<%If trim(lang_id) = "1" Then%>
	<option value=""><%=String(29,"-")%> בחר <%=String(29,"-")%></option>
	<%Else%>
	<option value=""><%=String(28,"-")%> select <%=String(28,"-")%></option>
	<%End If%>	<%  sqlstr = "Select COMPANY_ID, COMPANY_NAME FROM COMPANIES WHERE ORGANIZATION_ID = " & OrgID & " Order BY PRIVATE_FLAG DESC,COMPANY_NAME"
		set rs_comp = con.getRecordSet(sqlstr)
		While not rs_comp.eof
	%>
	<option value="<%=rs_comp(0)%>" <%If trim(company_id) = trim(rs_comp(0)) Then%> selected <%End If%>><%=rs_comp(1)%></option>
	<%
		rs_comp.moveNext
		Wend
		set rs_comp = Nothing
	%>
	</select></td>
</tr>
<!-- start fields dynamics -->		
<!--#INCLUDE FILE="chform_fields.asp"-->	
<!-- end fields dynamics -->				
</table></td></tr></table>
</form>				
<!-- end code -->  
</div></center>              
</td></tr></table>
</body>
</html>
