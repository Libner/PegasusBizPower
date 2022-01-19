<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%'on error resume next
	selTemplate = trim(Request("template_id"))
	If trim(selTemplate) = "" Then
		selTemplate = "1"
	End If	
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 46 Order By word_id"				
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
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script>
function updateTemplate()
{   
	var templateID;
	templateID = window.document.all("template_id").value;
	//window.alert(window.opener.document.all("template_id"));
	if(window.opener.document.all("template_id") != null)
	{
		window.opener.document.all("template_id").value = templateID;
		window.opener.Form1.submit();
	}		
	window.close();
}
function openPreview(templateId)
{
	page = window.open("../../../admin/templates/result.asp?templateId="+templateId,"Template","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=780, height=540, left=2, top=2");	
	return false;
}
function selectTemp(imgObj,tempId)
{
	imgC = document.getElementsByTagName("IMG");
	//window.alert(imgC.length);
	for(i=0;i<imgC.length;i++)
	{
		if(imgC(i).style.border)
			imgC(i).style.border = "2px solid #C0C0C0";
	}
	bordSt = new String(imgObj.style.border);
	//window.alert(bordSt.search("2px"))
	if(bordSt.search("2px") > 0)
		imgObj.style.border = "3px solid Red"
	else	
		imgObj.style.border = "2px solid #C0C0C0";	
	window.document.all("template_id").value = 	tempId;
	return false;	
}
</script>
</head>
<body style="margin:0px;background:#E5E5E5">
<table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
<tr><td class="page_title" colspan=2><span id=word1 name=word1><!--רשימת תבניות--><%=arrTitles(1)%></span>&nbsp;</td></tr>
<form name=form1 id=form1 action="templates_list.asp" method=post>
<tr><td><table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
<tr><td colspan=2 height="10" nowrap></td></tr>
<tr><td colspan=2 align="<%=align_var%>" class=title_table><span id="word2" name=word2><!--בחר אחת מהתבניות ולחץ אישור--><%=arrTitles(2)%></span>&nbsp;&nbsp;&nbsp;</td></tr>
<tr><td colspan=2 height="10" nowrap></td></tr>
<tr>
	<td colspan=2 align="center" nowrap>
		<A class="but_menu" href="#" onClick="window.close();" style="width:80"><span id="word3" name=word3><!--ביטול--><%=arrTitles(3)%></span></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<A class="but_menu" href="#" onclick="updateTemplate()" style="width:80"><span id="word4" name=word4><!--אישור--><%=arrTitles(4)%></span></a>
	</td>		 
</tr>
<%
   i = 0
   sqlstr="Select Templates.Template_ID, Templates.Template_Title, Template_Screenshot From Templates"&_
   " Inner Join Templates_To_Organizations ON Templates.Template_ID = Templates_To_Organizations.Template_ID"&_
   " WHERE Templates_To_Organizations.Organization_ID = " & OrgID &_
   " And IsNULL(Templates.Template_Active,0) = '1' Order BY Templates.Template_Id"
   set rst = con.getRecordSet(sqlstr)
   'Response.Write rst.recordcount
   Do While not rst.eof
   i = i + 1
   templateId = trim(rst(0))
   teplateTitle = trim(rst(1))
   templateSize= rst.Fields(2).ActualSize   

%>
<tr>			    					    
   <td>
   <table cellpadding=1 border=0 cellspacing=1 width=100% dir=rtl>
   <tr>
   <%If templateSize > 0 Then%>
   <td width=33% valign=top>		
   <table cellpadding=1 cellspacing=1 width=100% border=0>   
   <tr><td align="<%=align_var%>" class="form">&nbsp;<a href="" class="linkFaq" onClick="return openPreview('<%=templateId%>')"><%=trim(teplateTitle)%></a>&nbsp;&nbsp;</td></tr>
   <tr><td align="<%=align_var%>">       
   <IMG class="hand" src="../../GetImage.asp?DB=Template&amp;FIELD=Template_Screenshot&amp;ID=<%=templateId%>" border="0" <%If trim(selTemplate) <> trim(templateId) Then%> style="border: 2px solid #C0C0C0" <%Else%> style="border: 3px solid Red" <%End If%> onClick="return selectTemp(this,'<%=templateId%>')">
   </td></tr>  
   </table>
   </td>
   <%End If%>
   <%rst.moveNext     
     If not rst.eof Then
	 templateId = trim(rst(0))
     teplateTitle = trim(rst(1))
     templateSize= rst.Fields(2).ActualSize
   If templateSize > 0 Then
   %>     
   <td width=33% valign=top>		
   <table cellpadding=1 cellspacing=1 width=100% border=0>   
   <tr><td align="<%=align_var%>" class="form">&nbsp;<a href="" class="linkFaq" onClick="return openPreview('<%=templateId%>')"><%=trim(teplateTitle)%></a>&nbsp;&nbsp;</td></tr>
   <tr><td align="<%=align_var%>">     
   <IMG class="hand" src="../../GetImage.asp?DB=Template&amp;FIELD=Template_Screenshot&amp;ID=<%=templateId%>" border=0 <%If trim(selTemplate) <> trim(templateId) Then%> style="border: 2px solid #C0C0C0" <%Else%> style="border: 3px solid Red" <%End If%> onClick="return selectTemp(this,'<%=templateId%>')"> 
   </td></tr>
   </table>
   </td>
   <%End If%>
   <%End If%>     
   <%       
      If not rst.eof Then
      rst.moveNext  
      If not rst.eof Then
	  templateId = trim(rst(0))
      teplateTitle = trim(rst(1))
      templateSize= rst.Fields(2).ActualSize
   If templateSize > 0 Then
   %>     
   <td width=33% valign=top>		
   <table cellpadding=1 cellspacing=1 width=100% border=0>   
   <tr><td align="<%=align_var%>" class="form">&nbsp;<a href="" class="linkFaq" onClick="return openPreview('<%=templateId%>')"><%=teplateTitle%></a>&nbsp;&nbsp;</td></tr>
   <tr><td align="<%=align_var%>">     
   <IMG class="hand" src="../../GetImage.asp?DB=Template&amp;FIELD=Template_Screenshot&amp;ID=<%=templateId%>" <%If trim(selTemplate) <> trim(templateId) Then%> style="border: 2px solid #C0C0C0" <%Else%> style="border: 3px solid Red" <%End If%> onClick="return selectTemp(this,'<%=templateId%>')">
   </td></tr>
   </table>
   </td>
   <%End If%>
   <%End If%>
   <%Else%>
   <td>&nbsp;</td>
   <%End If%>     
   </tr></table></td>  
</tr>
<%
  If not rst.eof Then
   rst.moveNext   
  End If 
  Loop
  set rst = Nothing 
 %> 
 </table></td> 
	</tr>
	<tr><td colspan=2 height="15" nowrap></td></tr>
    <tr>
	<td colspan=2 align="center" nowrap>
		<A class="but_menu" href="#" onClick="window.close();" style="width:80"><span id="word5" name=word5><!--ביטול--><%=arrTitles(5)%></span></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<A class="but_menu" href="#" onclick=updateTemplate() style="width:80"><span id="word6" name=word6><!--אישור--><%=arrTitles(6)%></span></a>
	</td>		 
    </tr>
	<tr><td colspan=2 height="5" nowrap></td></tr>
	<input type="hidden" name="template_id" value="<%=selTemplate%>" ID="template_id">
  </form>	
</table>
</body>
</html>							
									