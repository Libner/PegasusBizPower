<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%      
    If Request.QueryString("OrgId") = nil Then
		OrgId = trim(Request.Cookies("bizpegasus")("OrgId")) 
	Else
		OrgId = trim(Request.QueryString("OrgId"))
	End If	
    lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))  
	If lang_id = "2" Then
			dir_var = "rtl"
			align_var = "left"
			dir_obj_var = "ltr"
	Else
			dir_var = "ltr"
			align_var = "right"
			dir_obj_var = "rtl"
	End If            					
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" messangers="text/css">
<script language=javascript>
<!--
	
	function go_back(objmessanger)
	{			
		if(window.opener.document.all("messangerName") != null)
		{
			window.opener.document.all("messangerName").value = objmessanger.innerHTML;
		}
	    window.opener.focus();
	    window.close();
	    return false;	
	}
	
	function selectmessanger(objChk,messangerId)
	{
		all_radios = document.getElementsByName("messanger")
	
		for(count=0;count<all_radios.length;count++)
		{			
		    currID = all_radios[count].value;
			if(currID != messangerId)
			{
				window.document.all("tr"+currID).style.background = "#E6E6E6";
			}
			else
			{
				if(objChk.checked == true)
					window.document.all("tr"+messangerId).style.background = "#B1B0CF";
				else
					window.document.all("tr"+messangerId).style.background = "#E6E6E6";
			}		
		}		
	}	
		
//-->
</script>
</head>
<BODY style="margin:0px;background-color:#E4E4E4" onload="window.focus()">
<table width="100%" align=center cellpadding=0 cellspacing=0 border=0>
	<tr><td class="page_title" align=center>&nbsp;<%If trim(lang_id) = "1" Then%>רשימת עיסוקים<%Else%>Positions list<%End If%>&nbsp;</td></tr>
	<tr><td height=15 nowrap>&nbsp;</td></tr>
	<tr>
	<td width="100%" align="center" valign=top>
			<table border=0 width=90% align=center cellpadding=2 cellspacing=1 bgcolor="#FFFFFF" dir="<%=dir_var%>">
			<tr id="tr0" <%If checked <> "" Then%> style="background:#B1B0CF" <%Else%> style="background:#E6E6E6" <%End If%>>		    
				<td align="<%=align_var%>" dir=rtl>&nbsp;<a class="link1" href="javascript:go_back(messanger0)" target=_self id="messanger0">
				<%If trim(lang_id) = "1" Then%>ללא עיסוק
				<%Else%>Without position<%End If%></a>
				&nbsp;</td>				
			</tr>			
			<%
				sqlStr = "select item_Id, item_Name from messangers WHERE ORGANIZATION_ID=" & OrgId & " order by item_Name"
				set rs_messangers = con.GetRecordSet(sqlStr)
				If not rs_messangers.eof Then												
			%>									
			<%
				do while not rs_messangers.eof
				messangerId = trim(rs_messangers(0))				
				messanger_Name = trim(rs_messangers(1))							
			%>		
			<tr id="tr<%=messangerId%>" <%If checked <> "" Then%> style="background:#B1B0CF" <%Else%> style="background:#E6E6E6" <%End If%>>		    
				<td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;
				<a class="link1" href="javascript:go_back(messanger<%=messangerId%>)" target=_self id="messanger<%=messangerId%>"><%=messanger_Name%></a>&nbsp;</td>				
			</tr>					
			<%
		rs_messangers.movenext
		loop
		set rs_messangers = nothing		
		End If
		%>					
		</table>						
	  </td></tr>
	<tr><td colspan=2 height="15" nowrap></td></tr>
</table>
</body>
</html>
