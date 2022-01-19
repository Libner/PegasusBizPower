<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	templateId = trim(Request("templateId"))
	If isNumeric(trim(templateId)) Then
	sqlpg="SELECT Template_Screenshot FROM Templates WHERE Template_Id="&templateId&""
	'Response.Write sqlpg
	'Response.End
	set pg=con.getRecordSet(sqlpg)	
	perSize=pg.Fields("Template_Screenshot").ActualSize
	set pg = Nothing
	End If
	
	If Request.QueryString("delScreen") <> nil Then
		con.executeQuery("UPdate Templates set Template_screenshot = null where Template_Id = " & templateId)
		'Response.Write "UPdate Templates set page_screenshot = null where page_id = " & pageId
		'Response.End
		Response.Redirect "editScreen.asp?templateId=" & templateId
	End If
%>
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
<script>
function ifFieldEmpty(field){
	  if (field.value=='')
		{
		alert('חובה לבחור את התמונה');
		return false;
		}
	  else
		return true;
}
function CheckDel(field,templateId) 
//For picture deleting
{
  if(confirm("?האם ברצונך למחוק את התמונה") == true )
  {
	document.location.href='editScreen.asp?templateId='+templateId+'&'+field+'=1';
	return true;
  }
  return false;
}
</script>
</head>
<body>
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr><td colspan="4" class="page_title">דף ניהול תבניות דפי מבצע</td></tr>
  <tr>     
     <td width="5%" align="center"><a class="button_admin_1" href="default.asp">חזרה לניהול תבנית</a></td>  
     <td width="5%" align="center"></td> 
     <td width="*%" align="center"></td>      
  </tr>
</table>
<br>
<table width="90%" cellspacing="1" cellpadding="2" align=center border="0" bgcolor="#ffffff">
<tr><td height="10" nowrap></td></tr>
		<form ACTION="Aimgadd.asp?C=1&F=Template_Screenshot&templateId=<%=templateId%>" METHOD="post" ENCTYPE="multipart/form-data" ID="Form1">						
		<%If perSize = 0 then %>
		<tr>
			<td align="right" width="530" nowrap bgcolor="#dbdbdb">&nbsp;</td>
			<td align="center" width="100" nowrap rowspan="2" class="10normalB" bgcolor="#DDDDDD"><b>Print Screen</b></td>
		</tr>
		<tr valign=middle>
			<td align="right" bgcolor="#dbdbdb" valign="middle">
			<table border="0" cellspacing="0" cellpadding="0" align="right" ID="Table1">
				<tr valign=middle>        
				<td align="right" class="td_admin_1" style="padding-right:2px;"><INPUT type="submit" class="but_browse" onclick="return ifFieldEmpty('UploadFile2')" value="Print Screen העלאת" id=submit1 name=submit1></td>
				<td align="right" class="td_admin_1" style="padding-right:2px;"><INPUT TYPE="FILE" NAME="UploadFile2" size=30 ID="File1"></td>
				</tr>
			</table>       
			</td>
		</tr>
		<%Else%>
		<tr>
			<td align=right  bgcolor="#dbdbdb"><img id="imgPict" name="imgPict" src="..\..\netcom/GetImage.asp?DB=Template&amp;FIELD=Template_Screenshot&amp;ID=<%=templateId%>" border="0" hspace=2 ></td>						
			<td align=center rowspan="2" bgcolor="#b0b0b0" nowrap valign=top><font style="font-size:10pt;color:#ffffff;"><b>Print Screen</b></font></td>
		</tr>
		<tr>
			<td align=right  bgcolor="#b0b0b0" >
			<table border="0" cellspacing="0" cellpadding="0" align="right" ID="Table2">
				<tr valign=middle>        
				<td align="right" class="td_admin_1" style="padding-right:2px;"><INPUT type="button" class="but_browse"  ONCLICK="return CheckDel('delScreen','<%=templateId%>');" value="Print Screen מחיקת " id=button1 name=button1></td>
				<td align="right" class="td_admin_1" style="padding-right:2px;"><INPUT type="submit" class="but_browse" onclick="return ifFieldEmpty('UploadFile2')" value="Print Screen החלפת " id=submit2 name=submit2></td>
				<td align="right" class="td_admin_1" style="padding-right:2px;"><INPUT TYPE="FILE" NAME="UploadFile2" size=30 ID="File2"></td>
				</tr>
			</table>
			</td>
		</tr>
		<%End If%>		
		</form>
	</table>
	</body>
</html>		