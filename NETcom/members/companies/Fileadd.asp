<%SERVER.ScriptTimeout=3000000%>
<%Response.Buffer = False%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<style>
.normalB
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 11pt;
    COLOR: #2e2380;
    FONT-FAMILY: Arial
}
</style>
</head>
<body style="margin:0px;background-color:#E5E5E5">
<div style="background-color:#E6E6E6;align:center" id="loadMessage">
<table border="0" width="400" cellspacing="1" cellpadding="0" align=center>
    <tr><td height="30"></td></tr>
	<tr>
		<td align="right" nowrap class="normalB" dir=rtl>תהליך טעינת הנתונים לאתר מתבצעת כעת, התהליך יארך מספר דקות</td>
	</tr>		
	<tr><td height="10"></td></tr>	
	<tr>
		<td  align="right" nowrap class="normalB" dir=rtl><font color="red">אין לסגור חלון זה ואין להפעיל ישומים אחרים במחשב בזמן התהליך.</font></td>
	</tr>	
	<tr><td height="20"></td></tr>
	<tr><td align=center><img src="../../images/butt_act.gif" border=0 vspace=0 hspace=0></td></tr>
	<tr><td height="10"></td></tr>
</table>
</div>
<%
    UserId = trim(Request.Cookies("bizpegasus")("UserID"))
    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))   
  
	set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
	filestring="../../../download/import/"
	'Response.Write server.mappath(filestring)
	'Response.End
	'fs.DeleteFile server.mappath(filestring) & "/*.*" ,False
	
	set upl = Server.CreateObject("SoftArtisans.FileUp") 
	upl.path=server.mappath("../../../download/import/")
	newFileName=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
	newFileName= "upload_customers_"&OrgID&".csv"
	
	''Response.Write newFileName
	upl.Form("UploadFile").SaveAs newFileName	
	set upl = Nothing	
%>
<script language=javascript>
<!--
	document.location.href = "fileRead.asp?newFileName=<%=newFileName%>";
//-->
</script>
</body>
</html>




