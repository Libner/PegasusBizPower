<% Server.ScriptTimeout=10000 %>
<% Response.Buffer = False %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%mailId = trim(Request.querystring("mailId"))%>
<%If isNumeric(trim(mailId)) = true and isNumeric(trim(UserID)) = true And IsNumeric(trim(OrgID)) Then%>
<HTML>
<HEAD>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</HEAD>
<BODY class="body_admin">
<div id="div_save"  style="position:absolute; left:0px; top:0px; width:100%; height:100%; " >  												
  <table  height="100%" width="100%" cellspacing="10" cellpadding="10">  
     <tr><td  align="center" >
         <table  border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr>  
              <td align="right" dir="rtl">
              <%
	
	sqlSel="select Content_Email from MarketingMailing  where Mail_Id=" & mailId
    'Response.Write(sqlIns)
    'Response.End     
	Set rs_tmp = con.getRecordSet(sqlSel)
		Content_Email = rs_tmp.Fields("Content_Email").value
	Set rs_tmp = Nothing	  
	%>
	<%=Content_Email%>
	<%
  Set con = Nothing%>
              </td>
            </tr>
         </table>
         </td>
     </tr>
  </table>
</div>

<%End If%>
</BODY>
</HTML>