<%SERVER.ScriptTimeout=3000%>
<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//HE">
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=windows-1255">

<%'on error resume next
 
templateID=trim(Request("templateID"))

If isNumeric(templateID) And trim(templateID) <> "0" Then
  set pg=con.getRecordSet("SELECT * FROM Templates WHERE Template_Id="&templateID&" ")	 
  If not pg.eof Then
	 pertitle=pg("Template_Title")
	 persource=pg("Template_Source")	 
  End If 	
  Set pg = Nothing
  poll_message = poll_message & strLocal & "seker/seker.asp?" & Encode("P=" & trim(Request("prodId")) & "&C=" & trim(Request("C")) & "&S=" & trim(Request("S")))
%>

<title><%=pertitle%></title>
</head>
<body  topmargin=0 leftmargin=0 rightmargin=0> 
<div align="center">
<table>
<tr><td>
<%=persource%>
</td></tr></table>
</div>
</body>
</html>
<%End If%>
