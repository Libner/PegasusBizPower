<% OrgName = trim(Request.Cookies("bizpegasus")("ORGNAME"))%>
<%If trim(OrgID) <> "" Then
	sqlstring = "SELECT UseBizLogo FROM organizations WHERE ORGANIZATION_ID = " & OrgID
	'Response.Write sqlstring
	'Response.End
	set rs_tmp=con.GetRecordSet(sqlstring)
	if not rs_tmp.EOF  then				
		UseBizLogo = trim(rs_tmp("UseBizLogo"))
	else
	    UseBizLogo = 0	
	end if
	set rs_tmp = Nothing	
End If%>
<table cellpadding="1" cellspacing="0" width="620" border="0" align="center">
<tr><td height="10" nowrap></td></tr>
<tr><td align="left" style="font-family:Arial; font-size:8pt; font-weight:500; color:#6A6A6A; vertical-align:top;" 
width="160" nowrap dir="ltr"><%If trim(UseBizLogo) = "1" Then%><A style="font-family:Arial; font-size:8pt; font-weight:500; color:#6A6A6A; 
vertical-align:top;" href="http://pegasus.bizpower.co.il" target="_blank" dir="ltr"><%End If%><img 
src="<%=Application("VirDir")%>/images/Powered-By-BP2.gif" border="0"><%If trim(UseBizLogo) = "1" Then%></A><%End If%></td>
<td width="460" valign="top" nowrap>&nbsp;</td></tr>
<tr><td height=10 nowrap></td></tr></table>
</td></tr></table>