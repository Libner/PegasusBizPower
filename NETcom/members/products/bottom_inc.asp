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
<tr><td align="center" width=100% height="5" nowrap></td></tr>
<%If trim(UseBizLogo) = "1" Then%>
	<tr>
		<td align="center" width=100%>
			<table width=100% cellpadding=5 cellspacing=1 border=0 style="font-family:Arial; font-size:9pt; font-weight:500; color:#575757" align=center >
				<tr>
					<td nowrap>
						<a href="http://pegasus.bizpower.co.il" target="_blank"><img src="<%=Application("VirDir")%>/images/Powered-By-BP2.gif" border="0"></a> 
					</td>
					<%If trim(lang_id) = "1" Then%>
					<td align=right dir=ltr ><A style="color:#575757; font-size: 7pt;" href="http://pegasus.bizpower.co.il" target="_blank">BIZPOWER</a>&nbsp;<A style="color:#575757; font-size: 7pt; text-decoration: none;" href="http://pegasus.bizpower.co.il" target="_blank">CRM</a>-טופס זה נוצר באמצעות מחולל הטפסים של מערכת ה</td>
					<%Else%>
					<td align=right dir=ltr ></td>
					<%End If%>
				</tr>				
			</table>
		</td>
	</tr>
<%End If%>	