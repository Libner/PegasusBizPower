<!--#include file="include/connect.asp"-->
<!--#include file="include/reverse.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<link href="<%=Application("VirDir")%>/dynamic_style.css" rel="STYLESHEET" type="text/css">
		<title>Bizpower - ���� ����� �������</title>
	<!--#include file="include/top.asp"-->
	<table dir="rtl" border="1" style="border-collapse:collapse" bordercolor="#999999" cellspacing="0" cellpadding="0" width="780" align="center">
	<tr><td colspan=4 bgcolor="#EBEBEB" align="center" width="780"><img src=images/pegasus_biz.jpg></td></tr>
	<%
		sqlStr="SELECT new_ID,New_Title,New_Date FROM News"
		sqlStr=sqlStr& " WHERE New_Home_Visible=1 and category_Id=1" 
		sqlStr=sqlStr& " ORDER BY New_Date DESC,New_ID DESC" 
		set newsList=con.EXECUTE(sqlStr)
		is_news = false
	    if not newsList.eof Then
	    is_news = true
	%>	
	<tr>	
	<td width="600" nowrap colSpan="3" align="right" bgcolor="#DADADA" dir="rtl" style="padding:5px">	
		<marquee direction=right  scrolldelay=120>
		<%		
			do while not newsList.EOF 
			newID=newsList("New_ID")
			newTitle=newsList("New_Title")
			newDate=newsList("New_Date")
		%>
		&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
		<a class="homeNews" href="<%=Application("VirDir")%>/news/news.asp?ID=<%=newID%>">>> <%=newTitle%></a>
		<%
			newsList.MoveNext
			loop				
		%>											
		</marquee>	
	</td>		
	<td rowspan=3 align="right" valign="top" width="180" nowrap bgcolor="#FFFFFF">
	<!--#include file="include/ashafim_inc.asp"-->
	</td>
	</tr>
	<%	End If	%>
	<% newsList.close
	   set newsList=Nothing							
	%>
	<tr>
		<td width="600" nowrap colSpan="3" align="right" class="subtitle" dir="rtl" style="padding:10px">
			����� Bizpower ����� �� ����� �� ���� ���������� ������, �����, ������ 
			������ ���� ������.<br>
				������ ����� ������, ���� ����� �����, ������� 
			������ 24 ���� ����� ��� ��������.
		</td>
		<% if not is_news Then	%>	
		<td rowspan=2 align="right" valign="top" width="180" nowrap bgcolor="#FFFFFF">
		<!--#include file="include/ashafim_inc.asp"-->
		</td>
		<% end if %>				
	</tr>
	<tr>
		<td align="right" width="200" nowrap bgcolor="#F9F9F9" valign=top>
		<%'//start of ����� �����%>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="ltr">
			<tr>
				<td align="left" style="padding:10px;padding-top:30px">
					<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="ltr">
					<tr>
						<td align="center" dir="rtl" class="titleB">
						<!--Bizpower CRM--><a class="details" href="<%=Application("VirDir")%>/template/default.asp?PageId=6&amp;catId=3&amp;maincat=1"><img src="images/title-CRM.gif" border=0 vspace=0 hspace=0></a>
						</td>
					</tr>
					<tr>
					<td height="10" nowrap></td>
					</tr>									
					<tr>
					<td align="right" dir="rtl">
					<b>��� ���� ����� ����� �������� ���� ������ ����� ���� ������ �������</b>
					</td>
					</tr>
					<tr>
					<td height="5" nowrap></td>
					</tr>																		
					<tr>
						<td align="right" dir="rtl">
						����� Bizpower CRM ������ �� ����� ������ ������ ����� �����, ���� ���� ������, ���� �����, ������ ������ ����� ���� �������� �������   
						</td>
					</tr>
					<tr>
						<td height="10" nowrap></td>
					</tr>
						<tr>
							<td>
								<table width="100%" cellpadding="0" cellspacing="0" ID="Table32">
									<tr>
										<td height="7" nowrap><a class="details" href="<%=Application("VirDir")%>/template/default.asp?PageId=6&amp;catId=3&amp;maincat=1">������&nbsp;&laquo;</a></td>
										<td width="100%" height="7"></td>
										<td height="7" valign="bottom"></td>
									</tr>
								</table>
							</td>
						</tr>
					</table></td></tr></table>
				</td>				

				<td align="right" width="200" nowrap bgcolor="#F9F9F9" valign=top>
					<%'//start of ���� ������%>
					<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="ltr">
						<tr>
							<td align="left" style="padding:10px;padding-top:30px">
								<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="ltr">
									<tr>
										<td align="center" dir="rtl" class="titleB">
										<!--Bizpower Feedback--><a class="details" href="<%=Application("VirDir")%>/template/default.asp?PageId=7&amp;catId=3&amp;maincat=1"><img src="images/title-feedback.gif" vspace=0 hspace=0 border=0></a>
										</td>
									</tr>
									<tr>
										<td height="10" nowrap></td>
									</tr>									
									<tr>
										<td align="right" dir="rtl">
										<b>���� ����� ������� �����, ����� ������� ����� �� ����� ��� ���� ���.</b>
										</td>
									</tr>
									<tr>
										<td height="5" nowrap></td>
									</tr>																		
									<tr>
										<td align="right" dir="rtl">
                                     ����� Bizpower Feedback ������ �� ����� ���� ����� ���, ����� �����, ����� ��������, ����� ����� ����, ������, ����� ����� ����.
                                       </td>
									</tr>
									<tr>
										<td height="10" nowrap></td>
									</tr>
								<tr>
								<td>
								<table width="100%" cellpadding="0" cellspacing="0">
									<tr>
										<td height="7" nowrap><a class="details" href="<%=Application("VirDir")%>/template/default.asp?PageId=7&amp;catId=3&amp;maincat=1">������&nbsp;&laquo;</a></td>
										<td width="100%" height="7"></td>
										<td height="7" valign="bottom"></td>
									</tr>
								</table>
							</td>
						</tr>
					</table></td></tr></table></td>
					<%'//end of ���� ������%>
				</td>	
				<td align="right" width="200" valign="top" nowrap bgcolor="#F9F9F9">
					<%'//start of ����� ������%>
					<table border="0" width="100%" style="padding:10px;padding-top:30px" cellspacing="0" cellpadding="0" dir="ltr">
						<tr>
							<td align="left">
								<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="ltr">
									<tr>
										<td align="center" dir="rtl" class="titleB">
										<!--Bizpower Campaign--><a class="details" href="<%=Application("VirDir")%>/template/default.asp?PageId=5&amp;catId=3&amp;maincat=1"><img src="images/title-campaigns.gif" border=0 vspace=0 hspace=0></a>
										</td>
									</tr>
									<tr>
										<td height="10" nowrap></td>
									</tr>
									<tr>
										<td align="right" dir="rtl">
										<b>����� ������ �������� ������� ������ ����� ������ ���� ���.</b>
										</td>
									</tr>
									<tr>
										<td height="5" nowrap></td>
									</tr>									
									<tr>
										<td align="right" dir="rtl">
											����� Bizpower Campaign ������ �� ����� ������ �������� �����, ������� �����-���, ����� ����� ������ ����� �� �������.
										</td>
									</tr>
									<tr>
										<td height="25" nowrap></td>
									</tr>																											
									<tr>
										<td >
											<table width="100%" cellpadding="0" cellspacing="0" border=0>
												<tr>
													<td height="7" nowrap><a class="details" href="<%=Application("VirDir")%>/template/default.asp?PageId=5&amp;catId=3&amp;maincat=1">������&nbsp;&laquo;</a></td>
													<td width="100%" height="7"></td>
													<td height="7" valign="bottom"></td>
												</tr>
											</table>
										</td>
									</tr>
					</table>
					</td></tr></table></td>
					<%'//end of ����� �����%>
				</td>
					
			</tr>			
		</table>		
		<%'//end of ����� �������%>
		<!--#include file="include/bottom.asp"-->
	</body>
<%If Request("error") <> nil Then%>
	<%if request("ip")<>nil then   '����� ����� ����� �� �� IP%>
	<script language="javascript" type="text/javascript">
	<!--
		window.alert("����� ����� �� ��-IP");	
	//-->
	</script>
	<%end if %>
	<%if request("blocked")<>nil then '����� ���� ��� 4 ����� %>
	<script language="javascript" type="text/javascript">
	<!--
		window.alert("����� �� ��� ����� ���� ������ ��� ���� �������� ������,\n ��� ��� ������ ���� ������ ��� ����� �� ����� �����");	
	//-->
	</script>
	<%end if %>
	<%if request("userError")<>nil then   '����� �� ���� �� ����� �� ���� USERS.ACTIVE = 0%>
	<script language="javascript" type="text/javascript">
	<!--
		window.alert("�� ����� �� ����� ���� ������ \n\n<%=Space(36)%>��� ��� ���");	
	//-->
	</script>
	<%end if %>
	<%if request("PasswordError")<>nil then '����� ������%>
	<script language="javascript" type="text/javascript">
	<!--
		window.alert("�� ����� �� ����� ���� ������ \n\n<%=Space(36)%>��� ��� ���");	
	//-->
	</script>
	<%End If%>
<%End If%>
</html>
<%Set con=Nothing%>