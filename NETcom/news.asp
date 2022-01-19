<!-- #include file="connect.asp" -->
<!-- #include file="reverse.asp" -->
<!-- #include file="members/checkWorker.asp" -->
<html>
<head>
<!-- #include file="title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="IE4.css" rel="STYLESHEET" type="text/css">
</head>

<body marginwidth="10" marginheight="0" hspace="10" vspace="0" topmargin="0"
leftmargin="10" rightmargin="10" bgcolor="white">
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">        
<tr><td width=100% colspan=3>
<!--#include file="logo_top.asp"-->
</td></tr>			
 <tr>				
  <td width=100% valign=top>
  <table cellpadding=0 cellspacing=0 width=100% border=0 dir="<%=dir_var%>">
  <tr><td>
  <!--#include file="top.asp"-->
  </td></tr>			
  <tr><td height="2" bgcolor="#FFFFFF"></td></tr>
  <tr>
  <td align="<%=align_var%>" valign="top" width="100%">
  <table border="0" width="100%" bgcolor="#6F6DA6" height="302" cellspacing="0" cellpadding="0" class="bgrHome"  dir="<%=dir_var%>">
  <tr>
    <td width="100%" valign="top" align="<%=align_var%>" style="padding-right:0px;" bgcolor=#6B69A5>
		<table border="0" cellspacing="0" cellpadding="0" width="95%" align="<%=align_var%>" dir="<%=dir_var%>">														
							
<%
ID=Request.QueryString("ID")

	if trim(ID)<>"" then
		sqlStr = "SELECT New_Title,Page_Width,New_Date,new_Time_on,new_Time_off,New_Content,New_Desc, Category_ID  FROM News  WHERE new_ID="& ID 
		set rs_news = con.getRecordSet(sqlStr)
		if not rs_news.eof then
			Page_Width = rs_news("Page_Width")
			New_Content = rs_news("New_Content")
			New_Title = rs_news("New_Title")	 
			Page_Width = rs_news("Page_Width")
			New_Desc = rs_news("New_Desc") 
			New_Date = rs_news("New_Date")
			Category_ID = rs_news("Category_ID")								
		end if
		
		set rs_news = nothing
	end if
%>
							
							<tr>
								<td align="<%=align_var%>" style="padding-right:20px; padding-left:20px;" width=100%>
									<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
									<%'//start of centent%>
										<tr>
											<td width="100%" valign="top" align="<%=align_var%>" height="5"></td>
										</tr>		
																			
										<tr>
											<td align="<%=align_var%>" class="title_page" dir="rtl">
														<span class="headDate"><%=New_Date%></span>
														<br>
														<span><%=New_Title%></span>																					
											</td>
										</tr>
										

										<tr>
											<td width="100%" valign="top" align="<%=align_var%>" height="16"></td>
										</tr>								
      
      
										<tr>
											<td width="100%" height="214" valign="top" align="<%=align_var%>">
											<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="ltr">
											<tr>
												<td height="7"><img  border="0" src="<%=Application("VirDir")%>/images/pina2.gif" width="7" height="7"></td>
												<td bgcolor="#FFFFFF" width="100%" height="7"></td>
												<td height="7"><img  border="0" src="<%=Application("VirDir")%>/images/pina1.gif" width="7" height="7"></td>
											</tr>
											<tr>
												<td bgcolor="#FFFFFF"></td>
												<td bgcolor="#FFFFFF" width="100%" align="center">
												<table border="0" width="96%" cellspacing="0" cellpadding="0"  dir="<%=dir_var%>">
													<tr>
														<td align="left">
														
														<table border="0" width="100%" cellspacing="0" cellpadding="0"  dir="<%=dir_var%>">
															<tr><td height="5"></td></tr>
															<tr>
																<td align="<%=align_var%>" dir="rtl">
																
																	<table width="100%" align="<%=align_var%>" border="0" cellspacing="0" cellpadding="0" ID="Table9">
																		<tr>
																		<td width="100%">
																			<%'//start of content%>			

																							<table width="100%" align="<%=align_var%>" border="0" cellspacing="0" cellpadding="0"  dir="<%=dir_var%>">
																								<tr>
																								<td align="<%=align_var%>" width="100%" >
																									<table align="<%=align_var%>" width="100%" border="0" cellspacing="0" cellpadding="0"  dir="<%=dir_var%>">
																										
																										<tr><td width="100%" height="2"></td></tr>
																										<tr><td align="<%=align_var%>" dir=rtl width="100%" valign="top"><b><%=Replace(New_Desc,vbCrLf,"<br>")%></b></td></tr> 
																										<tr><td width="100%" height="10"></td></tr>
																										<tr>
																										<td align="<%=align_var%>" width="100%" valign="top"><%=New_Content%></td>
																										</tr> 
																										<tr><td width="100%" height="20"></td></tr>
																										<tr><td align="left" width="100%" align="<%=align_var%>" dir="ltr"><a href="" onclick="javascript:history.back(); return false;" class="btnLink" style="width:50px;">
																										<%If trim(lang_id) = "1" Then%>
																										&#0139;&#0139; חזרה
																										<%Else%>
																										Back &#0155;&#0155;
																										<%End If%>
																										</a></td></tr>
																										<tr><td width="100%" height="10"></td></tr>
																									</table>
																								</td>											
																								</tr>
																							</table>
																
																			<%'//end of content%>																					
																		</td>
																		</tr>
																	</table>
																							
																</td>
															</tr>
															<tr><td height="5"></td></tr>

														</table>
														</td>
													</tr>
										<tr><td height="5"></td></tr>
											</table>
												</td>
												<td bgcolor="#FFFFFF"></td>
											</tr>
											<tr>
												<td height="7" dir=ltr><img  border="0" src="<%=Application("VirDir")%>/images/pina3.gif" width="7" height="7"></td>
												<td bgcolor="#FFFFFF" width="100%" height="7" dir=ltr></td>
												<td height="7" dir=ltr><img  border="0" src="<%=Application("VirDir")%>/images/pina4.gif" width="7" height="7"></td>
											</tr>

											</table>
											</td>
										</tr>      		
									<%'//end of content%>
									</table>
								</td>
							</tr>
							<tr><td height="20"></td></tr>							
				</div>
			</table></td></tr>
		</table></td></tr></table>		
</body>
</html>
