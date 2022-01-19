<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="<%=Application("VirDir")%>/dynamic_style.css" rel="STYLESHEET" type="text/css">
<!--#include file="../include/title_meta_inc.asp"-->
</head>

<body marginwidth="10" marginheight="0" hspace="10" vspace="0" topmargin="0" leftmargin="10" rightmargin="10" bgcolor="white">

<!--#include file="../include/top.asp"-->

<table border="0" width="100%" bgcolor="#6F6DA6" height="302" cellspacing="0" cellpadding="0" class="bgrHome" ID="Table1">
  <tr>
    <td width="100%" valign="top" align="right" style="padding-right:0px;"><div align="right">
		<table border="0" cellspacing="0" cellpadding="0" ID="Table2">
			<tr>
			
				<td align=right valign="top" width="240" nowrap bgcolor="#DEDFF7">
<!--#include file="../include/ashafim_inc.asp"-->				
				</td>
							
				<td width="2" nowrap bgcolor="#FFFFFF"></td>
							
				<td align=right valign="top" width="100%">
						<table border="0" width="100%" cellpadding="0" cellspacing="0" ID="Table3">
							<tr>
								<td  align="right" width="100%" height="120" background="<%=Application("VirDir")%>/images/home_pict_bgr.jpg"><img src="<%=Application("VirDir")%>/images/home_pict_right<%=Request.QueryString("catId")%>.jpg"></td>
							</tr>
							<tr><td height="2" bgcolor="#FFFFFF"></td></tr>						
							
							<tr>
								<td align="right" style="padding-right:20px; padding-left:20px;">
									<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table4">
									<%'//start of centent%>
										<tr>
											<td width="100%" valign="top" align="right" height="5"></td>
										</tr>		
																			
										<tr>
											<td align="right" class="title_page">
												הודעות וחדשות													
											</td>
										</tr>
										

										<tr>
											<td width="100%" valign="top" align="right" height="16"></td>
										</tr>								
      
      
										<tr>
											<td width="100%" height="214" valign="top" align="right">
											<table border="0" width="100%" 
											cellspacing="0" cellpadding="0" ID="Table6">
											<tr>
												<td height="7"><img  border="0" src="<%=Application("VirDir")%>/images/pina2.gif" width="7" height="7"></td>
												<td bgcolor="#FFFFFF" width="100%" height="7"></td>
												<td height="7"><img  border="0" src="<%=Application("VirDir")%>/images/pina1.gif" width="7" height="7"></td>
											</tr>
											<tr>
												<td bgcolor="#FFFFFF"></td>
												<td bgcolor="#FFFFFF" width="100%" align="center">
												<table border="0" width="96%" cellspacing="0" cellpadding="0" ID="Table7">
													<tr>
														<td align="left">
														
														<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table8">
															<tr><td height="5"></td></tr>
															<tr>
																<td align="right" dir="rtl">
																
																	<table width="100%" align="right" border="0" cellspacing="0" cellpadding="0" ID="Table9">
																		<tr>
																		<td width="100%">
						<%'//start of content%>			
															<table border="0" width="100%" align="center" cellspacing="0" cellpadding="0" ID="Table5">
																<tr>
																	<td valign="top" width="100%">
																		<table align="right" border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table10">       
																			<tr>
																			<%'//start of news%>
																				<td align="right" style="padding-left:15px" width="100%" valign="top"> 
																					<table align="right" border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table11">              
																					<%
																						sqlStr = "Select New_Id,New_Title,CONVERT(VARCHAR(8),New_Date,3) AS New_Date_Date from News where category_ID=1 AND New_Site_Visible=1 ORDER BY New_Date desc"
																						set rs_newsList=con.execute(sqlStr)
																						if rs_newsList.EOF then
																						%>
																						<tr>
																							<td width="100%" align="right" valign="top" class="head1">אין הודעות</td>
																						</tr>       
																						<%
																						end if
																						%>
																						<%
																						do while not rs_newsList.eof 
																						New_Date_Date=rs_newsList("New_Date_Date")
																						New_ID=rs_newsList("New_ID")
																						New_Title=rs_newsList("New_Title")
																						%>
																							<tr><td height="5"></td></tr>																																							
																							<tr>
																								<td width="100%" align="right" valign="top">
																									<table border="0" cellspacing="0" cellpadding="0" ID="Table12">
																									<tr><td align="right" class="headDate" style="color:#B53C52" valign="top"><%=New_Date_Date%></td></tr>
																									<tr><td align="right" valign="top" dir=rtl><a class="link_news" href="../news/news.asp?id=<%=New_ID%>"><u><%=New_Title%></u></a></td></tr> 
																									</table>
																								</td>
																							</tr>  
																							<tr><td height="5"></td></tr>																				
																							<tr><td height="1" bgcolor="#FFFFFF"></td></tr>																				
																						<%
																						rs_newsList.MoveNext
																						loop
																						rs_newsList.close
																						set rs_newsList = nothing
																					%>
																					</table>
																				</td>
																			<%'//end of news%>																		
																			</tr>	
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
												<td height="7"><img  border="0" src="<%=Application("VirDir")%>/images/pina3.gif" width="7" height="7"></td>
												<td bgcolor="#FFFFFF" width="100%" height="7"></td>
												<td height="7"><img  border="0" src="<%=Application("VirDir")%>/images/pina4.gif" width="7" height="7"></td>
											</tr>

											</table>
											</td>
										</tr>      		
									<%'//end of content%>
									</table>
								</td>
							</tr>
														
						</table>
				</td>
							
			</tr>

	      
		</table>
    </div></td>
  </tr>
</table>

<!--#include file="../include/bottom.asp"-->

</body>
</html>
