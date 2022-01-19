<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%  
	If Request.QueryString("CompetitorId") <> nil Then
		if cint(trim(Request.QueryString("CompetitorId")))>0 then
			Competitor_Id = trim(Request.QueryString("CompetitorId"))
		end if
	end if
	If Request.QueryString("delpict") ="1" and Competitor_Id<>"" Then
		set fso=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
		set pr=con.getRecordSet("SELECT Competitor_Logo FROM Competitors where Competitor_Id="& Competitor_Id)
		if not pr.eof then
			fileName1=pr("Competitor_Logo")
			if fileName1<>"" then
				fileString= Server.MapPath("../../../download/competitors/"& fileName1 ) 'deleting of existing file
				if fso.FileExists(fileString) then
					set f=fso.GetFile(fileString) 
					f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
					sqlstr="Update Compare_Competitors set Competitor_Logo = '" & sFix(upl.UserFileName) & "'"
					sqlstr=sqlstr & " Where Competitor_Id = " & Competitor_Id
					con.executeQuery(sqlstr)
				end if
			end if	
		end if
	end if
	


	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 32 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing
	  
	  sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	  set rsbuttons = con.getRecordSet(sqlstr)
	  If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	  End If
	  set rsbuttons=nothing	 	  
%>
<html>
	<head>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
			<SCRIPT LANGUAGE="javascript">
<!--
		function checkForm()
		{
		
			if(window.document.form1.Competitor_Name.value == "")
			{
			
				<%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס שם מתחרה "
				Else
					str_alert = "Please insert the competitor name !!"
				End If   
				%>
				window.alert("<%=str_alert%>");
				window.document.form1.Competitor_Name.focus();
				return false;
			}
			if (window.document.form1.Competitor_Rank.selectedIndex==0) {
				<%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא לבחור דרגת תחרות כללית "
				Else
					str_alert = "Please insert the competitor rank !!"
				End If   
				%>
				window.alert("<%=str_alert%>");
				window.document.form1.Competitor_Rank.focus();
				return false;
			}
			return true;			
		}
//-->
			</SCRIPT>
	</head>
	<body style="margin:0px;background:#E5E5E5" onload="window.focus()">
		<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>" ID="Table4">
			<tr>
				<td align="left" valign="middle" nowrap>
					<table width="100%" border="0" cellpadding="1" cellspacing="0" ID="Table5">
						<tr>
							<td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(Competitor_Id) > 0 Then%>עדכון<%Else%>הוספת<%End If%>&nbsp;מתחרה&nbsp;</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td height="15" nowrap></td>
			</tr>
			<tr>
				<td width="100%">
					<form name="form1" id="form1" action="uploadFormCompetitor.asp?add=1&Competitor_Id=<%=Competitor_Id%>" target="_self" method="post" ENCTYPE="multipart/form-data">
						<table dir="<%=dir_var%>" border="0" width="100%" cellspacing="0" cellpadding="0" align="center" ID="Table1">
							<tr>
								<td width="100%">
									<table width="100%" cellspacing="1" cellpadding="2" align="center" border="0" ID="Table3">
										<!--site Competitor-->
										<tr>
											<td align="right" nowrap class="td_admin_4" valign="top">
												<table border="0" cellpadding="3" cellspacing="0">
													<%	
											If Len(Competitor_Id) > 0 Then
												sqlstr="Select Competitor_Id, Competitor_Name, Competitor_Phone, Competitor_Rank, Competitor_SearchTerms,Competitor_Logo From Compare_Competitors Where Competitor_Id = " & Competitor_Id
												set rssub = con.getRecordSet(sqlstr)
												If not rssub.eof Then
													Competitor_Name = trim(rssub("Competitor_Name"))
													Competitor_Phone = trim(rssub("Competitor_Phone"))
													Competitor_Rank = cint(rssub("Competitor_Rank"))
													Competitor_SearchTerms = trim(rssub("Competitor_SearchTerms"))
													Competitor_Pict= trim(rssub("Competitor_Logo"))
												End If
												set rssub=Nothing
											End If
										%>
													<tr>
														<TD align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" name="Competitor_Name" ID="Competitor_Name"  style="width:500px" value="<%=vFix(Competitor_Name)%>" maxlength=150 dir="<%=dir_obj_var%>"></TD>
														<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>"><b>&nbsp;שם מתחרה&nbsp;</b></td>
													</tr>
													<tr>
														<TD align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" name="Competitor_Phone" ID="Competitor_Phone"  style="width:500px" value="<%=vFix(Competitor_Phone)%>" maxlength=150 dir="<%=dir_obj_var%>"></TD>
														<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>"><b>&nbsp;מספר טלפון 
																מתחרה&nbsp;</b></td>
													</tr>
													<tr>
														<TD align="<%=align_var%>" dir="<%=dir_obj_var%>">
															<select name="Competitor_Rank" id="Competitor_Rating" class="norm" dir="<%=dir_obj_var%>"  style="width:500px">
																<option value="0" id="word13">--בחר דרגה--</option>
																<%for i=1 to 5%>
																<option value="<%=i%>" <%If Competitor_Rank = i Then%> selected <%End if%>><%=i%></option>
																<%next%>
															</select>
														</TD>
														<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>"><b>&nbsp;דרגת תחרות 
																כללית&nbsp;</b></td>
													</tr>
													<tr>
														<TD align="<%=align_var%>" dir="<%=dir_obj_var%>"><textarea type="text" name="Competitor_SearchTerms" ID="Competitor_SearchTerms" rows="7" style="width:500px" dir="<%=dir_obj_var%>"><%=Competitor_SearchTerms%></textarea></TD>
														<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>"><b>&nbsp;שמות תגיות 
																נוספים&nbsp;</b><br><font color="#ff0000">יש להפריד בין כל תגית עם פסיק ורווח</font></td>
													</tr>
										<tr>
											<td height="5" colspan="2" nowrap></td>
										</tr>
										<tr>
											<td >
												<table width="100%" cellspacing="1" cellpadding="2" align="center" border="0" ID="Table2">
													<%if Competitor_Pict<>"" and Competitor_Pict<>nil then %>
													<tr>
														<td align="<%=align_var%>" dir="<%=dir_obj_var%>"><img src="../../../download/competitors/<%=Competitor_Pict%>" style="max-width:200px;max-height:70px" border="0"><a href="addCompetitor.asp?delpict=1&competitorID=<%=competitorID%>" ONCLICK="return CheckDel()" class="but_menu" style="width:100px;">מחיקת 
																תמונה</a></td>
													</tr>
													<%end if%>
													<tr>
														<td align="<%=align_var%>" dir="<%=dir_obj_var%>"><INPUT TYPE="FILE" NAME="UploadFile1" style="width:500px;" ID="UploadFile1"></td>
													</tr>
												</table>
											</td>
											<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>"><b>&nbsp;קובץ לוגו חברה<br><font color="#ff0000">(90x90px)</font></b></td>
										</tr>
										<tr>
											<td colspan="2" height="10"></td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td height="5" nowrap colspan="2"></td>
							</tr>
							<tr>
								<td height="1" nowrap colspan="2" bgcolor="#808080"></td>
							</tr>
							<tr>
								<td height="10" nowrap colspan="2"></td>
							</tr>
							<tr>
								<td align="center" colspan="2">
									<input type="button" value="<%=arrButtons(2)%>" class="but_menu" style="width:90" onclick="window.close();" id="Button3" name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<input type="submit" value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onclick="return checkForm()" id="Submit1" name=Button1></td>
							</tr>
						</table>
					</form>
				</td>
			</tr>
		</table>
	</body>
</html>
