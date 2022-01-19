<%@ Page Language="vb" AutoEventWireup="false" Codebehind="updTiming.aspx.vb" Inherits="bizpower_pegasus2018.updateTimingAutoSending" %>
<%@ Register TagPrefix="DS" TagName="metaInc" Src="../../title_meta_inc.ascx" %>
<html>
	<head>
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<DS:metaInc id="metaInc" runat="server"></DS:metaInc>
		<SCRIPT LANGUAGE="javascript">
<!--
		function checkForm()
		{

			return true;			
		}
		function changeDayNumber(frequancy)
		{
			if(frequancy==0)
			{
				document.getElementById("trDayNumber").style.display="none"
				document.getElementById("selDayNumber").innerHTML = "";
			}
			if(frequancy==1)
			{
				document.getElementById("trDayNumber").style.display="none"
				document.getElementById("selDayNumber").innerHTML = "";
			}
			if(frequancy==7)
			{
				document.getElementById("trDayNumber").style.display="block"
				document.getElementById("titleDayNumber").innerText="יום בשבוע"
				document.getElementById("selDayNumber").innerHTML = "";
				for(i=1;i<=7;i++)
				{
					document.getElementById("selDayNumber").options
					document.getElementById("selDayNumber").options[i-1] = new Option(i, i);
				}
			}
			if(frequancy==30)
			{
				document.getElementById("titleDayNumber").innerText="יום בחודש"
				document.getElementById("trDayNumber").style.display="block"
				document.getElementById("selDayNumber").innerHTML = "";
				for(i=1;i<=30;i++)
				{
					document.getElementById("selDayNumber").options
					document.getElementById("selDayNumber").options[i-1] = new Option(i, i);
				}
				document.getElementById("selDayNumber").options[30] = new Option("יום אחרון בכל חודש",-1);
			}
			
		}
//-->
		</SCRIPT>
	</head>
	<body style="margin:0px;background:#E5E5E5" onload="window.focus()">
		<%if UpdateTimingPermitted  then%>
		<table border="0" width="480" cellspacing="0" cellpadding="0" align="center">
			<tr>
				<td align="left" valign="middle" nowrap>
					<table width="100%" border="0" cellpadding="1" cellspacing="0">
						<tr>
							<td class="page_title" align="rtl" valign="middle" width="100%">הגדרת תזמון שליחת 
								טפסי רישום אוטומטית&nbsp;</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td bgcolor="#e6e6e6" background="../../images/Line_3D.gif" height="2" width="100%"></td>
			</tr>
			<tr>
				<td height="15" nowrap></td>
			</tr>
			<tr>
				<td width="100%">
					<form id="Form1" method="post" runat="server" autocomplete="off">
					
						<table width="480" cellspacing="1" cellpadding="2" align="center" border="0">
						<tr>
								<th align="right">
									תאור התזמון</th>
							</tr>
							<tr>
								<td class="td_admin_5" align="right" nowrap dir="ltr">
									<table cellspacing="0" cellpadding="0">
										<tr>
											<td colspan="2" align="right"><textarea id="timingTitle" name="timingTitle" rows="2"dir="rtl" style="width:400px"><%=timingTitle%></textarea></td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<th align="right">
									מועד השליחה</th>
							</tr>
							<tr>
								<td class="td_admin_5" align="right" nowrap dir="ltr">
									<table cellspacing="0" cellpadding="0">
										<tr>
											<td colspan="2" align="right"> ללא שליחה אוטומטית <input type="radio" name="checkFrequancy" value="0" <%if Frequancy=0 then%>checked<%end if%> onclick="changeDayNumber(0)"></td>
										</tr>
										<tr>
											<td colspan="2" align="right"> אחת ליום <input type="radio" name="checkFrequancy" value="1" <%if Frequancy=1 then%>checked<%end if%> onclick="changeDayNumber(1)"></td>
										</tr>
										<tr>
											<td colspan="2" align="right"> אחת לשבוע <input type="radio" name="checkFrequancy" value="7"  <%if Frequancy=7 then%>checked<%end if%> onclick="changeDayNumber(7)"></td>
										</tr>
										<tr>
											<td colspan="2" align="right"> אחת לחודש <input type="radio" name="checkFrequancy" value="30"  <%if Frequancy=30 then%>checked<%end if%> onclick="changeDayNumber(30)"></td>
										</tr>
									</table>
								</td>
							</tr>
							<tr >
								<td class="td_admin_5" align="right" nowrap dir="ltr" id="trDayNumber" <%=iif(Frequancy<7,"style='display:none'","")%>>
									<table cellspacing="0" cellpadding="0" align="right" >
									<tr><td align="right" id="titleDayNumber" style="font-weight:bold"> יום ב<%=iif(Frequancy=7,"שבוע","חודש")%></td></tr>
										<tr>
											<td align="right">
											<select id="selDayNumber" name="selDayNumber">
											<%if Frequancy=7 then
											Dim HeCultureInfo As Globalization.CultureInfo = New Globalization.CultureInfo("he-IL")
											
											for i as integer=1 to Frequancy%>
											<option value="<%=i%>" <%if DayNumber=i%>selected<%end if%>><%=HeCultureInfo.DateTimeFormat.GetDayName(i)%></option>											 
											<%next
											else
											for i as integer=1 to Frequancy%>
											<option value="<%=i%>" <%if DayNumber=i%>selected<%end if%>><%=i%></option>											 
											<%next
											end if%>
											<option value="-1" <%=iif(Frequancy<30,"style='display:none'","")%> <%if DayNumber=-1%>selected<%end if%>>יום אחרון בחודש</option>
											</select>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<th align="right">
									טווח תאריכי פתיחת טפסי הרישום ביחס למועד ההפצה</th>
							</tr>
							<tr>
								<td class="td_admin_5" align="right" nowrap dir="ltr">
									<table cellspacing="0" cellpadding="0">
										<tr>
											<td colspan="2" align="right">השבוע אחרון<input type="radio" name="checkInterval" value="7" <%if DateInterval=7 then%>checked<%end if%>></td>
										</tr>
										<tr>
											<td colspan="2" align="right">חודש אחרון<input type="radio" name="checkInterval" value="30" <%if DateInterval=30 then%>checked<%end if%>></td>
										</tr>
										<tr>
											<td colspan="2" align="right">חודשיים אחרונים<input type="radio" name="checkInterval" value="60" <%if DateInterval=60 then%>checked<%end if%>></td>
										</tr>
										<tr>
											<td colspan="2" align="right">שלושה חודשים אחרונים<input type="radio" name="checkInterval" value="90" <%if DateInterval=90 then%>checked<%end if%>></td>
										</tr>
										<tr>
											<td colspan="2" align="right">חצי שנה אחרונה<input type="radio" name="checkInterval" value="180" <%if DateInterval=180 then%>checked<%end if%>></td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<th align="right">
									יעדים</th>
							</tr>
							<tr>
								<td class="td_admin_5" align="right" nowrap dir="ltr">
									<table cellspacing="0" cellpadding="0">
										<tr>
											<td class="td_admin_4" nowrap align="right">
												<select runat="server" id="sCountries" dir="rtl" name="sCountries" multiple style="height:150px">
												</select>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td height="35" nowrap></td>
							</tr>
							<tr>
								<td align="center">
									<table cellspacing="0" cellpadding="0">
										<tr>
											<td class="td_admin_4" nowrap >
												<input type="button" value="סגור" class="button_edit_1" style="width:90px;height:28px"
													onclick="window.close();window.opener.location.href='jobTimingList.aspx'" id="Button2" name="Button2"></td>
											
											<td class="td_admin_4" nowrap align="left">		
												<asp:LinkButton runat="server" ID="btnSubmit" CssClass="button_edit_1" style="width:90px;height:28px;background-color:#ff9900">שלח</asp:LinkButton>
									</table>
								</td>
							</tr>
					</table>
					</form>
				</td>
			</tr>
		</table>
		<%end if%>
	</body>
</html>
